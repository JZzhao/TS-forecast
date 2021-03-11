library(ggplot2)
library(reshape2)
library(lubridate)
library(forecast)
library(dplyr)
library(caret)
library(bsts)
library(stringr)
library(prophet)
library(data.table)
library(zoo)
library(tidyverse)
#library(XLConnect)
library(xlsx)

#Define fresh year-month to automate file save-in and save-out
fresh_year_month<-paste0(year(Sys.time()),"_",month(Sys.time()),"_")

#Set working folder
Working_Folder<-"C:/ForMe/CPLM/CONS_test/"
setwd(Working_Folder)
setwd(paste0(Working_Folder,year(Sys.time()),"_",month(Sys.time()),"_Refresh"))
Output_Folder<-paste0(Working_Folder,year(Sys.time()),"_",month(Sys.time()),"_Refresh")

suffix<-''
forecast_to<-'2022-12-01'

#---Define Index Calculation Period---#
HIST_START_DATE <- '2020-04-01'
HIST_END_DATE <- '2020-09-01'

FCST_START_DATE <- '2022-04-01'
FCST_END_DATE <- '2022-09-01'

#######################################################################################
###--------------------------Import Data and Reformatting---------------------------###
#######################################################################################
input_data<-read.csv("./L2_forecast_input.csv")
case_select<-read.csv("./model_training.csv")

#--transfer to character--#
for (name in names(input_data)){
  if(name %in% c("Sales.USD","Past.12m.Rev","Past.12m.MS")){
    input_data[,name]<-as.numeric(as.character(input_data[,name]))
  }else{
    input_data[,name]<-as.character(input_data[,name])
  }
}

#--merge key--#
input_data$Base_Feature<-paste0(input_data$Base,'_',input_data$Feature)

#--Only keep data for forecast--#
input_data<-input_data[paste0(input_data$Product.Category,input_data$Base_Feature) %in% paste0(case_select$Product.Category,case_select$Base),]

#--!!!!!!date format!!!!!!--#
#input_data$Date<-parse_date_time(sub('-','',input_data$Date), orders = c("%m-%d-%Y"))
input_data$Date<-parse_date_time(sub('-','',input_data$Date), orders = c("%Y-%m-%d"))
head(input_data$Date)

#--check missing value--#
input_data_new<-input_data %>% group_by(Product.Category,Base_Feature) %>% mutate(date_lag = as.period(lag(Date) %--% Date)$year * 12 + as.period(lag(Date) %--% Date)$month)

#--Get the stop month--#
na_over_4<-as.data.frame(input_data_new[is.na(input_data_new$date_lag)==FALSE & input_data_new$date_lag>4,])
#too many NAs in continuous month, need remove all values before it

#--Drop missing value over 3 consecutive months--#
input_data_new<-as.data.frame(input_data_new)
input_data_model<-data.frame()
for (product in unique(input_data$Product.Category)){
  tmp_1<-input_data[input_data$Product.Category==product,]
  for (base in unique(tmp_1$Base_Feature)){
    tmp_2<-tmp_1[tmp_1$Base_Feature==base,]
    if(paste0(product,base) %in% paste0(na_over_4$Product.Category,na_over_4$Base_Feature)){
      tmp_out<-tmp_2[tmp_2$Date>= max(na_over_4[na_over_4$Product.Category==product & na_over_4$Base_Feature==base,'Date']),]
    }else{
      tmp_out<-tmp_2
    }
    input_data_model<-rbind(input_data_model,tmp_out)
  }
}

input_data_model_new<-input_data_model %>% group_by(Product.Category,Base_Feature) %>% mutate(date_lag = as.period(lag(Date) %--% Date)$year * 12 + as.period(lag(Date) %--% Date)$month)

#--Filter out forecast cases still with missing value--#
na_case<-input_data_model_new[is.na(input_data_model_new$date_lag)==FALSE & input_data_model_new$date_lag>=2,]

model_data_complete<-data.frame()
for (product in unique(input_data_model$Product.Category)){
  tmp_1<-input_data_model[input_data_model$Product.Category==product,]
  for (base in unique(tmp_1$Base_Feature)){
    tmp_2<-tmp_1[tmp_1$Base_Feature==base,]
    if(paste0(product,base) %in% paste0(na_case$Product.Category,na_case$Base_Feature)){
      date_frame<-data.frame(Date=seq(min(tmp_2$Date),max(tmp_2$Date),by='month'))
      tmp_2<-merge(date_frame,tmp_2,all = TRUE)
      na_row<-tmp_2[is.na(tmp_2$Sales.USD),]
      na_range<-tmp_2[tmp_2$Date>= min(na_row$Date) %m-% months(2) & tmp_2$Date<= max(na_row$Date) %m+% months(2),]
      na_range$Sales.USD<-na.approx(na_range$Sales.USD,1:nrow(na_range))
      tmp_2[is.na(tmp_2$Sales.USD)==TRUE,'Sales.USD']<-na_range[as.character(na_range$Date) %in%  as.character(na_row$Date),'Sales.USD']
      tmp_2[is.na(tmp_2$Product.Category)==TRUE,'Product.Category']<-unique(tmp_2[is.na(tmp_2$Product.Category)==FALSE,'Product.Category'])
      tmp_2[is.na(tmp_2$Base)==TRUE,'Base']<-unique(tmp_2[is.na(tmp_2$Base)==FALSE,'Base'])
      tmp_2[is.na(tmp_2$Feature.type)==TRUE,'Feature.type']<-unique(tmp_2[is.na(tmp_2$Feature.type)==FALSE,'Feature.type'])
      tmp_2[is.na(tmp_2$Feature)==TRUE,'Feature']<-unique(tmp_2[is.na(tmp_2$Feature)==FALSE,'Feature'])
      tmp_2[is.na(tmp_2$Base_Feature)==TRUE,'Base_Feature']<-unique(tmp_2[is.na(tmp_2$Base_Feature)==FALSE,'Base_Feature'])
      tmp_out<-tmp_2
    }else{
      tmp_out<-tmp_2
    }
    model_data_complete<-rbind(model_data_complete,tmp_out)
  }
}

#--Filter out cases with 0 in latest month--#
max_date<-max(model_data_complete$Date)

all_0_case<-data.frame()
na_latest_3m<-data.frame()
case_0_latest<-data.frame()
for (product in unique(model_data_complete$Product.Category)){
  for (base in unique(model_data_complete[model_data_complete$Product.Category==product,'Base_Feature'])){
  sub_data<-model_data_complete[model_data_complete$Base_Feature==base & model_data_complete$Product.Category==product,]
  case_summary<-data.frame()
  latest_date_temp<-max(sub_data$Date)
  month_diff <- as.period(latest_date_temp %--%  max_date )$year * 12 + as.period( latest_date_temp %--%  max_date )$month
  if(month_diff>3){
    all_0_case=rbind(all_0_case,sub_data)
    case_summary[1,'Product.Category']<-product
    case_summary[1,'Base_Feature']<-base
    case_summary[1,'latest_month']<-as.character(as.Date(latest_date_temp))
  }else if(month_diff>0){
      na_latest_3m=rbind(na_latest_3m,sub_data)
      case_summary[1,'Product.Category']<-product
      case_summary[1,'Base_Feature']<-base
      case_summary[1,'latest_month']<-as.character(as.Date(latest_date_temp))
  }
  case_0_latest<-rbind(case_0_latest,case_summary)
  }
}

model_data_complete<-model_data_complete[!paste0(model_data_complete$Product.Category,model_data_complete$Base_Feature) %in% 
                                           paste0(case_0_latest$Product.Category, case_0_latest$Base_Feature),]

#---Highlight cases less than 24 month data points---#
short_case<-data.frame()
for (product in unique(input_data$Product.Category)){
  for (base in unique(input_data[input_data$Product.Category==product,'Base_Feature'])){
  sub_data<-model_data_complete[model_data_complete$Base_Feature==base & model_data_complete$Product.Category==product,]
  date_length<-length(sub_data$Date)
  if(date_length<12){
    if(nrow(short_case)==0){
      short_case=sub_data
    }else{
      short_case=rbind(short_case,sub_data)
    }
  }
  }
}

if(nrow(short_case)>=1){
  print("############----------Short Case Alert Highlight----------############")
  write.csv(short_case,'Short_case_Less_Than_12_Months.csv',row.names = FALSE)
}

model_data_complete<-model_data_complete[!paste0(model_data_complete$Product.Category,model_data_complete$Base_Feature) %in% 
                                           paste0(short_case$Product.Category, short_case$Base_Feature),]

accuracy_mape <- function(test, fcst)
{
  diff <- abs(test-fcst)
  mape <- mean(diff / (test + 1))
  
  return(1-mape)
}


##########################################################################
###--------------------------Model Container---------------------------###
##########################################################################


#auto.arima
forecast_auto_arima <- function(training_data, forecasting_period)
{
  fit <- auto.arima(training_data, D=1, stepwise=FALSE, approximation=FALSE, parallel=TRUE)
  fcst <- forecast(fit, h=forecasting_period)
  return(fcst$mean)
}

forecast_auto_arima_fit <- function(training_data, forecasting_period)
{
  fit <- auto.arima(training_data, D=1, stepwise=FALSE, approximation=FALSE, parallel=TRUE)
  return(fitted(fit))
}

#ets
forecast_ets <- function(training_data, forecasting_period)
{
  fit <- ets(training_data)
  fcst <- forecast(fit, h=forecasting_period)
  return(fcst$mean)
}

#hw
forecast_hw <- function(training_data, forecasting_period)
{
  fit <- tbats(training_data, seasonal.periods = 12)
  fcst <- forecast(fit, h=forecasting_period)
  return(fcst$mean)
}

#stl
forecast_stl <- function(training_data, forecasting_period)
{
  fit <- stl(training_data, s.window='periodic')
  fcst <- forecast(fit, h=forecasting_period)
  return(fcst$mean)
}

#prophet
forecast_prophet <- function(training_data, forecasting_period)
{
  training_date_list <- as.Date(time(training_data))
  training_data <- as.data.frame(training_data)
  training_data <- cbind(training_data, training_date_list)
  colnames(training_data) <- c('y', 'ds')
  
  model <- prophet(seasonality.mode='additive', daily.seasonality=FALSE, weekly.seasonality=FALSE)
  model <- add_seasonality(model, 'Quarterly', period=4, fourier.order=2)
  model <- fit.prophet(model, training_data)
  future <- make_future_dataframe(model, periods=forecasting_period, freq='month')
  
  future_start<-max(training_date_list) %m+% months(1)
  fcst <- predict(model, future)
  fcst <- ts(fcst[fcst$ds>= future_start,'yhat'],frequency = 12, start = c(year(future_start),month(future_start)))
  
  return(fcst)
}


#combined: arima, ets
forecast_combined <- function(training_data, forecasting_period)
{
  fcst_arima <- forecast_auto_arima(training_data, forecasting_period)
  fcst_ets <- forecast_ets(training_data, forecasting_period)
  
  output <- (fcst_ets+fcst_arima)/2
  return(output)
}

#combined_2: stl, hw
forecast_combinedv2 <- function(training_data, forecasting_period)
{
  fcst_stl <- forecast_stl(training_data, forecasting_period)
  fcst_hw <- forecast_hw(training_data, forecasting_period)
  
  output <- (fcst_stl+fcst_hw)/2
  return(output)
}


#-----Batch Run Forecast model + boxcox tranformation----#

#auto.arima
forecast_auto_arima_box_cox <- function(training_data, forecasting_period)
{
  transform_coder <- BoxCox.lambda(training_data)
  train_data_temp <- BoxCox(training_data, transform_coder)
  fit <- auto.arima(train_data_temp, D=1, stepwise=FALSE, approximation=FALSE, parallel=TRUE)
  fcst <- forecast(fit, h=forecasting_period)
  fcst <- InvBoxCox(fcst$mean, transform_coder)
  return(fcst)
}

forecast_auto_arima_box_cox_fit <- function(training_data, forecasting_period)
{
  transform_coder <- BoxCox.lambda(training_data)
  train_data_temp <- BoxCox(training_data, transform_coder)
  fit <- auto.arima(train_data_temp, D=1, stepwise=FALSE, approximation=FALSE, parallel=TRUE)
  return(InvBoxCox(fitted(fit),transform_coder))
  
}
#ets
forecast_ets_box_cox <- function(training_data, forecasting_period)
{
  transform_coder <- BoxCox.lambda(training_data)
  train_data_temp <- BoxCox(training_data, transform_coder)
  fit <- ets(train_data_temp)
  fcst <- forecast(fit, h=forecasting_period)
  fcst <- InvBoxCox(fcst$mean, transform_coder)
  return(fcst)
}

#hw
forecast_hw_box_cox <- function(training_data, forecasting_period)
{
  transform_coder <- BoxCox.lambda(training_data)
  train_data_temp <- BoxCox(training_data, transform_coder)
  fit <- tbats(train_data_temp, seasonal.periods = 12)
  fcst <- forecast(fit, h=forecasting_period)
  fcst <- InvBoxCox(fcst$mean, transform_coder)
  return(fcst)
}

#stl
forecast_stl_box_cox <- function(training_data, forecasting_period)
{
  transform_coder <- BoxCox.lambda(training_data)
  train_data_temp <- BoxCox(training_data, transform_coder)
  fit <- stl(train_data_temp, s.window='periodic')
  fcst <- forecast(fit, h=forecasting_period)
  fcst <- InvBoxCox(fcst$mean, transform_coder)
  return(fcst)
}

#prophet
forecast_prophet_box_cox <- function(training_data, forecasting_period)
{
  transform_coder <- BoxCox.lambda(training_data)
  train_data_temp <- BoxCox(training_data, transform_coder)
  training_date_list <- as.Date(time(training_data))
  train_data_temp <- as.data.frame(train_data_temp)
  train_data_temp <- cbind(train_data_temp, training_date_list)
  colnames(train_data_temp) <- c('y', 'ds')
  
  model <- prophet(seasonality.mode='additive', daily.seasonality=FALSE, weekly.seasonality=FALSE)
  model <- add_seasonality(model, 'Quarterly', period=4, fourier.order=2)
  model <- fit.prophet(model, train_data_temp)
  future <- make_future_dataframe(model, periods=forecasting_period, freq='month')
  
  fcst <- predict(model, future)
  fcst$yhat<-InvBoxCox(fcst$yhat,transform_coder)
  
  future_start<-max(training_date_list) %m+% months(1)
  fcst <- ts(fcst[fcst$ds>= future_start,'yhat'],frequency = 12, start = c(year(future_start),month(future_start)))
  
  return(fcst)
}

#combined: arima, ets
forecast_combined_box_cox <- function(training_data, forecasting_period)
{
  fcst_arima <- forecast_auto_arima_box_cox(training_data, forecasting_period)
  fcst_ets <- forecast_ets_box_cox(training_data, forecasting_period)
  
  output <- (fcst_ets+fcst_arima)/2
  return(output)
}

######################################################################
###-------------------------Explore Data---------------------------###
######################################################################

dir.create("./model_training_segment_EDA/")
dir.create("./model_training_segment_EDA/EDA_chart_archived/")
EDA_wb <- xlsx::createWorkbook('xlsx')
trend_sheet <- createSheet(EDA_wb, sheetName = 'EDA pic')
row_start<-1
for (product in unique(model_data_complete$Product.Category)){
  for (base in unique(model_data_complete[model_data_complete$Product.Category==product,'Base_Feature'])){
    case_sales <- model_data_complete[model_data_complete$Product.Category==product & model_data_complete$Base_Feature==base,c('Date', 'Sales.USD')]
    #trend chart
    p <- ggplot()+  geom_line(aes_string(x='Date', 'Sales.USD'),case_sales)
    p <- p + labs(title=product, subtitle=base)+ theme(plot.title = element_text(color = "brown3", size = 20, face = "bold", hjust = 0.5),axis.text.x = element_text(size=18),
                                                       plot.subtitle = element_text(color = "blue", size = 18, hjust = 0.5))
    ggsave(paste0("./model_training_segment_EDA/EDA_chart_archived/", product,'_',base, '_trend.png'), width = 20, height = 7, dpi = 100, device = 'png')
    pic_file <- paste0("./model_training_segment_EDA/EDA_chart_archived/", product,'_',base, '_trend.png')
    addPicture(pic_file, sheet=trend_sheet, scale=0.4, startRow = row_start,startColumn = 1)
    
    #seasonal chart
    p <- case_sales %>% mutate( year = factor(year(Date)), Date = update(Date, year = 1))  %>% 
      ggplot(aes(as.Date(Date,format = "%Y-%m-%d"), `Sales.USD`, color = year)) + scale_x_date(date_breaks = "1 month", date_labels = "%b") + labs(x='Date')
    
    p + geom_line() +labs(title=product, subtitle=base)+ theme(plot.title = element_text(color = "brown3", size = 20, face = "bold", hjust = 0.5),axis.text.x = element_text(size=18),
                                                               plot.subtitle = element_text(color = "blue", size = 18, hjust = 0.5), legend.title = element_text(size=18), legend.text = element_text(size=18))
    ggsave(paste0("./model_training_segment_EDA/EDA_chart_archived/", product,'_',base, '_seasonal.png'), width = 18, height = 7, dpi = 100, device = 'png')
    
    
    pic_file <- paste0("./model_training_segment_EDA/EDA_chart_archived/", product,'_',base, '_seasonal.png')
    addPicture(pic_file, sheet=trend_sheet, scale=0.4, startRow = row_start, startColumn = 14)
    
    row_start<-row_start + 15
  }
}
xlsx::saveWorkbook(EDA_wb,paste0("./model_training_segment_EDA/EDA_before_model.xlsx"))

######################################################################
###--------------------------Train Model---------------------------###
######################################################################

forecast_methods <- c('forecast_auto_arima','forecast_ets', 'forecast_hw','forecast_stl','forecast_combined','forecast_combinedv2','forecast_prophet',
                      'forecast_auto_arima_box_cox','forecast_ets_box_cox', 'forecast_hw_box_cox','forecast_stl_box_cox','forecast_combined_box_cox','forecast_prophet_box_cox',
                      'forecast_auto_arima_fit','forecast_auto_arima_box_cox_fit')
Backtesting_periods <- c(6,12)
#ts_index <- seq(as.Date("2015-01-01"), as.Date(forecast_to), "month")

##----Generate Backtest Accuracy result for Step-1 methods----##
dir.create("./Step_1/")
accuracy_record <- data.frame()
for (product in unique(model_data_complete$Product.Category)){
  backtest_wb <- xlsx::createWorkbook('xlsx')
  accuracy_record <- data.frame()
  for (base in unique(model_data_complete[model_data_complete$Product.Category==product,'Base_Feature'])){
  case_sales <- model_data_complete[model_data_complete$Product.Category==product & model_data_complete$Base_Feature==base,c('Date', 'Sales.USD')]
  bk_record<-data.frame()
  
  #---Back Test for different algorithms---#
  for (backtesting_period in Backtesting_periods)
  {
    backtest_train_data <- head(case_sales$Sales.USD, nrow(case_sales)-backtesting_period)
    backtest_test_data <- tail(case_sales$Sales.USD, backtesting_period)
    actual <- data.frame('Actual'=backtest_train_data, stringsAsFactors=FALSE)
    
    hist_start_date<-case_sales$Date[1];hist_end_date<-case_sales$Date[nrow(case_sales)]
    #hist_length<-nrow(case_sales)
    fcst_start_date<-hist_end_date %m+% months(1)
    fcst_end_date<-parse_date_time(sub('-','',forecast_to), orders = c("%Y-%m-%d"))
    #fcst_length<-as.period( hist_end_date %--% fcst_end_date)$year * 12 + as.period( hist_end_date %--% fcst_end_date)$month
    
    backtest_train_ts <- try(ts(backtest_train_data, start = c(year(hist_start_date),month(hist_start_date)), frequency=12),silent=T)
    fit_test_ts<-try(ts(case_sales$Sales.USD, start = c(year(hist_start_date),month(hist_start_date)), frequency=12),silent=T)
    if('try-error' %in% class(backtest_train_ts)) {next}
    
    for(method in forecast_methods)
    { 
      if(grepl('fit',method)==FALSE){
      #Use train + test method
      fcst <- try(do.call(method, list(backtest_train_ts, backtesting_period)), silent = TRUE,outFile = getOption("try.outFile", default = stderr()))
      if('try-error' %in% class(fcst)) {next}
      
      if(any(is.na(fcst)==T)){
        fcst <- as.data.frame(fcst)
        colnames(fcst) <- method
        k<-which(is.na(fcst),arr.ind = T)
        fcst[k] <- colMeans(fcst,na.rm = T)
        fcst <- as.numeric(unlist(fcst))
      }
      
      accuracy <- accuracy_mape(backtest_test_data, fcst)
      
      if(nrow(accuracy_record)==0){
        accuracy_record = data.frame('Product.Category'=product,'Base'=base, 'Backtesting_period'=backtesting_period, 'Method'=method, 'Accuracy'=accuracy, stringsAsFactors=FALSE)
      }else{
        accuracy_record = rbind(accuracy_record, data.frame('Product.Category'=product,'Base'=base, 'Backtesting_period'=backtesting_period, 'Method'=method, 'Accuracy'=accuracy, stringsAsFactors=FALSE))
      }
      accuracy_record <- accuracy_record[with(accuracy_record, order(Base, -Accuracy)),]
      
      backtest_output_temp <- data.frame('Fcst'=append(as.vector(unlist(actual)),fcst))
      colnames(backtest_output_temp) <- c(paste(method,backtesting_period,sep = "_"))
      
      if(nrow(bk_record)==0){
        bk_record <- backtest_output_temp
      }else{
        bk_record <- cbind(bk_record,backtest_output_temp)
      }
      }else{
     #use fitted accuracy method
        if('try-error' %in% class(fit_test_ts)) {next}
          fcst <- try(do.call(method, list(fit_test_ts, backtesting_period)), silent = TRUE,outFile = getOption("try.outFile", default = stderr()))
          if('try-error' %in% class(fcst)) {next}
          
          if(any(is.na(fcst)==T)){
            fcst <- as.data.frame(fcst)
            colnames(fcst) <- method
            k<-which(is.na(fcst),arr.ind = T)
            fcst[k] <- colMeans(fcst,na.rm = T)
            fcst <- as.numeric(unlist(fcst))
          }
          
          accuracy <- accuracy_mape(backtest_test_data, tail(fcst,backtesting_period))
          accuracy_record = rbind(accuracy_record, data.frame('Product.Category'=product,'Base'=base, 'Backtesting_period'=backtesting_period, 'Method'=method, 'Accuracy'=accuracy, stringsAsFactors=FALSE))
          
          accuracy_record <- accuracy_record[with(accuracy_record, order(Base, -Accuracy)),]
          
          backtest_output_temp <- data.frame('Fcst'=fcst)
          colnames(backtest_output_temp) <- c(paste(method,backtesting_period,sep = "_"))
          
          if(nrow(bk_record)==0){
            bk_record <- backtest_output_temp
          }else{
            bk_record <- cbind(bk_record,backtest_output_temp)
          }
      }
    }
  }
  
  bk_record <- cbind(case_sales, bk_record)
  names(bk_record)[2] <- 'Actual'
  
  temp_sheet<-xlsx::createSheet(backtest_wb, sheetName=base)
  xlsx::addDataFrame(bk_record , sheet=temp_sheet, startColumn=1, row.names=FALSE)

  }
  temp_sheet<-xlsx::createSheet(backtest_wb, sheetName='Accuracy')
  xlsx::addDataFrame(accuracy_record , sheet=temp_sheet, startColumn=1, row.names=FALSE)
  xlsx::saveWorkbook(backtest_wb, paste0("./Step_1/Step_1_",product,'_',"Backtest_record",suffix,".xlsx"))
}


##----Generate Forecast result for Step-1 methods----##

for (product in unique(model_data_complete$Product.Category)){
  forecast_wb <- xlsx::createWorkbook('xlsx')
  for (base in unique(model_data_complete[model_data_complete$Product.Category==product,'Base_Feature'])){
    case_sales <- model_data_complete[model_data_complete$Product.Category==product & model_data_complete$Base_Feature==base,c('Date', 'Sales.USD')]
    fcst_output<-data.frame()
    
    
      forecast_train_data <- case_sales$Sales.USD
      actual_temp <- data.frame(forecast_train_data, stringsAsFactors=FALSE)
      
      hist_start_date<-case_sales$Date[1];hist_end_date<-case_sales$Date[nrow(case_sales)]
      #hist_length<-nrow(case_sales)
      fcst_start_date<-hist_end_date %m+% months(1)
      fcst_end_date<-parse_date_time(sub('-','',forecast_to), orders = c("%Y-%m-%d"))
      #fcst_length<-as.period( hist_end_date %--% fcst_end_date)$year * 12 + as.period( hist_end_date %--% fcst_end_date)$month
      ts_index <- seq(as.Date(case_sales$Date[1]), as.Date(forecast_to), "month")
      
      forecast_train_ts <- ts(forecast_train_data, start=c(year(hist_start_date),month(hist_start_date)), frequency=12)
      
      
      for (method in forecast_methods){
        if(grepl('fit',method)){
        fcst <- try(do.call(sub('_fit','',method), list(forecast_train_ts, forecasting_period)), silent = TRUE,outFile = getOption("try.outFile", default = stderr()))
        }else{
          fcst <- try(do.call(method, list(forecast_train_ts, forecasting_period)), silent = TRUE,outFile = getOption("try.outFile", default = stderr()))
        }
        if('try-error' %in% class(fcst)) {next}
        fcst <- as.data.frame(fcst)
        colnames(fcst) <- method
        
        #Replace fcst NA value with colMeans
        if(any(is.na(fcst)==T)){
          k<-which(is.na(fcst),arr.ind = T)
          fcst[k] <- colMeans(fcst,na.rm = T)
        }
        
        if(nrow(fcst_output)==0){
          fcst_output <- fcst
        }else{
          fcst_output <- cbind(fcst_output,fcst)
        }
      }
    
    
    actual_temp <- as.data.frame(rep(actual_temp, ncol(fcst_output)))
    colnames(actual_temp) <- colnames(fcst_output)
    # if(nrow(actual_temp)!=(length(ts_index)-forecasting_period)){
    #   zero_table <- as.data.frame(matrix(0,nrow=(length(ts_index)-forecasting_period-nrow(actual_temp)),ncol=ncol(fcst_output)))
    #   colnames(zero_table) <- colnames(fcst_output)
    #   actual_temp <- rbind(zero_table,actual_temp)
    # }
    
    fcst_output <- cbind(ts_index,rbind(actual_temp, fcst_output))
    
    colnames(fcst_output)[1] <- 'Date'

    temp_sheet<-xlsx::createSheet(forecast_wb, sheetName=base)
    xlsx::addDataFrame(fcst_output , sheet=temp_sheet, startColumn=1, row.names=FALSE)

  }
  xlsx::saveWorkbook(forecast_wb, paste0("./Step_1/Step_1_",product,'_',"Forecast_record",suffix,".xlsx"))
}

metric_table <- data.frame()
for (product in unique(model_data_complete$Product.Category)){
  for (base in unique(model_data_complete[model_data_complete$Product.Category==product,'Base_Feature'])){
    sheet=base
    
  ###=====1.Metric: Accuracy=======###
  accuracy_record <-xlsx::read.xlsx(paste0("./Step_1/Step_1_",product,'_',"Backtest_record",suffix,".xlsx"),sheetName = 'Accuracy')
  accuracy_table <- spread(accuracy_record,Backtesting_period,Accuracy)

  colnames(accuracy_table)[colnames(accuracy_table)=='6']<-'Accuracy_6'
  colnames(accuracy_table)[colnames(accuracy_table)=='12']<-'Accuracy_12'
  accuracy_table <- accuracy_table[paste0(accuracy_table$Product.Category,'_',accuracy_table$Base)==paste0(product,"_",base),]
  
  fcst_df <-xlsx::read.xlsx(paste0("./Step_1/Step_1_",product,'_',"Forecast_record",suffix,".xlsx"),sheetName = sheet)
  
  fcst_growth_rate_table <- data.frame()
  fcst_cv_table <- data.frame()
  fcst_sumValue_table <- data.frame()
  fcst_slope_table<-data.frame()
  
  for (col in colnames(fcst_df)) {
    if(col=='Date'){next}
    fcst_ts <-as.vector(unlist(fcst_df[fcst_df$Date>=FCST_START_DATE & fcst_df$Date<=FCST_END_DATE,col]))
    if(NA %in% fcst_ts){next}
    hist_ts <-as.vector(unlist(fcst_df[fcst_df$Date>=HIST_START_DATE & fcst_df$Date<=HIST_END_DATE,col]))
    hist_ts_N_1<-as.vector(unlist(fcst_df[fcst_df$Date>=as.Date(HIST_START_DATE)%m-%months(12) & fcst_df$Date<=as.Date(HIST_END_DATE)%m-%months(12),col]))
    
    ###====2.Metric: Growth rate====###
    growth_rate <- (sum(fcst_ts)/sum(hist_ts))^(1/(as.period(HIST_END_DATE %--% FCST_END_DATE)$year))-1
    hist_growth_rate<-ifelse(length(hist_ts_N_1)<length(hist_ts),NA,(sum(hist_ts)/sum(hist_ts_N_1))-1)
    fcst_growth_rate_table<-rbind(fcst_growth_rate_table,data.frame('Product.Category'=product,'Base'=base,'Method'=col,
                                                                    'Hist_Growth_rate'=hist_growth_rate,'Growth_rate'=growth_rate))

    
    ####====3.Metric: Variability====####
    fcst_cv <- sd(fcst_ts)/mean(fcst_ts)
    hist_cv <- sd(hist_ts)/mean(hist_ts)
    hist_cv_N_1 <- ifelse(length(hist_ts_N_1)<length(hist_ts),NA,sd(hist_ts_N_1)/mean(hist_ts_N_1))

    fcst_cv_table<-rbind(fcst_cv_table,data.frame('Product.Category'=product,'Base'=base,'Method'=col,'Hist_cv_N_1'=hist_cv_N_1,'Hist_cv'=hist_cv,'Fcst_cv'=fcst_cv))
    
    
    ####====4.Metric: Sum Value====#### 
    fcst_sumValue <- sum(fcst_ts)
    hist_sumValue <- sum(hist_ts)

    fcst_sumValue_table<-rbind(fcst_sumValue_table,data.frame('Product.Category'=product,'Base'=base,'Method'=col,'Hist_sumValue'=hist_sumValue,'Fcst_sumValue'=fcst_sumValue))

    ####====5.Metric: Slope====####
    fcst_slope<-as.vector(lm(fcst_ts~c(1:length(fcst_ts)))$coefficients[2])
    hist_slope<-as.vector(lm(hist_ts~c(1:length(hist_ts)))$coefficients[2])
    hist_slope_N_1<-as.vector(lm(hist_ts_N_1~c(1:length(hist_ts_N_1)))$coefficients[2])
    
    fcst_slope_table<-rbind(fcst_slope_table,data.frame('Product.Category'=product,'Base'=base,'Method'=col,'Hist_Slope_N_1'=hist_slope_N_1,
                                                        'Hist_Slope'=hist_slope,'Fcst_Slope'=fcst_slope))
    
    
  }

  metric_tmp <- Reduce(function(x,y,...) merge(x,y,all=T,...), list(accuracy_table,fcst_growth_rate_table,fcst_cv_table,fcst_sumValue_table,fcst_slope_table))
  metric_table <- rbind(metric_table,metric_tmp)
  
  }
}

metric_table<-metric_table[is.na(metric_table$Accuracy_6)==FALSE,]

###---Filter out the best method with best accuracy---###

accuracy_threshold<-0.8
best_method <- data.frame()
for (product in unique(model_data_complete$Product.Category)){
  for (base in unique(model_data_complete[model_data_complete$Product.Category==product,'Base_Feature'])){
  metric_table_tmp <- metric_table[metric_table$Product.Category==product &metric_table$Base==base,]
  metric_table_tmp$Accuracy_12<-ifelse(is.na(metric_table_tmp$Accuracy_12)==TRUE,0,metric_table_tmp$Accuracy_12)
  if(any(metric_table_tmp$Accuracy_12>=accuracy_threshold)){
    best_method_record <- metric_table_tmp[which.max(metric_table_tmp$Accuracy_12),]
    best_method_tmp <- data.frame('Product.Category'=product,'Base'=base,'Method'=best_method_record$Method,'Accuracy'=best_method_record$Accuracy_12,'Backtesting_period'=12)
  }else if(any(metric_table_tmp$Accuracy_6>=accuracy_threshold)){
    best_method_record <- metric_table_tmp[which.max(metric_table_tmp$Accuracy_6),]
    best_method_tmp <- data.frame('Product.Category'=product,'Base'=base,'Method'=best_method_record$Method,'Accuracy'=best_method_record$Accuracy_6,'Backtesting_period'=6)
  }else{
    if(max(metric_table_tmp$Accuracy_6)>max(metric_table_tmp$Accuracy_12)){
      max_col<-'Accuracy_6'
      best_accuracy<-max(metric_table_tmp$Accuracy_6)
    }else{
      max_col<-'Accuracy_12'
      best_accuracy<-max(metric_table_tmp$Accuracy_12)
    }
    best_method_record <- metric_table_tmp[which.max(metric_table_tmp[[max_col]]),]
    best_method_tmp <- data.frame('Product.Category'=product,'Base'=base,'Method'=best_method_record$Method,'Accuracy'=best_accuracy,'Backtesting_period'=sub('.*_','',max_col))
  }
  
  if(nrow(best_method)==0){
    best_method<-best_method_tmp
  }else{
    best_method<-rbind(best_method,best_method_tmp)
  }
  }
}

best_method$Method<-sub('_fit','',best_method$Method)

write.csv(best_method,paste0('./Step_1/Step_1_master_file_backup',suffix,'.csv'),row.names = F)
if(nrow(best_method[best_method$Accuracy>=accuracy_threshold,])>=1){
  write.csv(best_method[best_method$Accuracy>=accuracy_threshold,],paste0('./Step_1/Step_1_Best_Method',suffix,'.csv'),row.names = F)
}

if(nrow(best_method[best_method$Accuracy<accuracy_threshold,])>=1){
  write.csv(best_method[best_method$Accuracy<accuracy_threshold,],paste0('./Step_1/Segments_for_Step_2',suffix,'.csv'),row.names = F)
}


###---Generate the plots and save out----###

Backtest_periods <- c(6,12)
dir.create(paste0("./Step_1/Step_1_pic_output",suffix))
for (product in unique(model_data_complete$Product.Category)){
  trend_wb <- xlsx::createWorkbook('xlsx')
  for (base in unique(model_data_complete[model_data_complete$Product.Category==product,'Base_Feature'])){
    sheet=base
    accuracy_normal<-xlsx::read.xlsx(paste0("./Step_1/Step_1_",product,'_',"Backtest_record",suffix,".xlsx"),sheetName = 'Accuracy')

    fcst_df <-xlsx::read.xlsx(paste0("./Step_1/Step_1_",product,'_',"Forecast_record",suffix,".xlsx"),sheetName = sheet)
    bk_df <-xlsx::read.xlsx(paste0("./Step_1/Step_1_",product,'_',"Backtest_record",suffix,".xlsx"),sheetName = sheet)
    fcst_df$Date <- as.Date(fcst_df$Date)
    
    hist_start_date<-fcst_df$Date[1];hist_end_date<-fcst_df$Date[nrow(bk_df)]
    #hist_length<-nrow(fcst_df)
    fcst_start_date<-hist_end_date %m+% months(1)
    fcst_end_date<-parse_date_time(sub('-','',forecast_to), orders = c("%Y-%m-%d"))
    #fcst_length<-as.period( hist_end_date %--% fcst_end_date)$year * 12 + as.period( hist_end_date %--% fcst_end_date)$month
    ts_index <- seq(as.Date(fcst_df$Date[1]), as.Date(forecast_to), "month")
    actual_period <- as.period(hist_start_date %--% max(model_data_complete$Date))$year * 12 + as.period(hist_start_date %--% max(model_data_complete$Date))$month+1
    
    
    #create a flag for different line colors
    fcst_df[1:actual_period,"Flag"] <- "Actual"
    fcst_df[(actual_period+1):(nrow(fcst_df)),"Flag"] <- "Forecast"
    
    ##Create the worksheet
    trend_sheet <- createSheet(trend_wb, sheet)
    
    for (backtest_period in Backtest_periods) {
      bk_df <-xlsx::read.xlsx(paste0("./Step_1/Step_1_",product,'_',"Backtest_record",suffix,".xlsx"),sheetName = sheet)
      bk_df<-bk_df[,grepl(paste0("_", backtest_period),colnames(bk_df))]
      colnames(bk_df) <- gsub(paste0("_", backtest_period), "", colnames(bk_df))
      
      for(name in names(bk_df)){
        bk_df[,name]<-as.numeric(as.character(bk_df[,name]))
      }
      
      bk_record <- tail(bk_df, backtest_period)
      bk_record$Flag <- "Backtest"
      
      bk_index<-ts_index[(actual_period-backtest_period+1):actual_period]
      bk_record <- cbind(bk_index, bk_record)
      bk_record <- rename(bk_record, Date=bk_index)
      
      for (col in colnames(fcst_df)) {
        if(col %in% c('Date','Flag')){next}
        accuracy_value <- round(accuracy_normal[(accuracy_normal$Product.Category==product & accuracy_normal$Base==base) &
                                                  (accuracy_normal$Method==col)&(accuracy_normal$Backtesting_period==backtest_period), "Accuracy"],4)
        
        sub_metric_table<-metric_table[metric_table$Product.Category==product & metric_table$Base==base,]
        sub_metric_table<-sub_metric_table[with(sub_metric_table,order(-Accuracy_12,-Accuracy_6)),]
        
        sub_best_method<-best_method[best_method$Product.Category==product & best_method$Base==base,]
        #Save the graph
        p<-try(ggsave(paste("./Step_1/Step_1_pic_output",suffix,"/", sheet, col,backtest_period, '.png', sep=''),
                      ggplot()+
                        geom_line(aes_string(x="Date", y=col, color="Flag"),fcst_df)+
                        geom_line(aes_string(x="Date", y=col), bk_record)+
                        labs(title=paste("Forecast_method:", col), subtitle=paste("Accuracy:", accuracy_value,"\n","Backtesting_period:",backtest_period))+
                        theme(plot.title = element_text(color = "brown3", size = 20, face = "bold", hjust = 0.5),
                              plot.subtitle = element_text(color = "blue", size = 18, hjust = 0.5)),
                      width = 12, height = 5, dpi = 100, device = 'png'),silent = TRUE)
        if('try-error' %in% class(p)) {next}
        
        pic_file <- paste("./Step_1/Step_1_pic_output",suffix,"/", sheet, col, backtest_period,'.png',sep='')
        addPicture(pic_file, sheet=trend_sheet, scale=0.4, startRow = 2+9.5*(which(colnames(fcst_df)==col)-2), 
                   startColumn = 18+8*(which(Backtest_periods==backtest_period)-1))
        
      }
    }
    
    xlsx::addDataFrame(sub_metric_table,sheet=trend_sheet, startRow = 1,startColumn = 1)
    xlsx::addDataFrame(sub_best_method,sheet=trend_sheet,startRow = 18,startColumn = 1)
    print(paste0(product,' ',sheet,': done'))
  }
  xlsx::saveWorkbook(trend_wb,paste0("./Step_1/Step_1_",product,'_', "Forecast_Summary_Plot",suffix,".xlsx"))
}

