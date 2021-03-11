Sys.time()
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
library(xlsx)

#Define fresh year-month to automate file save-in and save-out
fresh_year_month<-paste0(year(Sys.time()),"_",month(Sys.time()),"_")

#Set working folder
Working_Folder<-"C:/ForMe/CPLM/CONS_test/"
setwd(Working_Folder)
setwd(paste0(Working_Folder,year(Sys.time()),"_",month(Sys.time()),"_Refresh"))
Output_Folder<-paste0(Working_Folder,year(Sys.time()),"_",month(Sys.time()),"_Refresh")

#######################################################################################
###--------------------------Import Data and Reformatting---------------------------###
#######################################################################################

input_data<-read.csv("./L2_forecast_input.csv")
master_file<-read.csv("C:/ForMe/CPLM/CONS_test/L2_forecast_method.csv")

forecast_to<-'2022-12-01'
target_start<-'2022-01-01'
target_end<-forecast_to

#--transfer to character--#
for (name in names(input_data)){
  if(name %in% c("Sales.USD","Past.12m.Rev","Past.12m.MS")){
    input_data[,name]<-as.numeric(as.character(input_data[,name]))
  }else{
    input_data[,name]<-as.character(input_data[,name])
  }
}

for (name in names(master_file)){
  if(name %in% c("Backtesting.period","Accuracy")){
    master_file[,name]<-as.numeric(as.character(master_file[,name]))
  }else{
    master_file[,name]<-as.character(master_file[,name])
  }
}


#--merge key--#
input_data$Base_Feature<-paste0(input_data$Base,'_',input_data$Feature)

#--Only keep data for forecast--#
input_data<-input_data[paste0(input_data$Product.Category,input_data$Base_Feature) %in% paste0(master_file$Product.Category,master_file$Base),]
#--!!!!!!date format!!!!!!--#
#input_data$Date<-parse_date_time(sub('-','',input_data$Date), orders = c("%m-%d-%Y"))
input_data$Date<-parse_date_time(sub('-','',input_data$Date), orders = c("%Y-%m-%d"))
head(input_data$Date)

#----------------------------#
#--Missing value Imputation--#
#----------------------------#

#--check missing value--#
input_data_new<-input_data %>% group_by(Product.Category,Base_Feature) %>% mutate(date_lag = as.period(lag(Date) %--% Date)$year * 12 + as.period(lag(Date) %--% Date)$month)
#--All_0 data should not be considerer--#
input_data_new<-input_data_new[!paste0(input_data_new$Product.Category,input_data_new$Base_Feature) %in% paste0(master_file[master_file$Method=='All_0','Product.Category'],
                                                                                                                master_file[master_file$Method=='All_0','Base']),]

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
    if(paste0(product,base) %in% paste0(master_file[master_file$Method=='All_0','Product.Category'],
                                        master_file[master_file$Method=='All_0','Base'])){
      tmp_out<-tmp_2
    } else if(paste0(product,base) %in% paste0(na_over_4$Product.Category,na_over_4$Base_Feature)){
      tmp_out<-tmp_2[tmp_2$Date>= max(na_over_4[na_over_4$Product.Category==product & na_over_4$Base_Feature==base,'Date']),]
    }else{
      tmp_out<-tmp_2
    }
    input_data_model<-rbind(input_data_model,tmp_out)
  }
}

input_data_model_new<-input_data_model %>% group_by(Product.Category,Base_Feature) %>% mutate(date_lag = as.period(lag(Date) %--% Date)$year * 12 + as.period(lag(Date) %--% Date)$month)
input_data_model_new<-input_data_model_new[!paste0(input_data_model_new$Product.Category,input_data_model_new$Base_Feature) %in% paste0(master_file[master_file$Method=='All_0','Product.Category'],
                                                                                                                                        master_file[master_file$Method=='All_0','Base']),]

#--Filter out forecast cases still with missing value--#
na_case<-input_data_model_new[is.na(input_data_model_new$date_lag)==FALSE & input_data_model_new$date_lag>=2,]

model_data_complete<-data.frame()
for (product in unique(input_data_model$Product.Category)){
  tmp_1<-input_data_model[input_data_model$Product.Category==product,]
  for (base in unique(tmp_1$Base_Feature)){
    tmp_2<-tmp_1[tmp_1$Base_Feature==base,]
    if(paste0(product,base) %in% paste0(master_file[master_file$Method=='All_0','Product.Category'],
                                        master_file[master_file$Method=='All_0','Base'])){
      tmp_out<-tmp_2
    }else if(paste0(product,base) %in% paste0(na_case$Product.Category,na_case$Base_Feature)){
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

#write.csv(model_data_complete,'C:/ForMe/CPLM/Automation_Production/Data_Input/Data_auto_test.csv',row.names = FALSE)

accuracy_mape <- function(test, fcst)
{
  diff <- abs(test-fcst)
  mape <- mean(diff / (test + 1))
  
  return(1-mape)
}

######################################################################################
###--------------------------Building forecast container---------------------------###
######################################################################################

#-----Batch Run Pure Forecast model-----#

gen_to_gen_GR_extrapolate <-function(case_data,forecasting_period){
  case_ref<-model_data_complete[model_data_complete$Product.Category== unique(case_data$Product.Category) & model_data_complete$Base_Feature== unique(case_method$Regressor_Method),]
  case_ref<-case_ref%>% mutate(Monthly_diff=(Sales.USD - lag(Sales.USD ,1))/lag(Sales.USD ,1))
  
  diff_ref<-case_ref[case_ref$Date>=as.Date(case_method$Parameter_1,'%m/%d/%Y'),c('Date','Monthly_diff')]
  diff_ref$Week_index<-seq(as.numeric(case_method$Parameter_2),nrow(diff_ref)+as.numeric(case_method$Parameter_2)-1,1)
  
  fcst_file<-data.frame(Date=seq(fcst_start_date,fcst_end_date,by='month'),Fcst=0, Base=unique(case_data$Base_Feature),Product.Category=unique(case_data$Product.Category))
  hist_file<-data.frame(Date=case_data$Date,Fcst=as.numeric(case_data$Sales.USD), Base=unique(case_data$Base_Feature),Product.Category=unique(case_data$Product.Category))
  
  fcst_file$Week_index<-seq(as.numeric(case_method$Parameter_3),nrow(fcst_file)+as.numeric(case_method$Parameter_3)-1,1)
  hist_file$Week_index<-NA
  case_out<-rbind(hist_file,fcst_file)
  
  for(i in 1:nrow(case_out)){
    if(is.na(case_out[i,'Week_index'])==TRUE){
      next
    }
    else{
      diff<-diff_ref[diff_ref$Week_index==case_out[i,'Week_index'],'Monthly_diff']
      new<-case_out[i-1,'Fcst'] * (1+diff)
      case_out[i,'Fcst']<-new
    }
  }
  case_out<-case_out[names(case_out) != 'Week_index']
  return(case_out)
}

month_to_month_GR_extrapolate <- function(case_data,forecasting_period){
  fcst_file<-data.frame(Date=seq(fcst_start_date,fcst_end_date,by='month'),Fcst=0, Base=unique(case_data$Base_Feature),Product.Category=unique(case_data$Product.Category))
  hist_file<-data.frame(Date=case_data$Date,Fcst=as.numeric(case_data$Sales.USD), Base=unique(case_data$Base_Feature),Product.Category=unique(case_data$Product.Category))
  case_ref<-hist_file[hist_file$Date >= as.Date(case_method$Parameter_1,'%m/%d/%Y') & hist_file$Date<= as.Date(case_method$Parameter_2,'%m/%d/%Y'),]
  
  case_out<-rbind(hist_file,fcst_file)
  
  for(i in 1:nrow(case_out)){
    date<-case_out[i,'Date']
    if(date %in% hist_file$Date){
      next
    }else if(month(date) %in% as.numeric(unlist(strsplit(case_method$Parameter_3,',')))){
      case_out[i,'Fcst']<-case_out[case_out$Date== date %m-% months(1),'Fcst'] * 
        case_ref[month(case_ref$Date)==month(date),'Fcst']/case_ref[month(case_ref$Date)==month(date %m-% months(1)),'Fcst']
    }else{
      case_out[i,'Fcst']<-mean(case_out[case_out$Date <= date %m-% months(1) & case_out$Date >= date %m-% months(3),'Fcst'])
    }
  }
  return(case_out)
}

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

forecast_prophet_fit <- function(training_data, forecasting_period)
{
  training_date_list <- as.Date(time(training_data))
  training_data <- as.data.frame(training_data)
  training_data <- cbind(training_data, training_date_list)
  colnames(training_data) <- c('y', 'ds')
  
  model <- prophet(seasonality.mode='additive', daily.seasonality=FALSE, weekly.seasonality=FALSE)
  model <- add_seasonality(model, 'Quarterly', period=4, fourier.order=2)
  model <- fit.prophet(model, training_data)
  future <- make_future_dataframe(model, periods=forecasting_period, freq='month')
  
  fcst <- predict(model, future)
  fcst <-ts(fcst[,'yhat'],frequency = 12, start = c(year(hist_start_date),month(hist_start_date)))
  fcst <- window(fcst, end = c(year(hist_end_date),month(hist_end_date)))
  
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

forecast_prophet_box_cox_fit <- function(training_data, forecasting_period)
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
  
  fcst <-ts(fcst[,'yhat'],frequency = 12, start = c(year(hist_start_date),month(hist_start_date)))
  fcst <- window(fcst, end = c(year(hist_end_date),month(hist_end_date)))
  
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

#-----Batch Run Forecast model + regressor----#

regressor_extract<-function(case_method){
  regressor_file<-data.frame()
  reg_list<-unlist(strsplit(case_method$Regressor_Method,','))
  for (reg in reg_list){
    if(reg=='Specific'){reg_col<-xlsx::read.xlsx(paste0('C:/ForMe/CPLM/CONS_test/L2_forecast_method_Regressor_',product,'.xlsx'),sheetName = base)}
    else if(reg=='Activation'){
      reg_col<-xlsx::read.xlsx(paste0('C:/ForMe/CPLM/CONS_test/L2_forecast_method_Regressor_',product,'.xlsx'),sheetName = reg)
      reg_col<-reg_col[,c('Date',base)]
    }else{reg_col<-xlsx::read.xlsx(paste0('C:/ForMe/CPLM/CONS_test/L2_forecast_method_Regressor_',product,'.xlsx'),sheetName = reg)}
    if(nrow(regressor_file)==0){
      regressor_file<-reg_col
    }else{regressor_file<-merge(regressor_file,reg_col)}
  }
  regressor_file$Date<-as.Date(regressor_file$Date)
  return(regressor_file)
}


#auto.arima
forecast_auto_arima_regressor <- function(training_data, forecasting_period)
{
  regressor_file<-regressor_extract(case_method)
  reg_hist<-regressor_file[regressor_file$Date >= as.Date(hist_start_date) & regressor_file$Date <= as.Date(hist_end_date %m-% months(forecasting_period)),
                           names(regressor_file)[names(regressor_file)!='Date']]
  reg_new<-regressor_file[regressor_file$Date > as.Date(hist_end_date %m-% months(forecasting_period)) & 
                            regressor_file$Date <= as.Date(hist_end_date %m+% months(ifelse(forecasting_period==0,fcst_length,0)))
                          ,names(regressor_file)[names(regressor_file)!='Date']]
  fit <- auto.arima(training_data, D=1, stepwise=FALSE, approximation=FALSE, parallel=TRUE,xreg = data.matrix(reg_hist))
  fcst <- forecast(fit, h=ifelse(forecasting_period==0,fcst_length,forecasting_period), xreg=data.matrix(reg_new))
  return(fcst$mean)
}

forecast_auto_arima_regressor_fit <- function(training_data, forecasting_period)
{
  regressor_file<-regressor_extract(case_method)
  reg_hist<-regressor_file[regressor_file$Date >= as.Date(hist_start_date) & regressor_file$Date <= as.Date(hist_end_date),
                           names(regressor_file)[names(regressor_file)!='Date']]
  fit <- auto.arima(training_data, D=1, stepwise=FALSE, approximation=FALSE, parallel=TRUE,xreg = data.matrix(reg_hist))
  return(fitted(fit))
}

forecast_auto_arima_box_cox_regressor <- function(training_data, forecasting_period)
{
  regressor_file<-regressor_extract(case_method)
  reg_hist<-regressor_file[regressor_file$Date >= as.Date(hist_start_date) & regressor_file$Date <= as.Date(hist_end_date %m-% months(forecasting_period)),
                           names(regressor_file)[names(regressor_file)!='Date']]
  reg_new<-regressor_file[regressor_file$Date > as.Date(hist_end_date %m-% months(forecasting_period)) & 
                            regressor_file$Date <= as.Date(hist_end_date %m+% months(ifelse(forecasting_period==0,fcst_length,0)))
                          ,names(regressor_file)[names(regressor_file)!='Date']]
  transform_coder <- BoxCox.lambda(training_data)
  train_data_temp <- BoxCox(training_data, transform_coder)
  fit <- auto.arima(train_data_temp, D=1, stepwise=FALSE, approximation=FALSE, parallel=TRUE,xreg = data.matrix(reg_hist))
  fcst <- forecast(fit, h=ifelse(forecasting_period==0,fcst_length,forecasting_period), xreg=data.matrix(reg_new))
  fcst <- InvBoxCox(fcst$mean, transform_coder)
  return(fcst)
}

forecast_auto_arima_box_cox_regressor_fit <- function(training_data, forecasting_period)
{
  regressor_file<-regressor_extract(case_method)
  reg_hist<-regressor_file[regressor_file$Date >= as.Date(hist_start_date) & regressor_file$Date <= as.Date(hist_end_date),
                           names(regressor_file)[names(regressor_file)!='Date']]
  transform_coder <- BoxCox.lambda(training_data)
  train_data_temp <- BoxCox(training_data, transform_coder)
  fit <- auto.arima(train_data_temp, D=1, stepwise=FALSE, approximation=FALSE, parallel=TRUE,xreg = data.matrix(reg_hist))
  return(InvBoxCox(fitted(fit),transform_coder))
}

#stl
forecast_stl_regressor <- function(training_data, forecasting_period)
{
  regressor_file<-regressor_extract(case_method)
  reg_hist<-regressor_file[regressor_file$Date >= as.Date(hist_start_date) & regressor_file$Date <= as.Date(hist_end_date %m-% months(forecasting_period)),
                           names(regressor_file)[names(regressor_file)!='Date']]
  reg_new<-regressor_file[regressor_file$Date > as.Date(hist_end_date %m-% months(forecasting_period)) & 
                            regressor_file$Date <= as.Date(hist_end_date %m+% months(ifelse(forecasting_period==0,fcst_length,0)))
                          ,names(regressor_file)[names(regressor_file)!='Date']]
  fit <- stlm(training_data, s.window='periodic',method='arima',xreg = data.matrix(reg_hist))
  fcst <- forecast(fit, h=ifelse(forecasting_period==0,fcst_length,forecasting_period), newxreg = data.matrix(reg_new))
  return(fcst$mean)
}

forecast_stl_regressor_fit <- function(training_data, forecasting_period)
{
  regressor_file<-regressor_extract(case_method)
  reg_hist<-regressor_file[regressor_file$Date >= as.Date(hist_start_date) & regressor_file$Date <= as.Date(hist_end_date),
                           names(regressor_file)[names(regressor_file)!='Date']]
  fit <- stlm(training_data, s.window='periodic',method='arima',xreg = data.matrix(reg_hist))
  return(fitted(fit))
}

forecast_stl_box_cox_regressor <- function(training_data, forecasting_period)
{
  regressor_file<-regressor_extract(case_method)
  reg_hist<-regressor_file[regressor_file$Date >= as.Date(hist_start_date) & regressor_file$Date <= as.Date(hist_end_date %m-% months(forecasting_period)),
                           names(regressor_file)[names(regressor_file)!='Date']]
  reg_new<-regressor_file[regressor_file$Date > as.Date(hist_end_date %m-% months(forecasting_period)) & 
                            regressor_file$Date <= as.Date(hist_end_date %m+% months(ifelse(forecasting_period==0,fcst_length,0)))
                          ,names(regressor_file)[names(regressor_file)!='Date']]
  transform_coder <- BoxCox.lambda(training_data)
  train_data_temp <- BoxCox(training_data, transform_coder)
  fit <- stlm(train_data_temp, s.window='periodic',method='arima',xreg = data.matrix(reg_hist))
  fcst <- forecast(fit, h=ifelse(forecasting_period==0,fcst_length,forecasting_period), newxreg = data.matrix(reg_new))
  fcst <- InvBoxCox(fcst$mean, transform_coder)
  return(fcst$mean)
}

forecast_stl_box_cox_regressor_fit <- function(training_data, forecasting_period)
{
  regressor_file<-regressor_extract(case_method)
  reg_hist<-regressor_file[regressor_file$Date >= as.Date(hist_start_date) & regressor_file$Date <= as.Date(hist_end_date),
                           names(regressor_file)[names(regressor_file)!='Date']]
  transform_coder <- BoxCox.lambda(training_data)
  train_data_temp <- BoxCox(training_data, transform_coder)
  fit <- stlm(train_data_temp, s.window='periodic',method='arima',xreg = data.matrix(reg_hist))
  return(InvBoxCox(fitted(fit),transform_coder))
}

#prophet
forecast_prophet_regressor <- function(training_data, forecasting_period)
{
  regressor_file<-regressor_extract(case_method)
  reg_hist<-regressor_file[regressor_file$Date >= as.Date(hist_start_date) & regressor_file$Date <= as.Date(hist_end_date %m-% months(forecasting_period)),
                           names(regressor_file)[names(regressor_file)!='Date']]
  reg_new<-regressor_file[regressor_file$Date > as.Date(hist_end_date %m-% months(forecasting_period)) & 
                            regressor_file$Date <= as.Date(hist_end_date %m+% months(ifelse(forecasting_period==0,fcst_length,0)))
                          ,names(regressor_file)[names(regressor_file)!='Date']]
  training_date_list <- as.Date(time(training_data))
  training_data <- as.data.frame(training_data)
  training_data <- cbind(training_data, training_date_list,reg_hist)
  colnames(training_data)[c(1,2)] <- c('y', 'ds')
  
  model <- prophet(seasonality.mode='additive', daily.seasonality=FALSE, weekly.seasonality=FALSE)
  model <- add_seasonality(model, 'Quarterly', period=4, fourier.order=2)
  
  for (i in 3:ncol(training_data)){
    model <- add_regressor(model, colnames(training_data)[i], standardize = FALSE)
  }
  model <- fit.prophet(model, training_data)
  
  future <- make_future_dataframe(model, periods=ifelse(forecasting_period==0,fcst_length,forecasting_period), freq='month')
  future[,colnames(training_data)[3:ncol(training_data)]]<-rbind(as.matrix(reg_hist),as.matrix(reg_new))
  
  future_start<-max(training_date_list) %m+% months(1)
  fcst <- predict(model, future)
  fcst <- ts(fcst[fcst$ds>= future_start,'yhat'],frequency = 12, start = c(year(future_start),month(future_start)))
  
  return(fcst)
}

forecast_prophet_regressor_fit <- function(training_data, forecasting_period)
{
  regressor_file<-regressor_extract(case_method)
  reg_hist<-regressor_file[regressor_file$Date >= as.Date(hist_start_date) & regressor_file$Date <= as.Date(hist_end_date),
                           names(regressor_file)[names(regressor_file)!='Date']]
  reg_new<-regressor_file[regressor_file$Date > as.Date(hist_end_date %m-% months(forecasting_period)) & 
                            regressor_file$Date <= as.Date(hist_end_date %m+% months(ifelse(forecasting_period==0,fcst_length,0)))
                          ,names(regressor_file)[names(regressor_file)!='Date']]
  training_date_list <- as.Date(time(training_data))
  training_data <- as.data.frame(training_data)
  training_data <- cbind(training_data, training_date_list,reg_hist)
  colnames(training_data)[c(1,2)] <- c('y', 'ds')
  
  model <- prophet(seasonality.mode='additive', daily.seasonality=FALSE, weekly.seasonality=FALSE)
  model <- add_seasonality(model, 'Quarterly', period=4, fourier.order=2)
  
  for (i in 3:ncol(training_data)){
    model <- add_regressor(model, colnames(training_data)[i], standardize = FALSE)
  }
  model <- fit.prophet(model, training_data)
  
  future <- make_future_dataframe(model, periods=ifelse(forecasting_period==0,fcst_length,forecasting_period), freq='month')
  future[,colnames(training_data)[3:ncol(training_data)]]<-rbind(as.matrix(reg_hist),as.matrix(reg_new))
  
  fcst <- predict(model, future)
  fcst <- ts(fcst[fcst$ds<= hist_end_date,'yhat'],frequency = 12, start = c(year(hist_start_date),month(hist_start_date)))
  return(fcst)
}

forecast_prophet_box_cox_regressor <- function(training_data, forecasting_period)
{
  regressor_file<-regressor_extract(case_method)
  reg_hist<-regressor_file[regressor_file$Date >= as.Date(hist_start_date) & regressor_file$Date <= as.Date(hist_end_date %m-% months(forecasting_period)),
                           names(regressor_file)[names(regressor_file)!='Date']]
  reg_new<-regressor_file[regressor_file$Date > as.Date(hist_end_date %m-% months(forecasting_period)) & 
                            regressor_file$Date <= as.Date(hist_end_date %m+% months(ifelse(forecasting_period==0,fcst_length,0)))
                          ,names(regressor_file)[names(regressor_file)!='Date']]
  transform_coder <- BoxCox.lambda(training_data)
  train_data_temp <- BoxCox(training_data, transform_coder)
  training_date_list <- as.Date(time(training_data))
  train_data_temp<- as.data.frame(train_data_temp)
  train_data_temp <- cbind(train_data_temp, training_date_list,reg_hist)
  colnames(train_data_temp)[c(1,2)] <- c('y', 'ds')
  
  model <- prophet(seasonality.mode='additive', daily.seasonality=FALSE, weekly.seasonality=FALSE)
  model <- add_seasonality(model, 'Quarterly', period=4, fourier.order=2)
  
  for (i in 3:ncol(train_data_temp)){
    model <- add_regressor(model, colnames(train_data_temp)[i], standardize = FALSE)
  }
  model <- fit.prophet(model, train_data_temp)
  
  future <- make_future_dataframe(model, periods=ifelse(forecasting_period==0,fcst_length,forecasting_period), freq='month')
  future[,colnames(training_data)[3:ncol(training_data)]]<-rbind(as.matrix(reg_hist),as.matrix(reg_new))
  
  future_start<-max(training_date_list) %m+% months(1)
  fcst <- predict(model, future)
  fcst$yhat<-InvBoxCox(fcst$yhat,transform_coder)
  fcst <- ts(fcst[fcst$ds>= future_start,'yhat'],frequency = 12, start = c(year(future_start),month(future_start)))
  
  return(fcst)
}

forecast_prophet_box_cox_regressor_fit <- function(training_data, forecasting_period)
{
  regressor_file<-regressor_extract(case_method)
  reg_hist<-regressor_file[regressor_file$Date >= as.Date(hist_start_date) & regressor_file$Date <= as.Date(hist_end_date),
                           names(regressor_file)[names(regressor_file)!='Date']]
  reg_new<-regressor_file[regressor_file$Date > as.Date(hist_end_date %m-% months(forecasting_period)) & 
                            regressor_file$Date <= as.Date(hist_end_date %m+% months(ifelse(forecasting_period==0,fcst_length,0)))
                          ,names(regressor_file)[names(regressor_file)!='Date']]
  transform_coder <- BoxCox.lambda(training_data)
  train_data_temp <- BoxCox(training_data, transform_coder)
  training_date_list <- as.Date(time(training_data))
  train_data_temp<- as.data.frame(train_data_temp)
  train_data_temp <- cbind(train_data_temp, training_date_list,reg_hist)
  colnames(train_data_temp)[c(1,2)] <- c('y', 'ds')
  
  model <- prophet(seasonality.mode='additive', daily.seasonality=FALSE, weekly.seasonality=FALSE)
  model <- add_seasonality(model, 'Quarterly', period=4, fourier.order=2)
  
  for (i in 3:ncol(train_data_temp)){
    model <- add_regressor(model, colnames(train_data_temp)[i], standardize = FALSE)
  }
  model <- fit.prophet(model, train_data_temp)
  
  future <- make_future_dataframe(model, periods=ifelse(forecasting_period==0,fcst_length,forecasting_period), freq='month')
  future[,colnames(training_data)[3:ncol(training_data)]]<-rbind(as.matrix(reg_hist),as.matrix(reg_new))
  
  fcst <- predict(model, future)
  fcst$yhat<-InvBoxCox(fcst$yhat,transform_coder)
  fcst <- ts(fcst[fcst$ds<= hist_end_date,'yhat'],frequency = 12, start = c(year(hist_start_date),month(hist_start_date)))
  
  return(fcst)
}

#-----enhanced stl----#

forecast_enhance_stl<-function(training_data,forecasting_period){
  STL_forecast<-stlf(training_data,t.window = as.numeric(case_method$Parameter_1), s.window = as.numeric(case_method$Parameter_2),
                     method='arima', robust = TRUE, h=forecasting_period, lambda = as.numeric(case_method$Parameter_3), biasadj = TRUE,
                     allow.multiplicative.trend=TRUE)
  return(STL_forecast$mean)
}


forecast_enhance_stl_fit<-function(training_data,forecasting_period){
  STL_forecast<-stlf(training_data,t.window = as.numeric(case_method$Parameter_1), s.window = as.numeric(case_method$Parameter_2),
                     method='arima', robust = TRUE, h=forecasting_period, lambda = as.numeric(case_method$Parameter_3), biasadj = TRUE,
                     allow.multiplicative.trend=TRUE)
  return(fitted(STL_forecast))
}


forecast_enhance_stl_regressor<-function(training_data,forecasting_period){
  regressor_file<-regressor_extract(case_method)
  reg_hist<-regressor_file[regressor_file$Date >= as.Date(hist_start_date) & regressor_file$Date <= as.Date(hist_end_date %m-% months(forecasting_period)),
                           names(regressor_file)[names(regressor_file)!='Date']]
  reg_new<-regressor_file[regressor_file$Date > as.Date(hist_end_date %m-% months(forecasting_period)) & 
                            regressor_file$Date <= as.Date(hist_end_date %m+% months(ifelse(forecasting_period==0,fcst_length,0)))
                          ,names(regressor_file)[names(regressor_file)!='Date']]
    STL_forecast<-stlf(training_data,t.window = as.numeric(case_method$Parameter_1), s.window = as.numeric(case_method$Parameter_2),
                     method='arima', robust = TRUE, h=ifelse(forecasting_period==0,fcst_length,forecasting_period), lambda = as.numeric(case_method$Parameter_3), biasadj = TRUE,
                     allow.multiplicative.trend=TRUE, xreg=data.matrix(reg_hist),newxreg = data.matrix(reg_new))
  return(STL_forecast$mean)
}

forecast_enhance_stl_regressor_fit<-function(training_data,forecasting_period){
  regressor_file<-regressor_extract(case_method)
  reg_hist<-regressor_file[regressor_file$Date >= as.Date(hist_start_date) & regressor_file$Date <=as.Date(hist_end_date),
                           names(regressor_file)[names(regressor_file)!='Date']]
  reg_new<-regressor_file[regressor_file$Date > as.Date(hist_end_date %m-% months(forecasting_period)) & 
                            regressor_file$Date <= as.Date(hist_end_date %m+% months(ifelse(forecasting_period==0,fcst_length,0)))
                          ,names(regressor_file)[names(regressor_file)!='Date']]
  STL_forecast<-stlf(training_data,t.window = as.numeric(case_method$Parameter_1), s.window = as.numeric(case_method$Parameter_2),
                     method='arima', robust = TRUE, h=forecasting_period, lambda = as.numeric(case_method$Parameter_3), biasadj = TRUE,
                     allow.multiplicative.trend=TRUE,  xreg=data.matrix(reg_hist),newxreg = data.matrix(reg_new))
  return(fitted(STL_forecast))
}

###############################################################################################
###--------------------------Forecast Batch Run and Accuracy Test---------------------------###
###############################################################################################

Forecast_out<-data.frame()
Accuracy_Update<-data.frame()
for (product in unique(model_data_complete$Product.Category)){
  for (base in unique(model_data_complete[model_data_complete$Product.Category==product,'Base_Feature'])){
    case_data<-model_data_complete[model_data_complete$Product.Category==product & model_data_complete$Base_Feature==base,]
    case_out<-data.frame()
    Accuracy_Test<-data.frame()
    case_method<-master_file[master_file$Product.Category==product & master_file$Base==base,]
    
    hist_start_date<-case_data$Date[1];hist_end_date<-case_data$Date[nrow(case_data)]
    hist_length<-nrow(case_data)
    fcst_start_date<-hist_end_date %m+% months(1)
    fcst_end_date<-parse_date_time(sub('-','',forecast_to), orders = c("%Y-%m-%d"))
    fcst_length<-as.period( hist_end_date %--% fcst_end_date)$year * 12 + as.period( hist_end_date %--% fcst_end_date)$month
    
    Accuracy_Test[1,'Product.Category']<-product
    Accuracy_Test[1,'Base']<-base
    
    if(case_method$Method== 'All_0'){
           
      fcst_file<-data.frame(Date=seq(fcst_start_date,fcst_end_date,by='month'),Fcst=0, Base=unique(case_data$Base_Feature),Product.Category=unique(case_data$Product.Category))
      hist_file<-data.frame(Date=case_data$Date,Fcst=as.numeric(case_data$Sales.USD), Base=unique(case_data$Base_Feature),Product.Category=unique(case_data$Product.Category))
      
      case_out<-rbind(hist_file,fcst_file)
     
      Accuracy_Test[1,'back_test_6']<-NA
      Accuracy_Test[1,'back_test_12']<-NA
      Accuracy_Test[1,'fitted_test_6']<-NA
      Accuracy_Test[1,'fitted_test_12']<-NA
      
    }else if(case_method$Method %in% c('month_to_month_GR_extrapolate','gen_to_gen_GR_extrapolate') ){
      case_out<-do.call(case_method$Method,list(case_data,fcst_length))
      
      Accuracy_Test[1,'back_test_6']<-NA
      Accuracy_Test[1,'back_test_12']<-NA
      Accuracy_Test[1,'fitted_test_6']<-NA
      Accuracy_Test[1,'fitted_test_12']<-NA
    }else{
      case_data$Sales.USD<-ts(case_data$Sales.USD,start = c(year(hist_start_date),month(hist_start_date)),frequency = 12)
      Fcst<-if(grepl('regressor',case_method$Method)){do.call(case_method$Method,list(case_data$Sales.USD,0))}else{
        do.call(case_method$Method,list(case_data$Sales.USD,fcst_length))
      }
      
      fcst_file<-data.frame(Date=seq(fcst_start_date,fcst_end_date,by='month'),Fcst=as.numeric(Fcst), Base=unique(case_data$Base_Feature),Product.Category=unique(case_data$Product.Category))
      hist_file<-data.frame(Date=seq(hist_start_date,hist_end_date,by='month'),Fcst=as.numeric(case_data$Sales.USD), Base=unique(case_data$Base_Feature),Product.Category=unique(case_data$Product.Category))
      
      case_out<-rbind(hist_file,fcst_file)

      #---back test---#
      if(
        case_method$Method %in% c('forecast_ets', 'forecast_hw','forecast_stl','forecast_combined','forecast_combinedv2','forecast_prophet',
                                  'forecast_ets_box_cox', 'forecast_hw_box_cox','forecast_stl_box_cox','forecast_combined_box_cox','forecast_prophet_box_cox')
        ){
          train_6<-window(case_data$Sales.USD,end=c(year(hist_end_date %m-% months(6)),month(hist_end_date %m-% months(6))))
          test_6<-window(case_data$Sales.USD,start=c(year(hist_end_date %m-% months(5)),month(hist_end_date %m-% months(5))))
          fcst_6<-do.call(case_method$Method,list(train_6,6))
          
          train_12<-window(case_data$Sales.USD,end=c(year(hist_end_date %m-% months(12)),month(hist_end_date %m-% months(12))))
          test_12<-window(case_data$Sales.USD,start=c(year(hist_end_date %m-% months(11)),month(hist_end_date %m-% months(11))))
          if(length(train_12)>12){fcst_12<-do.call(case_method$Method,list(train_12,12))}else{fcst_12<-NA}
          

          Accuracy_Test[1,'back_test_6']<-accuracy_mape(test_6,fcst_6)
          Accuracy_Test[1,'back_test_12']<-accuracy_mape(test_12,fcst_12)
          
          Accuracy_Test[1,'fitted_test_6']<-NA
          Accuracy_Test[1,'fitted_test_12']<-NA

      }
      else if(
        grepl('stl',case_method$Method)
      ){
        train_6<-window(case_data$Sales.USD,end=c(year(hist_end_date %m-% months(6)),month(hist_end_date %m-% months(6))))
        test_6<-window(case_data$Sales.USD,start=c(year(hist_end_date %m-% months(5)),month(hist_end_date %m-% months(5))))
        
        train_12<-window(case_data$Sales.USD,end=c(year(hist_end_date %m-% months(12)),month(hist_end_date %m-% months(12))))
        test_12<-window(case_data$Sales.USD,start=c(year(hist_end_date %m-% months(11)),month(hist_end_date %m-% months(11))))
        
        if(length(train_6)<=24){
          Accuracy_Test[1,'back_test_6']<-NA
          Accuracy_Test[1,'back_test_12']<-NA
          fitted_6<-window(do.call(paste0(case_method$Method,'_fit'),list(case_data$Sales.USD,fcst_length)),start=c(year(hist_end_date %m-% months(5)),month(hist_end_date %m-% months(5))))
          fitted_12<-window(do.call(paste0(case_method$Method,'_fit'),list(case_data$Sales.USD,fcst_length)),start=c(year(hist_end_date %m-% months(11)),month(hist_end_date %m-% months(11))))
          Accuracy_Test[1,'fitted_test_6']<-accuracy_mape(fitted_6,window(case_data$Sales.USD,start=c(year(hist_end_date %m-% months(5)),month(hist_end_date %m-% months(5)))))
          Accuracy_Test[1,'fitted_test_12']<-accuracy_mape(fitted_12,window(case_data$Sales.USD,start=c(year(hist_end_date %m-% months(11)),month(hist_end_date %m-% months(11)))))
          
        }else{
          train_6<-window(case_data$Sales.USD,end=c(year(hist_end_date %m-% months(6)),month(hist_end_date %m-% months(6))))
          test_6<-window(case_data$Sales.USD,start=c(year(hist_end_date %m-% months(5)),month(hist_end_date %m-% months(5))))
          fcst_6<-do.call(case_method$Method,list(train_6,6))
          
          train_12<-window(case_data$Sales.USD,end=c(year(hist_end_date %m-% months(12)),month(hist_end_date %m-% months(12))))
          test_12<-window(case_data$Sales.USD,start=c(year(hist_end_date %m-% months(11)),month(hist_end_date %m-% months(11))))
          fcst_12<-do.call(case_method$Method,list(train_12,12))
          
          fitted_6<-window(do.call(paste0(case_method$Method,'_fit'),list(case_data$Sales.USD,fcst_length)),start=c(year(hist_end_date %m-% months(5)),month(hist_end_date %m-% months(5))))
          fitted_12<-window(do.call(paste0(case_method$Method,'_fit'),list(case_data$Sales.USD,fcst_length)),start=c(year(hist_end_date %m-% months(11)),month(hist_end_date %m-% months(11))))
          
          Accuracy_Test[1,'back_test_6']<-accuracy_mape(test_6,fcst_6)
          Accuracy_Test[1,'back_test_12']<-accuracy_mape(test_12,fcst_12)
          
          Accuracy_Test[1,'fitted_test_6']<-accuracy_mape(fitted_6,window(case_data$Sales.USD,start=c(year(hist_end_date %m-% months(5)),month(hist_end_date %m-% months(5)))))
          Accuracy_Test[1,'fitted_test_12']<-accuracy_mape(fitted_12,window(case_data$Sales.USD,start=c(year(hist_end_date %m-% months(11)),month(hist_end_date %m-% months(11)))))
          
        }
      }
      
        else{
          train_6<-window(case_data$Sales.USD,end=c(year(hist_end_date %m-% months(6)),month(hist_end_date %m-% months(6))))
          test_6<-window(case_data$Sales.USD,start=c(year(hist_end_date %m-% months(5)),month(hist_end_date %m-% months(5))))
          if(length(train_6)>12){fcst_6<-do.call(case_method$Method,list(train_6,6))}else{fcst_6<-NA}
          
          train_12<-window(case_data$Sales.USD,end=c(year(hist_end_date %m-% months(12)),month(hist_end_date %m-% months(12))))
          test_12<-window(case_data$Sales.USD,start=c(year(hist_end_date %m-% months(11)),month(hist_end_date %m-% months(11))))
          if(length(train_12)>12){fcst_12<-do.call(case_method$Method,list(train_12,12))}else{fcst_12<-NA}
          

          fitted_6<-window(do.call(paste0(case_method$Method,'_fit'),list(case_data$Sales.USD,fcst_length)),start=c(year(hist_end_date %m-% months(5)),month(hist_end_date %m-% months(5))))
          fitted_12<-window(do.call(paste0(case_method$Method,'_fit'),list(case_data$Sales.USD,fcst_length)),start=c(year(hist_end_date %m-% months(11)),month(hist_end_date %m-% months(11))))
          
          Accuracy_Test[1,'back_test_6']<-accuracy_mape(test_6,fcst_6)
          Accuracy_Test[1,'back_test_12']<-accuracy_mape(test_12,fcst_12)
           
          Accuracy_Test[1,'fitted_test_6']<-accuracy_mape(fitted_6,window(case_data$Sales.USD,start=c(year(hist_end_date %m-% months(5)),month(hist_end_date %m-% months(5)))))
          Accuracy_Test[1,'fitted_test_12']<-accuracy_mape(fitted_12,window(case_data$Sales.USD,start=c(year(hist_end_date %m-% months(11)),month(hist_end_date %m-% months(11)))))
         
        }
    }
    Forecast_out<-rbind(Forecast_out,case_out)
    Accuracy_Update<-rbind(Accuracy_Update,Accuracy_Test)
  }
}
Sys.time()

#--Adjust format and ouput--#
Forecast_out$Fcst<-abs(Forecast_out$Fcst)
Output_full<-merge(Forecast_out,input_data[,c('Product.Category','Base_Feature','Date','Sales.USD')],by.x = c('Date','Product.Category','Base'),
                   by.y = c('Date','Product.Category','Base_Feature'),all = TRUE)

names(Output_full)[names(Output_full)=='Sales.USD']<-'Actual'

#--Write Out files--#
write.csv(Output_full,paste0("./L2_forecast_out_",year(Sys.time()),'_',month(Sys.time()),'.csv'),row.names = FALSE)


##########################################################################################
###--------------------------Compare with last month refresh---------------------------###
##########################################################################################
Last_update<-read.csv(paste0("C:/ForMe/CPLM/CONS_test/", year(Sys.time() %m-% months(1)),'_',month(Sys.time() %m-% months(1)),"_Refresh/L2_forecast_out_",
                             year(Sys.time() %m-% months(1)),'_',month(Sys.time() %m-% months(1)),".csv"))

Last_update$Date<-as.Date(Last_update$Date,'%m/%d/%Y')
names(Last_update)[names(Last_update)=='Fcst']<-'Fcst_Last'

dir.create('./refresh_compare/')

for (product in unique(Last_update$Product.Category)){
  if(product=='TRA'){next}
  for( base in unique(Last_update[Last_update$Product.Category==product,'Base'])){
    if(!base %in% unique(Output_full[Output_full$Product.Category==product,'Base'])){next}
    Last_Case<-Last_update[Last_update$Product.Category==product & Last_update$Base==base,]
    Current_Case<-Output_full[Output_full$Product.Category==product & Output_full$Base==base,]
    Current_Case<-Current_Case[is.na(Current_Case$Fcst)==FALSE,]
    Last_Case<-Last_Case[is.na(Last_Case$Fcst_Last)==FALSE,]
    Current_Case$Fcst<-ts(Current_Case$Fcst,start = c(year(Current_Case[1,'Date']),month(Current_Case[1,'Date'])),frequency = 12)
    Last_Case$Fcst_Last<-ts(Last_Case$Fcst_Last,start = c(year(Last_Case[1,'Date']),month(Last_Case[1,'Date'])),frequency = 12)
    jpeg(file=paste0('./refresh_compare/',product,'_',base,'.jpeg'),width = 900,height = 450)
    print(autoplot(Current_Case$Fcst)+ autolayer(Last_Case$Fcst_Last) + ggtitle(paste0(product,'_',base)))
    dev.off()
  }
}


#################################################################################################
###--------------------------Filiter Out Low Accuracy Forecast Case---------------------------###
#################################################################################################
accuracy_threshold<-0.8
Accuracy_Highlight<-data.frame()
for (i in 1:nrow(Accuracy_Update)){
  base<-Accuracy_Update[i,'Base']
  product<-Accuracy_Update[i,'Product.Category']
  tmp<-data.frame()
  if(all(is.na(Accuracy_Update[i,grepl('_test_',names(Accuracy_Update))]))){
      next
  }else{
    Accuracy_ref<-master_file[master_file$Product.Category==product & master_file$Base==base,'Accuracy']
    Accuracy_Best<-max(Accuracy_Update[i,grepl('_test_',names(Accuracy_Update))],na.rm = TRUE)
    if(Accuracy_Best< min(Accuracy_ref,accuracy_threshold)){
      tmp[1,'Product.Category']<-product
      tmp[1,'Base']<-base
      tmp[1,'Accuracy_run']<-Accuracy_Best
      tmp[1,'Accuracy_ref']<-Accuracy_ref
    }else{
      next
    }
  }
  Accuracy_Highlight<-rbind(Accuracy_Highlight,tmp)
}

if(nrow(Accuracy_Highlight)>1){
  print('=======++++++++++++++++++Alert Low Accuracy Case++++++++++++++++++=======')
  print(Accuracy_Highlight)
  write.csv(Accuracy_Highlight,paste0('./L2_Accuracy_Highlight_',year(Sys.time()),'_',month(Sys.time()),'.csv'),row.names = FALSE)
  print('=========================+++++++++++++++++++++++=========================')
  }


