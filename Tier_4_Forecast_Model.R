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

suffix<-'_test'
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
case_select<-read.csv("./step_3/Segments_for_Step_4.csv")

case_select$Regressor_Method<-'Specific'

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
    if(date_length<=24){
      if(nrow(short_case)==0){
        short_case=sub_data
      }else{
        short_case=rbind(short_case,sub_data)
      }
    }
  }
}

if(nrow(short_case)>=1){
  print("############----------Short Case Alert----------############")
  write.csv(short_case,'./Step_4/Short_case_Less_Than_24_Months_retrain.csv',row.names = FALSE)
}

model_data_complete<-model_data_complete[!paste0(model_data_complete$Product.Category,model_data_complete$Base_Feature) %in% 
                                           paste0(short_case$Product.Category, short_case$Base_Feature),]

accuracy_mape <- function(test, fcst)
{
  diff <- abs(test-fcst)
  mape <- mean(diff / (test + 1))
  
  return(1-mape)
}

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

options("scipen"=100, "digits"=4)

####################################################################################
###-------------------Build Model and Save all Accruacy Result-------------------###
####################################################################################
dir.create("./Step_4/")
t_window_list<-c(1,3,5,7,9,11,13,15,23,25)
s_window_list<-c(1,3,5,7,9,11,13,25,100000)
lambda_list<-c(0.01,0.05,0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9,0.95,1)

Backtesting_periods <- c(6,12)
forecast_methods<-c('forecast_enhance_stl_regressor','forecast_enhance_stl_regressor_fit')
for (product in unique(model_data_complete$Product.Category)){
  accuracy_wb <- xlsx::createWorkbook('xlsx')
  for (base in unique(model_data_complete[model_data_complete$Product.Category==product,'Base_Feature'])){
    accuracy_record <- data.frame()
    case_data<-model_data_complete[model_data_complete$Product.Category==product & model_data_complete$Base_Feature==base,]
    case_data$Sales.USD<-ts(case_data$Sales.USD,start = c(year(min(case_data$Date)),month(min(case_data$Date))),frequency = 12)
    case_method<-case_select[case_select$Product.Category==product & case_select$Base==base,]
    
    for (backtesting_period in Backtesting_periods){
      train_data<-window(case_data$Sales.USD, end =c(year(max(case_data$Date)%m-% months(backtesting_period)),month(max(case_data$Date)%m-% months(backtesting_period))))
      test_data<-window(case_data$Sales.USD,start=c(year(max(case_data$Date)%m-% months(backtesting_period-1)),month(max(case_data$Date)%m-% months(backtesting_period-1))))
      for (method in forecast_methods){
        if(grepl('fit',method)==FALSE){
          
          hist_start_date<-case_data$Date[1];hist_end_date<-case_data$Date[nrow(case_data)]
          regressor_file<-regressor_extract(case_method)
          reg_hist<-regressor_file[regressor_file$Date >= as.Date(hist_start_date) & regressor_file$Date <= as.Date(hist_end_date %m-% months(backtesting_period)),
                                   names(regressor_file)[names(regressor_file)!='Date']]
          reg_new<-regressor_file[regressor_file$Date > as.Date(hist_end_date %m-% months(backtesting_period)) & 
                                    regressor_file$Date <= as.Date(hist_end_date %m+% months(0)),names(regressor_file)[names(regressor_file)!='Date']]
          
          for(t_win in t_window_list){
            for (s_win in s_window_list){
              for(l in lambda_list){
                temp<-data.frame()
                set.seed(123)
                temp[1,'Product.Category']<-product
                temp[1,'Base']<-base
                temp[1,'t_window']<-t_win
                temp[1,'s_window']<-s_win
                temp[1,'lambda']<-l
                temp[1,'Backtesting_period']<-backtesting_period
                temp[1,'Method']<-method
                if(length(train_data)<=24){next}
                STL_forecast<-stlf(train_data,
                                   t.window = t_win, s.window = s_win, method='arima', robust = TRUE, h=backtesting_period, lambda = l, biasadj = TRUE, allow.multiplicative.trend=TRUE,
                                   xreg=data.matrix(reg_hist),newxreg = data.matrix(reg_new))
                temp[1,'Accuracy']<-accuracy_mape(test_data,STL_forecast$mean)
                #temp<-temp[sort(names(temp))]
                accuracy_record<-rbind(accuracy_record,temp)
            }
            }
          }
        }else{
          #fitted accuracy
          hist_start_date<-case_data$Date[1];hist_end_date<-case_data$Date[nrow(case_data)]
          regressor_file<-regressor_extract(case_method)
          reg_hist<-regressor_file[regressor_file$Date >= as.Date(hist_start_date) & regressor_file$Date <= as.Date(hist_end_date),
                                   names(regressor_file)[names(regressor_file)!='Date']]
          #new regressor has no actual impact here, just to impute newreg in the function, coz we need fitted accuracy ONLY
          reg_new<-regressor_file[regressor_file$Date > as.Date(hist_end_date %m-% months(backtesting_period)) & 
                                    regressor_file$Date <= as.Date(hist_end_date %m+% months(0)),names(regressor_file)[names(regressor_file)!='Date']]
          for(t_win in t_window_list){
            for (s_win in s_window_list){
              for(l in lambda_list){
                temp<-data.frame()
                set.seed(123)
                temp[1,'Product.Category']<-product
                temp[1,'Base']<-base
                temp[1,'t_window']<-t_win
                temp[1,'s_window']<-s_win
                temp[1,'lambda']<-l
                temp[1,'Backtesting_period']<-backtesting_period
                temp[1,'Method']<-method
                if(length(case_data$Sales.USD)<=24){next}
                STL_forecast_fit<-stlf(case_data$Sales.USD,
                                   t.window = t_win, s.window = s_win, method='arima', robust = TRUE, h=backtesting_period, lambda = l, biasadj = TRUE, allow.multiplicative.trend=TRUE,
                                   xreg=data.matrix(reg_hist),newxreg = data.matrix(reg_new))
                temp[1,'Accuracy']<-accuracy_mape(window(case_data$Sales.USD,start=c(year(max(case_data$Date)%m-% months(backtesting_period-1)),month(max(case_data$Date)%m-% months(backtesting_period-1)))),   
                                                  window(fitted(STL_forecast_fit),start=c(year(max(case_data$Date)%m-% months(backtesting_period-1)),month(max(case_data$Date)%m-% months(backtesting_period-1)))))
                #temp<-temp[sort(names(temp))]
                accuracy_record<-rbind(accuracy_record,temp)
          }
          }
          }
        }
      }
    }
    temp_sheet<-xlsx::createSheet(accuracy_wb, sheetName=base)
    xlsx::addDataFrame(accuracy_record , sheet=temp_sheet, startColumn=1, row.names=FALSE)
  }
  xlsx::saveWorkbook(accuracy_wb, paste0("./Step_4/Step_4_",product,'_',"Accuracy_ALL_Models",suffix,".xlsx"))
}


####################################################################################
###------------------Select Parameter Combo by Ranking Accuracy------------------###
####################################################################################

rank_list<-c(1,0.999,0.975,0.974,0.95,0.949,0.925,0.924,0.9,0.899,0.85,0.849,0.8,0.799,0.75,0.749,
             0.7,0.699,0.6,0.599,0.5,0.499,0.4,0.399,0.2,0.199,0.001,0)
accuracy_rank<-data.frame()
for (product in unique(model_data_complete$Product.Category)){
  for (base in unique(model_data_complete[model_data_complete$Product.Category==product,'Base_Feature'])){
    accuracy_record <-openxlsx::read.xlsx(paste0("./Step_4/Step_4_",product,'_',"Accuracy_ALL_Models",suffix,".xlsx"),sheet= base)
    accuracy_record <- accuracy_record[is.na(accuracy_record$Accuracy)==FALSE,]
    accuracy_list<-unique(accuracy_record$Accuracy)

    for(i in 1:length(rank_list)){
      rank_temp<-data.frame()
      perct<-rank_list[i]
      rank_temp<-accuracy_record[accuracy_record$Accuracy==quantile(accuracy_list,perct,type=1),][1,]
      rownames(rank_temp)<-NULL
      accuracy_rank<-rbind(accuracy_rank,rank_temp)
    }
  }
}

####################################################################################
###--------------Backtest Result Documentation (align with Step-1,2)-------------###
####################################################################################
for (product in unique(accuracy_rank$Product.Category)){
  backtest_wb <- xlsx::createWorkbook('xlsx')

  for (base in unique(accuracy_rank[accuracy_rank$Product.Category==product,'Base'])){
    case_sales <- model_data_complete[model_data_complete$Product.Category==product & model_data_complete$Base_Feature==base,c('Date', 'Sales.USD')]
    bk_record<-data.frame()
    rank_list_tmp<-accuracy_rank[accuracy_rank$Product.Category==product & accuracy_rank$Base==base,]
    case_method<-case_select[case_select$Product.Category==product & case_select$Base==base,]
    
    for (i in 1:nrow(rank_list_tmp)){
      
      backtesting_period<-rank_list_tmp[i,'Backtesting_period']
      method<-rank_list_tmp[i,'Method']
      t_win<-rank_list_tmp[i,'t_window']
      s_win<-rank_list_tmp[i,'s_window']
      l<-rank_list_tmp[i,'lambda']
      
      backtest_train_data <- head(case_sales$Sales.USD, nrow(case_sales)-backtesting_period)
      backtest_test_data <- tail(case_sales$Sales.USD, backtesting_period)
      actual <- data.frame('Actual'=backtest_train_data, stringsAsFactors=FALSE)
      hist_start_date<-case_sales$Date[1];hist_end_date<-case_sales$Date[nrow(case_sales)]
      fcst_start_date<-hist_end_date %m+% months(1)
      fcst_end_date<-parse_date_time(sub('-','',forecast_to), orders = c("%Y-%m-%d"))
      backtest_train_ts <- try(ts(backtest_train_data, start = c(year(hist_start_date),month(hist_start_date)), frequency=12),silent=T)
      fit_test_ts<-try(ts(case_sales$Sales.USD, start = c(year(hist_start_date),month(hist_start_date)), frequency=12),silent=T)
      
      if(grepl('fit',method)==FALSE){
        #Use train + test method
        regressor_file<-regressor_extract(case_method)
        reg_hist<-regressor_file[regressor_file$Date >= as.Date(hist_start_date) & regressor_file$Date <= as.Date(hist_end_date %m-% months(backtesting_period)),
                                 names(regressor_file)[names(regressor_file)!='Date']]
        reg_new<-regressor_file[regressor_file$Date > as.Date(hist_end_date %m-% months(backtesting_period)) & 
                                  regressor_file$Date <= as.Date(hist_end_date %m+% months(0)),names(regressor_file)[names(regressor_file)!='Date']]
        
        fcst <- try(stlf(backtest_train_ts,
                         t.window = t_win, s.window = s_win, method='arima', robust = TRUE, h=backtesting_period, lambda = l, biasadj = TRUE, allow.multiplicative.trend=TRUE,
                         xreg=data.matrix(reg_hist),newxreg = data.matrix(reg_new))
                    , silent = TRUE,outFile = getOption("try.outFile", default = stderr()))
        if('try-error' %in% class(fcst)) {next}
        fcst<-fcst$mean
        
        if(any(is.na(fcst)==T)){
          fcst <- as.data.frame(fcst)
          colnames(fcst) <- method
          k<-which(is.na(fcst),arr.ind = T)
          fcst[k] <- colMeans(fcst,na.rm = T)
          fcst <- as.numeric(unlist(fcst))
        }
        
        #accuracy <- accuracy_mape(backtest_test_data, fcst)
        
        backtest_output_temp <- data.frame('Fcst'=append(as.vector(unlist(actual)),fcst))
        colnames(backtest_output_temp) <- c(paste(method,backtesting_period,'t',t_win,'s',s_win,'la',l,sep = "_"))

        if(nrow(bk_record)==0){
          bk_record <- backtest_output_temp
        }else{
          bk_record <- cbind(bk_record,backtest_output_temp)
        }
      }else{
        #use fitted accuracy method
        regressor_file<-regressor_extract(case_method)
        reg_hist<-regressor_file[regressor_file$Date >= as.Date(hist_start_date) & regressor_file$Date <= as.Date(hist_end_date),
                                 names(regressor_file)[names(regressor_file)!='Date']]
        #new regressor has no actual impact here, just to impute newreg in the function, coz we need fitted accuracy ONLY
        reg_new<-regressor_file[regressor_file$Date > as.Date(hist_end_date %m-% months(backtesting_period)) & 
                                  regressor_file$Date <= as.Date(hist_end_date %m+% months(0)),names(regressor_file)[names(regressor_file)!='Date']]
        if('try-error' %in% class(fit_test_ts)) {next}
        fcst <- try(stlf(fit_test_ts,
                        t.window = t_win, s.window = s_win, method='arima', robust = TRUE, h=backtesting_period, lambda = l, biasadj = TRUE, allow.multiplicative.trend=TRUE,
                        xreg=data.matrix(reg_hist),newxreg = data.matrix(reg_new))
                    , silent = TRUE,outFile = getOption("try.outFile", default = stderr()))
        if('try-error' %in% class(fcst)) {next}
        fcst<-fitted(fcst)
        
        if(any(is.na(fcst)==T)){
          fcst <- as.data.frame(fcst)
          colnames(fcst) <- method
          k<-which(is.na(fcst),arr.ind = T)
          fcst[k] <- colMeans(fcst,na.rm = T)
          fcst <- as.numeric(unlist(fcst))
        }
        
        accuracy <- accuracy_mape(backtest_test_data, tail(fcst,backtesting_period))
               
        backtest_output_temp <- data.frame('Fcst'=fcst)
        colnames(backtest_output_temp) <- c(paste(method,backtesting_period,'t',t_win,'s',s_win,'la',l,sep = "_"))
        
        if(nrow(bk_record)==0){
          bk_record <- backtest_output_temp
        }else{
          bk_record <- cbind(bk_record,backtest_output_temp)
        }
      }
    }
    
    bk_record <- cbind(case_sales, bk_record)
    names(bk_record)[2] <- 'Actual'
    
    temp_sheet<-xlsx::createSheet(backtest_wb, sheetName=base)
    xlsx::addDataFrame(bk_record , sheet=temp_sheet, startColumn=1, row.names=FALSE)
    
  }
  temp_sheet<-xlsx::createSheet(backtest_wb, sheetName='Accuracy')
  xlsx::addDataFrame(accuracy_rank , sheet=temp_sheet, startColumn=1, row.names=FALSE)
  xlsx::saveWorkbook(backtest_wb, paste0("./Step_4/Step_4_",product,'_',"Backtest_record",suffix,".xlsx"))
}


####################################################################################
###--------------Forecast Result Documentation (align with Step-1,2)-------------###
####################################################################################

for (product in unique(accuracy_rank$Product.Category)){
  forecast_wb <- xlsx::createWorkbook('xlsx')
  
  for (base in unique(accuracy_rank[accuracy_rank$Product.Category==product,'Base'])){
    case_sales <- model_data_complete[model_data_complete$Product.Category==product & model_data_complete$Base_Feature==base,c('Date', 'Sales.USD')]
    case_method<-case_select[case_select$Product.Category==product & case_select$Base==base,]
    fcst_output<-data.frame()
    rank_list_tmp<-accuracy_rank[accuracy_rank$Product.Category==product & accuracy_rank$Base==base,]
    
    for (i in 1:nrow(rank_list_tmp)){
      
      backtesting_period<-rank_list_tmp[i,'Backtesting_period']
      method<-rank_list_tmp[i,'Method']
      t_win<-rank_list_tmp[i,'t_window']
      s_win<-rank_list_tmp[i,'s_window']
      l<-rank_list_tmp[i,'lambda']

      forecast_train_data <- case_sales$Sales.USD
      actual_temp <- data.frame(forecast_train_data, stringsAsFactors=FALSE)
      
      hist_start_date<-case_sales$Date[1];hist_end_date<-case_sales$Date[nrow(case_sales)]
      fcst_start_date<-hist_end_date %m+% months(1)
      fcst_end_date<-parse_date_time(sub('-','',forecast_to), orders = c("%Y-%m-%d"))
      ts_index <- seq(as.Date(case_sales$Date[1]), as.Date(forecast_to), "month")
      fcst_length<-as.period( hist_end_date %--% fcst_end_date)$year * 12 + as.period( hist_end_date %--% fcst_end_date)$month
      forecast_train_ts <- ts(forecast_train_data, start=c(year(hist_start_date),month(hist_start_date)), frequency=12)
      
      regressor_file<-regressor_extract(case_method)
      reg_hist<-regressor_file[regressor_file$Date >= as.Date(hist_start_date) & regressor_file$Date <= as.Date(hist_end_date %m-% months(0)),
                               names(regressor_file)[names(regressor_file)!='Date']]
      reg_new<-regressor_file[regressor_file$Date > as.Date(hist_end_date %m-% months(0)) & 
                                regressor_file$Date <= as.Date(hist_end_date %m+% months(fcst_length))
                              ,names(regressor_file)[names(regressor_file)!='Date']]
      
      fcst <- try(stlf(forecast_train_ts,
                       t.window = t_win, s.window = s_win, method='arima', robust = TRUE, h=fcst_length, lambda = l, biasadj = TRUE, allow.multiplicative.trend=TRUE,
                       xreg=data.matrix(reg_hist),newxreg = data.matrix(reg_new))
                  , silent = TRUE,outFile = getOption("try.outFile", default = stderr()))
      
      if('try-error' %in% class(fcst)) {next}
      fcst<-fcst$mean
      fcst <- as.data.frame(fcst)
      colnames(fcst) <- paste(method,backtesting_period,'t',t_win,'s',s_win,'la',l,sep = "_")
      
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
      fcst_output <- cbind(ts_index,rbind(actual_temp, fcst_output))
      
      colnames(fcst_output)[1] <- 'Date'
      
      temp_sheet<-xlsx::createSheet(forecast_wb, sheetName=base)
      xlsx::addDataFrame(fcst_output , sheet=temp_sheet, startColumn=1, row.names=FALSE)
  }
  xlsx::saveWorkbook(forecast_wb, paste0("./Step_4/Step_4_",product,'_',"Forecast_record",suffix,".xlsx"))
}
      

####################################################################################
###-------------------Create Metric Table (align with Step-1,2)------------------###
####################################################################################

metric_table <- data.frame()
for (product in unique(model_data_complete$Product.Category)){
  for (base in unique(model_data_complete[model_data_complete$Product.Category==product,'Base_Feature'])){
    sheet=base
    
    ###=====1.Metric: Accuracy=======###
    accuracy_record <-xlsx::read.xlsx(paste0("./Step_4/Step_4_",product,'_',"Backtest_record",suffix,".xlsx"),sheetName = 'Accuracy')
    accuracy_table<-accuracy_record
    accuracy_table <- accuracy_table[paste0(accuracy_table$Product.Category,'_',accuracy_table$Base)==paste0(product,"_",base),]
    
    fcst_df <-xlsx::read.xlsx(paste0("./Step_4/Step_4_",product,'_',"Forecast_record",suffix,".xlsx"),sheetName = sheet)
    
    fcst_growth_rate_table <- data.frame()
    fcst_cv_table <- data.frame()
    fcst_sumValue_table <- data.frame()
    fcst_slope_table<-data.frame()
    
    for (col in colnames(fcst_df)) {
      if(col=='Date'){next}
      
      l<-as.numeric(unlist(strsplit(col,'_'))[length(unlist(strsplit(col,'_')))])
      s_win<-as.numeric(unlist(strsplit(col,'_'))[length(unlist(strsplit(col,'_')))-2])
      t_win<-as.numeric(unlist(strsplit(col,'_'))[length(unlist(strsplit(col,'_')))-4])
      method<-ifelse(grepl('fit',col),'forecast_enhance_stl_regressor_fit','forecast_enhance_stl_regressor')
      
      fcst_ts <-as.vector(unlist(fcst_df[fcst_df$Date>=FCST_START_DATE & fcst_df$Date<=FCST_END_DATE,col]))
      if(NA %in% fcst_ts){next}
      hist_ts <-as.vector(unlist(fcst_df[fcst_df$Date>=HIST_START_DATE & fcst_df$Date<=HIST_END_DATE,col]))
      hist_ts_N_1<-as.vector(unlist(fcst_df[fcst_df$Date>=as.Date(HIST_START_DATE)%m-%months(12) & fcst_df$Date<=as.Date(HIST_END_DATE)%m-%months(12),col]))
      
      ###====2.Metric: Growth rate====###
      growth_rate <- (sum(fcst_ts)/sum(hist_ts))^(1/(as.period(HIST_END_DATE %--% FCST_END_DATE)$year))-1
      hist_growth_rate<-ifelse(length(hist_ts_N_1)<length(hist_ts),NA,(sum(hist_ts)/sum(hist_ts_N_1))-1)
      fcst_growth_rate_table<-rbind(fcst_growth_rate_table,data.frame('Product.Category'=product,'Base'=base,'Method'=method,'t_window'=t_win, 's_window'=s_win,
                                                                      'lambda'=l,'Hist_Growth_rate'=hist_growth_rate,'Growth_rate'=growth_rate))
      
      
      ####====3.Metric: Variability====####
      fcst_cv <- sd(fcst_ts)/mean(fcst_ts)
      hist_cv <- sd(hist_ts)/mean(hist_ts)
      hist_cv_N_1 <- ifelse(length(hist_ts_N_1)<length(hist_ts),NA,sd(hist_ts_N_1)/mean(hist_ts_N_1))
      
      fcst_cv_table<-rbind(fcst_cv_table,data.frame('Product.Category'=product,'Base'=base,'Method'=method,'t_window'=t_win, 's_window'=s_win,
                                                    'lambda'=l,'Hist_cv_N_1'=hist_cv_N_1,'Hist_cv'=hist_cv,'Fcst_cv'=fcst_cv))
      
      
      ####====4.Metric: Sum Value====#### 
      fcst_sumValue <- sum(fcst_ts)
      hist_sumValue <- sum(hist_ts)
      
      fcst_sumValue_table<-rbind(fcst_sumValue_table,data.frame('Product.Category'=product,'Base'=base,'Method'=method,'t_window'=t_win, 's_window'=s_win,
                                                                'lambda'=l,'Hist_sumValue'=hist_sumValue,'Fcst_sumValue'=fcst_sumValue))
      
      ####====5.Metric: Slope====####
      fcst_slope<-as.vector(lm(fcst_ts~c(1:length(fcst_ts)))$coefficients[2])
      hist_slope<-as.vector(lm(hist_ts~c(1:length(hist_ts)))$coefficients[2])
      hist_slope_N_1<-as.vector(lm(hist_ts_N_1~c(1:length(hist_ts_N_1)))$coefficients[2])
      
      fcst_slope_table<-rbind(fcst_slope_table,data.frame('Product.Category'=product,'Base'=base,'Method'=method,'t_window'=t_win, 's_window'=s_win,
                                                          'lambda'=l,'Hist_Slope_N_1'=hist_slope_N_1,'Hist_Slope'=hist_slope,'Fcst_Slope'=fcst_slope))
      
      
    }
    
    metric_tmp <- Reduce(function(x,y,...) merge(x,y,all=T,...), list(accuracy_table,fcst_growth_rate_table,fcst_cv_table,fcst_sumValue_table,fcst_slope_table))
    metric_table <- rbind(metric_table,metric_tmp)
    
  }
}
metric_table<-metric_table[!duplicated(metric_table),]


#########################################################################################
###-------------------Filter Out Best Accuracy (align with Step-1,2)------------------###
#########################################################################################
accuracy_threshold<-0.8
best_method <- data.frame()
for (product in unique(model_data_complete$Product.Category)){
  for (base in unique(model_data_complete[model_data_complete$Product.Category==product,'Base_Feature'])){
    metric_table_tmp <- metric_table[metric_table$Product.Category==product &metric_table$Base==base,]

    best_method_record <- metric_table_tmp[which.max(metric_table_tmp$Accuracy),][1,]
    best_method_record <- data.frame('Product.Category'=product,'Base'=base,'Method'=best_method_record$Method,
                                     'Accuracy'=best_method_record$Accuracy,'Backtesting_period'=best_method_record$Backtesting_period,
                                     'Parameter_1'=best_method_record$t_window, 'Parameter_2'=best_method_record$s_window, 'Parameter_3'=best_method_record$lambda)
    best_method<-rbind(best_method,best_method_record)
    
  }
}

best_method$Method<-sub('_fit','',best_method$Method)
best_method<-merge(best_method,case_select[,c("Product.Category","Base","Regressor_Method")])

write.csv(best_method,paste0('./Step_4/Step_4_master_file_backup',suffix,'.csv'),row.names = F)

if(nrow(best_method[best_method$Accuracy>=accuracy_threshold,])>=1){
  write.csv(best_method[best_method$Accuracy>=accuracy_threshold,],paste0('./Step_4/Step_4_Best_Method',suffix,'.csv'),row.names = F)
}

if(nrow(best_method[best_method$Accuracy<accuracy_threshold,])>=1){
  write.csv(best_method[best_method$Accuracy<accuracy_threshold,],paste0('./Step_4/Segments_after_4Steps',suffix,'.csv'),row.names = F)
}


##########################################################################################
###-------------------Generate Plots & Save out (align with Step-1,2)------------------###
##########################################################################################


dir.create(paste0("./Step_4/Step_4_pic_output",suffix))
for (product in unique(model_data_complete$Product.Category)){
  trend_wb <- xlsx::createWorkbook('xlsx')
  for (base in unique(model_data_complete[model_data_complete$Product.Category==product,'Base_Feature'])){
    sheet=base
    
    fcst_df <-xlsx::read.xlsx(paste0("./Step_4/Step_4_",product,'_',"Forecast_record",suffix,".xlsx"),sheetName = sheet)
    bk_df <-xlsx::read.xlsx(paste0("./Step_4/Step_4_",product,'_',"Backtest_record",suffix,".xlsx"),sheetName = sheet)
    fcst_df$Date <- as.Date(fcst_df$Date)
    
    hist_start_date<-fcst_df$Date[1];hist_end_date<-fcst_df$Date[nrow(bk_df)]
    fcst_start_date<-hist_end_date %m+% months(1)
    fcst_end_date<-parse_date_time(sub('-','',forecast_to), orders = c("%Y-%m-%d"))
    ts_index <- seq(as.Date(fcst_df$Date[1]), as.Date(forecast_to), "month")
    actual_period <- as.period(hist_start_date %--% max(model_data_complete$Date))$year * 12 + as.period(hist_start_date %--% max(model_data_complete$Date))$month +1
    
    
    #create a flag for different line colors
    fcst_df[1:actual_period,"Flag"] <- "Actual"
    fcst_df[(actual_period+1):(nrow(fcst_df)),"Flag"] <- "Forecast"
    
    ##Create the worksheet
    trend_sheet <- xlsx::createSheet(trend_wb, sheet)
    
   
      #require(XLConnect)
      bk_df <-xlsx::read.xlsx(paste0("./Step_4/Step_4_",product,'_',"Backtest_record",suffix,".xlsx"),sheetName = sheet)
      bk_df$Date<-as.Date(bk_df$Date)
      
      for(name in names(bk_df)){
        if(name %in% c('Date','Flag')){next}
        bk_df[,name]<-as.numeric(as.character(bk_df[,name]))
      }
      
      bk_df$Flag<-"Backtest"

      rank_list_tmp<-accuracy_rank[accuracy_rank$Product.Category==product & accuracy_rank$Base==base,]
      
      for (i in 1:nrow(rank_list_tmp)){
        accuracy_value <- round(rank_list_tmp[i, "Accuracy"],4)
        backtesting_period<-rank_list_tmp[i,'Backtesting_period']
        method<-rank_list_tmp[i,'Method']
        t_win<-rank_list_tmp[i,'t_window']
        s_win<-rank_list_tmp[i,'s_window']
        l<-rank_list_tmp[i,'lambda']
        
        sub_metric_table<-metric_table[metric_table$Product.Category==product & metric_table$Base==base,]
        sub_best_method<-best_method[best_method$Product.Category==product & best_method$Base==base,]
        bk_record <- tail(bk_df, backtesting_period)
        col<-paste(method,backtesting_period,'t',t_win,'s',s_win,'la',l,sep = "_")
        #Save the graph
        p<-try(ggsave(paste("./Step_4/Step_4_pic_output", suffix,"/", sheet, col, '.png', sep=''),
                      ggplot()+
                        geom_line(aes_string(x="Date", y=col, color="Flag"),fcst_df)+
                        geom_line(aes_string(x="Date", y=col), bk_record)+
                        labs(title=paste("Forecast_method:", method), subtitle=paste("Accuracy:", accuracy_value,"\n","Backtesting_period:",backtesting_period,"\n",
                                                                                     "t_win:",t_win," s_win:",s_win, " lambda:",l))+
                        theme(plot.title = element_text(color = "brown3", size = 20, face = "bold", hjust = 0.5),
                              plot.subtitle = element_text(color = "blue", size = 18, hjust = 0.5)),
                      width = 12, height = 5, dpi = 100, device = 'png'),silent = TRUE)
        if('try-error' %in% class(p)) {next}
        
        pic_file <- paste("./Step_4/Step_4_pic_output", suffix, "/",sheet, col, '.png',sep='')
        addPicture(pic_file, sheet=trend_sheet, scale=0.4, startRow = 2+9.5*(which(colnames(fcst_df)==col)-2), 
                   startColumn = ncol(sub_metric_table)+4)
        
      
    }
    
      xlsx::addDataFrame(sub_metric_table,sheet=trend_sheet, startRow = 1,startColumn = 1)
      xlsx::addDataFrame(sub_best_method,sheet=trend_sheet,startRow = nrow(rank_list_tmp) + 5,startColumn = 1)
    print(paste0(product," ",sheet,':done'))
  }
  xlsx::saveWorkbook(trend_wb,paste0("./Step_4/Step_4_",product,'_', "Forecast_Summary_Plot",suffix,".xlsx"))
}



##############################################################################
###-------------------Collect best methods from all steps------------------###
##############################################################################
Step_list<-4:1
best_method_all<-data.frame()
for (i in Step_list){
  tmp_step<-read.csv(paste0('./Step_',i,'/Step_',i,'_Best_Method',suffix,'.csv'))
  col_add<-setdiff(colnames(best_method_all),colnames(tmp_step))
  for (col in col_add){
    tmp_step[,col]<-''
  }
  tmp_step<-tmp_step[!paste0(tmp_step$Product.Category,tmp_step$Base) %in% paste0(best_method_all$Product.Category,best_method_all$Base),]
  best_method_all<-rbind(best_method_all,tmp_step)
}


#add comment if needed for this case
best_method_all$Comment<-''

write.csv(best_method_all,paste0("L2_forecast_method",suffix,"_model_training.csv"))

major_master_file<-read.csv("C:/ForMe/CPLM/CONS_test/L2_forecast_method.csv")
master_file_keep<-major_master_file[!paste0(major_master_file$Product.Category,major_master_file$Base) %in% paste0(best_method_all$Product.Category, best_method_all$Base),]
names(best_method_all)[names(best_method_all)=='Backtesting_period']<-"Backtesting.period"
master_file_new<-rbind(master_file_keep,best_method_all)
write.csv(master_file_new,paste0("L2_forecast_method_",year(Sys.time()),"_",month(Sys.time()),"_Refresh",".csv"))


#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!#
# Double confirm before writing into major master file#
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!#


write.csv(master_file_new,"C:/ForMe/CPLM/CONS_test/L2_forecast_method.csv")















