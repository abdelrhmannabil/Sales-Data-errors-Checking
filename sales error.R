

#loading packages
library(tidyverse) #loading mother packages for all below packages
library(googledrive) #for opening all files in my drive
library(googlesheets4) #for authentication of files
library(dplyr) #for data manipulation
library(lubridate) #for working with dates

library(readxl)
setwd("M:/تحليل البيانات/statistical analysis with r/sales/employee sales report")
emp_links <- read_excel("Employees Links.xlsx")

report_me <- function(link) {
  



googledrive::drive_auth(email = "data.analysis@mashroo3k.com")
for_power_bi <- read_sheet(link, #link of sheet
                           sheet = 'For power bi', #name of sheet need to read
                           col_names = TRUE # to take column name as column name in R)
)

sales_DB <- read_sheet(link, #link of sheet
                       sheet = 'Sales DB', #name of sheet need to read
                       col_names = TRUE # to take column name as column name in R 
)

Dashboard_sales_DB <- read_sheet(link, #link of sheet
                                 sheet = 'Dashboard Sales DB', #name of sheet need to read
                                 col_names = TRUE # to take column name as column name in R
)




#renaming of columns from arabic to english for easy handling

Dashboard_sales_DB_id <- Dashboard_sales_DB %>%                     #selecting data frame of interest
                                    select(deal_id = `Deal ID`)%>%  #selecting deal id column
                                    filter(!is.na(deal_id))         #filtering NA values in deal id
for_power_bi_id <- for_power_bi %>%                       #selecting data frame of interest          
                          select(deal_id = `Deal ID`)%>%  #selecting column of interest
                          filter(!is.na(deal_id))         # filtering Na values in deal id column
sales_DB_id <- sales_DB %>%                          #selecting data frame of interest 
                    select(deal_id = `Deal Id`) %>%  #selecting column of interest
                    filter(!is.na(deal_id))          # filtering Na values in deal id column



#checking duplication in Sales DB sheet

duplicated_id_sales_db_list <- data.frame(paste("duplicated id in sales DB"                         #making data frame
                                                ,sales_DB_id[duplicated(sales_DB_id$deal_id ),]) )  #sub-setting duplicated values in Sales DB sheet


#counting total unique IDs between different sheets to check if they are the same


uid_sales_DB_count <- sales_DB_id %>%           #selecting data frame of interest
                            select(deal_id )%>% #selecting deal id column
                            summarise(n = n())  #counting deal_id

uid_for_power_bi_count <- for_power_bi_id %>%                       #selecting data frame of interest 
                                select(deal_id ) %>%                #selecting deal id column
                                filter(!is.na(deal_id)) %>%         #filtering NA values in deal id column
                                summarise(n = n_distinct(deal_id))  #counting distinct values of deal id

uid_dashboard_sales_DB_count <- Dashboard_sales_DB_id %>%                    #selecting data frame of interest
                                      select(deal_id ) %>%                   #selecting deal id column
                                      filter(!is.na(deal_id)) %>%            #filtering NA values in deal id column
                                      summarise(n = n_distinct(deal_id))     #counting distinct values of deal id
deal_id_count_list <- data.frame(                                   #making data frame
  paste("sales DB: ", uid_sales_DB_count,                           #naming output and printing value
  " for Power BI sheet: ",                                          #printing text for naming output
  uid_for_power_bi_count,                                           #printing output
  "Dashboard sales DB: " ,uid_dashboard_sales_DB_count,sep = " ")   #naming output and printing value
)


#checking lost deals in any sheet and getting its ID


unfound_deals_list <- data.frame(c(
  paste("what is found in dashboard sales DB but not in sales DB is : ", anti_join(Dashboard_sales_DB_id, sales_DB_id, by = "deal_id" )),
  
  paste("what is found in dashboard sales DB but not in for power BI is : ", anti_join(Dashboard_sales_DB_id, for_power_bi_id, by = "deal_id" )),
  
  paste("what is found in sales DB but not in for power BI is : ", anti_join(sales_DB_id, for_power_bi_id, by = "deal_id" )),
  
  paste("what is found in sales DB but not in for dashboard sales DB is : ", anti_join(sales_DB_id, Dashboard_sales_DB_id, by = "deal_id" )),
  
  paste("what is found in for power BI but not in for  sales DB is : ", anti_join(for_power_bi_id, sales_DB_id, by = "deal_id" ))
  ,
  paste("what is found in for power BI but not in for  dashboard sales DB is : ", anti_join(for_power_bi_id, Dashboard_sales_DB_id, by = "deal_id" )))
)



#subsetting small data frames for further investigation in installments

installments_sales_db <- sales_DB %>%                                                  #data frame of interest
                                select(deal_id = `Deal Id`,                            # selecting deal id
                                       price_after_discount = `Price after discount` , #selecting price after discount column
                                       frst_install = `1st Installment`,               #selecting first installment column
                                       remaining1 = `Remaining Amount` ,               #selecting remaining after first installment column
                                       sec_install = `2nd Installment`,                #selecting second installment column
                                       remaining2 = `Remaining Amount 2`,              #selecting remaining after second installment
                                       thrd_install = `3rd Installment`,               #selecting third installment column
                                       remaining3 = `Remaining Amount 3`,              #selecting remaining after third installment column
                                       fourth_install = `4th Installment`,             #selecting fourth installment column
                                       remaining4 = `Remaining Amount 4` )             #selecting remaining after fourth installment column

installments_for_power_bi <- for_power_bi %>%                                                 #data frame of interest                             
                                    select(deal_id = `Deal ID`,                               #selecting deal id column
                                           price_after_discount =  `Price After Discount`,    #selecting price after discount column and renaming it
                                           paid = Paid,                                       #selecting paid column and renaming it
                                           remaining = Remainin,                              #selecting remaining column and renaming it
                                           installment = Installment  )                       #selecting installment column and renaming it

installments_dashboard_sales_db <- Dashboard_sales_DB %>%                                           #data frame of interest
                                            select(deal_id = `Deal ID`,                             #selecting deal id column
                                                   price_after_discount =  `Price After Discount`,  #selecting price after discount column and renaming it
                                                   paid = Paid,                                     #selecting paid column and renaming it
                                                   remaining = Remainin  )                          #selecting remaining column and renaming it


##checking the total remaining between different sheets if they are equal or not


sales_db_remaining <- sum(installments_sales_db$price_after_discount, na.rm = TRUE) - sum(sum(installments_sales_db$frst_install, na.rm = TRUE) , sum(installments_sales_db$sec_install, na.rm = TRUE) ,sum(installments_sales_db$thrd_install, na.rm = TRUE) , sum(installments_sales_db$fourth_install, na.rm = TRUE)) #calculating total remaining for Sales DB sheet



dashboard_sales_db_remaining <-  sum(installments_dashboard_sales_db$price_after_discount, na.rm = TRUE) - sum(installments_dashboard_sales_db$paid, na.rm = TRUE)  #calculating total remaining for Dashboard Sales DB sheet


for_power_bi_remaining <-  sum(installments_for_power_bi$price_after_discount,  #calculating total price 
                               na.rm = TRUE) -                                  #removing NA values
                           sum(installments_for_power_bi$paid,                  #calculating total paid for Power BI sheet
                               na.rm = TRUE)                                    #removing NA values

remaining_sheets_list <- data.frame(
                                paste("Remaining", "Sales DB: ",sales_db_remaining, 
                                      "Dashboard Sales DB: ", dashboard_sales_db_remaining, 
                                      "For Power BI:", for_power_bi_remaining )
) #making data frame of total remaining in each sheet

#checking duplicated installments in for power BI sheet
duplicated_installment_for_power_bi <-  installments_for_power_bi %>%     #data frame of interest
                                              select(deal_id, installment ) %>%                                       #selecting columns of interests
                                              filter(!is.na(deal_id)) %>%                                             #filtering NA values in deal_id
                                              group_by(deal_id, installment) %>%                                      #grouping by deal_id and installments to check duplication of any installments
                                              tally() %>%                                                             #counting each installment in each deal_id
                                              filter(n > 1) %>%                                                       #filtering duplicated installments only
                                              data.frame()                                                            #convert output from tibble to data frame



#getting Deal ID of different remaining between Sales DB and For power BI sheet



r_sales_db <- installments_sales_db %>% #data frame of interest
  filter(!is.na(deal_id ) )%>%  #filtering NA values in deal_id column
  replace(is.na(.), 0) %>% #replacing Na values in numeric column with 0
  mutate(rem = price_after_discount - (frst_install+ sec_install+ thrd_install+ fourth_install)) #calculate remaining for each row in anew column

r_for_power_bi <- installments_for_power_bi  %>% #data frame of interest
  filter(!is.na(deal_id ) )  %>% #filtering NA values in deal_id column
  replace(is.na(.), 0) %>%  #replacing Na values in numeric column with 0
  group_by(deal_id) %>%  #grouping by deal_id to prepare for calculation of  remaining for each deal_id
  mutate(rem =sum(price_after_discount)   - sum(paid)) %>% #calculation of remaining for each deal_id in a new column
  filter(!duplicated(deal_id)) #filter unique deal_id due to duplication as different installments are in different rows

r_all <- r_sales_db %>% inner_join(r_for_power_bi, by = 'deal_id',suffix =c('_sales_Db','_for_power_BI') ) #joining two data frames of different sheets

r_all <- r_all %>%  #joined data frame
  filter(rem_sales_Db != rem_for_power_BI) %>%  #filtering remamining that are not equal in both sheets
  select(deal_id, rem_sales_Db, rem_for_power_BI) #selecting columns of interest



#making list of different data frames for exporting into excel
list_of_lists <- list('deal id count' = deal_id_count_list[1,],'duplicated deal id' = duplicated_id_sales_db_list[1,] ,'duplicated installments' =duplicated_installment_for_power_bi,'missed deal ids' = unfound_deals_list[1:6,], 'remaining comprison' =remaining_sheets_list[1,], 'deals with different remaining' = r_all)


#exporting different data frames to an excel on our network with the name of the employee
library(openxlsx) #loading package that will help in exporting 'openxlsx' 

setwd("M:/تحليل البيانات/statistical analysis with r/sales/employee sales report") #path of network where we will save the file
write.xlsx(list_of_lists, paste(sales_DB[2,5],".xlsx")) #naming the file of employee name and different sheets


}

emp_links <- emp_links[-1:-7,]
emp_links <- emp_links[-1:-19,]
emp_links <- emp_links[-1:-3,]
emp_links <- emp_links[-1:-21,]
map(emp_links$Link, report_me)

