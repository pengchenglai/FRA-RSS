# getwd() #check current working directory
# setwd() #reset wd if not ok. Remember to replace'\ by'/'.

# # install packages in case you do not have
# install.packages(c("tidyverse","data.table",'readxl','foreach',"scales"))

# 1.Loading packages====
pack <- list("tidyverse","data.table",'readxl','foreach',"scales",'writexl')
lapply(pack,require,character.only=T)

# 2. read global dataset ====
# Specify the path to your .xlsx file
xlsx_file_path <- "C:/Users/Laip/OneDrive - Food and Agriculture Organization/Desktop/Project/RSS/Simplified_calculation/Script/Summary_tables_rss/Data/World ESTIMATES Subreg trying_6Dec21 JG_ak_el_pv_V1.xlsx"
# Read the first sheet with original column names
data <- read_excel(xlsx_file_path, sheet = 1, col_names = TRUE)
# View(data) #see the orignial global dataset of the sampling plot

# 3.load overall population of the collective ====
# (combo of 3 er,str and country,) for each country
csv_path <- "C:/Users/Laip/OneDrive - Food and Agriculture Organization/Desktop/Project/RSS/Simplified_calculation/Script/Summary_tables_rss/Data/stratumbyGEZbycountry_total.csv"
countrypop <- read.csv(csv_path,header= T)
countrypop <- countrypop %>%
  group_by(countryname, stratum, SubRegion,gGEZS) %>%
  summarize(n=sum(n))
# assign the collective code(Sreg_EGZ_str)
stratum<- unique(countrypop$stratum)#get unique values of stratum
stratum <- stratum[!is.na(stratum)] # remove the NA
# create a numeric mark for each stratum
str_class<- data.table(str=c('1','4','3','2'),
                       stratum=stratum)
SubRegion<- unique(countrypop$SubRegion)#get unique values of SubRegion
# create a abbreviation for each subregion
subregion<- data.table(Subreg=c('WAS','EUR','NAF','OCE','EAF',
                                'CAB','SAM','SSA','CAM','WAF',
                                'NAM','EAS'),
                       SubRegion=SubRegion)
# Left join data frames based on the common column 
countrypop <- left_join(countrypop ,str_class, by = "stratum")
countrypop <- left_join(countrypop ,subregion, by = "SubRegion")
# finally create the new column for the collective 
countrypop$Sreg_GEZ_str <- paste0(countrypop$Subreg,
                                 countrypop$gGEZS,
                                 countrypop$str
                                 )
# View(countrypop) # see the global population of the plot


# detect the missing countries
a<- getwd()
country=sort(unique(data$Country_n))
country_done <- list.files(path=paste0(a,"/Table"),pattern = 'tab1.xlsx',
                           recursive = T) # path contains the file
country_done <- str_split(country_done,"_")
country_done <- lapply(country_done, function(x) x[1])
country_done <- str_split(country_done,"/")
country_done <- lapply(country_done, function(x) x[2])
country_done <- unlist(country_done)
country_to_be_done <- country[!(country %in% country_done)]

# For one single country each time ====
foreach( i= country_to_be_done) %do% {
countryglos <- data[data$Country_n== i,] # extract data from global sampling plot dataset
countrypop2 <- countrypop[countrypop$countryname== i,] # extract info of total plot for each country
# Re-order the countrypop2
# Extract numeric values from the column
numeric_values <- as.numeric(gsub("[^0-9]", "", countrypop2$Sreg_GEZ_str))
# Order the countrypop2 based on the numeric values
countrypop2 <- countrypop2[order(numeric_values), ]

# 4.First table: Area from centroid ====
# count appearance of the collective for one country from global sampling plot
table1_sample_weight <- countryglos %>%
  group_by(Sreg_GEZ_str) %>%
  summarise('Nsamp' = n())

# join the info of global sampling plot info and total pop info at country level
table1_sample_weight <- countrypop2 %>% 
  left_join(table1_sample_weight, by = "Sreg_GEZ_str") %>%
  select(Sreg_GEZ_str,Nsamp, n) %>%
  rename("Npop"= "n")
# remove row which contains NA value if this country does not have this collective of sampling plot
table1_sample_weight <- na.omit(table1_sample_weight)

# calculate other variables (table a) ====
table1_sample_weight$km2 <- table1_sample_weight$Npop*0.3962
table1_sample_weight$weight_km2 <- table1_sample_weight$km2/table1_sample_weight$Nsamp

# View(table1_sample_weight)

# table b ====
# account the frequency for each land use among one set of the collective
country_landtype <- countryglos %>%
  group_by(Sreg_GEZ_str, LU18CEN) %>%
  summarise(count = n())%>%
  pivot_wider(names_from = LU18CEN, values_from = count, values_fill = 0) # pivot version and auto-fill zero for corresponding combo without value 
# Reorder the rest of the columns alphabetically
col_order <- order(colnames(country_landtype)[-1])
col_order <- col_order+1
country_landtype <- country_landtype[, c(1,col_order )]
# check if the country has the land type, if not assign them as 0
if (!"Forest" %in% names(country_landtype)) {
  # If not, create the column and fill with zeros
  country_landtype<- country_landtype %>%
    mutate(Forest = 0)%>%
  select(1, Forest, everything())
}
if (!"Other Land" %in% names(country_landtype)) {
  # If not, create the column and fill with zeros
  country_landtype<- country_landtype %>%
    mutate("Other Land" = 0)%>%
    select(1, Forest,"Other Land",everything())
}

if (!"Other Wooded Land" %in% names(country_landtype)) {
  # If not, create the column and fill with zeros
  country_landtype<- country_landtype %>%
    mutate("Other Wooded Land" = 0)%>%
    select(1, Forest,"Other Land","Other Wooded Land",everything())
}
if (!"Water" %in% names(country_landtype)) {
  # If not, create the column and fill with zeros
  country_landtype<- country_landtype %>%
    mutate("Water" = 0)%>%
    select(1, Forest,"Other Land","Other Wooded Land",Water,everything())
}
# Add a new column for the sum of each row
country_landtype$Grand_Total= rowSums(country_landtype[,-1])
# Add a new row for the sum of each column
column_sums= colSums(country_landtype[,-1])
a <- data.table(t(column_sums))
b <- data.table(Sreg_GEZ_str='Total')
Total <- cbind(b,a)
country_landtype <- rbind(country_landtype, Total)
# View(country_landtype)

# calculate the proportion for each land type (table c) ====
proportions <- foreach(col = names(country_landtype)[2:5], .combine = cbind) %do% {
  (country_landtype[,col] / country_landtype[,6])
}    

## table d.1 calculate the area for each land use type
table1_land_type_area <- foreach(col2 = names(proportions[, 1:4]), .combine = cbind) %do% {
  (proportions[(1:nrow(proportions) - 1),col2] * table1_sample_weight[,"km2"]) # exclude the last row the total
}   
# rename the column names
colnames(table1_land_type_area) <-colnames(proportions)
colnames(table1_land_type_area) <- paste0("area_", colnames(table1_land_type_area), sep = "")
table1_land_type_area <- cbind(table1_sample_weight[,c('Sreg_GEZ_str' ,'Npop','km2')],
                               table1_land_type_area)

# table d.2 calculate variance
var <- cbind(country_landtype[1:nrow(country_landtype) - 1,'Grand_Total'],
             proportions[(1:nrow(proportions) - 1),])
var <- cbind(var,table1_sample_weight[,'km2'])

# table d calculate the variance 
variance <- foreach(col3 = names(var[, 2:5]), .combine = cbind) %do% {
  ( ifelse(var$Grand_Total>1,var[,length(var)]^2 * var[,col3]
           *(1-var[,col3])/(var$Grand_Total-1),0))
}   

# rename the column names
colnames(variance) <-colnames(proportions)
colnames(variance) <- paste0("V_", colnames(variance), sep = "")

# final table d====
final <- cbind(table1_land_type_area,variance)
final$GEZ <-as.numeric(str_extract(final$Sreg_GEZ_str, '\\d{2}'))
final$Str_n <-as.numeric(str_extract(final$Sreg_GEZ_str, '\\d$')) 
final$Subreg <-str_extract(final$Sreg_GEZ_str, "[A-Z]+")
# colnames(final)

#  Summary at subregion level table e ====
sum_subreg <- final %>%
  group_by(Subreg) %>%
  summarise(across(names(final[,4:11] ), sum, na.rm = TRUE))
# Add a new row for the sum of each column
column_sums2= colSums(sum_subreg[,-1])
a2 <- data.table(t(column_sums2))
b2 <- data.table(Subreg='Total')
Total2 <- cbind(b2,a2)
sum_subreg <- rbind(sum_subreg, Total2)
# names(sum_subreg)
# std error
std_err <- foreach(col4 = names(sum_subreg[, 6:9]), .combine = cbind) %do% {
  sqrt(sum_subreg[,col4])
}
#  change the column names by adding std_
colnames(std_err) <- str_replace_all(colnames(std_err), "V_", "std_")
sub=cbind(sum_subreg,std_err)
colnames(sub)
# summary for Coef V
coef_V <- foreach(col5 = 2:5, .combine = cbind) %do% {
  sub[,col5+8]/sub[,col5]
}
# change the column names
new_column_names <- c("Coef_V_Forest", "Coef_V_Other Land", "Coef_V_Other Wooded Land", "Coef_V_Water")
colnames(coef_V) <- new_column_names
sub= cbind(sub,coef_V)

#  Summary at ecozone level table f ====
sum_GEZ <- final %>%
  group_by(GEZ) %>%
  summarise(across(names(final[,4:11] ), sum, na.rm = TRUE))
# Add a new row for the sum of each column
column_sums3= colSums(sum_GEZ[,-1])
a3 <- data.table(t(column_sums3))
b3 <- data.table(GEZ='Total')
Total3<- cbind(b3,a3)
sum_GEZ <- rbind(sum_GEZ, Total3)
# std error
std_err <- foreach(col4 = names(sum_GEZ[, 6:9]), .combine = cbind) %do% {
  sqrt(sum_GEZ[,col4])
}
#  change the column names by adding std_
colnames(std_err) <- str_replace_all(colnames(std_err), "V_", "std_")
GEZ=cbind(sum_GEZ,std_err)
# colnames(GEZ)
# summary for Coef V
coef_V <- foreach(col5 = 2:5, .combine = cbind) %do% {
  GEZ[,col5+8]/GEZ[,col5]
}
# change the column names
new_column_names <- c("Coef_V_Forest", "Coef_V_Other Land", "Coef_V_Other Wooded Land", "Coef_V_Water")
colnames(coef_V) <- new_column_names
GEZ= cbind(GEZ,coef_V)

#  Summary at stratum level table g ====
sum_str <- final %>%
  group_by(Str_n) %>%
  summarise(across(names(final[,4:11] ), sum, na.rm = TRUE))
# Add a new row for the sum of each column
column_sums4= colSums(sum_str[,-1])
a4 <- data.table(t(column_sums4))
b4 <- data.table(Str_n='Total')
Total4<- cbind(b4,a4)
sum_str <- rbind(sum_str, Total4)
# std error
std_err <- foreach(col4 = names(sum_str[, 6:9]), .combine = cbind) %do% {
  sqrt(sum_str[,col4])
}
#  change the column names by adding std_
colnames(std_err) <- str_replace_all(colnames(std_err), "V_", "std_")
STR=cbind(sum_str,std_err)
colnames(STR)
# summary for Coef V
coef_V <- foreach(col5 = 2:5, .combine = cbind) %do% {
  STR[,col5+8]/STR[,col5]
}
# change the column names
new_column_names <- c("Coef_V_Forest", "Coef_V_Other Land", "Coef_V_Other Wooded Land", "Coef_V_Water")
colnames(coef_V) <- new_column_names
STR= cbind(STR,coef_V)

#  table h final summary table for each land use type / the subregion ====
table_summary_1<- data.table()
table_summary_1$subreg <- sub$Subreg
table_summary_1$Forest <- sub$area_Forest/10000
table_summary_1[, 'Forest_±'] <- 1.96*sub$Coef_V_Forest
table_summary_1$OWL <- sub$`area_Other Wooded Land`/10000
table_summary_1[, 'OWL_±'] <- 1.96*sub$`Coef_V_Other Wooded Land`
table_summary_1$Other_Land <- sub$`area_Other Land`/10000
table_summary_1[, 'Other Land_±'] <- 1.96*sub$`Coef_V_Other Land`
table_summary_1$Water <- sub$area_Water/10000
table_summary_1[, 'Water_±'] <- 1.96*sub$Coef_V_Water


# table i final summary table for each land use type / the global ecological zone ====
table_summary_2<- data.table()
table_summary_2$GEZ <- GEZ$GEZ
table_summary_2$Forest <- GEZ$area_Forest/10000
table_summary_2[, 'Forest_±'] <- 1.96*GEZ$Coef_V_Forest
table_summary_2$OWL <- GEZ$`area_Other Wooded Land`/10000
table_summary_2[, 'OWL_±'] <- 1.96*GEZ$`Coef_V_Other Wooded Land`
table_summary_2$Other_Land <- GEZ$`area_Other Land`/10000
table_summary_2[, 'Other Land_±'] <- 1.96*GEZ$`Coef_V_Other Land`
table_summary_2$Water <- GEZ$area_Water/10000
table_summary_2[, 'Water_±'] <- 1.96*GEZ$Coef_V_Water
table_summary_2$GEZ <- as.numeric(table_summary_2$GEZ)
table_summary_2[is.na(GEZ), GEZ := 99]
code_gez <- read.csv2('Data/GEZ.csv',header=T, sep=",")
table_summary_2 <- left_join(table_summary_2,code_gez, by= join_by(GEZ==GEZ))
# Rename column names
colnames(table_summary_2)[1] <- 'Mha'
colnames(table_summary_2)[length(table_summary_2)] <- 'GEZ'
# Move the last column to the second position)
table_summary_2=table_summary_2[,c(1,10,2:9)]


# table j final summary table for each land use type / the stratum====
table_summary_3<- data.table()
table_summary_3$STR <- STR$Str_n
table_summary_3$Forest <- STR$area_Forest/10000
table_summary_3[, 'Forest_±'] <- 1.96*STR$Coef_V_Forest
table_summary_3$OWL <- STR$`area_Other Wooded Land`/10000
table_summary_3[, 'OWL_±'] <- 1.96*STR$`Coef_V_Other Wooded Land`
table_summary_3$Other_Land <- STR$`area_Other Land`/10000
table_summary_3[, 'Other Land_±'] <- 1.96*STR$`Coef_V_Other Land`
table_summary_3$Water <- STR$area_Water/10000
table_summary_3[, 'Water_±'] <- 1.96*STR$Coef_V_Water
table_summary_3$STR <- as.numeric(table_summary_3$STR)
table_summary_3[is.na(STR), STR := 99]
# Add a new row for the total
new_row <- data.frame(str = 99, stratum = "Total")
code_STR <- rbind(str_class, new_row)
code_STR$str= as.numeric(code_STR$str)
table_summary_3 <- left_join(table_summary_3,code_STR, by= join_by(STR==str))
# Rename column names
colnames(table_summary_3)[1] <- 'Mha'
# Move the last column to the second position
table_summary_3=table_summary_3[,c(1,10,2:9)]

# save sub as sub1 for the calculation of efficiency hexagons
sub1 <- sub
# 5.Save the file ====
path <- getwd()
dir.create(paste0(path,"/Table"))
dir.create(paste0(path,"/Table",'/',i))
write_xlsx(list(table1_sample_weight=table1_sample_weight,
                country_landtype=country_landtype,
                proportions=proportions,
                Area_Variance=final,
                Sub_std_Er= sub,
                GEZ_std_Er= GEZ,
                Str_std_Er=STR,
                Summary_Sub= table_summary_1,
                Summary_GEZ= table_summary_2,
                Summary_STR= table_summary_3
), path=paste0(path,'/Table/',i,"/" ,i,"_Area from centroid_tab1.xlsx"))

countryglos <- data[data$Country_n==i,] # extract data from global sampling plot dataset
countrypop2 <- countrypop[countrypop$countryname==i,] # extract info of total plot for each country

# Re-order the countrypop2
# Extract numeric values from the column
numeric_values <- as.numeric(gsub("[^0-9]", "", countrypop2$Sreg_GEZ_str))
# Order the countrypop2 based on the numeric values
countrypop2 <- countrypop2[order(numeric_values), ]

# 6.Second table: Area from hex ====
# count appearance of the collective for one country from global sampling plot
table1_sample_weight <- countryglos %>%
  group_by(Sreg_GEZ_str) %>%
  summarise('Nsamp' = n())

# join the info of global sampling plot info and total pop info at country level
table1_sample_weight <- countrypop2 %>% 
  left_join(table1_sample_weight, by = "Sreg_GEZ_str") %>%
  select(Sreg_GEZ_str,Nsamp, n) %>%
  rename("Npop"= "n")
# remove row which contains NA value if this country does not have this collective of sampling plot
table1_sample_weight <- na.omit(table1_sample_weight)
# calculate other variables (table a) ====
table1_sample_weight$km2 <- table1_sample_weight$Npop*0.3962
table1_sample_weight$weight_km2 <- table1_sample_weight$km2/table1_sample_weight$Nsamp

# table b ====
# calculate the average for each land use among one set of the collective
options(digits=10)
country_landmean <- countryglos %>%
  select("Sreg_GEZ_str","FOREST%","OWL%","OL%","WATER%")%>%
  group_by(Sreg_GEZ_str)%>%
  summarise_all(mean)%>%# the default digits are 3 using summarise
  as.data.frame() # the function is used to show 'digits' as much as getOption("digits")
# Remove the percentage sign from each column name
colnames(country_landmean) <- sub("%", "", colnames(country_landmean) )

# Add a new row for the sum of each column
column_mean= colMeans(country_landmean[,-1])
a <- data.table(t(column_mean))
b <- data.table(Sreg_GEZ_str='Total')
Total <- cbind(b,a)
country_landmean <- rbind(country_landmean, Total)

# calculate the VARP for each land type (table c) ====
vap_data <- countryglos %>%
  group_by(Sreg_GEZ_str) %>% 
  select("Sreg_GEZ_str","FOREST%","OWL%","OL%","WATER%")%>%
  summarise(across("FOREST%":"WATER%", ~ var(.,na.rm=TRUE))) %>%
  as.data.frame() 
# rename the column names
colnames(vap_data)[-1] <- paste0("Vap_", colnames(vap_data)[-1], sep = "")
# Remove the percentage sign from each column name
colnames(vap_data) <- sub("%", "", colnames(vap_data) )

# Add a new row for the sum of each column
column_mean2= colMeans(vap_data[,-1])
a2 <- data.table(t(column_mean2))
b2 <- data.table(Sreg_GEZ_str='Total')
Total2 <- cbind(b2,a2)
vap_data <- rbind(vap_data, Total2)

# first summary table b-c====
land_mean <- left_join(country_landmean,vap_data, by= "Sreg_GEZ_str" )

# Area and variance d ====
area_var <- data.table()
area_var$Sreg_GEZ_str <- table1_sample_weight$Sreg_GEZ_str
area_var$Subreg <-str_extract(table1_sample_weight$Sreg_GEZ_str, "[A-Z]+")
area_var$GEZ <-as.numeric(str_extract(table1_sample_weight$Sreg_GEZ_str, '\\d{2}'))
area_var$Str_n <-as.numeric(str_extract(table1_sample_weight$Sreg_GEZ_str, '\\d$'))
area_var$area <- table1_sample_weight$km2
area_var$Nsamp <- table1_sample_weight$Nsamp
## calculate the area for each land use type d1====
area <- foreach(col = names(land_mean)[2:5], .combine = cbind) %do% {
  (land_mean[1:(nrow(land_mean)-1),col] * area_var[,"area"])
}  
# rename the columns
colnames(area) <-colnames(land_mean)[2:5]
colnames(area) <- gsub("%", "",colnames(area))
colnames(area) <- paste0("Area_", colnames(area), sep = "")
#  join the table
area_var <- cbind(area_var,area)

## calculate the variance d2 ====
land_mean2 <- land_mean[1:nrow(land_mean)-1,]
land_mean2 <- area_var %>%
  select(Sreg_GEZ_str,Nsamp,area) %>% # include the area
  left_join(land_mean, by= 'Sreg_GEZ_str')
variance<- land_mean2 %>%
  mutate(across(starts_with("Vap_"), ~ifelse(Nsamp > 1, . * area^2 / (Nsamp - 1), 0)))%>%
  rename_with(~sub("^Vap_", "Var_", .), starts_with("Vap_"))
Var <- variance[ ,(ncol(variance)-3):ncol(variance)]

# summary table d =====
area_var <- cbind(area_var,Var)

#  Summary at subregion level table e ====
sum_subreg <- area_var %>%
  group_by(Subreg) %>%
  summarise(across(names(area_var[,7:ncol(area_var)] ), sum, na.rm = TRUE))
# Add a new row for the sum of each column
column_sums2= colSums(sum_subreg[,-1])
a2 <- data.table(t(column_sums2))
b2 <- data.table(Subreg='Total')
Total2 <- cbind(b2,a2)
sum_subreg <- rbind(sum_subreg, Total2)
# names(sum_subreg)
# std error
std_err <- foreach(col4 = names(sum_subreg[, 6:ncol(sum_subreg)]), .combine = cbind) %do% {
  sqrt(sum_subreg[,col4])
}
#  change the column names by adding std_
colnames(std_err) <- str_replace_all(colnames(std_err), "Var_", "std_")
sub=cbind(sum_subreg,std_err)
colnames(sub)
# summary for Coef V
coef_V <- foreach(col5 = 2:5, .combine = cbind) %do% {
  sub[,col5+8]/sub[,col5]
}
# change the column names
new_column_names <- c("Coef_V_Forest",  "Coef_V_Other Wooded Land", "Coef_V_Other Land","Coef_V_Water")
colnames(coef_V) <- new_column_names
sub= cbind(sub,coef_V)


#  Summary at ecozone level table f ====
sum_GEZ <- area_var %>%
  group_by(GEZ) %>%
  summarise(across(names(area_var[,7:ncol(area_var)] ), sum, na.rm = TRUE))
# Add a new row for the sum of each column
column_sums3= colSums(sum_GEZ[,-1])
a3 <- data.table(t(column_sums3))
b3 <- data.table(GEZ='Total')
Total3<- cbind(b3,a3)
sum_GEZ <- rbind(sum_GEZ, Total3)
# std error
std_err <- foreach(col4 = names(sum_GEZ[, 6:9]), .combine = cbind) %do% {
  sqrt(sum_GEZ[,col4])
}
#  change the column names by adding std_
colnames(std_err) <- str_replace_all(colnames(std_err), "Var_", "std_")
GEZ=cbind(sum_GEZ,std_err)
# colnames(GEZ)
# summary for Coef V
coef_V <- foreach(col5 = 2:5, .combine = cbind) %do% {
  GEZ[,col5+8]/GEZ[,col5]
}
# change the column names
new_column_names <- c("Coef_V_Forest",  "Coef_V_Other Wooded Land", "Coef_V_Other Land","Coef_V_Water")
colnames(coef_V) <- new_column_names
GEZ= cbind(GEZ,coef_V)


#  Summary at stratum level table g ====
sum_str <- area_var %>%
  group_by(Str_n) %>%
  summarise(across(names(area_var[,7:ncol(area_var)] ), sum, na.rm = TRUE))
# Add a new row for the sum of each column
column_sums4= colSums(sum_str[,-1])
a4 <- data.table(t(column_sums4))
b4 <- data.table(Str_n='Total')
Total4<- cbind(b4,a4)
sum_str <- rbind(sum_str, Total4)
# std error
std_err <- foreach(col4 = names(sum_str[, 6:9]), .combine = cbind) %do% {
  sqrt(sum_str[,col4])
}
#  change the column names by adding std_
colnames(std_err) <- str_replace_all(colnames(std_err), "Var_", "std_")
STR=cbind(sum_str,std_err)
colnames(STR)
# summary for Coef V
coef_V <- foreach(col5 = 2:5, .combine = cbind) %do% {
  STR[,col5+8]/STR[,col5]
}
# change the column names
new_column_names <- c("Coef_V_Forest",  "Coef_V_Other Wooded Land", "Coef_V_Other Land","Coef_V_Water")
colnames(coef_V) <- new_column_names
STR= cbind(STR,coef_V)

#  table h final summary table for each land use type / the subregion ====
table_summary_1<- data.table()
table_summary_1$subreg <- sub$Subreg
table_summary_1$Forest <- sub$Area_FOREST/10000
table_summary_1[, 'Forest_±'] <- ifelse(sub$Coef_V_Forest<0.5,sub$Coef_V_Forest*1.96,'NP')
table_summary_1$OWL <- sub$Area_OWL / 10000
table_summary_1[, 'OWL_±'] <- ifelse(sub$`Coef_V_Other Wooded Land`<0.5,sub$`Coef_V_Other Wooded Land`*1.96,'NP')
table_summary_1$Other_Land <- sub$Area_OL/10000
table_summary_1[, 'Other Land_±'] <- ifelse(sub$`Coef_V_Other Land`<0.5,sub$`Coef_V_Other Land`*1.96,'NP')
table_summary_1$Water <- sub$Area_WATER/10000
table_summary_1[, 'Water_±'] <- ifelse(sub$Coef_V_Water<0.5,sub$Coef_V_Water*1.96,'NP')


# table i final summary table for each land use type / the global ecological zone ====
table_summary_2<- data.table()
table_summary_2$GEZ <- GEZ$GEZ
table_summary_2$Forest <- GEZ$Area_FOREST/10000
table_summary_2[, 'Forest_±'] <- ifelse(GEZ$Coef_V_Forest<0.5,GEZ$Coef_V_Forest*1.96,'NP')
table_summary_2$OWL <- GEZ$Area_OWL / 10000
table_summary_2[, 'OWL_±'] <- ifelse(GEZ$`Coef_V_Other Wooded Land`<0.5,GEZ$`Coef_V_Other Wooded Land`*1.96,'NP')
table_summary_2$Other_Land <- GEZ$Area_OL/10000
table_summary_2[, 'Other Land_±'] <- ifelse(GEZ$`Coef_V_Other Land`<0.5,GEZ$`Coef_V_Other Land`*1.96,'NP')
table_summary_2$Water <- GEZ$Area_WATER/10000
table_summary_2[, 'Water_±'] <- ifelse(GEZ$Coef_V_Water<0.5,GEZ$Coef_V_Water*1.96,'NP')
table_summary_2$GEZ <- as.numeric(table_summary_2$GEZ)
table_summary_2[is.na(GEZ), GEZ := 99]
code_gez <- read.csv2('Data/GEZ.csv',header=T, sep=",")
table_summary_2 <- left_join(table_summary_2,code_gez, by= join_by(GEZ==GEZ))
# Rename column names
colnames(table_summary_2)[1] <- 'Mha'
colnames(table_summary_2)[length(table_summary_2)] <- 'GEZ'
# Move the last column to the second position)
table_summary_2=table_summary_2[,c(1,10,2:9)]


# table j final summary table for each land use type / the stratum====
table_summary_3<- data.table()
table_summary_3$Str_n <- STR$Str_n
table_summary_3$Forest <- STR$Area_FOREST/10000
table_summary_3[, 'Forest_±'] <- ifelse(STR$Coef_V_Forest<0.5,STR$Coef_V_Forest*1.96,'NP')
table_summary_3$OWL <- STR$Area_OWL / 10000
table_summary_3[, 'OWL_±'] <- ifelse(STR$`Coef_V_Other Wooded Land`<0.5,STR$`Coef_V_Other Wooded Land`*1.96,'NP')
table_summary_3$Other_Land <- STR$Area_OL/10000
table_summary_3[, 'Other Land_±'] <- ifelse(STR$`Coef_V_Other Land`<0.5,STR$`Coef_V_Other Land`*1.96,'NP')
table_summary_3$Water <- STR$Area_WATER/10000
table_summary_3[, 'Water_±'] <- ifelse(STR$Coef_V_Water<0.5,STR$Coef_V_Water*1.96,'NP')
table_summary_3$Str_n <- as.numeric(table_summary_3$Str_n)
table_summary_3[is.na(Str_n), Str_n := 99]
# Add a new row for the total
new_row <- data.frame(str = 99, stratum = "Total")
code_STR <- rbind(str_class, new_row)
code_STR$str= as.numeric(code_STR$str)
table_summary_3 <- left_join(table_summary_3,code_STR, by= join_by(Str_n==str))
# Rename column names
colnames(table_summary_3)[1] <- 'Mha'
# Move the last column to the second position
table_summary_3=table_summary_3[,c(1,10,2:9)]

# new table for Efficiency_hexagons
Efficiency_hexagons <- data.table()
Efficiency_hexagons$country <- i
Efficiency_hexagons$Forest_area <- sub1$V_Forest[1]/sub$Var_FOREST[1]


# 7. Save the file ====
path <- getwd()
dir.create(paste0(path,"/Table"))
dir.create(paste0(path,"/Table",'/',i))
write_xlsx(list(table1_sample_weight=table1_sample_weight,
                land_mean=land_mean,
                area_var=area_var,
                Sub_std_Er= sub,
                GEZ_std_Er= GEZ,
                Str_std_Er=STR,
                Summary_Sub= table_summary_1,
                Summary_GEZ= table_summary_2,
                Summary_STR= table_summary_3,
                Efficiency_hexagons = Efficiency_hexagons		
), path=paste0(path,'/Table/',i,"/" ,i,"_Area from hex_tab2.xlsx"))
}

