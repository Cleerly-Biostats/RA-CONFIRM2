#RA analysis
#Matching of RA cases to controls (1:3)
#2/11/26
#Author Michelle Wiest

#On Feb 9th, Maya and Brittany agreed the matching variables are:
#Age sex race BMI hypertension(or hypertension meds) diabetes dyslipidemia smoking
#Variable names are age, sexc, race_new bmi, htnc, dm, dyslipidn, smoker

#Install packages
library(haven)
library(MatchIt)
library(mice)

#data location: "G:\Shared drives\Biostats\CONFIRM2\Programs\AnalysisDatasets\MPAC2 LOCKED Programs\programs\output\adsl_stand_vars.sas7bdat"

#Read in sas data
dat=read_sas("G:/Shared drives/Biostats/CONFIRM2/Programs/AnalysisDatasets/MPAC2 LOCKED Programs/programs/output/adsl_stand_vars.sas7bdat")
df=as_factor(dat)
head(df)

#MPAC2 only
ra <- subset(dat, pop_mpac2 == 1)

#free up memory
rm(dat)
rm(df)
gc()

#Change values of RA so missing is considered "not RA"
ra$RA <- ifelse(!is.na(ra$pop_RA) & ra$pop_RA == 1, 1, 0)
#columns to keep for matching: 
keepem= c("cleerly_id", 'RA', 'age', 'sexc', 'race_new', 'bmi', 'htnc', 'dm', 'dyslipidn', 'smoker')
ra=ra[, keepem]

#We have missing covariates. So we will impute a dataset using mice and then just use that.
imputed_temp <- mice(ra, m = 1, maxit = 50, seed = 123)
rai <- complete(imputed_temp)

#free up memory
rm(ra)
rm(imputed_temp)
gc()

#match
m.out <- matchit(RA ~ age+ sexc+ race_new+ bmi+ htnc+ dm+ dyslipidn+ smoker,  
                 data = rai,           
                 method = "cem") #,       # cem because otherwise crashes (ow nearest neighbor matching (nearest)
                 #ratio = 3,                  # 1:3 matching
                #exact = ~sexc + race_new + smoker)   # must match (this help prevent R from crashing)             

summary(m.out)

matched_data <- match.data(m.out)
write.csv(matched_data, "SOMEWHERE?! ANYWHERE?!!!")
