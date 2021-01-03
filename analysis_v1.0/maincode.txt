# File:     DataAnalysis.R
# Project:  THESIS


# *************************Environment Cleaning*********************************
# Clear environment
rm(list = ls()) 

# Clear console
cat("\014")  # ctrl+L

# Clear mind :)

library(xlsx)
library(dplyr)
library(tidyr)
library(writexl)

options(scipen = 999)

# *************************Workbooks Reading*********************************

#mcafee_file = "data/McAfee.xlsx"

#kaspersky_file = "data/Kaspersky.xlsx"

#symantec_file = "data/Symantec.xlsx"

comparison_file = "data/comparison.xlsx"
 
# *************************McAfee Sheets*********************************

# mcafee_MainFigures = read.xlsx (mcafee_file, sheetName = "MainFigures")

# mcafee_PubliclyDisclosedAttacks = read.xlsx (mcafee_file, sheetName = "PubliclyDisclosedAttacks")

# mcafee_TopTargetedIndustries = read.xlsx (mcafee_file, sheetName = "TopTargetedIndustries")
 
# mcafee_TopAttackVectors = read.xlsx (mcafee_file, sheetName = "TopAttackVectors")
 
# *************************Kaspersky Sheets*********************************

# kaspersky_MainFigures = read.xlsx (kaspersky_file, sheetName = "MainFigures")
 
# kaspersky_BankingTargetedCountries = read.xlsx (kaspersky_file, sheetName = "BankingTargetedCountries")
 
# kaspersky_BankingMalwareFamilies = read.xlsx (kaspersky_file, sheetName = "BankingMalwareFamilies")
 
# kaspersky_RansomwareTargetedCountries = read.xlsx (kaspersky_file, sheetName = "RansomwareTargetedCountries")
 
# kaspersky_TopRansomwareFamilies = read.xlsx (kaspersky_file, sheetName = "TopRansomwareFamilies")
 
# kaspersky_OnlineTargetedCountries = read.xlsx (kaspersky_file, sheetName = "OnlineTargetedCountries")
 
# kaspersky_VulnerableAppsPercentage = read.xlsx (kaspersky_file, sheetName = "VulnerableAppsPercentage")
 
# *************************Symantec Sheets*********************************

# symantec_MainFigures = read.xlsx (symantec_file, sheetName = "MainFigures")
 
# symantec_TopMalwares = read.xlsx (symantec_file, sheetName = "TopMalwares")
 
# symantec_TotalMalwaresbyOS = read.xlsx (symantec_file, sheetName = "TotalMalwaresbyOS")
 
# symantec_TopMacMalwaresPercentage = read.xlsx (symantec_file, sheetName = "TopMacMalwaresPercentage")
 
# symantec_RansomwareCountriesPercentage = read.xlsx (symantec_file, sheetName = "RansomwareCountriesPercentage")
 
# symantec_TopRansomwareTrojans = read.xlsx (symantec_file, sheetName = "TopRansomwareTrojans")

# *************************Comparison Sheets*********************************

#comparison_MainFigures = read.xlsx (comparison_file, sheetName = "MainFigures")
 
# *************************Variables Used*********************************
 
MMR_data <- read.xlsx (comparison_file, sheetName = "MMR")

OC_data <- read.xlsx (comparison_file, sheetName = "OrgCategories")

num_iterations <- 50000

MMR_rows <- round(nrow(MMR_data), 0)

OC_rows <- round(nrow(OC_data), 0)

# *************************Monte Carlo Simulation*********************************

# Dataframe Creation to do Columns Mutation
df_main <- data.frame() 

# for loop to incorporate MMR_data
for (i in 1:MMR_rows){
  
  industry = MMR_data[i,]$Industry
  
  var_MMR = MMR_data[i,]$MMR
  
  # for loop to incorporate OC_data
  for (j in 1:OC_rows){ 
    
    # *************************Variables for Simulation*********************************
    
    # variables for employees and revenues
    oc = OC_data[j,]$OrgCategories 
    minNE = OC_data[j,]$MinNE
    maxNE = OC_data[j,]$MaxNE
    minAR = OC_data[j,]$MinAR
    maxAR = OC_data[j,]$MaxAR
    
    # variables for BCC
    minIH = OC_data[j,]$MinIH
    maxIH = OC_data[j,]$MaxIH
    
    # variables for IRC
    minIRCPH = OC_data[j,]$MinIRCPH
    maxIRCPH = OC_data[j,]$MaxIRCPH
    
    # variables for RPC
    minRPCPM = OC_data[j,]$MinRPCPM
    maxRPCPM = OC_data[j,]$MaxRPCPM
    
    # variables for NIC
    minNIC = OC_data[j,]$MinNIC 
    maxNIC = OC_data[j,]$MaxNIC 
    
    # variables for DRC
    minDRCPM = OC_data[j,]$MinDRCPM  
    maxDRCPM = OC_data[j,]$MaxDRCPM  
    
    # CPC = (AR * 4) / 100 As Per EU General Data Protection Regulation
    
    # variables for RDC
    minRDC = OC_data[j,]$MinRDC 
    maxRDC = OC_data[j,]$MaxRDC 
    
    # *************************Column Mutation for Simulation*********************************
    
    # Dataframe Mutation for Industry, MMR, OrgCategories (OC), Employees and Annual Revenue (AR)
    mutation <- expand.grid(Iterations = 1 : num_iterations) %>% 
      rowwise() %>% 
      mutate(Industry = industry) %>% 
      mutate(MMR = var_MMR) %>% 
      mutate(OrgCategories = oc) %>% 
      mutate(Employees = round(runif(1, minNE, maxNE), 0)) %>% 
      mutate(AR = round(runif(1, minAR, maxAR), 2)) 
    
    # Dataframe Mutation for Total Machines(TM), Infection Rate (IR), Infected Machine (IM)
    mutation <- mutation %>%  
      mutate(TM = round((Employees * var_MMR), 0)) %>% 
      mutate(IR = round(runif(1, 0.0, 1.0), 2)) %>% 
      mutate(IM = round((TM * IR), 0))
      
    # Dataframe Mutation for Hourly Revenue (HR), and Infected Hours (IH)
    mutation <- mutation %>%  
      mutate(HR = round((AR / (365 * 24)), 2)) %>% 
      mutate(IH = round(runif(1, minIH, maxIH), 0))
    
    # Dataframe Mutation for Business Continuity Cost (BCC)
    # BCCPM = (HR * IH) / TM, where BCCPM is Business Continuity Cost Per Machine.
    # BCC = BCCPM * IM
    mutation <- mutation %>%  
      mutate(BCCPM = round(((HR * IH) / TM), 2)) %>% 
      mutate(BCC = round((BCCPM * IM), 2))
    
    # Dataframe Mutation for Incident Response Cost (IRC)
    # IRC = IRCPH * IH, where IRCPH is Incident Response Cost Per Hour
    mutation <- mutation %>%  
      mutate(IRCPH = round(runif(1, minIRCPH, maxIRCPH), 2)) %>% 
      mutate(IRC = round((IRCPH * IH), 2))
    
    # Dataframe Mutation for Ransom Paid Cost (RPC)
    # RPC = RPCPM * IM, where RPCPM is Ransom Paid Cost Per Machine
    # RPM = Ransom Paid Machines
    # NRPM = Not Ransom Paid Machines
    mutation <- mutation %>%  
      mutate(RPCPM = round(runif(1, minRPCPM, maxRPCPM), 2)) %>% 
      mutate(RPM = round(IM * runif(1, 0.0, 1.0), 0)) %>% 
      mutate(NRPM = round((IM - RPM), 0)) %>% 
      mutate(RPC = round((RPCPM * RPM), 2))
  
    # Dataframe Mutation for New Installation Cost (NIC)
    mutation <- mutation %>%  
      mutate(NIC = round(runif(1, minNIC, maxNIC), 2)) 
    
    # Dataframe Mutation for Data Recovery Cost (DRC) 
    # DRC = DRCPM * NRPM, where DRCPM is Data Recovery Cost Per Machine
    mutation <- mutation %>%  
      mutate(DRCPM = round(runif(1, minDRCPM, maxDRCPM), 2)) %>% 
      mutate(DRC = round((DRCPM * NRPM), 2)) 
    
    # Dataframe Mutation for Compliance Penalty Cost (CPC)
    # CPC = (AR * 4) / 100 As Per EU General Data Protection Regulation
    mutation <- mutation %>%  
      mutate(CPC = round(((AR * 4) / 100), 2))
    
    # Dataframe Mutation for Reputation Damage Cost  (RDC)
    mutation <- mutation %>%  
      mutate(RDC = round(runif(1, minRDC, maxRDC), 2)) 
    
    # Dataframe Mutation for Estimated Loss (EL)
    mutation <- mutation %>%  
      mutate(EL = round((BCC + IRC + RPC + NIC + DRC + CPC + RDC), 2))
    
    # Dataframe Mutation for Annual Revenue After Estimated Loss (ARAEL)
    mutation <- mutation %>%  
      mutate(ARAEL = round((AR - EL), 2))
    
    # Dataframe Mutation for Percentage Loss In Annual Revenue (PLIAR)
    mutation <- mutation %>%  
      mutate(PLIAR = round((((EL * 100) / AR)), 2))
    
    df_main <- rbind(df_main, mutation)
    
    } # end inner for loop
  
} # end outer for loop

# *******************************Summaries********************************
# Summarize data for Employees
df_summary_Employees = df_main %>% 
  group_by(Industry, OrgCategories) %>% 
  summarize(count = n(), 
            meanVal = round(mean(Employees), 0),
            medVal = round(median(Employees), 0),
            minVal = round(min(Employees), 0),
            maxVal = round(max(Employees), 0),
            stdVal = round(sd(Employees), 0))

# Summarize data for Annual Review (AR) 
df_summary_AR = df_main %>% 
  group_by(Industry, OrgCategories) %>% 
  summarize(count = n(), 
            meanVal = round(mean(AR), 2),
            medVal = round(median(AR), 2),
            minVal = round(min(AR), 2),
            maxVal = round(max(AR), 2),
            stdVal = round(sd(AR), 2))

# Summarize data for Total Machines (TM)
df_summary_TM = df_main %>% 
  group_by(Industry, OrgCategories) %>% 
  summarize(count = n(), 
            meanVal = round(mean(TM), 0),
            medVal = round(median(TM), 0),
            minVal = round(min(TM), 0),
            maxVal = round(max(TM), 0),
            stdVal = round(sd(TM), 0))

# Summarize data for Infection Ratio (IR)
df_summary_IR = df_main %>% 
  group_by(Industry, OrgCategories) %>% 
  summarize(count = n(), 
            meanVal = round(mean(IR), 2),
            medVal = round(median(IR), 2),
            minVal = round(min(IR), 2),
            maxVal = round(max(IR), 2),
            stdVal = round(sd(IR), 2))

# Summarize data for Infected Machines (IM)
df_summary_IM = df_main %>% 
  group_by(Industry, OrgCategories) %>% 
  summarize(count = n(), 
            meanVal = round(mean(IM), 0),
            medVal = round(median(IM), 0),
            minVal = round(min(IM), 0),
            maxVal = round(max(IM), 0),
            stdVal = round(sd(IM), 0))

# Summarize data for Hourly Revenue (HR)
df_summary_HR = df_main %>% 
  group_by(Industry, OrgCategories) %>% 
  summarize(count = n(), 
            meanVal = round(mean(HR), 2),
            medVal = round(median(HR), 2),
            minVal = round(min(HR), 2),
            maxVal = round(max(HR), 2),
            stdVal = round(sd(HR), 2))

# Summarize data for Infected Hours (IH)
df_summary_IH = df_main %>% 
  group_by(Industry, OrgCategories) %>% 
  summarize(count = n(), 
            meanVal = round(mean(IH), 0),
            medVal = round(median(IH), 0),
            minVal = round(min(IH), 0),
            maxVal = round(max(IH), 0),
            stdVal = round(sd(IH), 0))

# Summarize data for Business Continuity Cost (BCC) 
df_summary_BCC = df_main %>% 
  group_by(Industry, OrgCategories) %>% 
  summarize(count = n(), 
            meanVal = round(mean(BCC), 2),
            medVal = round(median(BCC), 2),
            minVal = round(min(BCC), 2),
            maxVal = round(max(BCC), 2),
            stdVal = round(sd(BCC), 2))

# Summarize data for Incident Response Cost (IRC)
df_summary_IRC = df_main %>% 
  group_by(Industry, OrgCategories) %>% 
  summarize(count = n(), 
            meanVal = round(mean(IRC), 2),
            medVal = round(median(IRC), 2),
            minVal = round(min(IRC), 2),
            maxVal = round(max(IRC), 2),
            stdVal = round(sd(IRC), 2))

# Summarize data for Ransom Paid Cost (RPC)
df_summary_RPC = df_main %>% 
  group_by(Industry, OrgCategories) %>% 
  summarize(count = n(), 
            meanVal = round(mean(RPC), 2),
            medVal = round(median(RPC), 2),
            minVal = round(min(RPC), 2),
            maxVal = round(max(RPC), 2),
            stdVal = round(sd(RPC), 2))

# Summarize data for New Installation Cost (NIC)
df_summary_NIC = df_main %>% 
  group_by(Industry, OrgCategories) %>% 
  summarize(count = n(), 
            meanVal = round(mean(NIC), 2),
            medVal = round(median(NIC), 2),
            minVal = round(min(NIC), 2),
            maxVal = round(max(NIC), 2),
            stdVal = round(sd(NIC), 2))

# Summarize data for Data Recovery Cost (DRC)
df_summary_DRC = df_main %>% 
  group_by(Industry, OrgCategories) %>% 
  summarize(count = n(), 
            meanVal = round(mean(DRC), 2),
            medVal = round(median(DRC), 2),
            minVal = round(min(DRC), 2),
            maxVal = round(max(DRC), 2),
            stdVal = round(sd(DRC), 2))

# Summarize data for Compliance Penalty Cost (CPC)
df_summary_CPC = df_main %>% 
  group_by(Industry, OrgCategories) %>% 
  summarize(count = n(), 
            meanVal = round(mean(CPC), 2),
            medVal = round(median(CPC), 2),
            minVal = round(min(CPC), 2),
            maxVal = round(max(CPC), 2),
            stdVal = round(sd(CPC), 2))

# Summarize data for Reputation Damage Cost (RDC)
df_summary_RDC = df_main %>% 
  group_by(Industry, OrgCategories) %>% 
  summarize(count = n(), 
            meanVal = round(mean(RDC), 2),
            medVal = round(median(RDC), 2),
            minVal = round(min(RDC), 2),
            maxVal = round(max(RDC), 2),
            stdVal = round(sd(RDC), 2))

# Summarize data for Estimated Loss (EL)
df_summary_EL = df_main %>% 
  group_by(Industry, OrgCategories) %>% 
  summarize(count = n(), 
            meanVal = round(mean(EL), 2),
            medVal = round(median(EL), 2),
            minVal = round(min(EL), 2),
            maxVal = round(max(EL), 2),
            stdVal = round(sd(EL), 2))

# Summarize data for Annual Revenue After Estimated Loss (ARAEL)
df_summary_ARAEL = df_main %>% 
  group_by(Industry, OrgCategories) %>% 
  summarize(count = n(), 
            meanVal = round(mean(ARAEL), 2),
            medVal = round(median(ARAEL), 2),
            minVal = round(min(ARAEL), 2),
            maxVal = round(max(ARAEL), 2),
            stdVal = round(sd(ARAEL), 2))

# Summarize data for Percentage Loss In Annual Revenue (PLIAR)
df_summary_PLIAR = df_main %>% 
  group_by(Industry, OrgCategories) %>% 
  summarize(count = n(), 
            meanVal = round(mean(PLIAR), 2),
            medVal = round(median(PLIAR), 2),
            minVal = round(min(PLIAR), 2),
            maxVal = round(max(PLIAR), 2),
            stdVal = round(sd(PLIAR), 2))

# ********************Summarized Data Frames Exported as Excel Sheets*******************

# Summarized Employees
writexl::write_xlsx(df_summary_Employees,"data\\df_summary_Employees.xlsx")

# Summarized Annual Review (AR)
writexl::write_xlsx(df_summary_AR,"data\\df_summary_AR.xlsx")

# Summarized Total Machines (TM)
writexl::write_xlsx(df_summary_TM,"data\\df_summary_TM.xlsx")

# Summarized Infection Ratio (IR)
writexl::write_xlsx(df_summary_IR,"data\\df_summary_IR.xlsx")

# Summarized Infected Machines (IM)
writexl::write_xlsx(df_summary_IM,"data\\df_summary_IM.xlsx")

# Summarized Hourly Revenue (HR)
writexl::write_xlsx(df_summary_HR,"data\\df_summary_HR.xlsx")

# Summarized Infected Hours (IH)
writexl::write_xlsx(df_summary_IH,"data\\df_summary_IH.xlsx")

# Summarized Business Continuity Cost (BCC) 
writexl::write_xlsx(df_summary_BCC,"data\\df_summary_BCC.xlsx")

# Summarized Incident Response Cost (IRC)
writexl::write_xlsx(df_summary_IRC,"data\\df_summary_IRC.xlsx")

# Summarized Ransom Paid Cost (RPC)
writexl::write_xlsx(df_summary_RPC,"data\\df_summary_RPC.xlsx")

# Summarized New Installation Cost (NIC)
writexl::write_xlsx(df_summary_NIC,"data\\df_summary_NIC.xlsx")

# Summarized Data Recovery Cost (DRC)
writexl::write_xlsx(df_summary_DRC,"data\\df_summary_DRC.xlsx")

# Summarized Compliance Penalty Cost (CPC)
writexl::write_xlsx(df_summary_CPC,"data\\df_summary_CPC.xlsx")

# Summarized Reputation Damage Cost (RDC)
writexl::write_xlsx(df_summary_RDC,"data\\df_summary_RDC.xlsx")

# Summarized Estimated Loss (EL)
writexl::write_xlsx(df_summary_EL,"data\\df_summary_EL.xlsx")

# Summarized Annual Revenue After Estimated Loss (ARAEL)
writexl::write_xlsx(df_summary_ARAEL,"data\\df_summary_ARAEL.xlsx")

# Summarized Percentage Loss In Annual Revenue (PLIAR)
writexl::write_xlsx(df_summary_PLIAR,"data\\df_summary_PLIAR.xlsx")


# *************************Code End Here*********************************


