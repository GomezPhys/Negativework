library(readxl) ## To load excel sheet
library(dplyr) # Data grammar and manipulation
library(rstatix) # Shapiro Wilk and effect size
library(psych) #descriptives
library(kableExtra) #tables
library(lme4) #linear mixed effects models (LMM)
library(lmerTest) #anova like output for LMM
library(ggplot2) #data visualization
library(ggpubr)#data visualization
library(ggprism)##makes plots look like graphad
library(table1) #for descriptives

#####Max####

Max <- read_excel("~/ECC_MX.xlsx",
                           sheet = "max")
View(Max)

## Order conditions
Max$Condition <- ordered(Max$Condition,
                                  levels = c("CON", "ECC",
                                             "ECC2"))

####NormalityMax
Max %>% group_by(Condition) %>% 
  shapiro_test(VO2max)
Max %>% group_by(Condition) %>% 
  shapiro_test(HR)
Max %>% group_by(Condition) %>% 
  shapiro_test(Lactate)
Max %>% group_by(Condition) %>% 
  shapiro_test(RPE)
Max %>% group_by(Condition) %>% 
  shapiro_test(Workload)
Max %>% group_by(Condition) %>% 
  shapiro_test(Age)
Max %>% group_by(Condition) %>% 
  shapiro_test(Weight)

#####Descriptives with IQR since non-normally distributed####
table1(~ Age + Weight + VO2max + HR + Lactate + RPE + Workload | Condition,
           total=F,render.categorical="FREQ (PCTnoNA%)", na.rm = TRUE,data=Max,
              render.missing=NULL,topclass="Rtable1-grid Rtable1-shade Rtable1-times",
              overall=FALSE)
#####Ancovas#####
  #####VO2####
ancova_VO2max <- aov(VO2max ~ Condition, data = Max)
# Summary of the ANCOVA model for VO2
summary(ancova_VO2max)
# Perform Tukey's HSD post-hoc test for VO2max
tukey_vo2max <- TukeyHSD(ancova_VO2max)

# View the results of Tukey's HSD test
print(tukey_vo2max)
# Extract the Tukey HSD results for Condition
tukey_vo2max_df <- as.data.frame(tukey_vo2max$Condition)

# View the data frame
print(tukey_vo2max_df)

tukey_condition_vo2max <- TukeyHSD(ancova_VO2max, "Condition")

# Extract the Tukey HSD results for Condition
tukey_condition_vo2max_df <- as.data.frame(tukey_condition_vo2max$Condition)

# View the data frame (optional)
print(tukey_condition_vo2max_df)

ggplot(tukey_condition_vo2max_df, aes(x = row.names(tukey_condition_vo2max_df), y = diff)) +
  geom_bar(stat = "identity", fill = "steelblue", width = 0.7) +
  geom_errorbar(aes(ymin = lwr, ymax = upr), width = 0.2, color = "black") +
  labs(title = "Tukey HSD for VO2max by Condition (ECC, ECC2, CON)",
       x = "Condition Comparisons",
       y = "Difference in VO2max (Mean ± 95% CI)") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

tukey_condition_vo2max_df %>%
  kable("html", caption = "Tukey HSD for Condition Comparisons (ECC, ECC2, CON)") %>%
  kable_styling(full_width = FALSE, bootstrap_options = c("striped", "hover", "condensed", "responsive"))

#####HR#####
ancova_HR <- aov(HR ~ Condition, data = Max)
# Summary of the ANCOVA model for HR
summary(ancova_HR)
# Perform Tukey's HSD post-hoc test for HR
tukey_HR <- TukeyHSD(ancova_HR)

# View the results of Tukey's HSD test
print(tukey_HR)
# Extract the Tukey HSD results for Condition
tukey_HR_Max <- as.data.frame(tukey_HR$Condition)

# View the data frame
print(tukey_HR_Max)

tukey_condition_HR <- TukeyHSD(ancova_HR, "Condition")

# Extract the Tukey HSD results for Condition
tukey_condition_HR_Max <- as.data.frame(tukey_condition_HR$Condition)

# View the data frame (optional)
print(tukey_condition_HR_Max)

ggplot(tukey_condition_HR_Max, aes(x = row.names(tukey_condition_HR_Max), y = diff)) +
  geom_bar(stat = "identity", fill = "steelblue", width = 0.7) +
  geom_errorbar(aes(ymin = lwr, ymax = upr), width = 0.2, color = "black") +
  labs(title = "Tukey HSD for VO2max by Condition (ECC, ECC2, CON)",
       x = "Condition Comparisons",
       y = "Difference in VO2max (Mean ± 95% CI)") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

tukey_condition_HR_Max %>%
  kable("html", caption = "Tukey HSD for Condition Comparisons (ECC, ECC2, CON)") %>%
  kable_styling(full_width = FALSE, bootstrap_options = c("striped", "hover", "condensed", "responsive"))

#####Lactate#####
# Fit ANCOVA model for Lactate with Condition as a factor
ancova_Lactate <- aov(Lactate ~ Condition, data = Max)

# Summary of the ANCOVA model for Lactate
summary(ancova_Lactate)

# Perform Tukey's HSD post-hoc test for Lactate
tukey_Lactate <- TukeyHSD(ancova_Lactate)

# View the results of Tukey's HSD test
print(tukey_Lactate)

# Extract the Tukey HSD results for Condition
tukey_Lactate_Max <- as.data.frame(tukey_Lactate$Condition)

# View the data frame
print(tukey_Lactate_Max)

# Perform Tukey's HSD post-hoc test for Condition in Lactate
tukey_condition_Lactate <- TukeyHSD(ancova_Lactate, "Condition")

# Extract the Tukey HSD results for Condition
tukey_condition_Lactate_Max <- as.data.frame(tukey_condition_Lactate$Condition)

# View the data frame (optional)
print(tukey_condition_Lactate_Max)

# Create the bar plot showing Tukey HSD results for Lactate by Condition
ggplot(tukey_condition_Lactate_Max, aes(x = row.names(tukey_condition_Lactate_Max), y = diff)) +
  geom_bar(stat = "identity", fill = "steelblue", width = 0.7) +
  geom_errorbar(aes(ymin = lwr, ymax = upr), width = 0.2, color = "black") +
  labs(title = "Tukey HSD for Lactate by Condition (ECC, ECC2, CON)",
       x = "Condition Comparisons",
       y = "Difference in Lactate (Mean ± 95% CI)") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

# Display Tukey HSD results in a table format
tukey_condition_Lactate_Max %>%
  kable("html", caption = "Tukey HSD for Condition Comparisons (ECC, ECC2, CON)") %>%
  kable_styling(full_width = FALSE, bootstrap_options = c("striped", "hover", "condensed", "responsive"))


                      
#####RPE####
# Fit ANCOVA model for RPE with Condition as a factor
ancova_RPE <- aov(RPE ~ Condition, data = Max)

# Summary of the ANCOVA model for RPE
summary(ancova_RPE)

# Perform Tukey's HSD post-hoc test for RPE
tukey_RPE <- TukeyHSD(ancova_RPE)

# View the results of Tukey's HSD test
print(tukey_RPE)

# Extract the Tukey HSD results for Condition
tukey_RPE_Max <- as.data.frame(tukey_RPE$Condition)

# View the data frame
print(tukey_RPE_Max)

# Perform Tukey's HSD post-hoc test for Condition in RPE
tukey_condition_RPE <- TukeyHSD(ancova_RPE, "Condition")

# Extract the Tukey HSD results for Condition
tukey_condition_RPE_Max <- as.data.frame(tukey_condition_RPE$Condition)

# View the data frame (optional)
print(tukey_condition_RPE_Max)

# Create the bar plot showing Tukey HSD results for RPE by Condition
ggplot(tukey_condition_RPE_Max, aes(x = row.names(tukey_condition_RPE_Max), y = diff)) +
  geom_bar(stat = "identity", fill = "steelblue", width = 0.7) +
  geom_errorbar(aes(ymin = lwr, ymax = upr), width = 0.2, color = "black") +
  labs(title = "Tukey HSD for RPE by Condition (ECC, ECC2, CON)",
       x = "Condition Comparisons",
       y = "Difference in RPE (Mean ± 95% CI)") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

# Display Tukey HSD results in a table format
tukey_condition_RPE_Max %>%
  kable("html", caption = "Tukey HSD for Condition Comparisons (ECC, ECC2, CON)") %>%
  kable_styling(full_width = FALSE, bootstrap_options = c("striped", "hover", "condensed", "responsive"))
#####Workload####
# Fit ANCOVA model for Workload with Condition as a factor
ancova_Workload <- aov(Workload ~ Condition, data = Max)

# Summary of the ANCOVA model for Workload
summary(ancova_Workload)

# Perform Tukey's HSD post-hoc test for Workload
tukey_Workload <- TukeyHSD(ancova_Workload)

# View the results of Tukey's HSD test
print(tukey_Workload)

# Extract the Tukey HSD results for Condition
tukey_Workload_Max <- as.data.frame(tukey_Workload$Condition)

# View the data frame
print(tukey_Workload_Max)

# Perform Tukey's HSD post-hoc test for Condition in Workload
tukey_condition_Workload <- TukeyHSD(ancova_Workload, "Condition")

# Extract the Tukey HSD results for Condition
tukey_condition_Workload_Max <- as.data.frame(tukey_condition_Workload$Condition)

# View the data frame (optional)
print(tukey_condition_Workload_Max)

# Create the bar plot showing Tukey HSD results for Workload by Condition
ggplot(tukey_condition_Workload_Max, aes(x = row.names(tukey_condition_Workload_Max), y = diff)) +
  geom_bar(stat = "identity", fill = "steelblue", width = 0.7) +
  geom_errorbar(aes(ymin = lwr, ymax = upr), width = 0.2, color = "black") +
  labs(title = "Tukey HSD for Workload by Condition (ECC, ECC2, CON)",
       x = "Condition Comparisons",
       y = "Difference in Workload (Mean ± 95% CI)") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))

# Display Tukey HSD results in a table format
tukey_condition_Workload_Max %>%
  kable("html", caption = "Tukey HSD for Condition Comparisons (ECC, ECC2, CON)") %>%
  kable_styling(full_width = FALSE, bootstrap_options = c("striped", "hover", "condensed", "responsive"))


#####Table for all###########
# Create separate data frames for each variable's Tukey HSD results
tukey_vo2max_df <- as.data.frame(tukey_vo2max$Condition)
tukey_vo2max_df$Variable <- "VO2max"

tukey_lactate_df <- as.data.frame(tukey_Lactate$Condition)
tukey_lactate_df$Variable <- "Lactate"

tukey_rpe_df <- as.data.frame(tukey_RPE$Condition)
tukey_rpe_df$Variable <- "RPE"

tukey_workload_df <- as.data.frame(tukey_Workload$Condition)
tukey_workload_df$Variable <- "Workload"

# Combine all the data frames into a single data frame
combined_tukey_df <- rbind(tukey_vo2max_df, tukey_lactate_df, tukey_rpe_df, tukey_workload_df)

# Add row names as a column (for condition comparisons)
combined_tukey_df$Comparison <- row.names(combined_tukey_df)

# Rearrange columns for better readability
combined_tukey_df <- combined_tukey_df[, c("Variable", "Comparison", "diff", "lwr", "upr", "p adj")]

# Display the combined Tukey HSD results in a table format
library(kableExtra)

combined_tukey_df %>%
  kable("html", caption = "Tukey HSD Results for Multiple Variables (ECC, ECC2, CON)") %>%
  kable_styling(full_width = FALSE, bootstrap_options = c("striped", "hover", "condensed", "responsive"))


#####Tailored Intensities 
  Df <- read_excel("~/ECC_MX.xlsx",
                 sheet = "tailored")

  View(Df)
  Df %>% group_by(Condition,Modality) %>% 
    shapiro_test(VO2)
  
  Df %>% group_by(Condition,Modality) %>% 
    shapiro_test(HR)
  
  Df %>% group_by(Condition,Modality) %>% 
    shapiro_test(RPE)
  
  Df %>% group_by(Condition,Modality) %>% 
    shapiro_test(Workload)
  
  ## Order conditions
  Df$Condition <- ordered(Df$Condition,
                          levels = c("Baseline", "Low", "Moderate", "High"))
  
  Df$Modality <- ordered(Df$Modality, levels = c("CON_CON", "ECC_CON", "ECC_ECC"))

  ####Vo2####

  # Perform ANCOVA for VO2 including Condition and Modality interaction
  ancova_vo2 <- aov(VO2 ~ Condition * Modality, data = Df)
  
  # Summary of the ANCOVA model for VO2
  summary(ancova_vo2)
  
  # Perform Tukey's HSD post-hoc test for VO2 by Condition and Modality interaction
  tukey_vo2 <- TukeyHSD(ancova_vo2, which = "Condition:Modality")
  
  # Extract the Tukey HSD results for Condition and Modality interaction
  tukey_vo2_df <- as.data.frame(tukey_vo2$`Condition:Modality`)
  
  # Add a column for the variable name
  tukey_vo2_df$Variable <- "VO2"
  
  # View the data frame
  print(tukey_vo2_df)
  
  # Optional: Display the Tukey HSD results in a table format
  library(kableExtra)
  
  tukey_vo2_df %>%
    kable("html", caption = "Tukey HSD for VO2 by Condition and Modality Interaction") %>%
    kable_styling(full_width = FALSE, bootstrap_options = c("striped", "hover", "condensed", "responsive"))
  
  # Perform ANCOVA for VO2 including Condition and Modality interaction
  ancova_vo2 <- aov(VO2 ~ Condition * Modality, data = Df)
  
  # Perform Tukey's HSD post-hoc test for Condition and Modality interaction
  tukey_vo2 <- TukeyHSD(ancova_vo2, which = "Condition:Modality")
  
  # Convert Tukey HSD results to data frame
  tukey_vo2_df <- as.data.frame(tukey_vo2$`Condition:Modality`)
  
  # Add comparison names from rownames to the data frame
  tukey_vo2_df$Comparison <- rownames(tukey_vo2_df)
  
  # View the data frame (optional)
  print(tukey_vo2_df)
  
  between_modalities <- tukey_vo2_df %>%
    filter(grepl("CON_CON", Comparison) | grepl("ECC_ECC", Comparison) | grepl("ECC_CON", Comparison))
  
  # Filter results for within-modality comparisons (same modality, different conditions)
  within_modalities <- tukey_vo2_df %>%
    filter(!grepl("CON_CON", Comparison) & !grepl("ECC_ECC", Comparison) & !grepl("ECC_CON", Comparison))
  
  # View the tables
  print(between_modalities)
  print(within_modalities)
  
  between_modalities %>%
    kable("html", caption = "Tukey HSD for Between-Modality Comparisons") %>%
    kable_styling(full_width = FALSE, bootstrap_options = c("striped", "hover", "condensed", "responsive"))
  
  within_modalities %>%
    kable("html", caption = "Tukey HSD for Within-Modality Comparisons") %>%
    kable_styling(full_width = FALSE, bootstrap_options = c("striped", "hover", "condensed", "responsive"))
  
  #####HR####
  # Perform ANCOVA for HR including Condition and Modality interaction
  ancova_hr <- aov(HR ~ Condition * Modality, data = Df)
  
  # Summary of the ANCOVA model for HR
  summary(ancova_hr)
  
  # Perform Tukey's HSD post-hoc test for HR by Condition and Modality interaction
  tukey_hr <- TukeyHSD(ancova_hr, which = "Condition:Modality")
  
  # Extract the Tukey HSD results for Condition and Modality interaction
  tukey_hr_df <- as.data.frame(tukey_hr$`Condition:Modality`)
  
  # Add a column for the variable name
  tukey_hr_df$Variable <- "HR"
  
  # Add comparison names from rownames to the data frame
  tukey_hr_df$Comparison <- rownames(tukey_hr_df)
  
  # View the data frame
  print(tukey_hr_df)
  
  # Separate the results into between-modality and within-modality comparisons
  
  # Between-modality comparisons
  between_modalities_hr <- tukey_hr_df %>%
    filter(grepl("CON_CON", Comparison) | grepl("ECC_ECC", Comparison) | grepl("ECC_CON", Comparison))
  
  # Within-modality comparisons (same modality, different conditions)
  within_modalities_hr <- tukey_hr_df %>%
    filter(!grepl("CON_CON", Comparison) & !grepl("ECC_ECC", Comparison) & !grepl("ECC_CON", Comparison))
  
  # View the tables
  print(between_modalities_hr)
  print(within_modalities_hr)
  
  # Display between-modality comparisons in a table
  between_modalities_hr %>%
    kable("html", caption = "Tukey HSD for Between-Modality Comparisons (HR)") %>%
    kable_styling(full_width = FALSE, bootstrap_options = c("striped", "hover", "condensed", "responsive"))
  
  # Display within-modality comparisons in a table
  within_modalities_hr %>%
    kable("html", caption = "Tukey HSD for Within-Modality Comparisons (HR)") %>%
    kable_styling(full_width = FALSE, bootstrap_options = c("striped", "hover", "condensed", "responsive"))
  
  
  
  #####RPE#####
  # Perform ANCOVA for RPE including Condition and Modality interaction
  ancova_rpe <- aov(RPE ~ Condition * Modality, data = Df)
  
  # Summary of the ANCOVA model for RPE
  summary(ancova_rpe)
  
  # Perform Tukey's HSD post-hoc test for RPE by Condition and Modality interaction
  tukey_rpe <- TukeyHSD(ancova_rpe, which = "Condition:Modality")
  
  # Extract the Tukey HSD results for Condition and Modality interaction
  tukey_rpe_df <- as.data.frame(tukey_rpe$`Condition:Modality`)
  
  # Add a column for the variable name
  tukey_rpe_df$Variable <- "RPE"
  
  # Add comparison names from rownames to the data frame
  tukey_rpe_df$Comparison <- rownames(tukey_rpe_df)
  
  # View the data frame
  print(tukey_rpe_df)
  
  # Separate the results into between-modality and within-modality comparisons
  
  # Between-modality comparisons
  between_modalities_rpe <- tukey_rpe_df %>%
    filter(grepl("CON_CON", Comparison) | grepl("ECC_ECC", Comparison) | grepl("ECC_CON", Comparison))
  
  # Within-modality comparisons (same modality, different conditions)
  within_modalities_rpe <- tukey_rpe_df %>%
    filter(!grepl("CON_CON", Comparison) & !grepl("ECC_ECC", Comparison) & !grepl("ECC_CON", Comparison))
  
  # View the tables
  print(between_modalities_rpe)
  print(within_modalities_rpe)
  
  # Display between-modality comparisons in a table
  between_modalities_rpe %>%
    kable("html", caption = "Tukey HSD for Between-Modality Comparisons (RPE)") %>%
    kable_styling(full_width = FALSE, bootstrap_options = c("striped", "hover", "condensed", "responsive"))
  
  # Display within-modality comparisons in a table
  within_modalities_rpe %>%
    kable("html", caption = "Tukey HSD for Within-Modality Comparisons (RPE)") %>%
    kable_styling(full_width = FALSE, bootstrap_options = c("striped", "hover", "condensed", "responsive"))
  
  #####Workload####
  # Perform ANCOVA for Workload including Condition and Modality interaction
  ancova_workload <- aov(Workload ~ Condition * Modality, data = Df)
  
  # Summary of the ANCOVA model for Workload
  summary(ancova_workload)
  
  # Perform Tukey's HSD post-hoc test for Workload by Condition and Modality interaction
  tukey_workload <- TukeyHSD(ancova_workload, which = "Condition:Modality")
  
  # Extract the Tukey HSD results for Condition and Modality interaction
  tukey_workload_df <- as.data.frame(tukey_workload$`Condition:Modality`)
  
  # Add a column for the variable name
  tukey_workload_df$Variable <- "Workload"
  
  # Add comparison names from rownames to the data frame
  tukey_workload_df$Comparison <- rownames(tukey_workload_df)
  
  # View the data frame
  print(tukey_workload_df)
  
  # Separate the results into between-modality and within-modality comparisons
  
  # Between-modality comparisons
  between_modalities_workload <- tukey_workload_df %>%
    filter(grepl("CON_CON", Comparison) | grepl("ECC_ECC", Comparison) | grepl("ECC_CON", Comparison))
  
  # Within-modality comparisons (same modality, different conditions)
  within_modalities_workload <- tukey_workload_df %>%
    filter(!grepl("CON_CON", Comparison) & !grepl("ECC_ECC", Comparison) & !grepl("ECC_CON", Comparison))
  
  # View the tables
  print(between_modalities_workload)
  print(within_modalities_workload)
  
  # Display between-modality comparisons in a table
  between_modalities_workload %>%
    kable("html", caption = "Tukey HSD for Between-Modality Comparisons (Workload)") %>%
    kable_styling(full_width = FALSE, bootstrap_options = c("striped", "hover", "condensed", "responsive"))
  
  # Display within-modality comparisons in a table
  within_modalities_workload %>%
    kable("html", caption = "Tukey HSD for Within-Modality Comparisons (Workload)") %>%
    kable_styling(full_width = FALSE, bootstrap_options = c("striped", "hover", "condensed", "responsive"))
  
  #####Lactate#####
  # Perform ANCOVA for Lactate including Condition and Modality interaction
  ancova_lactate <- aov(Lactate ~ Condition * Modality, data = Df)
  
  # Summary of the ANCOVA model for Lactate
  summary(ancova_lactate)
  
  # Perform Tukey's HSD post-hoc test for Lactate by Condition and Modality interaction
  tukey_lactate <- TukeyHSD(ancova_lactate, which = "Condition:Modality")
  
  # Extract the Tukey HSD results for Condition and Modality interaction
  tukey_lactate_df <- as.data.frame(tukey_lactate$`Condition:Modality`)
  
  # Add a column for the variable name
  tukey_lactate_df$Variable <- "Lactate"
  
  # Add comparison names from rownames to the data frame
  tukey_lactate_df$Comparison <- rownames(tukey_lactate_df)
  
  # View the data frame
  print(tukey_lactate_df)
  
  # Separate the results into between-modality and within-modality comparisons
  
  # Between-modality comparisons
  between_modalities_lactate <- tukey_lactate_df %>%
    filter(grepl("CON_CON", Comparison) | grepl("ECC_ECC", Comparison) | grepl("ECC_CON", Comparison))
  
  # Within-modality comparisons (same modality, different conditions)
  within_modalities_lactate <- tukey_lactate_df %>%
    filter(!grepl("CON_CON", Comparison) & !grepl("ECC_ECC", Comparison) & !grepl("ECC_CON", Comparison))
  
  # View the tables
  print(between_modalities_lactate)
  print(within_modalities_lactate)
  
  # Display between-modality comparisons in a table
  between_modalities_lactate %>%
    kable("html", caption = "Tukey HSD for Between-Modality Comparisons (Lactate)") %>%
    kable_styling(full_width = FALSE, bootstrap_options = c("striped", "hover", "condensed", "responsive"))
  
  Filter between-modality comparisons
  between_modalities_lactate <- tukey_lactate_df %>%
    filter(grepl("CON_CON", Comparison) & (grepl("ECC_CON", Comparison) | grepl("ECC_ECC", Comparison)) |
             grepl("ECC_CON", Comparison) & grepl("ECC_ECC", Comparison))
  

  