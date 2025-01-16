# Load required libraries
library(dplyr)
library(lubridate)

set.seed(123)  # For reproducibility

# Generate the data
n <- 100

# Generate the variables
studyID <- paste0("S", 1:n)
age <- sample(18:70, n, replace = TRUE)
sex <- sample(c("Male", "Female"), n, replace = TRUE)
BMI <- round(runif(n, 18.5, 40), 1)  # BMI between 18.5 and 40
heart_rate <- sample(60:100, n, replace = TRUE)  # Heart rate between 60 and 100 bpm

# Random diagnosis date between 1 Jan 2020 and 31 Dec 2023
diagnosis_date <- sample(seq(ymd("2020-01-01"), ymd("2023-12-31"), by="day"), n, replace = TRUE)

# Random randomisation date between 1 Jan 2023 and 31 Dec 2023
randomisation_date <- sample(seq(ymd("2023-01-01"), ymd("2023-12-31"), by="day"), n, replace = TRUE)

# ECOG score (values 0 to 3)
ECOG <- sample(0:3, n, replace = TRUE)

# Treatment group (0 or 1)
treatment_group <- sample(0:1, n, replace = TRUE)

# Baseline score (0 to 100)
baseline_score <- sample(0:100, n, replace = TRUE)

# Final score (0 to 100)
final_score <- sample(0:100, n, replace = TRUE)

# Site (values 1, 2, or 3)
site <- sample(1:3, n, replace = TRUE)

# Diagnosis (values 1, 2, 3, or 4)
diagnosis <- sample(1:4, n, replace = TRUE)

# Combine all variables into a data frame
data <- data.frame(
  studyID = studyID,
  age = age,
  sex = sex,
  BMI = BMI,
  heart_rate = heart_rate,
  date_of_diagnosis = diagnosis_date,
  date_of_randomisation = randomisation_date,
  ECOG = ECOG,
  treatment_group = treatment_group,
  baseline_score = baseline_score,
  final_score = final_score,
  site = site,
  diagnosis = diagnosis
)

#
View(data)

# write to an Excel file - then run the data checker macros
openxlsx::write.xlsx(data,file="study_data.xlsx")
