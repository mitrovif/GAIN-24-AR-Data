# ======================================================
# GAIN 2024 Annual Report Data Processing Script
# - Uses a user-defined working directory
# - Saves output files with a date stamp for version control
# ======================================================

# Load required libraries
library(dplyr)
library(tidyr)
library(readr)
library(readxl)
library(writexl)
library(openxlsx)
library(officer)
library(flextable)
library(ggplot2)
library(sf)
library(rnaturalearth)
library(rnaturalearthdata)
library(ggrepel)

# ======================================================
# Set Working Directory Dynamically
# ======================================================
# Paste your copied Windows file path here
working_dir <- "C:\\Users\\mitro\\UNHCR\\EGRISS Secretariat - 905 - Implementation of Recommendations\\01_GAIN Survey\\Integration & GAIN Survey\\EGRISS GAIN Survey 2024\\10 Data\\Analysis Ready Files\\Backup_2025-02-27_13-42-02"

# Automatically replace backslashes (\) with forward slashes (/)
working_dir <- gsub("\\\\", "/", working_dir)

# Set the working directory
setwd(working_dir)

# Confirm the working directory
message("Working directory set to: ", getwd())

# Get current date in YYYY-MM-DD format
current_date <- format(Sys.Date(), "%Y-%m-%d")  # Define the missing object
# ======================================================
# Load Group Roster Dataset with Relative Path
# ======================================================
group_roster_file <- file.path(working_dir, "analysis_ready_group_roster.csv")
group_roster <- read.csv(group_roster_file)

# ======================================================
# Save Excel Output in the Same Folder
# ======================================================
output_excel_file <- file.path(working_dir, "Annual Report GAIN 2024.xlsx")

# Save output
write_xlsx(list(`Figure 6` = summary_table), path = output_excel_file)

message("Summary table exported to 'Annual Report GAIN 2024.xlsx'.")



# ======================================================
# Tabulates `g_conled` by `ryear` and `PRO09`, replaces numeric values with descriptive text,
# and exports the table to the specified folder with the filename "Annual Report GAIN 2024.xlsx".
# Renames the Excel sheet to "Figure 6".
# ======================================================



# Load the group roster dataset
group_roster_file <- file.path(working_dir, "analysis_ready_group_roster.csv")
group_roster <- read.csv(group_roster_file)

# Filter and group data to create a summary table
summary_table <- group_roster %>%
  group_by(ryear, g_conled, PRO09) %>%
  summarise(count = n(), .groups = "drop") %>%
  pivot_wider(names_from = PRO09, values_from = count, values_fill = 0, names_prefix = "PRO09_")

# Replace numeric values in `g_conled` with descriptive text
summary_table <- summary_table %>%
  mutate(
    g_conled = case_when(
      g_conled == 1 ~ "National Example",
      g_conled == 2 ~ "Institutional Example",
      g_conled == 3 ~ "CSO Example",
      TRUE ~ as.character(g_conled) # Handle unexpected or missing cases
    )
  )

# Clean the data and replace NA with blanks
summary_table <- summary_table %>%
  mutate(across(everything(), ~ ifelse(is.na(.), "", .))) # Replace NA with blanks

# Rename columns with descriptive labels
colnames(summary_table) <- c(
  "Year of Reporting in GAIN",                           # ryear
  "Type of Example (National, Institutional, CSO)",      # g_conled
  "Using EGRISS Recommendations in the Example",        # PRO09_1
  "Not Using EGRISS Recommendations in the Example",    # PRO09_2
  "Don't Know if EGRISS Recommendations are Used",       # PRO09_8
  "column_name"                                          # PRO_?
)

# Export to the specified folder with the updated name and sheet title
output_excel_file <- file.path(working_dir, "Annual Report GAIN 2024.xlsx")
# Save the table with the renamed sheet using writexl
write_xlsx(list(`Figure 6` = summary_table), path = output_excel_file)

message("Summary table exported to 'Annual Report GAIN 2024.xlsx' with the sheet renamed to 'Figure 6'.")

# ======================================================
# Creates a summary table for `g_recuse` (Use of Recommendations) with the condition `PRO09 == 1`.
# Calculates breakdown dynamically for each year and exports to Excel in a new sheet named "Figure 7".
# ======================================================

# Load required libraries
library(dplyr)
library(tidyr)
library(writexl)

# Load the group roster dataset
group_roster_file <- file.path(working_dir, "analysis_ready_group_roster.csv")
group_roster <- read.csv(group_roster_file)

# Filter examples where PRO09 == 1
group_roster <- group_roster %>% filter(PRO09 == 1)

# Create `g_recuse` based on PRO10.A, PRO10.B, and PRO10.C
group_roster <- group_roster %>%
  mutate(
    g_recuse = case_when(
      PRO10.A == 1 & PRO10.B != 1 & PRO10.C != 1 ~ "Use of IRRS Only",
      PRO10.A != 1 & PRO10.B == 1 & PRO10.C != 1 ~ "Use of IRIS Only",
      PRO10.A != 1 & PRO10.B != 1 & PRO10.C == 1 ~ "Use of IROSS Only",
      (PRO10.A + PRO10.B + PRO10.C) > 1 ~ "Any Combination of Recommendations", # Combined category for multiple recommendations
      TRUE ~ "No Use of Recommendations"
    )
  )

# Summarize the data by `ryear` and `g_recuse`
recuse_table <- group_roster %>%
  group_by(ryear, g_recuse) %>%
  summarise(count = n(), .groups = "drop") %>%
  pivot_wider(names_from = g_recuse, values_from = count, values_fill = 0)

# Rename `ryear` column
colnames(recuse_table)[1] <- "Year of GAIN Data Collection"

# Calculate breakdown of "Any Combination of Recommendations" dynamically for each year
breakdown_combination <- group_roster %>%
  mutate(
    combination = case_when(
      PRO10.A == 1 & PRO10.B == 1 & PRO10.C != 1 ~ "IRRS + IRIS",
      PRO10.A == 1 & PRO10.B != 1 & PRO10.C == 1 ~ "IRRS + IROSS",
      PRO10.A != 1 & PRO10.B == 1 & PRO10.C == 1 ~ "IRIS + IROSS",
      PRO10.A == 1 & PRO10.B == 1 & PRO10.C == 1 ~ "All 3 Combined",
      TRUE ~ NA_character_
    )
  ) %>%
  filter(!is.na(combination)) %>%
  group_by(ryear, combination) %>%
  summarise(count = n(), .groups = "drop")

# Append the breakdown dynamically for each year
recuse_table <- recuse_table %>%
  left_join(
    breakdown_combination %>%
      pivot_wider(names_from = combination, values_from = count, values_fill = 0),
    by = c("Year of GAIN Data Collection" = "ryear")
  ) %>%
  mutate(
    `Breakdown of Any Combination` = paste0(
      "IRRS + IRIS: ", `IRRS + IRIS`, "; ",
      "IRRS + IROSS: ", `IRRS + IROSS`, "; ",
      "IRIS + IROSS: ", `IRIS + IROSS`, "; ",
      "All 3 Combined: ", `All 3 Combined`
    )
  ) %>%
  select(-`IRRS + IRIS`, -`IRRS + IROSS`, -`IRIS + IROSS`, -`All 3 Combined`)

# Export the table to Excel in a new sheet named "Figure 7"
output_excel_file <- file.path(working_dir, "Annual Report GAIN 2024.xlsx")
# Add the new sheet to the Excel file
write_xlsx(
  list(
    `Figure 6` = readxl::read_excel(output_excel_file, sheet = "Figure 6"), # Keep existing Figure 6 sheet
    `Figure 7` = recuse_table                                              # Add new Figure 7 sheet
  ),
  path = output_excel_file
)

message("Summary table for `g_recuse` exported to 'Figure 7' in the Excel file.")
# ======================================================
# Filter country-led projects using recommendations, generate a user-friendly regions table, and add it to the GAIN report Excel.
# ======================================================

# Load required libraries
library(dplyr)
library(readxl)
library(writexl)

# Load the group roster dataset
group_roster_file <- file.path(working_dir, "analysis_ready_group_roster.csv")
group_roster <- read.csv(group_roster_file)

# Filter data based on conditions: PRO09 == 1 and g_conled == 1
filtered_data <- group_roster %>%
  filter(PRO09 == 1, g_conled == 1) %>%
  group_by(region, ryear) %>%
  summarise(count = n(), .groups = "drop") %>%
  arrange(region, ryear)

# Create a user-friendly table
summary_table <- filtered_data %>%
  pivot_wider(
    names_from = ryear,
    values_from = count,
    values_fill = 0
  ) %>%
  rename(
    `Region` = region
  )

# Add human-readable column names for years
colnames(summary_table)[2:ncol(summary_table)] <- paste0("Year: ", colnames(summary_table)[2:ncol(summary_table)])

# Load the existing Excel file
output_excel_file <- file.path(working_dir, "Annual Report GAIN 2024.xlsx")
existing_sheets <- readxl::excel_sheets(output_excel_file)
existing_data <- lapply(existing_sheets, function(sheet) readxl::read_excel(output_excel_file, sheet = sheet))
names(existing_data) <- existing_sheets

# Add the new table to a new sheet named "Text 1"
existing_data$`Text 1` <- summary_table

# Save the updated Excel file
write_xlsx(existing_data, output_excel_file)

message("Filtered data saved and regions table added to 'Text 1' in the Annual Report Excel.")

# ======================================================
# Create a table for country-led examples by source of data, year, and use of recommendations.
# Includes survey-only, census-only, administrative-only, data integration-only, and combined uses.
# ======================================================

# Load required libraries
library(dplyr)
library(tidyr)
library(readxl)
library(writexl)

# Load the group roster dataset
group_roster_file <- file.path(working_dir, "analysis_ready_group_roster.csv")
group_roster <- read.csv(group_roster_file)

# Filter for country-led examples (g_conled == 1)
filtered_data <- group_roster %>%
  filter(g_conled == 1) %>%
  select(ryear, PRO09, PRO08.A:PRO08.X)

# Aggregate PRO08 variables into the specified categories
aggregated_data <- filtered_data %>%
  mutate(
    SURVEY = PRO08.A,
    ADMINISTRATIVE_DATA = PRO08.B,
    CENSUS = PRO08.C,
    DATA_INTEGRATION = PRO08.D,
    OTHER = PRO08.E + PRO08.F + PRO08.G + PRO08.H + PRO08.X,
    Combined_Use = (SURVEY + ADMINISTRATIVE_DATA + CENSUS + DATA_INTEGRATION + OTHER) > 1 # Flag for combined use
  ) %>%
  mutate(
    Source_of_Data = case_when(
      SURVEY == 1 & Combined_Use == FALSE ~ "Survey Only",
      ADMINISTRATIVE_DATA == 1 & Combined_Use == FALSE ~ "Administrative Data Only",
      CENSUS == 1 & Combined_Use == FALSE ~ "Census Only",
      DATA_INTEGRATION == 1 & Combined_Use == FALSE ~ "Data Integration Only",
      Combined_Use == TRUE ~ "Combined Use",
      TRUE ~ "Other"
    )
  ) %>%
  select(ryear, PRO09, Source_of_Data) %>%
  filter(!is.na(PRO09))   # PRO09 includes NA that should be removed / renamed 

# Summarize data by year, recommendations, and source of data
summary_table <- aggregated_data %>%
  group_by(ryear, PRO09, Source_of_Data) %>%
  summarise(Count = n(), .groups = "drop") %>%
  pivot_wider(
    names_from = Source_of_Data,
    values_from = Count,
    values_fill = 0
  ) %>%
  mutate(
    PRO09 = case_when(
      PRO09 == 1 ~ "Using Recommendations",
      PRO09 == 2 ~ "Not Using Recommendations",
      PRO09 == 8 ~ "Don't Know Recommendations",
      TRUE ~ NA_character_
    )
  ) %>%
  arrange(ryear, PRO09)

# Rename columns for better readability
colnames(summary_table) <- c(
  "Year of GAIN Data Collection", "Use of Recommendations",
  "Survey Only", "Administrative Data Only", "Census Only",
  "Data Integration Only", "Combined Use", "Other"
)

# Load the existing Excel file
output_excel_file <- file.path(working_dir, "Annual Report GAIN 2024.xlsx")
existing_sheets <- readxl::excel_sheets(output_excel_file)
existing_data <- lapply(existing_sheets, function(sheet) readxl::read_excel(output_excel_file, sheet = sheet))
names(existing_data) <- existing_sheets

# Add the new table to a new sheet named "Figure 8"
existing_data$`Figure 8` <- summary_table

# Save the updated Excel file
write_xlsx(existing_data, output_excel_file)

message("Figure 8 table with breakdown of data sources added to the Annual Report Excel.")

# ======================================================
# Create a table of examples from 2024 by region, country, and use of recommendations.
# This script:
# - Filters for examples from 2024 (ryear == 2024).
# - Summarizes data for each region and country, showing:
#   - Total examples irrespective of recommendation use.
#   - Examples using EGRISS recommendations (PRO09 = 1).
#   - National examples using EGRISS recommendations (PRO09 = 1 and g_conled = 1).
# - Adds the table to the existing Annual Report GAIN 2024 Excel file as a new sheet.
# - The sheet is named "Additional Table 1" for reporting purposes.
# Load required libraries
# ======================================================

# Load required libraries
library(dplyr)
library(readxl)
library(writexl)
library(ggplot2)
library(openxlsx)
library(sf)
library(rnaturalearth)
library(rnaturalearthdata)
library(ggrepel)

# Load the group roster dataset
group_roster_file <- file.path(working_dir, "analysis_ready_group_roster.csv")
group_roster <- read.csv(group_roster_file)

# Filter for 2024 examples (ryear == 2024)
examples_2024 <- group_roster %>%
  filter(ryear == 2024) %>%
  select(region, mcountry, PRO09, g_conled)

# Create the columns for the table
summary_table <- examples_2024 %>%
  group_by(region, mcountry) %>%
  summarise(
    Total_Examples = n(),
    Using_Recommendations = sum(PRO09 == 1, na.rm = TRUE),
    National_Examples_Using_Recommendations = sum(PRO09 == 1 & g_conled == 1, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  arrange(region, mcountry)

# Rename columns for better readability
colnames(summary_table) <- c(
  "Region", "Country",
  "Total Examples (PRO09: Any)",
  "Using Recommendations (PRO09 = 1)",
  "National Examples Using Recommendations (PRO09 = 1 & g_conled = 1)"
)

# Create a world map visualization
# Aggregate data by country
country_data <- examples_2024 %>%
  group_by(mcountry) %>%
  summarise(Total_Examples = n(), .groups = "drop")

# Load world map data
world <- ne_countries(scale = "medium", returnclass = "sf")

# Join country data with world map
data_map <- world %>%
  left_join(country_data, by = c("name" = "mcountry"))

# Define EGRISS color palette
color_palette <- c("#f0f8ff", "#4cc3c9", "#3b71b3", "#072d62", "#f06667")

# Create the map
world_map_plot <- ggplot(data_map) +
  geom_sf(aes(fill = Total_Examples)) +
  geom_text_repel(data = data_map %>% filter(!is.na(Total_Examples)), 
                  aes(label = paste(name, Total_Examples, sep = ": "), geometry = geometry), 
                  stat = "sf_coordinates", size = 3, color = "black") +
  scale_fill_gradientn(colors = color_palette, na.value = "#CCCCCC") +  # suggest using a different color from the color palette above
  theme_minimal() +
  theme(panel.grid = element_blank()) +
  labs(
    title = "Global Distribution of Examples (2024)",
    fill = "Total Examples"
  )

# Save the map plot as an image
map_image_file <- file.path(working_dir, paste0("World_Map_", current_date, ".png"))

ggsave(map_image_file, world_map_plot, width = 10, height = 6)

# Load the existing Excel file
output_excel_file <- file.path(working_dir, "Annual Report GAIN 2024.xlsx")

wb <- loadWorkbook(output_excel_file)

# Add the new table to a new sheet named "Additional Table 1"
addWorksheet(wb, "Additional Table 1")
writeData(wb, "Additional Table 1", summary_table)

# Add the world map to a new sheet named "World Map 2024"
addWorksheet(wb, "World Map 2024")
insertImage(wb, "World Map 2024", file = map_image_file, width = 10, height = 6, startRow = 1, startCol = 1)

# Save the updated Excel file
saveWorkbook(wb, output_excel_file, overwrite = TRUE)

message("Summary table and world map added to the Annual Report Excel file.")


# ======================================================
# Generate a table of reported challenges for 2023 and 2024, based on country-led examples.
# This script:
# - Filters for years 2023 and 2024 (ryear).
# - Includes only country-led examples (g_conled == 1) and examples using recommendations (PRO09 == 1).
# - Counts reported challenges across variables PRO20.B to PRO20.Z for each year.
# - Ensures consistent data types for challenge variables before summarizing.
# - Adds the table to the Annual Report GAIN 2024 Excel file as "Figure 9".
# ======================================================

# Load required libraries
library(dplyr)
library(readxl)
library(writexl)
library(tidyr)

# Load the group roster dataset
group_roster_file <- file.path(working_dir, "analysis_ready_group_roster.csv")
group_roster <- read.csv(group_roster_file)

# Filter for relevant years, country-led examples, and use of recommendations
filtered_data <- group_roster %>%
  filter(ryear %in% c(2023, 2024), g_conled == 1, PRO09 == 1) %>%
  select(ryear, starts_with("PRO20."))

# Ensure consistent data types for all PRO20 columns (convert to numeric)
filtered_data <- filtered_data %>%
  mutate(across(starts_with("PRO20."), ~ as.numeric(.), .names = "clean_{col}"))

# Create labels for the challenges
challenge_labels <- c(
  "Non-Response Bias",
  "Sampling Errors",
  "Identification of Populations",
  "Data Confidentiality and Privacy",
  "Resource Constraints",
  "Political Issues",
  "Safety Concerns",
  "Timeliness and Data Quality",
  "Limited Technical Capacity",
  "Lack of Accessible Guidance",
  "Other",
  "Other (Specified)"
)

# Summarize challenges per year
summary_table <- filtered_data %>%
  pivot_longer(cols = starts_with("clean_PRO20."), names_to = "Challenge", values_to = "Reported") %>%
  mutate(Challenge = recode(Challenge,
                            `clean_PRO20.B` = "Non-Response Bias",
                            `clean_PRO20.C` = "Sampling Errors",
                            `clean_PRO20.D` = "Identification of Populations",
                            `clean_PRO20.E` = "Data Confidentiality and Privacy",
                            `clean_PRO20.F` = "Resource Constraints",
                            `clean_PRO20.G` = "Political Issues",
                            `clean_PRO20.H` = "Safety Concerns",
                            `clean_PRO20.I` = "Timeliness and Data Quality",
                            `clean_PRO20.J` = "Limited Technical Capacity",
                            `clean_PRO20.X` = "Lack of Accessible Guidance",
                            `clean_PRO20.Z` = "Other (Specified)",
                            `clean_PRO20.A` = "rename"  # rename this column or remove it
  )) %>%
  filter(Reported == 1) %>%
  group_by(ryear, Challenge) %>%
  summarise(Count = n(), .groups = "drop") %>%
  pivot_wider(names_from = Challenge, values_from = Count, values_fill = 0) %>%
  arrange(ryear)

# Rename columns for better readability
colnames(summary_table)[1] <- "Year"

# Load the existing Excel file
output_excel_file <- file.path(working_dir, "Annual Report GAIN 2024.xlsx")

existing_sheets <- readxl::excel_sheets(output_excel_file)
existing_data <- lapply(existing_sheets, function(sheet) readxl::read_excel(output_excel_file, sheet = sheet))
names(existing_data) <- existing_sheets

# Add the new table to a new sheet named "Figure 9"
existing_data$`Figure 9` <- summary_table

# Save the updated Excel file
write_xlsx(existing_data, output_excel_file)

message("Figure 9 table added to the Annual Report Excel file.")

# ======================================================
# This script generates tables for the Annual Report GAIN 2024.
# It includes:
# - A new table for Partnerships, based on PRO17 and disaggregated by year, region, and source.
# - The breakdown is structured using a function similar to SPSS CTABLES.
# Data and outputs are saved in designated folders.
# ======================================================

# Load required libraries
library(dplyr)
library(readxl)
library(writexl)
library(openxlsx)
library(tidyr)

# Load the group roster dataset
group_roster_file <- file.path(working_dir, "analysis_ready_group_roster.csv")
group_roster <- read.csv(group_roster_file)

# Function to create CTABLES-like breakdown for partnerships
generate_partnership_ctables <- function(data, year_var, region_var, source_vars, partnership_var) {
  data %>%
    filter(.data[[partnership_var]] == 1) %>%
    pivot_longer(
      cols = all_of(source_vars),
      names_to = "Source_Variable",
      values_to = "Value"
    ) %>%
    filter(Value == 1) %>%
    mutate(
      Source = case_when(
        Source_Variable == "PRO08.A" ~ "Survey",
        Source_Variable == "PRO08.B" ~ "Administrative",
        Source_Variable == "PRO08.C" ~ "Census",
        Source_Variable == "PRO08.D" ~ "Data Integration",
        TRUE ~ "Other"
      )
    ) %>%
    group_by(.data[[year_var]], .data[[region_var]], Source) %>%
    summarise(Total = n(), .groups = "drop") %>%
    pivot_wider(names_from = Source, values_from = Total, values_fill = 0) %>%
    arrange(.data[[year_var]], .data[[region_var]])
}

# Generate Partnerships breakdown table
partnerships_ctables <- generate_partnership_ctables(
  group_roster,
  year_var = "ryear",
  region_var = "region",
  source_vars = c("PRO08.A", "PRO08.B", "PRO08.C", "PRO08.D", "PRO08.E", "PRO08.F", "PRO08.G", "PRO08.H", "PRO08.X"),
  partnership_var = "PRO17"
)

# Load the existing Excel file
output_excel_file <- file.path(working_dir, "Annual Report GAIN 2024.xlsx")
wb <- loadWorkbook(output_excel_file)

# Add the Partnerships CTABLES to the "Annual Report" sheet
if (!"Annual Report" %in% names(wb)) {
  addWorksheet(wb, "Annual Report")
}
writeData(wb, "Annual Report", partnerships_ctables, startRow = 1, startCol = 1, withFilter = TRUE)

# Save the updated Excel file
saveWorkbook(wb, output_excel_file, overwrite = TRUE)

message("Partnerships CTABLES breakdown added to the Annual Report Excel file.")
# ======================================================
# This script generates a table for Institutional Implementation in the Annual Report GAIN 2024.
# It includes:
# - A new table for Institutional Implementation, based on g_conled = 2.
# - Breakdown by year (ryear), organization (morganization), and level (PRO03B).
# - Breakdown by data source (PRO08.*) and use of recommendations (PRO09).
# - The breakdown categorizes organizations as Global, Regional, or Country level.
# Data and outputs are saved in designated folders.
# ======================================================

# Load required libraries
library(dplyr)
library(readxl)
library(writexl)
library(openxlsx)
library(tidyr)

# Load the group roster dataset
group_roster_file <- file.path(working_dir, "analysis_ready_group_roster.csv")
group_roster <- read.csv(group_roster_file)

# Generate Institutional Implementation breakdown table
institutional_implementation_table <- group_roster %>%
  filter(g_conled == 2) %>%
  mutate(
    Organization_Level = case_when(
      PRO03B == "01" ~ "Global",
      PRO03B == "02" ~ "Regional",
      PRO03B == "03" ~ "Country",
      TRUE ~ "Unknown"
    ),
    Source = case_when(
      PRO08.A == 1 ~ "Survey",
      PRO08.B == 1 ~ "Administrative",
      PRO08.C == 1 ~ "Census",
      PRO08.D == 1 ~ "Data Integration",
      PRO08.E == 1 | PRO08.F == 1 | PRO08.G == 1 | PRO08.H == 1 | PRO08.X == 1 ~ "Other",
      TRUE ~ "Unknown"
    ),
    Use_of_Recommendations = case_when(
      PRO09 == 1 ~ "Yes",
      PRO09 == 2 ~ "No",
      PRO09 == 8 ~ "Don't Know",
      TRUE ~ "Unknown"
    )
  ) %>%
  group_by(ryear, morganization, Organization_Level, Source, Use_of_Recommendations) %>%
  summarise(Total_Examples = n(), .groups = "drop") %>%
  arrange(ryear, morganization, Organization_Level, Source, Use_of_Recommendations)

# Load the existing Excel file
output_excel_file <- file.path(working_dir, "Annual Report GAIN 2024.xlsx")
wb <- loadWorkbook(output_excel_file)

# Add the Institutional Implementation table to a new sheet
if (!"Institutional Implementation" %in% names(wb)) {
  addWorksheet(wb, "Institutional Implementation")
}
writeData(wb, "Institutional Implementation", institutional_implementation_table, startRow = 1, startCol = 1, withFilter = TRUE)

# Save the updated Excel file
saveWorkbook(wb, output_excel_file, overwrite = TRUE)

message("Institutional Implementation table with breakdowns added to the Annual Report Excel file.")
# ======================================================
# This script generates the Future Projects tables for the Annual Report GAIN 2024.
# - Three separate breakdowns: by total sources, by organization type, and by quarter.
# - The tables are combined into one output and exported to the same sheet.
# - Columns are always: Administrative Data, Census, Data Integration, Guidance/Toolkit, Non-Traditional, Other, Strategy, Survey, Workshop/Training.
# Data and outputs are saved in designated folders.
# ======================================================

# Load required libraries
library(dplyr)
library(readxl)
library(writexl)
library(openxlsx)
library(tidyr)

# Load the group roster dataset
group_roster2_file <- file.path(working_dir, "analysis_ready_group_roster2.csv")
group_roster2 <- read.csv(group_roster2_file)

# Ensure all relevant FPR05 columns are numeric before pivoting
fpr05_columns <- grep("^FPR05", names(group_roster2), value = TRUE)
group_roster2[fpr05_columns] <- lapply(group_roster2[fpr05_columns], as.numeric)

# Transform data to long format for processing
future_projects_long <- group_roster2 %>%
  pivot_longer(
    cols = all_of(fpr05_columns),
    names_to = "Source_Variable",
    values_to = "Value"
  ) %>%
  filter(Value == 1) %>%
  mutate(
    Source = case_when(
      grepl("SURVEY", Source_Variable) ~ "Survey",
      grepl("ADMINISTRATIVE.DATA", Source_Variable) ~ "Administrative Data",
      grepl("CENSUS", Source_Variable) ~ "Census",
      grepl("DATA.INTEGRATION", Source_Variable) ~ "Data Integration",
      grepl("NON.TRADITIONAL", Source_Variable) ~ "Non-Traditional",
      grepl("STRATEGY", Source_Variable) ~ "Strategy",
      grepl("GUIDANCE.TOOLKIT", Source_Variable) ~ "Guidance/Toolkit",
      grepl("H..WORKSHOP.TRAINING", Source_Variable) ~ "Workshop/Training",
      grepl("OTHER", Source_Variable) ~ "Other",
      TRUE ~ "Unknown"
    ),
    Organization_Type = case_when(
      LOC01 == "1" ~ "National Organization",
      LOC01 == "2" ~ "International Organization",
      LOC01 == "3" ~ "CSO",
      TRUE ~ "Unknown"
    )
  )

# Define fixed column order
fixed_columns <- c("Administrative Data", "Census", "Data Integration", "Guidance/Toolkit", "Non-Traditional", "Other", "Strategy", "Survey", "Workshop/Training")

# Generate total breakdown by source
future_projects_total <- future_projects_long %>%
  count(Source) %>%
  pivot_wider(names_from = Source, values_from = n, values_fill = 0) %>%
  select(all_of(fixed_columns)) %>%
  mutate(Category = "Total Projects by Source")

# Generate breakdown by organization type
future_projects_by_org <- future_projects_long %>%
  count(Organization_Type, Source) %>%
  pivot_wider(names_from = Source, values_from = n, values_fill = 0) %>%
  select(all_of(fixed_columns)) %>%
  mutate(Category = "Projects by Organization Type")

# Generate breakdown by quarter
future_projects_by_quarter <- group_roster2 %>%
  pivot_longer(
    cols = starts_with("FPR07"),
    names_to = "Quarter",
    values_to = "Value"
  ) %>%
  filter(Value == 1) %>%
  count(Quarter) %>%
  pivot_wider(names_from = Quarter, values_from = n, values_fill = 0) %>%
  mutate(Category = "Projects by Quarter")

# Combine all tables into one
combined_future_projects <- bind_rows(
  future_projects_total,
  future_projects_by_org,
  future_projects_by_quarter
)

# Load the existing Excel file
output_excel_file <- file.path(working_dir, "Annual Report GAIN 2024.xlsx")
wb <- loadWorkbook(output_excel_file)

# Add the combined Future Projects table to a new sheet
if (!"Future Projects" %in% names(wb)) {
  addWorksheet(wb, "Future Projects")
}
writeData(wb, "Future Projects", combined_future_projects, startRow = 1, startCol = 1, withFilter = TRUE)

# Save the updated Excel file
saveWorkbook(wb, output_excel_file, overwrite = TRUE)

message("Future Projects tables added to the Annual Report Excel file.")
# ======================================================
# R Script for Enhanced GAIN 2024 Annual Report (Word)
# ======================================================

# Load required libraries
library(dplyr)
library(flextable)
library(readr)
library(writexl)
library(officer)
library(tidyr)
library(ggplot2)
library(sf)
library(rnaturalearth)
library(rnaturalearthdata)
library(patchwork)  # For arranging plots

# EGRISS Color Scheme
primary_color <- "#4cc3c9"
secondary_color <- "#3b71b3"
accent_color <- "#072d62"
background_color <- "#f0f8ff"

# Load dataset
group_roster_file <- file.path(working_dir, "analysis_ready_group_roster.csv")
group_roster <- read.csv(group_roster_file)


# Function to create styled flextables
create_flextable <- function(data, title) {
  flextable(data) %>%
    theme_booktabs() %>%
    fontsize(size = 10) %>%
    bold(part = "header") %>%
    color(color = primary_color, part = "header") %>%
    bg(bg = background_color, part = "body") %>%
    border_outer(border = fp_border(color = accent_color, width = 2)) %>%
    border_inner_h(border = fp_border(color = secondary_color, width = 1)) %>%
    autofit() %>%
    add_footer_lines(values = "Source: GAIN 2024 Data") %>%
    set_caption(caption = title)
}

# ======================================================
# Summary of Country-Led Examples (Figure 6)
# ======================================================
summary_table <- group_roster %>%
  group_by(ryear, g_conled, PRO09) %>%
  summarise(count = n(), .groups = "drop") %>%
  pivot_wider(names_from = ryear, values_from = count, values_fill = 0) %>%
  arrange(g_conled, PRO09)

summary_table <- summary_table %>%
  mutate(
    `Example Lead` = case_when(
      g_conled == 1 ~ "Nationally Led Examples",
      g_conled == 2 ~ "Institutionally Led Examples",
      g_conled == 3 ~ "CSO Led Examples",
      g_conled == 8 ~ "Unknown",
      TRUE ~ ""
    ),
    `Use of Recommendations` = case_when(
      PRO09 == 1 ~ "Using EGRISS Recommendations",
      PRO09 == 2 ~ "Not Using EGRISS Recommendations",
      PRO09 == 8 ~ "Don't Know if EGRISS Recommendations Used",
      is.na(PRO09) ~ "Not reported Use of EGRISS Recommendations",
      TRUE ~ ""
    )
  ) %>%
  select(`Example Lead`, `Use of Recommendations`, everything(), -g_conled, -PRO09)

summary_table$`Example Lead` <- ifelse(duplicated(summary_table$`Example Lead`), "", summary_table$`Example Lead`)

figure6 <- create_flextable(summary_table, "Figure 6: Summary of Country-Led Examples")
# ======================================================
# Use of Recommendations (Figure 7) - Updated
# ======================================================
# Function to create styled flextables
create_flextable <- function(data, title) {
  flextable(data) %>%
    theme_booktabs() %>%
    fontsize(size = 10) %>%
    bold(part = "header") %>%
    color(color = primary_color, part = "header") %>%
    bg(bg = background_color, part = "body") %>%
    border_outer(border = fp_border(color = accent_color, width = 2)) %>%
    border_inner_h(border = fp_border(color = secondary_color, width = 1)) %>%
    autofit() %>%
    add_footer_lines(values = "Source: GAIN 2024 Data") %>%
    set_caption(caption = title)
}
# Convert PRO10.A, PRO10.B, PRO10.C, PRO10.Z, and PRO09 to numeric after cleaning non-numeric values
group_roster <- group_roster %>%
  mutate(
    PRO10.A = as.numeric(gsub("[^0-9]", "", PRO10.A)),
    PRO10.B = as.numeric(gsub("[^0-9]", "", PRO10.B)),
    PRO10.C = as.numeric(gsub("[^0-9]", "", PRO10.C)),
    PRO10.Z = as.numeric(gsub("[^0-9]", "", PRO10.Z)),
    PRO09 = as.numeric(gsub("[^0-9]", "", PRO09)),
    g_recuse = case_when(
      PRO10.A == 1 & PRO10.B != 1 & PRO10.C != 1 ~ "Use of IRRS Only",
      PRO10.A != 1 & PRO10.B == 1 & PRO10.C != 1 ~ "Use of IRIS Only",
      PRO10.A != 1 & PRO10.B != 1 & PRO10.C == 1 ~ "Use of IROSS Only",
      (PRO10.A + PRO10.B + PRO10.C) > 1 ~ "Any Combination of Recommendations",
      PRO10.Z == 1 ~ "Don't Know Which Recommendations Were Used",
      TRUE ~ "Not Reported Use of Recommendations"
    )
  ) 

# Use of Recommendations (Figure 7) - Updated
recuse_table <- group_roster %>%
  filter(PRO09 == 1) %>%
  group_by(g_conled, g_recuse, ryear) %>%
  summarise(Count = n(), .groups = "drop") %>%
  mutate(
    `Example Lead` = case_when(
      g_conled == 1 ~ "Nationally Led Examples",
      g_conled == 2 ~ "Institutionally Led Examples",
      g_conled == 3 ~ "CSO Led Examples",
      g_conled == 8 ~ "Unknown",
      TRUE ~ ""
    )
  ) %>%
  pivot_wider(names_from = ryear, values_from = Count, values_fill = 0) %>%
  mutate(Total = rowSums(across(`2021`:`2024`), na.rm = TRUE)) %>%
  select(`Example Lead`, `Use of Recommendations by Leads` = g_recuse, `2021`, `2022`, `2023`, `2024`, Total) # Rename column and recalculate total

# Ensure the year columns are numeric
recuse_table <- recuse_table %>%
  mutate(
    across(`2021`:`2024`, ~ as.numeric(.)),
    Total = rowSums(across(`2021`:`2024`), na.rm = TRUE)  # Recalculate total
  )

# Remove duplicated g_conled labels
recuse_table$`Example Lead` <- ifelse(duplicated(recuse_table$`Example Lead`), "", recuse_table$`Example Lead`)

# Create flextable with column highlights
figure7 <- flextable(recuse_table) %>%
  theme_booktabs() %>%
  bold(part = "header") %>%
  bg(bg = "#f4cccc", j = ~ `2024`) %>%   # Highlight 2024 column
  bg(bg = "#c9daf8", j = ~ Total) %>%   # Highlight Total column
  border_outer(border = fp_border(color = "black", width = 2)) %>%
  autofit() %>%
  add_footer_lines(values = "Source: GAIN 2024 Data") %>%
  set_caption(caption = "Figure 7: Use of Recommendations")
# ======================================================
# Figure 8 - Step 1: Aggregate PRO08 variables into specified categories and count each source by year
# ======================================================
# Step 1: Prepare the data
aggregated_data <- group_roster %>%
  filter(g_conled == 1) %>%  # Filter for g_conled == 1
  mutate(across(starts_with("PRO08."), as.integer)) %>%  # Convert all PRO08.* columns to integer
  pivot_longer(
    cols = starts_with("PRO08."),
    names_to = "Source_Variable",
    values_to = "Value"
  ) %>% 
  filter(Value == 1) %>% 
  mutate(
    Source = case_when(
      grepl("PRO08.A", Source_Variable) ~ "Survey",
      grepl("PRO08.B", Source_Variable) ~ "Administrative Data",
      grepl("PRO08.C", Source_Variable) ~ "Census",
      grepl("PRO08.D", Source_Variable) ~ "Data Integration",
      grepl("PRO08.E|PRO08.F|PRO08.G|PRO08.H|PRO08.X", Source_Variable) ~ "Other",
      TRUE ~ "Unknown"
    ),
    Recommendation_Status = case_when(
      PRO09 == 1 ~ "Using Recommendations",
      PRO09 %in% c(2, 8) ~ "Not Used Recommendations and Other",
      TRUE ~ "Not Used Recommendations and Other"
    )
  ) %>% 
  group_by(Recommendation_Status, Source, ryear) %>% 
  summarise(Count = n(), .groups = "drop") %>% 
  pivot_wider(
    names_from = ryear,
    values_from = Count,
    values_fill = 0
  )
# Step 2: Add Total Column
aggregated_data <- aggregated_data %>%
  mutate(
    Total = rowSums(select(., `2021`, `2022`, `2023`, `2024`), na.rm = TRUE)  # Correctly calculate the total
  )

# Step 3: Order rows and columns
aggregated_data <- aggregated_data %>%
  mutate(
    Recommendation_Status = factor(
      Recommendation_Status,
      levels = c("Using Recommendations", "Not Used Recommendations and Other")
    )
  ) %>%
  select(Recommendation_Status, Source, `2021`, `2022`, `2023`, `2024`, Total) %>%  # Ensure correct column order
  arrange(Recommendation_Status, factor(Source, levels = c("Survey", "Census", "Administrative Data", "Data Integration", "Other")))

# Step 4: Beautify and create FlexTable for Word
figure8_flextable <- flextable(aggregated_data) %>%
  theme_booktabs() %>%
  bold(part = "header") %>%
  merge_v(j = ~ Recommendation_Status) %>%  # Merge vertical cells for Recommendation_Status
  bg(bg = "#f4cccc", j = ~ `2024`) %>%   # Highlight the 2024 column
  bg(bg = "#c9daf8", j = ~ Total) %>%   # Highlight the Total column
  border_outer(border = fp_border(color = "black", width = 2)) %>%
  border_inner(border = fp_border(color = "gray", width = 0.5)) %>%
  autofit() %>%
  add_footer_lines(values = "Source: GAIN 2024 Data") %>%
  set_caption(caption = "Figure 8: Breakdown by Year, Use of Recommendations, and Source")

# ======================================================
# Regional Analysis (Text 1)
# ======================================================
regional_data <- group_roster %>%
  filter(PRO09 == 1, g_conled == 1) %>%
  group_by(region, ryear) %>%
  summarise(count = n(), .groups = "drop") %>%
  pivot_wider(names_from = ryear, values_from = count, values_fill = 0)

text1 <- create_flextable(regional_data, "Text 1: Regional Summary of Country-Led Projects")
# Create world map plot for the year 2024
world <- ne_countries(scale = "medium", returnclass = "sf")

# Filter data for 2024
year_data <- group_roster %>%
  filter(ryear == 2024, g_conled == 1) %>%
  group_by(mcountry) %>%
  summarise(Count = n(), .groups = "drop")

# Merge with world map data
year_data <- left_join(year_data, world, by = c("mcountry" = "name")) %>%
  filter(!is.na(geometry))  # Ensure geometries are valid

total_examples <- sum(year_data$Count, na.rm = TRUE)

# Create the map
map_plot <- ggplot() +
  geom_sf(data = world, fill = "gray90", color = "white") +
  geom_sf(data = year_data, aes(geometry = geometry, fill = Count), color = primary_color, alpha = 0.7, show.legend = FALSE) +
  scale_fill_continuous(low = "#e0f3db", high = "#43a2ca") +
  geom_text(data = year_data, aes(label = Count, geometry = geometry), stat = "sf_coordinates", size = 3, color = "black") +
  labs(title = paste("Examples in 2024 (Total:", total_examples, ")")) +
  theme_minimal()

# ======================================================
# Challenges Reported (Figure 9) - Transposed and with Labels
# ======================================================
challenge_labels <- c(
  "PRO20.A" = "NON-RESPONSE BIAS",
  "PRO20.B" = "SAMPLING ERRORS",
  "PRO20.C" = "IDENTIFICATION OF POPULATIONS",
  "PRO20.D" = "DATA CONFIDENTIALITY AND PRIVACY",
  "PRO20.E" = "RESOURCE CONSTRAINTS",
  "PRO20.F" = "POLITICAL ISSUES",
  "PRO20.G" = "SAFETY CONCERNS",
  "PRO20.H" = "TIMELINESS AND DATA QUALITY",
  "PRO20.I" = "LIMITED TECHNICAL CAPACITY",
  "PRO20.J" = "LACK OF ACCESSIBLE GUIDANCE",
  "PRO20.X" = "Other"
)

challenges_data <- group_roster %>%
  filter(ryear %in% c(2023, 2024), g_conled == 1, PRO09 == 1) %>%
  select(ryear, starts_with("PRO20.")) %>%
  pivot_longer(cols = starts_with("PRO20."), names_to = "Challenge", values_to = "Reported") %>%
  filter(Reported == 1) %>%
  mutate(Challenge = recode(Challenge, !!!challenge_labels)) %>%
  group_by(Challenge, ryear) %>%
  summarise(Count = n(), .groups = "drop") %>%
  pivot_wider(names_from = ryear, values_from = Count, values_fill = 0)

figure9 <- create_flextable(challenges_data, "Figure 9: Challenges Reported")
# ======================================================
# Generate Institutional Implementation breakdown table
# ======================================================
institutional_implementation_table <- group_roster %>%
  filter(g_conled == 2) %>%
  mutate(
    Source = case_when(
      PRO08.A == 1 ~ "Survey",
      PRO08.B == 1 ~ "Administrative Data",
      PRO08.C == 1 ~ "Census",
      PRO08.D == 1 ~ "Data Integration",
      PRO08.E == 1 | PRO08.F == 1 | PRO08.G == 1 | PRO08.H == 1 | PRO08.X == 1 ~ "Other",
      TRUE ~ "Unknown"
    ),
    Use_of_Recommendations = case_when(
      PRO09 == 1 ~ "Using Recommendations",
      PRO09 == 2 ~ "Not Using Recommendations",
      PRO09 == 8 ~ "Don't Know",
      TRUE ~ "Unknown"
    )
  ) %>%
  group_by(Use_of_Recommendations, Source, ryear) %>%
  summarise(Total_Examples = n(), .groups = "drop") %>%
  pivot_wider(names_from = ryear, values_from = Total_Examples, values_fill = 0) %>%
  rowwise() %>%
  mutate(Total = sum(c_across(`2021`:`2024`), na.rm = TRUE)) %>%
  ungroup() %>%
  arrange(
    factor(Use_of_Recommendations, levels = c("Using Recommendations", "Not Using Recommendations", "Don't Know", "Unknown")),
    factor(Source, levels = c("Survey", "Census", "Administrative Data", "Other"))
  ) %>%
  select(Use_of_Recommendations, Source, `2021`, `2022`, `2023`, `2024`, Total)  # Ensure correct column order

# Beautify and create FlexTable for Word
institutional_flextable <- flextable(institutional_implementation_table) %>%
  theme_booktabs() %>%
  bold(part = "header") %>%
  bg(bg = "#f4cccc", j = ~ `2024`) %>%   # Highlight the 2024 column
  bg(bg = "#c9daf8", j = ~ Total) %>%   # Highlight the Total column
  merge_v(j = ~ Use_of_Recommendations) %>%  # Merge vertical cells for Use_of_Recommendations
  merge_v(j = ~ Source) %>%  # Merge vertical cells for Source
  border_outer(border = fp_border(color = "black", width = 2)) %>%
  border_inner(border = fp_border(color = "gray", width = 0.5)) %>%
  autofit() %>%
  add_footer_lines(values = "Source: GAIN 2024 Data") %>%
  set_caption(caption = "Institutional Implementation Breakdown")
# ======================================================
# Add Future Projects 
# ======================================================

sapply(group_roster2[, fpr05_columns], class)
group_roster2 <- group_roster2 %>%
  mutate(across(all_of(fpr05_columns), ~ as.numeric(.)))

# Step 1: Create the Source column and count the number of times each source is used
source_summary <- group_roster2 %>%
  pivot_longer(
    cols = all_of(fpr05_columns),
    names_to = "Source_Variable",
    values_to = "Value"
  ) %>%
  filter(Value == 1) %>%
  mutate(
    Source = case_when(
      grepl("SURVEY", Source_Variable) ~ "Survey",
      grepl("ADMINISTRATIVE.DATA", Source_Variable) ~ "Administrative Data",
      grepl("CENSUS", Source_Variable) ~ "Census",
      grepl("DATA.INTEGRATION", Source_Variable) ~ "Data Integration",
      grepl("NON.TRADITIONAL", Source_Variable) ~ "Non-Traditional",
      grepl("STRATEGY", Source_Variable) ~ "Strategy",
      grepl("GUIDANCE.TOOLKIT", Source_Variable) ~ "Guidance/Toolkit",
      grepl("H..WORKSHOP.TRAINING", Source_Variable) ~ "Workshop/Training",
      grepl("OTHER", Source_Variable) ~ "Other",
      TRUE ~ "Unknown"
    )
  ) %>%
  count(Source) %>%
  rename(Count = n) %>%
  bind_rows(tibble(Source = "Total", Count = sum(.$Count)))


# Create a FlexTable for Word
source_summary_flextable <- flextable(source_summary) %>%
  theme_booktabs() %>%
  bold(part = "header") %>%
  bg(bg = "#c9daf8", j = ~ Count) %>%  # Highlight the Count column
  border_outer(border = fp_border(color = "black", width = 2)) %>%
  border_inner(border = fp_border(color = "gray", width = 0.5)) %>%
  autofit() %>%
  add_footer_lines(values = "Source: GAIN 2024 Data") %>%
  set_caption(caption = "Future Projects Breakdown by Source for 2024")
# ======================================================
# Unique Country Count for Use of Recommendations (PRO09 == 1) by Leadership Type
# ======================================================

# Function to calculate unique country counts by `g_conled`
calculate_unique_country_count <- function(group_roster, leadership_type) {
  group_roster %>%
    filter(PRO09 == 1, g_conled == leadership_type) %>%  # Filter for Use of Recommendations and Leadership Type
    group_by(region, ryear) %>%
    summarise(unique_countries = n_distinct(mcountry), .groups = "drop") %>%
    pivot_wider(names_from = ryear, values_from = unique_countries, values_fill = 0) %>%
    mutate(Total = rowSums(across(`2021`:`2024`), na.rm = TRUE)) %>%
    mutate(Leadership = if_else(leadership_type == 1, "Nationally Led", "Institutionally Led"))
}

# Calculate unique country counts for nationally and institutionally led examples
nationally_led_count <- calculate_unique_country_count(group_roster, 1)
institutionally_led_count <- calculate_unique_country_count(group_roster, 2)

# Combine both tables
combined_unique_country_count <- bind_rows(nationally_led_count, institutionally_led_count)

# Add a summary row for total unique countries across all regions
total_unique_summary <- group_roster %>%
  filter(PRO09 == 1) %>%
  group_by(ryear) %>%
  summarise(unique_countries = n_distinct(mcountry), .groups = "drop") %>%
  pivot_wider(names_from = ryear, values_from = unique_countries, values_fill = 0) %>%
  mutate(Total = rowSums(across(`2021`:`2024`), na.rm = TRUE)) %>%
  mutate(region = "Total Unique Countries", Leadership = "Total")

# Final table with combined counts and summary
final_unique_country_table <- bind_rows(combined_unique_country_count, total_unique_summary)

# Beautify and create FlexTable for Word
unique_country_flextable <- flextable(final_unique_country_table) %>%
  theme_booktabs() %>%
  bold(part = "header") %>%
  bg(bg = "#f4cccc", j = ~ `2024`) %>%   # Highlight the 2024 column
  bg(bg = "#c9daf8", j = ~ Total) %>%   # Highlight the Total column
  merge_v(j = ~ Leadership) %>%  # Merge Leadership column for repeated values
  border_outer(border = fp_border(color = "black", width = 2)) %>%
  border_inner(border = fp_border(color = "gray", width = 0.5)) %>%
  autofit() %>%
  add_footer_lines(values = "Source: GAIN 2024 Data") %>%
  set_caption(caption = "Unique Country Count by Leadership Type, Region, and Year for Use of Recommendations (PRO09 == 1)")
# ======================================================
# Load and process PRO11/PRO12 variables
# ======================================================
library(dplyr)
library(tidyr)
library(flextable)
library(officer)

# Step 1: Load the dataset
file_path <- file.path(working_dir, "analysis_ready_repeat_PRO11_PRO12.csv")
repeat_data <- read.csv(file_path)

# Step 2: Rename `_recommendation` to `recommendation`
repeat_data <- repeat_data %>%
  rename(recommendation = X_recommendation)

# Step 3: Convert to long format, classify categories, and aggregate
processed_data <- repeat_data %>%
  pivot_longer(
    cols = starts_with("PRO12"),
    names_to = "Category_Variable",
    values_to = "Value"
  ) %>%
  filter(Value == 1) %>%  # Filter where Value is 1
  mutate(
    Category = case_when(
      Category_Variable == "PRO12A" ~ "Statistical framework/population group",
      Category_Variable == "PRO12B" ~ "Recommendations on data sources",
      Category_Variable == "PRO12C" ~ "Coordination",
      Category_Variable == "PRO12D" ~ "Data sharing",
      Category_Variable == "PRO12E" ~ "Analysis",
      Category_Variable == "PRO12F" ~ "Indicator selection",
      Category_Variable == "PRO12G" ~ "Data integration",
      Category_Variable == "PRO12H" ~ "Dissemination",
      Category_Variable == "PRO12I" ~ "Institutional or sectoral strategy",
      Category_Variable == "PRO12J" ~ "Other (specify)",
      TRUE ~ NA_character_
    )
  ) %>%
  filter(!is.na(Category))  # Remove rows with NA in Category

# Step 4: Create separate tables for nationally led and institutionally led examples
nationally_led <- processed_data %>%
  filter(g_conled == 1) %>%
  group_by(Category, recommendation) %>%
  summarise(Count = n(), .groups = "drop") %>%
  pivot_wider(names_from = recommendation, values_from = Count, values_fill = 0)

institutionally_led <- processed_data %>%
  filter(g_conled == 2) %>%
  group_by(Category, recommendation) %>%
  summarise(Count = n(), .groups = "drop") %>%
  pivot_wider(names_from = recommendation, values_from = Count, values_fill = 0)

# Function to create FlexTable with multi-level header
create_flextable <- function(data, table_title) {
  
  # Get unique recommendations
  unique_recommendations <- unique(processed_data$recommendation)
  num_recommendations <- length(unique_recommendations)
  num_columns <- ncol(data)
  
  # Create header labels
  header1 <- c("Category", rep(table_title, num_recommendations))
  header2 <- c("Category", unique_recommendations)
  
  # Ensure headers match the number of columns
  while (length(header1) < num_columns) {
    header1 <- c(header1, "")
  }
  while (length(header2) < num_columns) {
    header2 <- c(header2, "")
  }
  
  # Create header mapping
  header_mapping <- data.frame(
    col_keys = colnames(data),
    header1 = header1,
    header2 = header2,
    stringsAsFactors = FALSE
  )
  
  # Create flextable
  flextable(data) %>%
    set_header_df(mapping = header_mapping) %>%
    merge_h(part = "header") %>%
    theme_booktabs() %>%
    bold(part = "header") %>%
    autofit() %>%
    bg(bg = "#f4cccc", i = nrow(data)) %>%  # Highlight the Total row
    set_caption(table_title)
}

# Step 5: Create separate tables for report
text2_flextable_nationally_led <- create_flextable(nationally_led, "Nationally Led Examples")
text2_flextable_institutionally_led <- create_flextable(institutionally_led, "Institutionally Led Examples")

# Display both tables on the same page
text2_flextable_nationally_led
text2_flextable_institutionally_led
# ======================================================
# Breakdown of Nationally Led Partnerships
# ======================================================

library(dplyr)
library(tidyr)
library(flextable)

# EGRISS Color Scheme
primary_color <- "#4cc3c9"
secondary_color <- "#3b71b3"
accent_color <- "#072d62"
background_color <- "#f0f8ff"

# Load dataset
file_path <- file.path(working_dir, "analysis_ready_group_roster.csv")
data <- read.csv(file_path)

# Define Ordered Partnership Type Labels
partnership_labels <- c(
  "PRO18.A" = "National Partnership",
  "PRO18.B" = "International Organization Partnership",
  "PRO18.C" = "Academia Partnership"
)

# Define Year Order
year_order <- c("2021", "2022", "2023", "2024")

# Count total nationally led projects
nationally_led_count <- data %>%
  filter(g_conled == 1) %>%
  count(ryear) %>%
  pivot_wider(names_from = ryear, values_from = n, values_fill = 0) %>%
  mutate(Partnership_Type = "Total Nationally Led Projects")

# Count total nationally led projects with partnerships
partnership_count <- data %>%
  filter(g_conled == 1, PRO17 == 1) %>%
  count(ryear) %>%
  pivot_wider(names_from = ryear, values_from = n, values_fill = 0) %>%
  mutate(Partnership_Type = "Total Nationally Led Projects with Partnerships")

# Filter for PRO17 == 1 and g_conled == 1
partnership_data <- data %>%
  filter(g_conled == 1, PRO17 == 1) %>%  # Only nationally led projects with partnerships
  select(ryear, PRO18.A, PRO18.B, PRO18.C) %>%  # Keep necessary columns
  mutate(ryear = as.character(ryear)) %>%  # Ensure ryear is treated as character
  pivot_longer(cols = starts_with("PRO18"), names_to = "Partnership_Type", values_to = "Value") %>%
  mutate(Partnership_Type = recode(Partnership_Type, !!!partnership_labels)) %>%  # Apply Partnership Labels
  filter(Value == 1) %>%  # Keep only rows where partnership exists (Value == 1)
  count(Partnership_Type, ryear) %>%  # Count occurrences per year
  pivot_wider(names_from = ryear, values_from = n, values_fill = 0)  # Convert to wide format

# Combine total count with detailed breakdown
partnership_data <- bind_rows(nationally_led_count, partnership_count, partnership_data)

# Ensure Year Order in Columns
partnership_data <- partnership_data %>%
  select(Partnership_Type, all_of(year_order))

# Create FlexTable with EGRISS Color Scheme
partnership_flextable <- flextable(partnership_data) %>%
  theme_booktabs() %>%
  bold(part = "header") %>%
  set_table_properties(width = 1, layout = "autofit") %>%
  bg(bg = background_color, part = "body") %>%  # Apply background color
  bg(bg = primary_color, part = "header") %>%  # Apply primary color to header
  color(color = "white", part = "header") %>%  # Set header text color to white
  bold(j = 1, part = "body") %>%  # Bold the first column (Partnership Type)
  border_remove() %>%
  border_outer(part = "all", border = fp_border(color = accent_color, width = 1.5)) %>%
  border_inner_h(part = "body", border = fp_border(color = secondary_color, width = 1)) %>%
  set_caption("Breakdown of Nationally Led Partnerships by Year and Type") %>%
  autofit()

# Display Table in RStudio Viewer (for verification)
partnership_flextable


# ======================================================
# Create Word Document
# ======================================================

word_doc <- read_docx()

# Add content to Word document
word_doc <- word_doc %>%
  body_add_par("GAIN 2024 Annual Report", style = "heading 1") %>%
  body_add_flextable(figure6) %>%
  body_add_break() %>%
  body_add_flextable(figure7) %>%
  body_add_break() %>%
  body_add_par("Figure 8: Breakdown by Year, Use of Recommendations, and Source", style = "heading 2") %>%
  body_add_flextable(figure8_flextable) %>%
  body_add_break() %>%
  body_add_flextable(text1) %>%
  body_add_break() %>%
  body_add_par("Breakdown by Category and Region for PRO11/PRO12 Data", style = "heading 2") %>%
  body_add_flextable(text2_flextable_nationally_led) %>%  # Nationally Led Examples
  body_add_break() %>%
  body_add_flextable(text2_flextable_institutionally_led) %>%  # Institutionally Led Examples
  body_add_break() %>%
  body_add_par("Unique Country Count by Region and Year", style = "heading 2") %>%
  body_add_flextable(unique_country_flextable) %>%  
  body_add_break() %>%
  body_add_par("Map of Examples (2024)", style = "heading 2") %>%
  body_add_gg(map_plot, width = 8, height = 6.4) %>%
  body_end_section_landscape() %>%
  body_add_break() %>%
  body_add_flextable(figure9) %>%
  body_add_break() %>%
  body_add_par("Institutional Implementation Breakdown", style = "heading 2") %>%
  body_add_flextable(institutional_flextable) %>%
  body_add_break() %>%
  body_add_par("Future Projects Breakdown by Source for 2024", style = "heading 2") %>%
  body_add_flextable(source_summary_flextable) %>%
  body_add_break() %>%
  body_add_par("Breakdown of Nationally Led Partnerships by Year and Type", style = "heading 2") %>%
  body_add_flextable(partnership_flextable) %>%
  body_add_break()

# ======================================================
# Save the Word Document
# ======================================================

# Get current date in YYYY-MM-DD format
current_date <- format(Sys.Date(), "%Y-%m-%d")

# Define relative output file path with date
word_output_file <- file.path(working_dir, paste0("Annual_Report_GAIN_2024_", current_date, ".docx"))

# Save the Word document
print(word_doc, target = word_output_file)

# Display message confirming successful save
message("Updated GAIN 2024 Annual Report saved successfully at: ", word_output_file)
