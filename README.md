# GAIN 2024 Annual Report Repository

This repository contains scripts, data, and outputs for the **GAIN 2024 Annual Report**. The R scripts generate summary tables, visualizations, and formatted reports to support data analysis and reporting.

## Repository Structure
```
GAIN_2024_Annual_Report/
│
├── README.md                  # Description of the project and instructions
├── LICENSE                    # Open-source license referencing GAIN and EGRISS
├── .gitignore                 # Files and directories to exclude from the repository
├── scripts/                   # R scripts for data analysis and reporting
│   ├── analysis.R             # Main data analysis script
│   └── generate_report.R      # Script to generate the annual report
├── data/                      # Data files required for analysis
│   └── analysis_ready_group_roster.csv
├── output/                    # Generated output files
│   ├── Annual_Report_GAIN_2024.xlsx
│   └── Annual_Report_GAIN_2024_Updated.docx
└── figures/                   # Figures and visualizations
    └── World_Map_2024.png
```

## Prerequisites
- R (version 4.0 or higher)
- R packages: `dplyr`, `tidyr`, `writexl`, `readxl`, `ggplot2`, `sf`, `rnaturalearth`, `rnaturalearthdata`, `openxlsx`, `officer`, `flextable`

## How to Run the Scripts
1. Clone the repository:
   ```bash
   git clone https://github.com/username/GAIN_2024_Annual_Report.git
   cd GAIN_2024_Annual_Report
   ```
2. Install the required R packages:
   ```R
   install.packages(c("dplyr", "tidyr", "writexl", "readxl", "ggplot2", "sf", "rnaturalearth", "rnaturalearthdata", "openxlsx", "officer", "flextable"))
   ```
3. Run the analysis script:
   ```R
   source("scripts/analysis.R")
   ```
4. Generate the annual report:
   ```R
   source("scripts/generate_report.R")
   ```

## Output
- **Annual_Report_GAIN_2024.xlsx**: Contains summary tables and figures.
- **Annual_Report_GAIN_2024_Updated.docx**: Word document with formatted tables, visualizations, and captions.

## License
This project is open source and must reference **GAIN** and **EGRISS** in any derived work. See the [LICENSE](LICENSE) file for details.

## Contact
For questions or collaboration opportunities, please reach out to [Filip Mitrovic](mailto:mitrovif@unhcr.org).
