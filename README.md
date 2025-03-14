Overview
This project processes and analyzes photovoltaic (PV) production and power consumption data. It fetches data from the Leneda API (https://leneda.eu/en/), processes it, and generates monthly summaries and invoices. The project is written in the context of Energiekooperativ Uelzechtdall (ecud.lu), a local renewable energy cooperative in Luxembourg.

Disclaimer
The code was written to be easy to understand and be easily adapted by basic Python users. It does not correspond to the standards I would have towards my code in a professional context. This project is licensed under Apache License 2.0.

Project Structure
Leneda API_v3.ipynb: Jupyter notebook for data processing. utils.py: Utility functions for data fetching, processing, and Excel formatting.

0_Raw_production_data_per_installation: Directory for raw production data files. 1_Monthly_summaries: Directory for monthly summary files. 2_Invoices: Directory for generated invoices.

0_credentials.json: Contains API credentials. 1_Installations.xlsx: Contains information about PV installations.

Setup
Install Dependencies: pip install requests pandas reportlab openpyxl numpy

API Credentials: Ensure 0_credentials.json contains valid API credentials and is stored only on your local machne.

Installations Data: Ensure 1_Installations.xlsx contains the necessary information about PV installations.

Usage
Load and Process Data: Open Leneda API_v3.ipynb in Jupyter Notebook and run the cells to fetch and process data.

Generate Monthly Summaries: The notebook will generate monthly summaries and save them in the 1_Monthly_summaries directory.

Generate Invoices: The notebook will generate invoices for PV installations with autoconsumption and save them in the 2_Invoices directory.

Functions
utils.py fetch_data(HEADERS, pod, obis_code, start_date, end_date): Fetch data from the API. calculate_monthly_summaries(df_site, autoconsumption_price): Calculate monthly summaries. process_api_data(HEADERS, dict_inst, start_date, end_date): Process API data into DataFrames. apply_excel_formatting(output_file, site_name, pod, capacity, is_injection=False): Apply formatting to Excel files. aggregate_dataframe(df, agg_rules, group_period): Aggregate DataFrame based on rules. process_sheet(site_name, site_info, df_site_month, sheet_type="autoconsumption"): Process a given sheet. format_monthly_data(df): Format monthly data. generate_invoice_for_site(key_pod, df, idx): Generate a PDF invoice.