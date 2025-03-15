# **Leneda PV Data Processor**

## **Overview**
This project processes and analyzes **photovoltaic (PV) production** and **power consumption data**. It fetches data from the [Leneda API](https://leneda.eu/en/), processes it, and generates **monthly summaries** and **invoices**.

This project was developed in the context of **Energiekooperativ Uelzechtdall (ECUD)**, a **local renewable energy cooperative in Luxembourg** ([ecud.lu](https://ecud.lu)).

---

## **⚠ Disclaimer**
This code is designed to be **easy to understand and adapt** for basic Python users. It does not reflect the **professional standards** I would apply in a corporate setting.

> **License:** This project is licensed under the [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0).

---

## **📁 Project Structure**

```
📂 Project Root
├── 📜 Leneda_API_v3.ipynb       # Jupyter notebook for data processing
├── 📜 utils.py                  # Utility functions (data fetching, processing, formatting)
│
├── 📂 0_Raw_production_data_per_installation  # Raw PV production data files
├── 📂 1_Monthly_summaries                      # Monthly summary reports
├── 📂 2_Invoices                               # Generated invoices
│
├── 📜 0_credentials.json       # API credentials (NOT to be pushed to GitHub)
├── 📜 1_Installations.xlsx      # PV installation information
```

---

## **🔧 Setup**

### **1️⃣ Install Dependencies**
Ensure you have Python installed, then install the required libraries:
```sh
pip install requests pandas reportlab openpyxl numpy
```

### **2️⃣ API Credentials**
- Ensure `0_credentials.json` contains valid API credentials.
- **Security Tip:** Keep this file **only on your local machine** and add it to `.gitignore`.

### **3️⃣ Installations Data**
- Ensure `1_Installations.xlsx` contains necessary PV installation details.

---

## **🚀 Usage**

### **1️⃣ Load and Process Data**
Open **`Leneda_API_v3.ipynb`** in Jupyter Notebook and run the cells to:
- Fetch data from the **Leneda API**.
- Process **PV production & consumption data**.

### **2️⃣ Generate Monthly Summaries**
- Summaries are automatically **generated and saved** in `1_Monthly_summaries/`.

### **3️⃣ Generate Invoices**
- Invoices for PV installations with **autoconsumption** are generated and saved in `2_Invoices/`.

---

## **🛠 Functions (utils.py)**

### **📥 Data Processing**
- `fetch_data(HEADERS, pod, obis_code, start_date, end_date)`: Fetches data from the API.
- `process_api_data(HEADERS, dict_inst, start_date, end_date)`: Converts API data into **Pandas DataFrames**.
- `calculate_monthly_summaries(df_site, autoconsumption_price)`: Computes **monthly summaries**.

### **📊 Data Aggregation & Formatting**
- `aggregate_dataframe(df, agg_rules, group_period)`: Aggregates data based on rules.
- `format_monthly_data(df)`: Formats monthly data for reports.

### **📜 Excel & Invoice Generation**
- `apply_excel_formatting(output_file, site_name, pod, capacity, is_injection=False)`: Formats Excel reports.
- `process_sheet(site_name, site_info, df_site_month, sheet_type="autoconsumption")`: Processes individual sheets.
- `generate_invoice_for_site(key_pod, df, idx)`: Creates **PDF invoices**.

---

## **📌 Notes**
- This project is designed for **Energiekooperativ Uelzechtdall (ECUD)** but can be adapted for **other PV energy cooperatives**.
- Future improvements may include **better error handling, API enhancements, and automation**.

---

### **💡 Contributing**
Feel free to contribute! Open a pull request or issue if you find a bug or have a suggestion.
