# Python Financial Model — 3-Statement Forecast, DCF & Sensitivity Analysis

A Python-based financial modeling project that creates a historical financial model, 5-year financial forecast, DCF valuation, and WACC / terminal-growth sensitivity analysis.

The model is built using **Python, pandas, and openpyxl**, with Excel used as the final output and presentation format.

The project is designed as a **portfolio and learning project**, demonstrating how financial statements, forecasting assumptions, valuation mechanics, and Excel reporting can be implemented programmatically.

> **Disclaimer:** This is an educational and portfolio project. The data is illustrative/fictitious and the assumptions are not intended to represent an actual company's financial outlook, investment recommendation, or professional valuation. The model demonstrates financial modeling methodology and Python implementation rather than producing an investment-grade valuation.

---

# Project Structure

The model follows a four-stage architecture:

```text
A → B → C → D

A = Raw Financial Data
B = Historical Financial Model
C = Forecast Financial Model
D = DCF Valuation
```

### A — Raw Data

The original historical financial dataset covering:

**2021–2025**

The raw dataset contains the historical financial inputs used to build the model.

---

### B — Historical Model

A separate historical DataFrame is created from the raw data.

The Historical model contains the original historical information plus calculated financial metrics, including:

* Cash
* EBIT
* EBT
* Tax Rate
* Operating Cash Flow
* Investing Cash Flow
* Financing Cash Flow
* Change in Cash
* Total Assets
* Total Debt
* Interest Rate
* Total Liabilities
* Total Liabilities & Equity
* Balance Check

It also contains PP&E and Equity validation checks.

### Cash

Cash is generated in the Historical model rather than being taken from the raw dataset.

For 2021, the model directly assumes:

```text
2021 Ending Cash = 1,100
```

Cash is then rolled forward from 2021 onward using:

```text
Ending Cash (t) =
Ending Cash (t−1) + Change in Cash (t)
```

Because the historical model begins in 2021 and does not include a prior-year cash balance, the model directly assumes 2021 ending cash of 1,100.

---

### C — Forecast Model

The forecast extends the historical model from:

```text
2026–2030
```

The forecast is formula-driven and uses historical financial relationships, recent growth, and longer-term historical trends.

---

### D — DCF Valuation

The DCF uses the forecast financial statements to calculate:

* Net Working Capital
* Change in Net Working Capital
* NOPAT
* Unlevered Free Cash Flow
* Present Value of UFCF
* Terminal Value
* Present Value of Terminal Value
* Enterprise Value
* Equity Value
* Implied Share Price

A WACC / terminal-growth sensitivity analysis is also included.

---

# Files

## `financial_model.py`

The main Python script.

It:

1. Reads the raw Excel data.
2. Creates the Historical model.
3. Creates the Forecast model.
4. Creates the DCF valuation.
5. Exports the results to Excel.
6. Applies Excel formatting.

---

## `raw_financials.xlsx`

The historical input dataset used by the Python script.

---

## `financial_model.xlsx`

The generated Excel workbook containing the completed model and its outputs.

The workbook contains:

* `Historical`
* `Forecast`
* `DCF`

---

# Before Running the Python Script

The script contains two variables near the beginning:

```python
RAW_FILE_PATH = "ENTER YOUR FILE PATH HERE"

OUTPUT_FILE_PATH = "ENTER YOUR OUTPUT FILE PATH HERE"
```

These need to be changed to the location of the files on your own computer.

For example:

```python
RAW_FILE_PATH = "../Data/raw_financials.xlsx"
OUTPUT_FILE_PATH = "../Excel/financial_model.xlsx"
```

These paths are examples only. Each user should replace them with paths appropriate to their own setup.

The folder specified in `OUTPUT_FILE_PATH` must already exist; the output Excel file itself will be created by the script.

---

# Forecast Methodology

The forecast covers:

```text
2026
2027
2028
2029
2030
```

The forecast begins from the latest historical year, 2025.

The model intentionally combines several historical signals rather than relying on one arbitrary growth assumption.

## Growth Forecasting

For the growth-based forecast assumptions, the model combines:

* Recent historical growth
* Three-year CAGR
* Historical average growth

The weighting is:

```text
50% × Recent Momentum
30% × 3-Year CAGR
20% × Historical Average Growth
```

Therefore:

```text
Forecast Growth Rate =
50% × Recent Growth
+ 30% × 3-Year CAGR
+ 20% × Historical Average Growth
```

The forecast value is then calculated as:

```text
Forecast Value =
Previous Year Value × (1 + Forecast Growth Rate)
```

The use of the three-year CAGR provides a longer-term historical growth perspective, while recent momentum captures the latest direction of the business. The historical average provides an additional longer-term reference point.

This methodology is used for Revenue and Total Debt.

---

# Historical Ratio-Based Forecasting

For forecast assumptions based on historical financial ratios, the model uses:

```text
50% × Historical Average Ratio
+
50% × Latest Historical Ratio
```

Therefore:

```text
Forecast Ratio =
50% × Historical Average Ratio
+ 50% × Latest Historical Ratio
```

The resulting ratio is then applied to the appropriate forecast driver.

The purpose is to balance the longer-term historical relationship with the most recent observed relationship.

---

# Forecast Roll-Forwards

Several forecast items are explicitly rolled forward from the previous year's balance.

## Cash

```text
Ending Cash (t) =
Ending Cash (t−1) + Change in Cash (t)
```

Where:

```text
Change in Cash =
Operating Cash Flow
+ Investing Cash Flow
+ Financing Cash Flow
```

---

## PP&E

```text
Ending PP&E =
Beginning PP&E
+ Capital Expenditure
− D&A
```

The next year's beginning PP&E is then:

```text
Beginning PP&E (t) =
Ending PP&E (t−1)
```

This creates a continuous PP&E roll-forward throughout the forecast period.

---

## Equity

Ending equity is rolled forward using:

```text
Ending Equity =
Beginning Equity
+ Net Income
− Dividends Paid
+ New Equity Issued
```

The next year's beginning equity is:

```text
Beginning Equity (t) =
Ending Equity (t−1)
```

Forecast new equity issuance is assumed to be zero.


---

# Interest Expense

Interest expense is linked directly to forecast debt.

The historical interest rate is calculated as:

```text
Interest Rate =
Interest Expense / Total Debt
```

The **final historical interest rate** is then used as the forecast assumption.

Therefore:

```text
Forecast Interest Expense =
Forecast Total Debt × Final Historical Interest Rate
```

The forecast **Interest Rate is therefore the same percentage in every forecast year**.

The amount of Interest Expense can still change from year to year because forecast debt changes.

In other words:

```text
Interest Rate
2026 = Final Historical Rate
2027 = Final Historical Rate
2028 = Final Historical Rate
2029 = Final Historical Rate
2030 = Final Historical Rate
```

while:

```text
Interest Expense =
Forecast Debt × Interest Rate
```

changes with the debt balance.

---

# Tax Expense

Historical tax rate is calculated as:

```text
Tax Rate =
Tax Expense / EBT
```

The **final historical tax rate** is then used as the forecast assumption.

Therefore:

```text
Forecast Tax Expense =
Forecast EBT × Final Historical Tax Rate
```

The forecast **Tax Rate is therefore the same percentage in every forecast year**.

The amount of Tax Expense changes because forecast EBT changes.

Therefore:

```text
Tax Rate
2026 = Final Historical Rate
2027 = Final Historical Rate
2028 = Final Historical Rate
2029 = Final Historical Rate
2030 = Final Historical Rate
```

while:

```text
Tax Expense =
Forecast EBT × Tax Rate
```

changes with forecast EBT.

---

# Dividends

Forecast dividends are assumed to equal:

```text
80% × Net Income
```

Therefore:

```text
Forecast Dividends =
80% × Forecast Net Income
```

---

# Validation Checks

The model includes internal checks for important roll-forwards and statement consistency.

## PP&E Check

```text
PP&E Check =
Beginning PP&E
+ Capital Expenditure
− D&A
− Ending PP&E
```

Expected result:

```text
0
```

---

## Equity Check

```text
Equity Check =
Beginning Equity
+ Net Income
− Dividends Paid
+ New Equity Issued
− Ending Equity
```

Expected result:

```text
0
```

---

## Balance Sheet Check

```text
Balance Check =
Total Assets
− Total Liabilities & Equity
```

Expected result:

```text
0
```

Small floating-point differences can occur internally in Python because of numerical precision. The model rounds the final balance check so that immaterial floating-point residuals do not appear as model errors.

---

# DCF Sensitivity Analysis

The DCF includes a two-variable sensitivity analysis using:

### WACC

```text
9.98%
10.98%
11.98%
```

### Terminal Growth

```text
2.00%
3.00%
4.00%
```

The base case is:

```text
WACC = 10.98%
Terminal Growth = 3.00%
```

The sensitivity table shows how the implied valuation and share price change across different combinations of WACC and terminal growth.

The sensitivity WACC values are scenario cases around the calculated base WACC; they are not separately recalculated WACCs.

---

# Excel Formatting

The Python script uses `openpyxl` to format the generated workbook after the financial calculations have been completed.

The formatting is intended to make the model easier to read and review while keeping the calculations themselves in Python.

---

# Project Workflow

```text
Raw Financial Data
        ↓
Historical Model
        ↓
Forecast Model
        ↓
DCF Valuation
        ↓
Sensitivity Analysis
        ↓
Excel Output
```

The Python script is the primary source of the model's calculations and logic, while the generated Excel workbook provides the final model output and presentation.

---

# Important Considerations

This model is intentionally simplified compared with professional finance models.

Before using the framework with real company data, users should independently review whether the assumptions and methodologies are appropriate for the company being analyzed.

A mechanically consistent model does not necessarily represent an economically realistic forecast.

---

# Disclaimer

This project is provided for **educational, demonstration, and portfolio purposes only**.

This project uses illustrative financial data and assumptions for educational and portfolio purposes. The forecasts and DCF valuation are not intended to be used as investment advice.

The model is not intended to represent the financial performance or valuation of a specific real company.

Anyone using this framework for real-world analysis is responsible for independently verifying the underlying data, assumptions, accounting treatment, and valuation methodology.
