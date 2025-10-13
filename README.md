# Oil-Gas-Drilling-Potential-Analyzer

# 1. Executive Summary

This dashboard and analysis identify and rank locations with the highest potential for future oil and gas drilling based on historical production, well activity, and geological formation. Using cleaned production records and a scoring system that combines oil, gas (converted to BOE), and well activity, the tool highlights counties, formations, and fields that consistently outperform others, enabling data-driven prioritization of capital allocation for drilling.

---

## Key Findings (High-Level)

- **Top County:** Chautauqua (highest combined oil + gas production in the dataset)
- **Top Formations:** Medina, Queenston, and Bradford show the strongest production/synergy
- **Top Field (by Production Score):** Brazos Field (score: 92) — highest drilling potential
- **Operational Signal:** ~87% of wells are active in the current filtered context, indicating a largely active asset base

---

# 2. Problem Statement & Objectives

## Problem Statement

We need to determine where to drill new wells to maximize production and avoid capital wasted on low-yield or inactive locations. Historical production, well activity, and geology must be combined into an objective ranking so the company can prioritize drilling locations.

---

## Primary Objective

Identify and rank high-potential locations (fields/counties/formations) for new drilling to maximize production output and optimize capital expenditure.

---

## Specific Analysis Objectives

- Identify top-performing geographic areas (county, town, field) by historical oil (bbl) and gas (Mcf) production.
- Pinpoint the most productive producing formations associated with high-yield wells.
- Analyze well status patterns (active vs inactive) to understand success factors.
- Create a location scoring system (**Production Score**) to rank drilling targets using production and activity metrics.

# 3. Data Overview & Dictionary

## Dataset

## Important Cleaning Notes (Applied)

- Standardized column names to.
- Converted inconsistent date formats into a single `production_date_entered` datetime field.  
  - Rows with unparseable dates were flagged for review.
- Coerced numeric columns to appropriate numeric types.  
  - Missing numeric values were set to `0` where appropriate.
- Parsed `location` field to extract **town** and **geographic coordinates** (latitude, longitude).

---

## Data Dictionary (Key Columns)

| Column Name (Clean)        | Description |
|-----------------------------|-------------|
| `production_year`           | Reporting year (e.g., 1995) |
| `production_date_entered`   | Date record was entered into system |
| `operator`                  | Operating company or name |
| `county`                    | County where the well is located |
| `town`                      | Town or municipality |
| `field`                     | Field name |
| `producing_formation`       | Geological producing formation (e.g., Medina) |
| `active_oil_wells`          | Count of oil wells actively producing in the year |
| `inactive_oil_wells`        | Count of oil wells not producing that year |
| `active_gas_wells`          | Count of gas wells actively producing in the year |
| `inactive_gas_wells`        | Count of gas wells not producing that year |
| `injection_wells`           | Count of injection wells |
| `disposal_wells`            | Count of disposal wells |
| `self_use_well`             | Indicator if production used on-site (YES/NO) |
| `oil_produced_bbl`          | Oil produced during reporting year (bbl) |
| `gas_produced_mcf`          | Gas produced during reporting year (Mcf) |
| `water_produced_bbl`        | Produced water (bbl) |
| `taxable_gas_mcf`           | Gas subject to tax (Mcf) |
| `purchaser_codes`           | Purchaser code (if present) |
| `location`                  | Town + coordinates text (if available) |
| `latitude`                  | Parsed latitude (nullable) |
| `longitude`                 | Parsed longitude (nullable) |
| `total_production_boe`      | Derived: `oil_produced_bbl + gas_produced_mcf * 0.178` |
| `total_wells`               | Derived: sum of active + inactive oil & gas wells |
| `active_well_ratio`         | Derived: `(active_oil_wells + active_gas_wells) / total_wells` |

---

### Conversion Note

> **1 Mcf ≈ 0.178 barrel of oil equivalent (BOE)**  

# 4. Tools, Environment & Versions

## Tools Used

- **Power BI Desktop** — Used to build the dashboard visuals and measures.  

- **Microsoft Excel / Power Query** — Performed initial data cleaning, transformation, and parsing.

- **DAX (Data Analysis Expressions)** — Used for measures, calculated columns, and the scoring system.  
  *(See DAX code examples below.)*

---

## Assumptions

- **Gas to BOE conversion factor:** `1 Mcf = 0.178 BOE`  
- **Production year:** Stored as a numeric integer.  
- **Missing numeric cells:** Replaced with `0` where values were clearly missing or null.

# 5. Data Preparation & Cleaning Steps (Summary)

1. **Load Data**  
   - Imported the original Excel sheet into **Power Query**.

2. **Standardize Headers**  
   - Trimmed whitespace, converted to lowercase, and replaced spaces with underscores.

3. **Normalize Dates**  
   - Converted `production_date_entered` to a consistent datetime format.  
   - Flagged rows with unparseable dates for quality review.

4. **Type Conversion**  
   - Enforced numeric types on production and well count columns.  
   - Converted text fields to trimmed string types.

5. **Parse Location**  
   - Extracted `latitude` and `longitude` where coordinate data was present.

6. **Handle Nulls**  
   - Filled numeric nulls with `0` **only where appropriate**.  
   - Retained but flagged missing geographic or formation values for QA review.

7. **Create Derived Columns**  
   - `total_production_boe` = oil + gas (converted using BOE factor).  
  
8. **Export Cleaned File**  
   - Final dataset saved as **`oil_and_gas_production_cleaned.xlsx`** for analysis.

---

# 6. Key Measures & DAX (Copy-Paste Ready)

## Core Production Measures

Total Oil Produced (bbl) = 
SUM('oil_and_gas_production_cleaned'[oil_produced_bbl])

Total Gas Produced (Mcf) = 
SUM('oil_and_gas_production_cleaned'[gas_produced_mcf])

Total Production BOE =
SUM('oil_and_gas_production_cleaned'[oil_produced_bbl]) +
SUM('oil_and_gas_production_cleaned'[gas_produced_mcf]) * 0.178

---

Average Oil per Active Well =
DIVIDE(
    SUM('oil_and_gas_production_cleaned'[oil_produced_bbl]),
    SUM('oil_and_gas_production_cleaned'[active_oil_wells]),
    0
)

Average Gas per Active Well =
DIVIDE(
    SUM('oil_and_gas_production_cleaned'[gas_produced_mcf]),
    SUM('oil_and_gas_production_cleaned'[active_gas_wells]),
    0
)

% Active Wells =
DIVIDE(
    SUM('oil_and_gas_production_cleaned'[active_oil_wells]) + 
    SUM('oil_and_gas_production_cleaned'[active_gas_wells]),
    SUM('oil_and_gas_production_cleaned'[active_oil_wells]) + 
    SUM('oil_and_gas_production_cleaned'[inactive_oil_wells]) + 
    SUM('oil_and_gas_production_cleaned'[active_gas_wells]) + 
    SUM('oil_and_gas_production_cleaned'[inactive_gas_wells]),
    0
)

---

Average Oil per Active Well =
DIVIDE(
    SUM('oil_and_gas_production_cleaned'[oil_produced_bbl]),
    SUM('oil_and_gas_production_cleaned'[active_oil_wells]),
    0
)

Average Gas per Active Well =
DIVIDE(
    SUM('oil_and_gas_production_cleaned'[gas_produced_mcf]),
    SUM('oil_and_gas_production_cleaned'[active_gas_wells]),
    0
)

% Active Wells =
DIVIDE(
    SUM('oil_and_gas_production_cleaned'[active_oil_wells]) + 
    SUM('oil_and_gas_production_cleaned'[active_gas_wells]),
    SUM('oil_and_gas_production_cleaned'[active_oil_wells]) + 
    SUM('oil_and_gas_production_cleaned'[inactive_oil_wells]) + 
    SUM('oil_and_gas_production_cleaned'[active_gas_wells]) + 
    SUM('oil_and_gas_production_cleaned'[inactive_gas_wells]),
    0
)

---

Production Score (Field) =
VAR OilScore =
    CALCULATE(
        SUM('oil_and_gas_production_cleaned'[oil_produced_bbl]),
        ALLEXCEPT('oil_and_gas_production_cleaned', 'oil_and_gas_production_cleaned'[field])
    )
VAR GasScore =
    CALCULATE(
        SUM('oil_and_gas_production_cleaned'[gas_produced_mcf]) * 0.178,
        ALLEXCEPT('oil_and_gas_production_cleaned', 'oil_and_gas_production_cleaned'[field])
    )
VAR ActiveWells =
    CALCULATE(
        SUM('oil_and_gas_production_cleaned'[active_oil_wells]) + 
        SUM('oil_and_gas_production_cleaned'[active_gas_wells]),
        ALLEXCEPT('oil_and_gas_production_cleaned', 'oil_and_gas_production_cleaned'[field])
    )
VAR Score = (0.5 * OilScore) + (0.5 * GasScore) + (0.2 * ActiveWells)
RETURN Score


---

Here’s your section written cleanly and professionally in **GitHub README.md** format:


# 7. Dashboard Layout & How to Use It (Page by Page)

---

## **Page 1 — Production Overview (Executive)**

### **Layout**
- **Top Row (KPIs):**
  - Total Oil Produced  
  - Total Gas Produced  
  - Oil Prod./Active Well  
  - Gas Prod./Active Well  
  - Top County by Production (name)  
  - % Active Wells  

- **Left Panel:**
  - Slicers — `Production Year`, `County`, `Town`, `Field`, `Producing Formation`, `Operator`, `Well Status`

- **Main Visuals:**
  - **Map:** Bubble map where bubble **size = Total Production BOE**, **color = Producing Formation**
  - **Top 10 Counties by Oil**
  - **Top 10 Counties by Gas**
  - **Treemap:** Production by Formation
  - **Stacked Column Chart:** Active vs. Inactive Wells by County

### **How to Use**
- Use slicers to filter by year, location, or operator.
- KPIs and visuals update dynamically.
- Click any **bubble**, **county bar**, or **formation tile** to cross-filter other visuals on the page.

---

## **Page 2 — Map showing total production (oil + gas) by counties and production formation**

---

## **Page 3 — Scoring & Ranking**

### **Layout**
- **Table:**  
  - Columns: Field, County, Formation, Production Score, Total Oil, Total Gas, Active Wells  
  - Sorted by **Production Score (Descending)**

- **Gauge:**  
  - Displays the **Top Field Score** (maximum across all fields)

- **Bar Chart:**  
  - Top N Fields ranked by **Production Score**

### **How to Use**
- Use this page to identify and prioritize **high-potential drilling targets**.  
- Filter or export ranked field lists for strategic and operational planning.

  

# 8. Insights & Interpretation 

---

### **Top Counties and Formations**
- **Chautauqua County** and the **Medina Formation** are consistently high producers of oil and gas across the dataset.  
- These represent **high-priority targets** for deeper geological and economic evaluation.

### **Operator Patterns**
- Operators with strong performance in **Medina** tend to replicate success across multiple fields.  
- This suggests **formation-driven productivity** rather than purely operator-specific efficiency.

### **Oil–Gas Synergy**
- **Medina**, **Queenston**, and **Bradford** formations demonstrate balanced oil and gas outputs.  
- These are prime candidates for **multi-resource development** strategies.

### **Field Ranking**
- **Lakeshore Field** ranks highest by *Production Score*, indicating top drilling potential.  
- Recommended for **expansion studies**, **reserve reevaluation**, and **infill drilling assessments**.

### **Operational Signal**
- High **% Active Wells (~87%)** indicates a healthy, active asset base.  
- However, certain high-production fields still contain **inactive wells** — suggesting opportunities for **recompletion**, **workovers**, or **reactivation** programs.

---

# 9. Recommendations (Actionable)

---

### **1. Targeted Appraisal Drilling**
Focus on the **top three formations**:
- **Medina**, **Queenston**, and **Bradford**
  
Prioritize within **top-performing counties**:
- **Chautauqua**, **Cattaraugus**, and **Allegany**

### **2. Field-Level Technical Audits**
Conduct detailed audits for top-ranked fields:
- Review **reservoir performance**, **wellbore integrity**, and **production decline trends**.
- Identify candidates for **infill drilling** or **enhanced recovery**.

### **3. Operator Engagement**
- Partner with or benchmark **operators performing well in top formations**.  
- Capture **best practices** and **operational efficiencies** for replication across other assets.































