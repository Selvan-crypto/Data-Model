# 🏥 Hospital Analytics – Power BI Data Model

A end-to-end **data modelling project in Power BI** built on a hospital dataset (patients, visits, services, invoices, payments, referrals and health campaigns).
The focus of this repository is the **model itself**: how the raw Excel tables were cleaned in Power Query, reshaped into a **star schema**, and connected with the right relationships (active/inactive, single/both direction, bridge tables) so that reports are accurate and fast.

![Overall data model](https://github.com/Selvan-crypto/Data-Model/blob/main/image/model_view.png)

---

##  Table of Contents
1. [Project Overview](#-project-overview)
2. [Tools & Skills Used](#-tools--skills-used)
3. [Source Data](#-source-data)
4. [Data Modelling Workflow](#-data-modelling-workflow-start-to-end)
5. [Step 1 – Data Cleaning in Power Query](#step-1--data-cleaning-in-power-query)
6. [Step 2 – Merge Queries (Joins)](#step-2--merge-queries-joins)
7. [Step 3 – Append Queries](#step-3--append-queries)
8. [Step 4 – Building Dimension & Fact Tables](#step-4--building-dimension--fact-tables)
9. [Step 5 – Date Table](#step-5--date-table)
10. [Step 6 – Relationships](#step-6--relationships)
11. [Step 7 – Active vs Inactive Relationship](#step-7--active-vs-inactive-relationship-userelationship)
12. [Step 8 – Many-to-Many & Bridge Tables](#step-8--solving-many-to-many-with-bridge-tables)
13. [Step 9 – Row-Level Security](#step-9--row-level-security-rls)
14. [DAX Measures](#-dax-used)
15. [Key Learnings](#-key-learnings)

---

##  Project Overview

| | |
|---|---|
| **Domain** | Healthcare / Hospital operations |
| **Goal** | Build a clean, scalable semantic model to analyse hospital revenue, visits, services, patient demographics, campaign performance and targets |
| **Schema** | Star schema (with bridge tables for many-to-many) |
| **Tables in model** | 8 dimensions/bridge + 5 fact tables + 1 security table = **14 tables, 16 relationships** |

**Business questions the model can answer**
- How much revenue/discount did each service, region and city generate?
- How do visits differ by admission type, status and priority?
- Which patients are linked to which visit lines?
- How did campaigns (reach, registrations, spend) perform, and which services did they promote?
- Actual vs target revenue by period.
- Each regional manager should only see their own region (Row-Level Security).

---

## 🛠 Tools & Skills Used
- **Power BI Desktop** – Power Query (M), Model view, DAX
- **Excel** – source workbook `hospital_raw_tables.xlsx`
- **Concepts:** data cleaning, data type validation, merge & append, star schema, surrogate keys, role-playing dimensions, active/inactive relationships, cross-filter direction, bridge tables, Row-Level Security

---

##  Source Data
All tables were loaded from a single Excel workbook (`hospital_raw_tables.xlsx`), one sheet per table.

| Source sheet | Used for |
|---|---|
| `PATIENTS`, `patient_contacts`, `patient_details`, `Address`, `cities`, `doctors` | `dim_patient` |
| `services` | `dim_service` |
| `VISITS_2025`, `VISITS_2026` | `dim_visit_flag` (appended) |
| `cities` | `dim_geo` |
| `health_campaigns` | `dim_campaigns`, `fact_campaign_spend` |
| `visit_line_items`, `visit`, `referrals` | `fact_visit`, `dim_reference`, `bridge_patient_visit` |
| `invoice_lines`, `INVOICES`, `payments` | `fact_invoice` |
| `campaign_services` | `fact_promotion_coverage` |
| `visit_targets` | `fact_visit_target` |
| `security` | `security` (RLS mapping) |

---

##  Data Modelling Workflow (Start to End)

```
Raw Excel sheets
      │
      ▼
1. Clean & validate types (Power Query)
      │
      ▼
2. Merge queries  → flatten snowflake tables, look up surrogate keys
      │
      ▼
3. Append queries → combine 2025 + 2026 visit data
      │
      ▼
4. Build DIMENSION tables (unique rows + surrogate key) and FACT tables (keys + measures)
      │
      ▼
5. Create dim_date (DAX)
      │
      ▼
6. Create relationships (cardinality, filter direction)
      │
      ▼
7. Handle role-playing dimension  → inactive relationship + USERELATIONSHIP
      │
      ▼
8. Resolve many-to-many           → bridge tables
      │
      ▼
9. Secure the model               → Row-Level Security
```

---

## Step 1 – Data Cleaning in Power Query

### 1.1 Data type validation
Every query starts with **Promote Headers** followed by **Changed Type**, so each column gets the correct type before any join or load.

| Type applied | Examples |
|---|---|
| Whole number (`Int64.Type`) | `PatientID`, `AddressID`, `Quantity`, `Budget`, `Reach`, `Registrations`, `TargetRevenue` |
| Decimal (`type number`) | `UnitPrice`, `UnitCost`, `DiscountPct`, `LineTotal`, `Amount`, `Spend` |
| Date (`type date`) | `VisitDate`, `StartDate`, `EndDate`, `Period`, `InvoiceDate`, `PayDate` |
| Text (`type text`) | IDs such as `LineID`, `VisitID`, `ServiceCode`, plus names, emails, phone numbers |

> Keys that are alphanumeric (e.g. `VIS00071`, `L00078`) are kept as **text**; numeric IDs are kept as **whole numbers**. Matching key types on both sides of a join is what makes merges and relationships work.

### 1.2 Handling missing, blank and duplicate data
| Technique | Where it was used | Why |
|---|---|---|
| **Remove Blank Rows** (`each not List.IsEmpty(... {"", null})`) | `dim_reference` | After the left-outer merge with `referrals`, visits with no referral produce empty rows, so they are removed to keep the dimension clean |
| **Remove Duplicates** (`Table.Distinct`) | `dim_geo`, `dim_campaigns`, `dim_visit_flag`, `dim_reference` | A dimension must have **one row per member**; this guarantees a unique key and a clean one-side for the relationship |
| **Remove unneeded columns** | All queries (e.g. `hash_key`, `source_id`, `ServiceDescription`, repeated name/city columns) | Reduces model size and avoids ambiguous duplicate columns |
| **Trim text** (`Text.Trim`) | `fact_promotion_coverage` (`PromotedServices`, `CampaignName`) | After splitting a comma-separated list, leading spaces would break the lookup to `dim_service` |
| **Left Outer joins** | All lookups of surrogate keys | Keeps every fact row even when no dimension match exists, so no transactions are silently dropped |

### 1.3 Other transformations
- **Rename columns** to a consistent `Snake_Case` (`PatientName → Patient_Name`, `VisitTotal → Visit_Total` …)
- **Reorder columns** so keys come first
- **Split column by delimiter** (comma) into rows – `fact_promotion_coverage`
- **Add Index column** to create surrogate keys (`Service_Key`, `Geo_Key`, `Campaign_Key`, `Visit_Flag_Key`)

---

## Step 2 – Merge Queries (Joins)

All merges are **Left Outer** (keep all rows from the first table).

| Resulting table | Left table | Right table | Join key(s) | Purpose |
|---|---|---|---|---|
| `dim_patient` | `PATIENTS` | `patient_contacts` | `PatientID` | Contact name & email |
| `dim_patient` | `PATIENTS` | `patient_details` | `PatientID` | Coverage limit & phone |
| `dim_patient` | `PATIENTS` | `Address` | `AddressID` | Street & city |
| `dim_patient` | `PATIENTS` | `cities` | `CityName` | Region |
| `dim_patient` | `PATIENTS` | `doctors` | `Physician = DoctorName` | Department, specialty, doctor contact, experience |
| `fact_visit` | `visit_line_items` | `visit` | `VisitID` | Visit date, city, status, totals |
| `fact_visit` | `visit_line_items` | `dim_service` | `ServiceName = Service_Name` | Get `Service_Key` |
| `fact_visit` | `visit_line_items` | `dim_visit_flag` | `AdmissionType + Status + Priority` | Get `Visit_Flag_Key` (**composite key**) |
| `fact_visit` | `visit_line_items` | `dim_geo` (twice) | `TreatmentCity = City`, `BillToCity = City` | Get `Treatment_City_Key` and `Bill_To_City_Key` |
| `fact_visit` | `visit_line_items` | `referrals` | `VisitID` | Get referral info |
| `fact_invoice` | `invoice_lines` | `INVOICES` | `InvoiceID` | Invoice date & visit id |
| `fact_invoice` | `invoice_lines` | `payments` | `InvoiceID` | Payment id & date |
| `fact_invoice` | `invoice_lines` | `dim_service` | `ServiceName` | Get `Service_Key` |
| `fact_campaign_spend` | `health_campaigns` | `dim_campaigns` | `Campaign + Channel + Start + End + Budget` | Get `Campaign_Key` |
| `fact_promotion_coverage` | `campaign_services` | `dim_campaigns`, `dim_service` | `Campaign_Name`, `Service_Name` | Get both keys |
| `bridge_patient_visit` | `visit_line_items` | `visit` → `dim_patient` | `VisitID`, then `Patient_Name` | Get `Patient_Id` for each line |

**Why merge?** The raw data is normalised across many sheets. Merging flattens it into dimensions (e.g. patient + address + city + doctor in one `dim_patient`) and replaces text columns in the facts with **numeric surrogate keys**, which are smaller and faster than text.

---

## Step 3 – Append Queries

| Resulting table | Appended sources | Why |
|---|---|---|
| `dim_visit_flag` | `VISITS_2025` + `VISITS_2026` | Visit data is split by year. `Table.Combine` stacks them into one table so that all distinct *Admission Type / Status / Priority* combinations from **both years** are captured |

```powerquery
Source = Table.Combine({VISITS_2025, VISITS_2026})
```

---

## Step 4 – Building Dimension & Fact Tables

### Star schema design

| Type | Table | Grain / description | Key |
|---|---|---|---|
| **Dimension** | `dim_patient` | One row per patient (flattened with contact, address, region, doctor) | `Patient_Id` |
| **Dimension** | `dim_service` | One row per service | `Service_Key` (index) |
| **Dimension** | `dim_geo` | One row per city/region | `Geo_Key` (index) |
| **Dimension** | `dim_campaigns` | One row per campaign | `Campaign_Key` (index) |
| **Dimension** | `dim_visit_flag` | One row per Admission Type + Status + Priority combination (a **junk dimension**) | `Visit_Flag_Key` (index) |
| **Dimension** | `dim_reference` | One row per referral | `Referral_Id` |
| **Dimension** | `dim_date` | One row per calendar day | `Date` |
| **Fact** | `fact_visit` | One row per visit line item | `Line_Id` |
| **Fact** | `fact_invoice` | One row per invoice line | `Invoice_Line_Id` |
| **Fact** | `fact_campaign_spend` | Daily spend/reach/registrations per campaign | `Campaign_Key + Date` |
| **Fact** | `fact_promotion_coverage` | Which campaign promoted which service (factless) | `Campaign_Key + Service_Key` |
| **Fact** | `fact_visit_target` | Target revenue per period | `Period` |
| **Bridge** | `bridge_patient_visit` | Links patients to visit lines | `Line_ID + Patient_Id` |
| **Security** | `security` | User email → region mapping for RLS | `Region` |

**Highlights**
- **Surrogate keys:** `Table.AddIndexColumn` was used to create `Service_Key`, `Geo_Key`, `Campaign_Key` and `Visit_Flag_Key`, after removing duplicates.
- **Junk dimension:** three low-cardinality text columns (Admission Type, Status, Priority) were moved out of the fact into `dim_visit_flag`, leaving only one key in the fact.
- **Fact tables are lean:** descriptive text columns were removed from facts and replaced with keys.

---

## Step 5 – Date Table

`dim_date` was created in the model using DAX:

```dax
dim_date = CALENDARAUTO()
```
Then `Year` and `Month` columns were added. It was marked as the date table and used to connect all date columns of the facts.

---

## Step 6 – Relationships


| # | From (Many side) | To (One side) | Cardinality | Cross-filter | Status |
|---|---|---|---|---|---|
| 1 | `bridge_patient_visit[Line_ID]` | `fact_visit[Line_Id]` | Many → One | **Both** | Active |
| 2 | `bridge_patient_visit[Patient_Id]` | `dim_patient[Patient_Id]` | Many → One | **Both** | Active |
| 3 | `dim_geo[Region]` | `security[Region]` | Many → Many | Single | Active |
| 4 | `fact_campaign_spend[Campaign_Key]` | `dim_campaigns[Campaign_Key]` | Many → One | Single | Active |
| 5 | `fact_campaign_spend[Date]` | `dim_date[Date]` | Many → One | Single | Active |
| 6 | `fact_invoice[Invoice_Date]` | `dim_date[Date]` | Many → One | Single | Active |
| 7 | `fact_invoice[Service_Key]` | `dim_service[Service_Key]` | Many → One | Single | Active |
| 8 | `fact_promotion_coverage[Campaign_Key]` | `dim_campaigns[Campaign_Key]` | Many → One | Single | Active |
| 9 | `fact_promotion_coverage[Service_Key]` | `dim_service[Service_Key]` | Many → One | Single | Active |
| 10 | `fact_visit[Bill_To_City_Key]` | `dim_geo[Geo_Key]` | Many → One | Single | **Active** |
| 11 | `fact_visit[Referral_Id]` | `dim_reference[Referral_Id]` | Many → One | Single | Active |
| 12 | `fact_visit[Service_Key]` | `dim_service[Service_Key]` | Many → One | Single | Active |
| 13 | `fact_visit[Treatment_City_Key]` | `dim_geo[Geo_Key]` | Many → One | Single | **Inactive** |
| 14 | `fact_visit[Visit_Flag_Key]` | `dim_visit_flag[Visit_Flag_Key]` | Many → One | Single | Active |
| 15 | `fact_visit[VisitDate]` | `dim_date[Date]` | Many → One | Single | Active |
| 16 | `fact_visit_target[Period]` | `dim_date[Date]` | One → One | **Both** | Active |

### Single vs Both filter direction

| Direction | Meaning | Used where in this model |
|---|---|---|
| **Single** | Filter flows from the **dimension (one side) → fact (many side)** only. This is the default and recommended setting: it is fast and avoids ambiguity. | Almost all fact-to-dimension relationships |
| **Both** | Filters flow in **both** directions. | (a) `bridge_patient_visit` – so that a selection on `dim_patient` reaches `fact_visit` *through* the bridge, and a selection on `fact_visit` reaches `dim_patient`. (b) `fact_visit_target ↔ dim_date` – a 1:1 relationship between two tables at the same grain. |

> Both-direction filters were used only where the model *needs* them (bridge and 1:1), because bi-directional filtering can slow the model and create ambiguous filter paths.

---

## Step 7 – Active vs Inactive Relationship (USERELATIONSHIP)

`fact_visit` has **two city columns** that both point to the same dimension, `dim_geo`:

| Column | Meaning |
|---|---|
| `Bill_To_City_Key` | City where the patient is billed |
| `Treatment_City_Key` | City where the treatment was given |

This is a **role-playing dimension**: one `dim_geo` plays two roles. Power BI allows only **one active relationship** between two tables, so:

- ✅ `Bill_To_City_Key → Geo_Key` is **Active** (default filter path)
- ⛔ `Treatment_City_Key → Geo_Key` is **Inactive** (dotted line in the model)

To use the inactive relationship, a measure activates it temporarily with `USERELATIONSHIP`:

```dax
Total_Discount_by_treatment =
CALCULATE (
    SUM ( fact_visit[DiscountPct] ),
    USERELATIONSHIP ( fact_visit[Treatment_City_Key], dim_geo[Geo_Key] )
)
```

**Result:** one geography table serves two analyses – by *billing city* (default) and by *treatment city* (via the measure) – without duplicating `dim_geo`.

---

## Step 8 – Solving Many-to-Many with Bridge Tables

A **many-to-many** relationship (e.g. one patient has many visit lines, and a visit line can be linked to a patient) cannot be modelled correctly with a direct relationship without ambiguity and wrong totals. The fix is a **bridge table** that sits in the middle and turns it into two **one-to-many** relationships.

### 8.1 Patient ↔ Visit  →  `bridge_patient_visit`

```
dim_patient (1) ──< bridge_patient_visit >── (1) fact_visit
   Patient_Id        Patient_Id | Line_ID        Line_Id
```

**How the bridge was built (Power Query)**
1. Start from `visit_line_items`
2. Merge with `visit` on `VisitID` to bring in `PatientName`
3. Merge with `dim_patient` on `PatientName = Patient_Name` to get `Patient_Id`
4. Remove all other columns → keep only **`Line_ID`** and **`Patient_Id`**

Relationships: `bridge[Line_ID] → fact_visit[Line_Id]` and `bridge[Patient_Id] → dim_patient[Patient_Id]`, both **Many-to-One with Both-direction filtering** so slicing by patient correctly filters visit lines.

### 8.2 Campaign ↔ Service  →  `fact_promotion_coverage`

A campaign promotes several services, and a service can be promoted by several campaigns (also many-to-many). In the source, the services were stored as one comma-separated text cell.

**Solution (Power Query)**
1. **Split** `PromotedServices` by comma **into rows**
2. **Trim** the text
3. Merge to `dim_campaigns` and `dim_service` to get `Campaign_Key` and `Service_Key`
4. Keep only the two keys

The result is a **factless bridge** with one row per (campaign, service) pair, connected one-to-many to both dimensions.

```
dim_campaigns (1) ──< fact_promotion_coverage >── (1) dim_service
```

---

## Step 9 – Row-Level Security (RLS)

Each user should only see data for their own region. A `security` table maps **UserEmail → Region**, and a role named **Regional Access** filters `dim_geo`:

```dax
dim_geo[Region] =
LOOKUPVALUE (
    security[Region],
    security[UserEmail],
    USERPRINCIPALNAME ()
)
```

Because `dim_geo` filters `fact_visit`, the restriction flows through to the visit data.

---

##  DAX Used

| Name | Type | Formula | Purpose |
|---|---|---|---|
| `dim_date` | Calculated table | `CALENDARAUTO()` | Auto-generates a date table from the min/max dates in the model |
| `Total_Discount_by_treatment` | Measure | `CALCULATE(SUM(fact_visit[DiscountPct]), USERELATIONSHIP(fact_visit[Treatment_City_Key], dim_geo[Geo_Key]))` | Discount by **treatment city** using the inactive relationship |
| RLS rule – *Regional Access* | Role filter | `dim_geo[Region] = LOOKUPVALUE(security[Region], security[UserEmail], USERPRINCIPALNAME())` | Restricts a user to their own region |

---

##  Key Learnings
- Cleaning in **Power Query before modelling** (types, blanks, duplicates, trimming) prevents broken joins and wrong totals.
- **Merges** flatten normalised data; **appends** combine same-structure tables across periods.
- A **star schema with surrogate keys** keeps facts lean and relationships fast.
- **Single-direction** filtering is the default; **Both** should only be used when truly needed (bridge, 1:1).
- Only **one relationship can be active** between two tables, so role-playing dimensions need an **inactive relationship + `USERELATIONSHIP`**.
- **Bridge tables** convert many-to-many relationships into two clean one-to-many relationships.
- **Row-Level Security** with `USERPRINCIPALNAME()` and a mapping table gives secure, per-user views from one report.

---
