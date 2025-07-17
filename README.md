# 📊 Excel Analysis Project – Data Validation, Power Query, Power Pivot, Pivot Table, and Macros

This project showcases a comprehensive Excel-based data analysis pipeline designed to address real-world business problems across procurement, health research, sales, and student performance using the powerful features of Excel.

---

## ✅ 1. Data Validation – Procurement Order Form

As a procurement specialist in a manufacturing company, a validated data entry form was created with the following specifications:

- **Order Table Columns**: Order Number, Raw Material, Quantity, Transportation Method, Order Status.
- **Data Validations**:
  - *Order Number*: Unique 8-digit numeric code.
  - *Raw Material*: Dropdown with predefined options (Aluminum, Steel, Plastic, Glass, Copper) with auto-updating capability.
  - *Quantity*: Integer values between 10 and 1000.
  - *Transportation*: Predefined options (Normal, Express, Priority). If "Priority" is selected, cell is auto-colored.
  - *Order Status*: Must be one of (Pending, Shipped, Delivered). If more than 10 days old and still not shipped, status cell turns red.
- **Input & Error Messages**: Defined for all relevant fields to guide correct data entry.

---

## 🔄 2. Power Query – Student Nutrition & Lifestyle Dataset

A research dataset on student nutrition was cleaned and transformed in Power Query:

- **Cleaning Steps**:
  - Removed redundant columns like `type_sports`, `comfort_food_reasons_coded.1`.
  - Removed rows missing key values in `GPA`, `life_rewarding`, or `self_perception_weight`.

- **Merging & Transforming**:
  - Merged `comfort_food` and `comfort_food_reasons` → `Comfort_Food_Info` (title-cased, comma-separated).
  - Merged `fav_cuisine` and `cuisine` → `Cuisine_Preference` (preserving non-missing values).
  - Converted `weight` to numeric.
  - Mapped `Gender`: 1 → M, 2 → F.

- **New Columns**:
  - `Life_Satisfaction`: "Satisfied" if above average `life_rewarding`, else "Unsatisfied".
  - `Weight_Perception`: Based on `self_perception_weight` → "Light", "Heavy", or "Average".

- **Final Columns**: `Gender`, `GPA`, `Cuisine_Preference`, `Comfort_Food_Info`, `life_rewarding`, `self_perception_weight`, `Life_Satisfaction`, `Weight_Perception`.

---

## 📈 3. Power Pivot – Analytical Insights

Data from Power Query was analyzed using Power Pivot:

- **Data Model Setup**: Categorized fields for analysis.
- **Pivot Report 1**: Avg `Life_Satisfaction` by `Gender`, sorted by `GPA` (descending).
- **Pivot Report 2**: Avg `Life_Satisfaction` by `Weight_Perception` – revealed weight perception impact.
- **Final Analysis**: Cross-tab of `Life_Satisfaction` vs `Weight_Perception` revealed trends in self-image and satisfaction.

---

## 🛒 4. Pivot Table – Supermarket Sales Analysis

Sales dataset from an online supermarket analyzed using Pivot Tables:

- **Sales by Category & Region**: Total sales summarized and ranked by region.
- **Profit by Order Date**: Monthly profit analysis grouped by year and month.
- **Top Regions by Profit**: Filtered top 3 profitable regions and visualized using column chart.
- **Category vs Year Analysis**: Line chart showed category-wise sales trends over years.

---

## 🤖 5. Macros – Student Scores Automation

Used VBA macros to streamline grading workflows:

- **Highlighting Grades**:
  - Green for scores ≥ 50 (Pass), Red for < 50 (Fail).
- **Average Calculator**:
  - Macro to compute and show average score in a message box.
- **Filter Passed Students**:
  - Macro filters students with passing scores and copies them to a new sheet `Passed_Students` with headers.

All macros are well-documented in the VBA editor and fully tested.

---

## 📁 Tools Used

- **Microsoft Excel** (Data Validation, Power Query, Power Pivot, Pivot Table, VBA)
- **Familiarity with Excel functions, conditional formatting, dropdowns, macros, and data models**

---
