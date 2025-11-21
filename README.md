# Mini People Analytics Dashboard: Headcount and Turnover

This project is a small, synthetic "Workday style" HR dataset built in Excel.  
It uses VBA, Power Query, and PivotTables to show:

- Current headcount by department  
- Monthly turnover by department  
- Basic tenure insights

The goal is to simulate the kind of people analytics work done in an HR data analyst role, without using any real employee data.

---

## Files and Structure

**Workbook**

- `people_analytics_turnover_demo.xlsx`  
  Main Excel file that contains:
  - `employees_raw` sheet  
  - `employees_model` sheet  
  - PivotTables for headcount and turnover  
  - Charts that can be placed on a Dashboard sheet

**Code**

- VBA macro `BuildEmployeesRaw` stored in the workbook  
  - Automatically generates a synthetic HR dataset in the `employees_raw` sheet  
  - Creates an official Excel Table named `EmployeesRaw`

You can regenerate the entire fake dataset at any time by running the macro.

---

## Data Design

### Raw table: `employees_raw` / `EmployeesRaw`

This sheet simulates a Workday export of employee records.

Columns:

- `EmployeeID`  
- `FirstName`  
- `LastName`  
- `Department`  
  - Example values: Nursing, Medical Assistants, Admin, Behavioral Health, IT / Analytics  
- `JobTitle`  
  - Generated based on department, for example RN, MA I, Data Analyst  
- `Location`  
  - Example values: Seattle Clinic, Bellevue Clinic, Renton Clinic, Tacoma Clinic  
- `HireDate`  
  - Random dates between 2022 01 01 and 2025 11 01  
- `TermDate`  
  - Blank if the employee is still active  
  - If populated, always after the HireDate  
- `EmploymentType`  
  - Full time, Part time, Per diem  
- `FTE`  
  - Typical values such as 1.0, 0.9, 0.8, 0.6, 0.5, 0.3 etc.  
- `TerminationReason`  
  - Voluntary, Involuntary, Retirement, or blank if still employed

The macro aims for a termination rate around one third, skewed toward voluntary separations, to imitate real world patterns.

---

## Power Query Model: `employees_model`

The raw table is loaded into Power Query and transformed into a cleaner model table.

Core steps in Power Query:

- Enforce data types  
  - Dates as Date  
  - FTE as Decimal Number  
  - IDs and Tenure as Whole Number  
- Add calculated columns:

1. **IsActive**  
   - `"Yes"` if `TermDate` is null  
   - `"No"` if `TermDate` has a value  

2. **HireMonth**  
   - Month bucket of the hire date  
   - Example: if `HireDate` = 2023 04 15, `HireMonth` = 2023 04 01  

3. **TermMonth**  
   - Month bucket of the termination date  
   - Example: if `TermDate` = 2024 06 17, `TermMonth` = 2024 06 01  
   - Used for grouping turnover counts by month

4. **TenureDays**  
   - Number of days between `HireDate` and either:
     - `TermDate` if the employee has left  
     - A fixed "as of" date (for example 2025 11 01) if still active  

After loading back to Excel as `employees_model`, one more helper column is added directly in the sheet:

5. **HeadcountFlag** (Excel formula)  
   - `1` if `IsActive = "Yes"`  
   - `0` otherwise  
   - Allows simple headcount sums in PivotTables

---

## Analytics Views

### 1. Current Headcount by Department

Source: `employees_model`

- PivotTable configuration:
  - Rows: `Department`
  - Values: `Sum of HeadcountFlag`  
- Filter (optional): `IsActive = "Yes"`

This produces a simple view of current headcount by department.  
A column chart is used to visualize this on the Dashboard.

---

### 2. Monthly Turnover by Department

Source: `employees_model`

- PivotTable configuration:
  - Rows: `TermMonth`
  - Columns: `Department`
  - Values: `Count of EmployeeID`
  - Filter: `TerminationReason`  

This view answers questions such as:

- How many people left each department per month  
- When turnover spikes occur  
- Whether a specific department has higher voluntary or involuntary turnover

A line or column chart can be built from this PivotTable to visualize turnover over time.

---

### 3. Tenure by Department (Optional)

Source: `employees_model`

- PivotTable configuration:
  - Rows: `Department`
  - Values: `Average of TenureDays`

This helps compare typical tenure across departments.  
It can be converted to years for easier interpretation, but days are acceptable for a simple view.

---

## Screenshots

<img width="1630" height="620" alt="image" src="https://github.com/user-attachments/assets/1ac8f0e6-84fe-4e27-b3cb-a1c417bade36" />

