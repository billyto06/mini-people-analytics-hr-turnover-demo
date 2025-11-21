# Mini People Analytics Dashboard – Headcount and Turnover

This project is a small, synthetic “Workday style” HR dataset built in Excel.  
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
  - Random dates between 2022-01-01 and 2025-11-01  
- `TermDate`  
  - Blank if the employee is still active  
  - If populated, always after the HireDate  
- `EmploymentType`  
  - Full-time, Part-time, Per diem  
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
   - Example: if `HireDate` = 2023-04-15, `HireMonth` = 2023-04-01  

3. **TermMonth**  
   - Month bucket of the termination date  
   - Example: if `TermDate` = 2024-06-17, `TermMonth` = 2024-06-01  
   - Used for grouping turnover counts by month

4. **TenureDays**  
   - Number of days between `HireDate` and either:
     - `TermDate` if the employee has left  
     - A fixed “as of” date (for example 2025-11-01) if still active  

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

## How to Use This Workbook

### 1. Generate or Regenerate Synthetic Data

1. Open the workbook.  
2. Press `Alt + F8` to open the Macro dialog.  
3. Run `BuildEmployeesRaw`.  
4. This will:
   - Clear and rebuild the `employees_raw` sheet  
   - Populate a new set of 80 to 120 synthetic employees  
   - Recreate the `EmployeesRaw` table

### 2. Refresh the Model

After rebuilding the raw data:

1. Go to the **Data** tab.  
2. Click **Refresh All**.  
3. The `employees_model` query and all PivotTables will update.

### 3. Explore the Dashboard

- Use the headcount view to see how many active staff each department has.  
- Use the turnover view to see who left and when.  
- Use slicers for `Department`, `TerminationReason`, or `EmploymentType` to slice the data interactively.

---

## Adapting This To Real HR Data

To swap in real data:

1. Replace the rows in `employees_raw` with an export from your HR system.  
   - Keep the same column names and basic structure.  
2. Make sure the table still has the name `EmployeesRaw`.  
3. Refresh the Power Query model and PivotTables.  

Without changing the visuals, the same workbook will now analyze real headcount and turnover.

---

## Tools Used

- Microsoft Excel  
- VBA for synthetic data generation  
- Power Query for data cleaning and calculated columns  
- PivotTables and charts for headcount and turnover visuals  

This project is meant as a compact example of people analytics using common tools that appear in HR data analyst roles.

