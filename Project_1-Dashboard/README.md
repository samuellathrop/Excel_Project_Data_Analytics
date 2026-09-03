


# Excel Salary Dashboard



## Overview
This project is an interactive Excel dashboard built to explore salary trends across data-related careers.

Users can filter the dashboard by job title, country, and employment type to compare median salaries, job counts, and commonly used job platforms. The goal was to take a large job-posting dataset and turn it into something that makes salary information easier to explore and understand.


## Dashboard
https://github.com/user-attachments/assets/ebd2ba84-42e8-4553-b0bb-a52ccc1dd671

The dashboard allows users to:

- Compare median salaries across data-related job titles
- Filter results by country
- Filter by employment type
- View the number of matching job postings
- Identify the most common job platform for the selected criteria

## Project File

The completed Excel workbook can be found here:

[1_Salary_Dashboard.xlsx](1_Salary_Dashboard.xlsx)

### Excel Skills Used

This project gave me hands-on practice with several Excel tools and techniques:

- Charts and data visualization
- Formulas and functions
- Data validation
- Dynamic filtering
- Array formulas
- Sorting and organizing data
- Dashboard design
- Working with structured Excel tables

## Dataset

The dataset contains real-world data science job postings from 2023 and includes information such as:

- Job titles
- Salaries
- Countries and locations
- Employment types
- Job platforms
- Required skills

I used this data to build the calculations, charts, and interactive controls used throughout the dashboard.


# Dashboard Build

## Job Title Salary Comparison

<img src="/0_Resources/Images/1_Salary_Dashboard_Chart1.png" width="850" height="550" alt="Salary Dashboard Chart1">

I created a horizontal bar chart to compare median salaries across different data-related job titles.

The job titles are sorted by salary to make comparisons easier. The selected job title is also highlighted so users can quickly see how it compares with other roles.

For example, the dashboard makes it easy to see that senior-level and engineering positions generally have higher median salaries than analyst-level positions.


## Country Salary Map

![1_Salary_Dashboard_Chart2.png](/0_Resources/Images/1_Salary_Dashboard_Country_Map.gif)

I used Excel's map chart to visualize median salaries by country.

This gives users a geographic view of salary differences and makes it easier to compare compensation across different regions.

The map updates based on the criteria selected in the dashboard.


# Formulas and Functions

## Median Salary Calculation

The dashboard calculates median salary dynamically based on the selected job title, country, and employment type.

```excel
=MEDIAN(
IF(
    (jobs[job_title_short]=A2)*
    (jobs[job_country]=country)*
    (ISNUMBER(SEARCH(type,jobs[job_schedule_type])))*
    (jobs[salary_year_avg]<>0),
    jobs[salary_year_avg]
)
)
```
This formula uses multiple criteria to return salary values that match the user's selections before calculating the median.

The calculation:

- Matches the selected job title
- Matches the selected country
- Matches the selected employment type
- Excludes blank or zero salary values
- Returns the median salary for the remaining records

This value is used throughout the dashboard, including the salary card and related charts.





### Job Title Dashboard

The calculated salary values feed directly into the dashboard's job title comparison. When a user selects a job title, the dashboard highlights that role and updates the median salary based on the selected country and employment type.

[INSERT CROPPED IMAGE FROM DASHBOARD 1 SHOWING:
JOB TITLE SELECTOR + SALARY BAR CHART + MEDIAN SALARY CARD]


## Employment Type Filtering

To create a clean list of employment types for the dashboard, I used Excel's `FILTER()` function.

```excel
=FILTER(
J2#,
(NOT(ISNUMBER(SEARCH("and",J2#)))+
ISNUMBER(SEARCH(",",J2#))))*
(J2#<>0)
)
```

This formula filters the source list to remove unwanted entries and create a usable set of employment-type options.

The filtered list is then used for the dashboard's employment type dropdown.

### Data Validation

I applied data validation to the Job Title, Country, and Type fields so users can select from predefined values rather than manually entering criteria.

This:

- Keeps user inputs consistent
- Prevents invalid selections
- Makes the dashboard easier to navigate
- Allows the charts and calculations to update dynamically



### Employment Type Dashboard

The selected employment type updates the related salary chart and job count automatically.

<img src="./0_Resources/Images/1_Salary_Dashboard_Type.png" width="350" alt="Employment Type Dashboard">

### Data Validation in Action

The Job Title, Country, and Type dropdowns make it easy to change the dashboard criteria without manually editing formulas or cells.

The dashboard responds to each selection and updates the relevant calculations and visualizations.

<img src="./0_Resources/Images/1_Salary_Dashboard_Data_Validation.gif" width="425" alt="Salary Dashboard Data Validation">

## What I Learned

This project helped me get more comfortable using Excel for more than basic spreadsheet work. I was able to combine formulas, data validation, charts, and structured data into one interactive dashboard.

Some of the main things I practiced were:

- Writing formulas with multiple criteria
- Building dynamic calculations
- Creating dropdown controls with data validation
- Connecting charts to changing data
- Organizing supporting tables behind a dashboard
- Presenting data in a way that is easy to explore

## Conclusion

This project gave me practical experience building an interactive dashboard from a real-world job postings dataset.

The finished dashboard lets users compare salaries across different job titles, countries, and employment types while also showing job counts and other useful information.

More importantly, it helped me understand how different Excel tools can work together to turn a large dataset into something much easier to analyze and use.
