


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


 Background Table

![1_Salary_Dashboard_Screenshot1.png](/0_Resources/Images/1_Salary_Dashboard_Screenshot1.png)

 Dashboard Implementation

<img src="/0_Resources/Images/1_Salary_Dashboard_Job_Title.png" width="400" height="500" alt="Salary Dashboard Title">

#### Count of Job Schedule Type

```
=FILTER(J2#,(NOT(ISNUMBER(SEARCH("and",J2#))+ISNUMBER(SEARCH(",",J2#))))*(J2#<>0))
```

-  **Unique List Generation:** This Excel formula below employs the `FILTER()` function to exclude entries containing "and" or commas, and omit zero values.
- ** Formula Purpose:** This formula populates the table below, which gives us a list of unique job schedule types.

Background Table

![1_Salary_Dashboard_Type.png](/0_Resources/Images/1_Salary_Dashboard_Screenshot2.png)

 Dashboard Implementation:

<img src="/0_Resources/Images/1_Salary_Dashboard_Type.png" width="350" height="500" alt="Salary Dashboard Type">

### Data Validation

#### Filtered List

-  **Enhanced Data Validation:** Implementing the filtered list as a data validation rule under the `Job Title`, `Country`, and `Type` option in the Data tab ensures:
    -  User input is restricted to predefined, validated schedule types
    -  Incorrect or inconsistent entries are prevented
    -  Overall usability of the dashboard is enhanced

<img src="/0_Resources/Images/1_Salary_Dashboard_Data_Validation.gif" width="425" height="400" alt="Salary Dashboard Data Validation">



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
