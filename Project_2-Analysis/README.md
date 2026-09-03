# Excel Salary Analysis

## Introduction

This project explores the data job market using Excel to better understand how salary, job title, location, and technical skills relate to one another.

Rather than only looking at overall salary averages, I wanted to answer a few practical questions about data careers: which skills employers request most often, which skills are associated with higher pay, how salaries vary across regions, and whether jobs requiring more skills tend to offer higher salaries.

## Questions I Wanted to Answer

I focused the analysis on four main questions:

1. Do jobs that require more skills tend to pay more?
2. How do salaries for data-related roles vary by location?
3. Which skills are most commonly requested for data professionals?
4. Which skills are associated with the highest salaries?

## Excel Skills Used

This project gave me hands-on experience with several advanced Excel tools:

- PivotTables
- PivotCharts
- Power Query
- Power Pivot
- DAX
- Data relationships
- Slicers and filters
- Scatter plots and trendlines
- Data cleaning and transformation

## Dataset

The dataset contains real-world data science job postings from 2023.

It includes information such as:

- Job titles
- Salaries
- Countries and locations
- Required technical skills
- Job posting details

I used this data to build several related analyses focused on salary and skill demand.

# Analysis

## 1. Do More Skills Lead to Higher Pay?

To investigate whether jobs requiring more skills tend to offer higher salaries, I first prepared the dataset using Power Query and then analyzed the relationship between median salary and the average number of skills requested per job posting.

### Power Query: Extract, Transform, Load

I used Power Query to prepare the raw data before building the analysis.

### Extract

I used Power Query to import the original `data_salary_all.xlsx` dataset and create two separate queries:

- `data_jobs_all` — contains the main job posting information
- `data_job_skills` — contains the individual skills associated with each job ID

### Transform

I cleaned and prepared both queries by updating data types, removing unnecessary columns, cleaning text values, and trimming extra whitespace.

- `data_jobs_all`

        ![2_Project_Analysis_Screenshot1.png](/0_Resources/Images/2_Project_Analysis_Screenshot1.png)

    - `data_job_skills`

        ![2_Project_Analysis_Screenshot2.png](/0_Resources/Images/2_Project_Analysis_Screenshot2.png)

#### Load

- Finally, I loaded both transformed queries into the workbook, setting the foundation for my subsequent analysis.
    - `data_jobs_all`

        ![2_Project_Analysis_Screenshot3.png](/0_Resources/Images/2_Project_Analysis_Screenshot3.png)

    - `data_job_skills`

        ![2_Project_Analysis_Screenshot4.png](/0_Resources/Images/2_Project_Analysis_Screenshot4.png)

###  Analysis

####  Insights

- The data shows a positive relationship between the number of skills requested and median salary.
- Roles such as Senior Data Engineer combine relatively high skill requirements with high median salaries.
- Roles such as Business Analyst and Data Analyst appear toward the lower end of both measures.
- The pattern suggests that roles requiring broader or more specialized skill sets may also command higher salaries, although skill count alone does not explain every salary difference.

    ![2_Project_Analysis_Chart1.png](/0_Resources/Images/2_Project_Analysis_Chart1.png)

### Why It Matters?

This analysis suggests that developing a broader set of relevant technical skills may improve access to higher-paying data roles. It also shows that job title and specialization matter, since some roles pay more than others even when they require a similar number of skills.

## 2. What's the Salary for Data Jobs in Different Regions?

### Skills: PivotTables & DAX

### PivotTable

I created a PivotTable using the Data Model built in Power Pivot.

- I added `job_title_short` to the Rows area.
- I added `salary_year_avg` to the Values area.
- I created a new measure to calculate median salary specifically for United States job postings.

```DAX
=CALCULATE(
    MEDIAN(data_jobs_all[salary_year_avg]),
    data_jobs_all[job_country] = "United States"
)
```

### DAX

I also created a DAX measure to calculate the overall median annual salary:

```DAX
Median Salary := MEDIAN(data_jobs_all[salary_year_avg])
```

### Analysis

I used these measures to compare median salaries across different job titles and regions.

### Insights

- Senior Data Engineer and Data Scientist roles show some of the highest median salaries in both U.S. and international job postings.
- Salary levels vary noticeably between U.S. and non-U.S. markets for several job titles.
- The results suggest that both role specialization and geographic market can have a meaningful impact on compensation.


    ![2_Project_Analysis_Chart2.png](/0_Resources/Images/2_Project_Analysis_Chart2.png)

### Why It Matters

Understanding how salaries vary by role and location can help job seekers set more realistic salary expectations and make better-informed career decisions.

It also highlights why compensation should be evaluated within the context of both job title and geographic market.

## 3. What Are the Top Skills for Data Professionals?

### Skill: Power Pivot

### Power Pivot

I used Power Pivot to combine the `data_jobs_all` and `data_job_skills` tables into a single Data Model.

Because the data had already been cleaned in Power Query, I was able to connect the two tables and analyze job postings alongside their associated skills.

### Data Model

I created a relationship between the two tables using the shared `job_id` field.

    ![2_Project_Analysis_Screenshot5.png](/0_Resources/Images/2_Project_Analysis_Screenshot5.png)

#### Power Pivot Menu

I used Power Pivot to manage the Data Model, review relationships, and create measures for the analysis.

    ![2_Project_Analysis_Screenshot6.png](/0_Resources/Images/2_Project_Analysis_Screenshot6.png)

### Analysis

I used the connected job and skills tables to identify which technical skills appear most frequently across data-related job postings.

### Insights

- SQL and Python appear among the most commonly requested skills, reinforcing their importance across many data-related roles.
- Cloud technologies such as AWS and Azure also appear frequently, showing the growing importance of cloud platforms in data work.
- The results highlight a mix of core analytical skills and newer infrastructure-focused technologies.

    ![2_Project_Analysis_Chart3.png](/0_Resources/Images/2_Project_Analysis_Chart3.png)

### Why It Matters

Understanding which skills employers request most often can help job seekers prioritize which technologies to learn and strengthen.

It also provides a clearer picture of the technical skills that are currently most relevant across data-related roles.

## 4. What's the Pay of the Top 10 Skills?

### Skill: Advanced Charts and PivotCharts

### PivotChart

I created a combo PivotChart to compare median salary with skill likelihood for the top skills in the dataset.

- The primary axis displays median salary as clustered columns.
- The secondary axis displays skill likelihood as markers.
- I formatted the chart to make it easier to compare salary and demand at the same time.

### Analysis

This visualization makes it possible to compare how much jobs associated with a skill tend to pay with how frequently that skill appears in job postings.

### Insights

- Python, Oracle, and Tableau are associated with some of the highest median salaries among the skills shown.
- SQL combines a relatively high median salary with one of the highest skill likelihood percentages.
- Excel and Word appear frequently enough to remain useful workplace skills, but they are associated with lower median salaries than several more specialized technical skills.
- The results show that the most frequently requested skill is not always the highest-paying skill.

    ![2_Project_Analysis_Chart4.png](/0_Resources/Images/2_Project_Analysis_Chart4.png)

### Why It Matters

This comparison highlights the value of building skills that offer both strong demand and earning potential.

For someone pursuing a career in data, skills such as SQL and Python stand out because they appear frequently in job postings while also being associated with relatively strong salaries.

## Conclusion

This project gave me hands-on experience using Excel to answer practical questions about the data job market.

Using Power Query, Power Pivot, PivotTables, DAX, and charts, I analyzed relationships between job titles, salaries, locations, and technical skills.

The analysis showed several useful patterns:

- Senior and engineering-focused roles tend to have higher median salaries.
- Salary levels vary across geographic markets.
- SQL and Python are among the most commonly requested skills.
- Some skills offer a stronger combination of salary and demand than others.
- Roles requiring more skills generally trend toward higher salaries, although skill count is only one factor.

Overall, this project helped me get more comfortable turning a large dataset into focused analyses and explaining the results in a way that could support real career decisions.
