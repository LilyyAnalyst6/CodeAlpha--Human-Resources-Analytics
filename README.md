# CodeAlpha- HR analytics dashboard

A two-page Power BI report on recruitment, turnover, employee satisfaction and performance, built on a 311-employee HR dataset.

<img width="882" height="489" alt="CodeAlpha HR Pg 1" src="https://github.com/user-attachments/assets/41cb6fec-3abf-406f-a934-400517a8f6f6" />
<img width="882" height="494" alt="CodeAlpha HR Pg 2" src="https://github.com/user-attachments/assets/5944e25e-ab53-4aab-969b-7179e917b3e5" />



| **Tool** | Power BI Desktop, DAX, Power Query |
| **Dataset** | HRDataset_v14 (311 employees, 36 columns) |
| **Report file** | `HR.pbix` |
| **Pages** | Recruitment and turnover; Satisfaction and performance |
| **Canvas** | 1920 x 1080, fit to page |

## Overview

HR teams are asked the same questions every quarter. Who are we losing? Where do we find people who stay? Are the people who stay engaged, and does engagement show up in performance?

This report answers those questions from one dataset. Page one covers hiring and attrition. Page two covers satisfaction, engagement, absences and the spread of performance scores by manager. Both pages share the same four slicers, so a department head or a single manager can be isolated in one click.

## Objectives

1. Track recruitment and turnover by department, by recruitment source and by year.
2. Measure employee satisfaction and engagement, and test how both relate to performance.
3. Show how performance scores are distributed across managers.
4. Provide a base for forecasting hiring needs (see Roadmap).

## Data

The source is HRDataset_v14, one row per employee: 311 rows, 36 columns. 104 employees have left and 207 are active. The report uses the cleaned version of the file (see Cleaning). It uses these fields:

| Field | Used for |
|---|---|
| `EmpID` | Employee-level points in the scatter chart |
| `Department` | Turnover comparison, slicer |
| `RecruitmentSource` | Turnover by hiring channel |
| `EmploymentStatus` | Slicer (active and terminated staff) |
| `ManagerName` | Performance matrix, slicer |
| `PerformanceScore` | Engagement comparison, matrix columns |
| `EmpSatisfaction` | Satisfaction average and distribution |
| `EngagementSurvey` | Engagement average, scatter y-axis |
| `Absences` | Scatter x-axis |
| Hire and termination dates | Hires and terminations by year, through `DimDate` |

### Cleaning

The report runs on `HRDataset_v14_cleaned.xlsx`. The original file is untouched. The cleaning followed one rule: inspect first, change only what is clearly wrong, and log every change. The cleaned workbook has three sheets: Cleaned data (311 rows, 36 columns, same layout as the original), Cleaning log, and Flagged not changed.

**Inspection.** I checked sheet names, row and column counts, data types and blanks. I listed the unique values of every text column and tested for duplicate rows and duplicate IDs. The unique-value lists exposed the inconsistent entries below.

| Issue | Fix | Rows affected |
|---|---|---|
| Leading and trailing spaces in `Department`, `Sex`, `Position`, `Employee_Name` | Trimmed. Double spaces in names collapsed. Missing space added in "Alagbe,Trina" | 209, 135, 1, 91 |
| Mixed case in `HispanicLatino` ("yes", "Yes", "no", "No") | Made uniform | All mixed entries |
| Lowercase starts in `TermReason` ("career change") | Capitalised the first letter only. Wording unchanged | All affected entries |
| `Zip` stored as numbers, leading zero lost (01960 became 1960) | Stored as text, padded to five digits | 287 |
| `ManagerID` blank | Filled with 39. All eight rows belong to Webster Butler, whose other rows use 39 | 8 |
| Department, performance and status IDs that contradicted their text labels | Changed the ID to match the label | 6 |
| Date and sheet formatting | Dates stored as real dates (yyyy-mm-dd). Header row frozen, filters on | All date columns |

**ID corrections.** Each text label has a numeric code (`Department` and `DeptID`, `PerformanceScore` and `PerfScoreID`, `EmploymentStatus` and `EmpStatusID`). A cross-tab showed six rows where the code and label disagreed:

| Employee | Change |
|---|---|
| Randy Dee | `DeptID` 6 to 5, `PerfScoreID` 1 to 3 |
| Sean Quinn | `DeptID` 1 to 4 |
| Estelle Howard | `EmpStatusID` 1 to 4 |
| Alex Forrest | `EmpStatusID` 1 to 4, `PerfScoreID` 3 to 1 |

Treating the label as correct is a judgement call. For Randy Dee, the position and manager agreed with the label. Each change is recorded in the Cleaning log with the old and new value.

**Blanks.** `DateofTermination` is blank for 207 employees. These are active staff, so the blanks are correct and stay. Blanks with a reason are left. Blanks that can be proven from other rows are filled. Blanks that cannot be proven are flagged.

**Logic checks.** Four rules were tested:

- Every employee with `Termd` = 1 has a termination date, and nobody else does.
- `Sex` agrees with `GenderID`, and `MaritalDesc` agrees with `MaritalStatusID`.
- Termination dates fall after hire dates.
- Review dates fall after hire dates.

The first three passed. Two employees (EmpID 10142 and 10246) have a last review date before their hire date. The correct date cannot be known, so nothing was changed.

**Flagged, not changed.**

| Item | Handling in analysis |
|---|---|
| `PositionID` 13 covers three IT Manager roles. 23 covers both Shared Services Manager and Software Engineer | Group by `Position`, not `PositionID` |
| Michael Albert (IDs 22, 30) and Brandon R. LeBlanc (IDs 1, 3) each appear under two `ManagerID`s | Group by `ManagerID` where manager identity matters |
| Janet King's salary of 250,000 is valid but far above the rest | Analyse separately when averaging salary |
| `EmpStatusID` 2 and 3 appear on active employees, with no data dictionary | Use `EmploymentStatus` text instead |
| "On-line Web application" and "Website" may be one recruitment channel | Merge if the analysis needs it |

## Model

The model has three tables.

| Table | Role |
|---|---|
| `HR` | Employee-level fact table |
| `DimDate` | Calendar table, 2006 to 2018, marked as the date table. Supplies `Year` for the slicer and the line chart |
| `Measure` | Holds the report measures, kept apart from the data columns |

`DimDate` is built in DAX:

```dax
DimDate =
ADDCOLUMNS (
    CALENDAR ( DATE ( 2006, 1, 1 ), DATE ( 2018, 12, 31 ) ),
    "Year", YEAR ( [Date] ),
    "Month number", MONTH ( [Date] ),
    "Month start", DATE ( YEAR ( [Date] ), MONTH ( [Date] ), 1 )
)
```

Two relationships run from `DimDate[Date]`, both many-to-one with a single filter direction:

| To | State | Used by |
|---|---|---|
| `HR[DateofHire]` | Active | Hires |
| `HR[DateofTermination]` | Inactive | Terminations, through `USERELATIONSHIP` |

The inactive link lets hires and exits share one date axis in the line chart. A separate measure table keeps the field list readable and makes each calculation easy to audit.

## Measures

| Measure | Definition |
|---|---|
| Total employees | Count of employee records in the current filter context |
| Terminated | Employees flagged `Termd` = 1 |
| Turnover % | Employees who have left, divided by everyone ever employed |
| Hires | Employees by hire date. Works on the date axis because the hire relationship is active |
| Terminations | Employees by termination date, through the inactive relationship |
| Avg satisfaction | Mean of `EmpSatisfaction` |
| Avg engagement | Mean of `EngagementSurvey` |

```dax
Total employees = COUNTROWS ( HR )

Terminated = CALCULATE ( [Total employees], HR[Termd] = 1 )

Turnover % = DIVIDE ( [Terminated], [Total employees] )

Hires = COUNTROWS ( HR )

Terminations =
CALCULATE (
    COUNTROWS ( HR ),
    USERELATIONSHIP ( HR[DateofTermination], DimDate[Date] )
)

Avg engagement = AVERAGE ( HR[EngagementSurvey] )

Avg satisfaction = AVERAGE ( HR[EmpSatisfaction] )
```

**Read Turnover % with care.** It is a cumulative share: 104 employees who have left out of 311 ever employed, or 33.4%. It is not the annual rate that HR teams usually quote. An annual rate needs a headcount at the end of each period. That measure is in the Roadmap.

## Pages

### Recruitment

**Page 1: Recruitment and turnover.** Headline cards run across the top: Total employees, Turnover %, Avg engagement and Avg satisfaction. Four slicers sit beside them: Department, Employment status, Manager and Year.

| Visual | Type | Question it answers |
|---|---|---|
| Turnover % by department | Clustered bar | Which departments lose the most people? |
| Turnover % by recruitment source | Clustered bar | Which hiring channels produce staff who leave? |
| Hires and terminations by year | Line | Is headcount growing, flat or shrinking? |
| Avg engagement by performance score | Clustered bar | Do higher performers report higher engagement? |

All bar charts are sorted in descending order, so the largest value sits first.

<img width="882" height="489" alt="CodeAlpha HR Pg 1" src="https://github.com/user-attachments/assets/a2f3bf13-72e7-4242-bc85-e2d3a1fbb01c" />


### Satisfaction

**Page 2: Satisfaction and performance.** The same cards and slicers carry over, so context does not change when the reader switches page.



| Visual | Type | Question it answers |
|---|---|---|
| Employees by manager and performance score | Matrix | How is each manager's team spread across performance ratings? |
| Employees by satisfaction score | Clustered column | How are employees spread across the 1 to 5 satisfaction scale? |
| Avg engagement by performance score | Clustered bar | Repeats the page 1 view under the page 2 filters |
| Absences and engagement by employee | Scatter | Do employees with more absences report lower engagement? |

<img width="882" height="494" alt="CodeAlpha HR Pg 2" src="https://github.com/user-attachments/assets/e68cbd4f-1967-49c1-8a7b-e1ae77fb578c" />

## Design

- Each page uses a 1920 x 1080 canvas set to fit the page.
- Headline numbers come first, then slicers, then charts. A reader gets the answer before the detail.
- Chart titles use Cambria. The page header uses Tahoma. The base theme is Fluent 2.
- Every bar chart has a title that states the measure and the dimension, so no chart needs a legend to be read.
- The matrix on page 2 doubles as a manager view. Selecting a manager in the slicer filters every chart on the page.

## Findings

Figures come from the report with no slicers applied.

| Question | Finding |
|---|---|
| Overall turnover | 104 of 311 employees (33.4%) have left. 207 are active. |
| Department | Production has the highest turnover at 39.7%. Sales is at 16.1%. |
| Recruitment source | Google Search hires leave at 61.2% and Diversity Job Fair hires at 55.2%. Employee Referral hires leave at 16.1%. Website hires leave at 7.7%, but there are only 13 of them. |
| Hiring and exits over time | Hires fell from 83 in 2011 to 1 in 2018. Terminations peaked at 23 in 2015. Headcount is shrinking. |
| Engagement and performance | Engagement falls from 4.48 for Exceeds to 2.22 for PIP. It falls from 4.24 for Fully Meets to 2.99 for Needs Improvement. |


Sample sizes in the smaller groups are low, so read those rates as direction, not precision.

## Recommendations

1. **Shift sourcing toward referrals.** Employee Referral hires leave at 16.1%. Google Search hires leave at 61.2%. Compare cost per hire by channel once that data exists.
2. **Run exit conversations in Production.** Turnover is 39.7%, against 16.1% in Sales. Find out whether the cause is pay, workload or management before spending on retention.
3. **Treat low engagement as an early flag.** Engagement drops with every step down in performance rating. The data shows a link, not a cause, so use it to prompt a conversation and not to judge a person.
4. **Set a headcount target.** Hiring has almost stopped while exits continued. Decide the size the business wants before forecasting hiring needs.

## Limitations

- The data has gaps that cleaning could not close. Two employees have a review date before their hire date, and `EmpStatusID` codes 2 and 3 have no dictionary. The report uses `EmploymentStatus` text for that reason.
- The dataset has no time-to-fill, cost-per-hire or vacancy data. Recruitment metrics here mean volume, source mix and retention by source.
- The dataset ends in 2018, and the 2018 hire count of 1 is partial. Read the trend up to 2017.
- The dataset has 311 employees. Department, source and manager splits produce small groups, so a single exit can move a rate by several points. Executive Office has one employee, so its rate is 0% or 100%.
- The report groups by `ManagerName`. Michael Albert and Brandon R. LeBlanc each appear under two `ManagerID`s, so each name is read as one manager.
- "On-line Web application" and "Website" may be one recruitment channel. Unless they are merged in Power Query, the turnover by source chart shows them as separate bars.
- Satisfaction and engagement are single scores per employee. They show a position at one time, not a trend. Their links to performance and absences are correlations.
- The report does not include a forecast.

## Roadmap

**Annual turnover rate.** Terminations divided by headcount at the end of the period gives a rate that can be compared with other companies.

```dax
Headcount end of period =
VAR d = MAX ( DimDate[Date] )
RETURN
    CALCULATE (
        COUNTROWS ( HR ),
        REMOVEFILTERS ( DimDate ),
        FILTER (
            HR,
            HR[DateofHire] <= d
                && ( ISBLANK ( HR[DateofTermination] ) || HR[DateofTermination] > d )
        )
    )

Annual turnover % = DIVIDE ( [Terminations], [Headcount end of period] )
```



The card reads as the last 12 months of exits plus growth. A manager can move the slider to see how 5% or 10% expansion changes the number. With 311 rows and few events per month, the forecast is an estimate and needs its confidence band shown.


## Repository

```
hr-analytics-dashboard/
├── HR.pbix
├── README.md
├── data/
│   ├── HRDataset_v14.xlsx            (original, untouched)
│   └── HRDataset_v14_cleaned.xlsx    (report source: data, cleaning log, flagged items)
└── images/
    ├── 01-recruitment-turnover.png
    └── 02-satisfaction-performance.png
```

## Author

[Alexander Lilian]
Data Analyst
[linkedin.com/in/lilian-alexander] 
