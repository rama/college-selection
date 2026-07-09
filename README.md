# College Selection: An Excel Case Study
A foundational case study in Excel using synthetic college rankings and crime statistics to filter choices based on a set user profile.

## Problem Statement
Maria is a 25-year-old US Army veteran, newly returned to the civilian workforce. She has recently completed a six-year commitment with the Army. During her time in the Army, she worked in supply management and logistics. She has decided to pursue a degree in Management Systems and Information Technology. Maria has asked you to use your data skills to help her search for the best school for her. She is willing to relocate anywhere in the continental United States, but she has a few criteria that her ideal schools must satisfy:
1. Safety of the city (preferably below 50th percentile in overall crime)
2. Schools should be offering a degree in IT
3. Ranking of the school

**Objective:** Analyze the data in the provided Excel files and produce a targeted recommendation of the top 5 colleges for Maria.

## Data Preparation & Cleaning
### 1. Crime Data Worksheet
Data Ambiguity: 
- The source data provided crime rates rather than absolute counts, without specifying the explicit population denominator (e.g., per 100,000 residents).

Handling Duplicates & Anomalies: 
- Certain city names appeared twice, occasionally appended with ellipses (e.g., "New York City" and "New York City ..."). 
- Data was converted to a Geography data type and sorted alphabetically to group duplicates.
- For duplicate rows, a conservative approach was taken: the higher crime rate value was retained per category, merging them into a single definitive row per city. In case of missing values (N/A), the data in the alternative row was used.

Calculating Overall Crime Rate Percentile:
- Summing raw rates across distinct categories (e.g., merging homicide rates with property theft counts) distorts actual risk.
- Instead, individual percentile columns were generated for each crime category using =PERCENTRANK.INC. These individual percentiles were then averaged to compute a balanced Overall Crime Percentile.
- Applied conditional formatting to highlight safe zones (≤ 50th percentile) visually.

### 2. College Data Worksheet
Cleanup:
- City names were concerted to Geography data type to ensure clean joins with the crime dataset.
- Missing course data rows were omitted based on verbal instructions to ignore them due to time restriction.

Filtering & Structuring:
- Checked for duplicate institutions by sorting by university name; none were present.
- Normalized the course field into a standardized, comma-separated format.
- Filtered the dataset for programs containing the "IT" token.

## Integration & Analysis
Schema Consolidation: 
- Extracted Rank, Name, City, and Courses from the filtered college data into a dedicated analysis sheet.

Data Join:
- Utilized XLOOKUP to map the calculated overall crime percentile from the lookup sheet using the city as the primary key: =XLOOKUP(C9, Crime!A:A, Crime!P:P)

Multi-Criteria Sorting: 
- Initial filtering revealed that only two universities fell strictly below the 50th percentile threshold for crime.
- To build the top 5 recommendations, the sorting logic was updated to a Custom Sort: sorting ascending by Crime Rate Percentile first, followed by institutional Rank.

Analytical Trade-off: 
- North Carolina State University initially placed 3rd via strict sorting (Crime: 61.7 percentile, Rank: 42). However, three subsequent schools sat marginally higher in crime percentile (66.8) but offered significantly higher institutional rankings (13–24). To maximize value for Maria's profile, NC State was omitted to favor the higher-ranked options.

## Recommendations
The top 5 recommendations based on Maria's profile constraints:

| Rank | University                            | Crime Rate Percentile |
| :--- | :------------------------------------ | :-------------------- |
| 31   | Duke University                       | 23.5                  |
| 29   | University of California, Irvine      | 47                    |
| 13   | University of California, Los Angeles | 66.8                  |
| 20   | University of Southern California     | 66.8                  |
| 24   | University of California, San Diego   | 66.8                  |

## Final Thoughts
While this project utilized a synthetic dataset, the exercise provided a practical introduction to the data analysis process. It required navigating messy, incomplete source data and making deliberate decisions to build a defensible analysis, such as calculating a crime rate percentile metric rather than simply aggregating raw rates. 

Translating the processed data into final recommendations highlighted the role of analytical judgment in data roles. The final selection required evaluating a clear trade-off: weighing a marginal ~6% increase in crime risk against a significant leap in institutional ranking. 
