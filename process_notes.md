# Process Notes
**Datasets:** Two excel files, one with college ranking data and another with crime statistics.
**Initial setup:** Copied the raw data from the two files into separate worksheets (Crime and Colleges) within a unified workbook.
## Crime data
### Observations
- The crime statistics appear to be rates, not counts. However, it is unclear what the denominator is, i.e., is it per 10,000, 100,000, etc?
- Some city names appear twice and one of the two is followed by an ellipsis. For example, "New York City" and "New York City ..." Pittsburg, however, appears as "Pittsburg ..." twice.
### Steps I took
1. Converted city to data type geography.
2. Sorted alphabetically based on city to easily identify the duplicate cities.
3. In case of multiple rows for same city, I opted to use the higher value for each type of crime and merged into one row per city. There were a couple missing values under Pittsburg but since there was a duplicate for Pittsburg, I considered the N/A to be a 0 and used the other value.
4. I considered adding up the individual category crime rates to calculate overall crime. However, adding, for example, 50 murders with 3000 incidents of property theft to arrive at 3050 is an inaccurate measure of risk in either category. 
5. So I created corresponding percentile columns for each crime category and calculated the percentile using PERCENTRANK.INC. I then averaged these percentiles to calculate the overall crime percentile.
6. A further consideration which I did not pursue at this time is to weight different crimes, possibly based on "customer" preference. For example, murder or violent crime could be weighted higher than property theft when calculating overall crime. 
7. Used Conditional Formatting to highlight the cities at below 50th percentile in green.
## College Data
### Observations
- Some city names are followed by ellipses, similar to the crime dataset.
- Course data is missing for some universities. Ideally, I would fill this in, but for the purposes of this exercise we were told to ignore them.
### Steps I took
1. Sorted by University Name to check for duplicates. None.
2. Changed the course field to be comma-separated for consistency. 
3. Filtered on the courses column for "IT". Ideally, I would like to treat each item as a distinct value for easy filtering, which would require splitting into multiple columns or tables. Since this is a simpler analysis, I opted to keep it as is.
In both excel sheets, I also did some formatting changes like making the headers bold, consistent capitalization, freezing the top row, etc.
## Putting everything together
1. Copied over relevant columns (Rank, Name, City, Courses) of my filtered data from the Colleges sheet into a new sheet.
2. Added a Crime Rate Percentile column and used XLOOKUP to pull the crime rate percentile from the Crime sheet: =XLOOKUP(C9, Crime!A:A, Crime!P:P)
3. Sorted ascending for Crime Rate Percentile. Two colleges are located in a city below 50th percentile in crime: Duke University in Durham, North Carolina, and University of California in Irvine, California.
4. Since the requirements also mention ranking of the school, I changed the sort to a custom sort, sorting on Crime Rate Percentile first and then Rank.
5. North Carolina State University places 3rd with a crime rate percentile of 61.7 and a ranking of 42 while the next 3 colleges place below it due to their slightly higher crime rate percentile of 66.8 even though the ranking is much higher at 13-24. So I opted to remove North Carolina State University from the top 5.

The top 5 schools at the end of the analysis are:

| Rank | University                            | Crime Rate Percentile |
| :--- | :------------------------------------ | :-------------------- |
| 31   | Duke University                       | 23.5                  |
| 29   | University of California, Irvine      | 47                    |
| 13   | University of California, Los Angeles | 66.8                  |
| 20   | University of Southern California     | 66.8                  |
| 24   | University of California, San Diego   | 66.8                  |

## Final Thoughts

While this isn't a fully real-world dataset, the exercise was a useful introduction to the data analysis process. I got to think about how to approach messy or incomplete data and the decisions required to build a defensible analysis, such as creating the crime rate percentile rather than just summing raw rates. The final step of providing recommendations based on the data was interesting as well, as it required analytical judgment to weigh a ~6% increase in crime rate percentile against a significantly higher college ranking. 
