# Excel Sales Dashboard


### Table of contents
- [project overview](#project-overview)
- [Data Sources](#data-sources)
- [Recommmendation](#recommendation)
  
  

### Project Overview
 This project is an interactive sales dashboard built in Microsoft Excel to analyze sales performance across different products, countries, 
and time periods. The dashboard transforms raw sales data into meaningful insights using Pivot Tables, Pivot Charts, Slicers, and KPI cards.

![Dashboard](bar plot,PNG)<img width="1296" height="538" alt="screenshortdash" src="https://github.com/user-attachments/assets/d1e46dcb-32fc-4459-a062-16d7e8318478" />


![Uploading bar plot.PNG]()




### Data Sources
Kaggle Sample Sales Dataset (used for educational and portfolio purposes).


### Tools Used
- Excel
- Pivot Tables
- Pivot Charts
- Slicers
- KPI Card
- Shapes and Icons for dashboard design
- Data Formatting and Custom Number Formats

### Data Cleaning/Preparation

Before creating the dashboard, the dataset was cleaned and prepared in Microsoft Excel through the following steps:

-  Imported the sales dataset into Excel.
-  Checked for missing values and data inconsistencies.
-    Removed duplicate records where necessary.
-   Formatted the Date column and grouped dates by month and year.
-       Standardized number formats for sales, prices, and quantities.
-  Verified data types to ensure numeric fields were correctly recognized by Excel.
-  Created Pivot Tables to summarize the data by month, country, and product.
-  Developed calculated metrics, including Total Revenue and Total Quantity Sold.
-  Designed and formatted interactive dashboard elements such as KPI cards, charts, and slicers.


### Exploratory Data Analysis
The dataset was explored to understand sales patterns and identify key business insights. The following analyses were performed:

-  Examined the distribution of sales across different months to identify trends and seasonal patterns.
-  Analyzed total revenue and quantity sold to measure overall business performance.
-  Compared sales performance across different countries.
-  Evaluated product performance to identify top-selling and low-performing products.
-  Used Pivot Tables to summarize and aggregate sales data by month, country, and product.
-  Applied slicers to create interactive filtering and enable dynamic exploration of the data.
-  Identified key performance indicators (KPIs), including Total Revenue and Total Quantity Sold, for dashboard reporting.


### Data Analysis
Interactive Elements: The dashboard uses Pivot Tables, Pivot Charts, Slicers, KPI cards, and linked text boxes to provide a dynamic and user-friendly reporting experience without requiring VBA or programming languages
=SUM(range)

=SUMIFS(sum_range, criteria_range, criteria)

=COUNT(range)

=COUNTA(range)

=pivot!B13

### Results /Findings
The interactive sales dashboard provided valuable insights into sales performance across products, countries, and time periods.
  Total revenue and quantity sold were successfully summarized using KPI cards.
* Monthly sales analysis revealed fluctuations in sales performance, indicating periods of higher and lower demand.
* Sales performance differed across countries, with some countries contributing significantly more revenue than others.
* Product analysis identified top-performing products based on revenue and quantity sold.
* Interactive slicers enabled users to dynamically filter data by date, country, and product, making it easier to explore specific business scenarios.
The dashboard transformed raw sales data into clear, interactive visualizations that support:
* Product and market performance evaluation
* Trend analysis and decision-making
* Faster access to key business metrics through interactive filtering and KPIs


  ### Recommendation

  
Based on the analysis and insights generated from the interactive sales dashboard, the following recommendations are suggested:

1. Focus on top-performing products by maintaining adequate inventory levels and implementing targeted marketing strategies.
2. Investigate low-performing products and countries to identify factors affecting sales performance and develop improvement plans.
3. Monitor monthly sales trends to identify seasonal patterns and prepare for periods of high and low demand.
4. Use interactive dashboard filters regularly to analyze sales performance across different products, countries, and time periods.
5. Continuously update the dashboard with new sales data to support real-time monitoring and decision-making.

   ###Limitations
   Dataset Scope: The analysis was conducted using a sample sales dataset, which may not fully represent the complexity and scale of real-world business operations.
-  Data Quality Dependence: Dashboard insights are highly dependent on the accuracy and completeness of the source data. Any missing, inconsistent, or erroneous records may affect the reliability of the results.
- Manual Data Refresh: The dashboard requires manual data updates and pivot table refreshes, limiting its ability to provide real-time reporting and automated data synchronization.
-  Excel Scalability Constraints: While Excel is effective for small- to medium-sized datasets, performance may decline when handling large volumes of data or complex calculations.
-  Limited Advanced Analytics: The dashboard focuses on descriptive analytics and KPI monitoring. Predictive analytics, forecasting, and machine learning capabilities were not implemented.
-  Tool Limitations: Compared to business intelligence platforms such as Power BI and Tableau, Excel offers fewer
-  
  ### References

1.  Microsoft Corporation. (2024). Microsoft Excel Documentation. Used for data cleaning, Pivot Tables, Pivot Charts, slicers, and dashboard development.
2. Kaggle. (2024). Sample Sales Dataset. Used as the primary data source for building and testing the interactive sales dashboard.
3. Microsoft Support. (2024). Create a PivotTable and PivotChart in Excel. Referenced for creating data summaries and interactive visualizations.




  
