# Detailed Project Report on Coffee Sales Analysis Useing Excel
## Executive Summary
This project report encapsulates the comprehensive analysis of coffee sales using a dataset composed of three distinct sheets: "Orders," "Product," and "Customer." The objective was to integrate the data effectively to derive insights into sales performance, customer behavior, and product popularity. Various Excel techniques, including Power Query, XLOOKUP, INDEX-MATCH, and Pivot Tables, were employed to create a cohesive overview of the coffee sales landscape. A dashboard was developed to visualize key metrics, making it easier for stakeholders to understand the findings.
## 1. Introduction
The primary goal of this project was to analyze coffee sales data to understand sales trends, customer demographics, and product performance. By merging information from the "Orders," "Product," and "Customer" sheets, insights can be derived that foster better decision-making and targeted marketing strategies.
## 2. Data Integration
### 2.1 XLOOKUP for Customer Information
To enhance the "Orders" sheet, I utilized the XLOOKUP function to pull relevant customer data from the "Customer" sheet. After linking each order with the respective customer, the following fields were integrated:
* Customer Name

  [=XLOOKUP(C2,customers!$A$1:$A$1001,customers!$B$1:$B$1001,,0)]
* Email ID

  [ =IF(XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0) =0,"",XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0)) ]
* Country

  [=XLOOKUP(C2,customers!$A$1:$A$1001,customers!$G$1:$G$1001,,0)]

This integration was crucial for analyzing customer preferences and sales patterns.
### 2.2 INDEX-MATCH for Product Details
Subsequently, I employed the INDEX-MATCH technique to retrieve product details from the "Product" sheet. The information extracted included:
* Coffee Type
* Roast Type
* Size
* Unit Price

  [=INDEX(products!$A$1:$G$49,MATCH(Orders!$D10,products!$A$1:$A$49,0),MATCH(Orders!I$1,products!$A$1:$G$1,0))]

This step allowed for a detailed understanding of which products were generating sales and at what price points.
## 3. Pivot Tables and Data Visualization
### 3.1 Sales Analysis by Coffee Type
I created a pivot table to summarize total sales of different coffee types by year and month. The following steps were taken:
* Selected the integrated data sheet.
* Inserted a Pivot Table and set the Row Labels as Coffee Type and the Column Labels as Months, with values aggregated as total sales.
### Line Chart Creation
A line chart was generated from this pivot table, providing a visual representation of sales trends over time.

### 3.2 Top Customers Analysis
In another pivot table, I analyzed the top 5 customers based on total sales. The process included:
* Highlighting the total sales data.
* Setting the Row Labels as Customer Names and values as total sales, filtering to show only the top 5 customers.
### Bar Chart Creation
From this pivot table, a bar chart was created to visually depict the contribution of each top customer to overall sales.

### 3.3 Country-wise Sales Analysis
A third pivot table was constructed to analyze total sales by country. The steps included:
* Using the integrated dataset and selecting Country as Row Labels and total sales as Values.
### Bar Chart Creation
A corresponding bar chart was produced to showcase the geographical distribution of coffee sales.

## 4. Dashboard Creation
A comprehensive dashboard was created to encapsulate all the visualizations and insights from the analysis:
### Visualizations Included:
* Line chart showing coffee sales trends by type.
* Bar charts displaying top customers and country-wise sales.
### Interactivity:
* Implemented slicers for Roast Type, Loyalty Card status, and Size of coffee in kilograms.
* Added a timeline slicer for year and month filtering.

The dashboard offers an interactive platform for stakeholders to explore sales data dynamically, thereby aiding in more informed decision-making.

## 5. Conclusion
The integration and analysis of the coffee sales dataset have yielded valuable insights into sales performance across various dimensions. The use of advanced Excel functions and visualization techniques has enabled a clearer understanding of customer preferences and sales trends.

## 6. Recommendations
Based on the analysis, I recommend:
* Targeted Marketing Campaigns: Focus on top customers and high-performing coffee types to enhance sales further.
* Geographic Expansion: Consider increasing marketing efforts in countries showing significant sales potential.
* Product Development: Explore customer preferences in roast types and sizes to guide product diversification.

By adhering to these recommendations, the business can optimize its coffee sales strategy and improve overall performance. This report provides a detailed overview of the methodology, processes, and outcomes of the coffee sales analysis project, ensuring all relevant data and insights are presented in a structured and actionable manner.
