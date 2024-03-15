# VBA Stock Challenge

## Objective:

This project involves creating a VBA script to analyze stock data from multiple worksheets representing different years. The script calculates various metrics such as yearly change, percentage change, and total stock volume for each stock. Additionally, it identifies the stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume. The analysis is performed on the provided dataset "alphabetical_testing.xlsx."

## Implementation:

## Retrieval of Data:

The VBA script loops through each year of stock data and retrieves the following values from each row: ticker symbol, volume of stock, opening price, and closing price.

## Column Creation:

A new worksheet is created for each year or columns are added to the existing worksheet to store the calculated metrics: ticker symbol, total stock volume, yearly change, and percent change.

## Conditional Formatting:

Conditional formatting is applied to the yearly change column to highlight positive change in green and negative change in red. Similar formatting is applied to the percent change column.

## Calculated Values:

The script accurately calculates and displays the following values:
Greatest Percentage Increase
Greatest Percentage Decrease
Greatest Total Volume

## Looping Across Worksheets:

The VBA script is designed to run successfully on all sheets, ensuring consistent analysis across different years of stock data.

## Results

The VBA script effectively analyzes stock data, providing valuable insights into each stock's performance over the years. Users can easily identify stocks with significant changes in price and trading volume, aiding investment decision-making processes.

## Conclusion:

Through this VBA stock analysis project, users can efficiently analyze and visualize stock performance across multiple years.
