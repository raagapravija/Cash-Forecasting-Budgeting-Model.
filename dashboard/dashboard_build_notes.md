# Dashboard Build Notes

## Excel

Use the included `Dashboard` sheet as the starting point. If rebuilding manually:

1. Create KPI cards linked to `KPI_Summary` and `Cash_Flow`.
2. Add a line chart for Cash_Flow Base Case `Total Cash Inflows` and `Total Cash Outflows`.
3. Add a line chart for Base/Best/Worst Closing Cash and a horizontal Minimum Reserve line.
4. Add a line chart for Revenue_Forecast Base/Best/Worst Revenue.
5. Add a clustered column chart for Budget vs Actual by department.
6. Add AR Aging and AP charts.
7. Add a scenario summary with Closing Cash, Minimum Cash and Alert.
8. Add conditional formatting: green, yellow and red for Healthy, Watch and Critical.
9. Add PivotTable slicers for Month, Quarter, Department, Cost Category, Scenario and Transaction Type.

## Power BI

Load `data/cash_flow_transactions.csv` and `forecast/monthly_forecast_snapshot.csv`. Create a Date table and measures for Inflow, Outflow, Net Cash Flow, Budget Variance %, AR, AP, DSO, DPO and Closing Cash. Use drill-through from department to transaction and add a page tooltip for management commentary.

## Design standard

Use deep navy `#17324D`, blue `#1F5A85`, teal `#0E9F9A`, amber `#F4B942`, red `#D64545` and light grey backgrounds. Every visual should have a title, unit and clear business purpose.
