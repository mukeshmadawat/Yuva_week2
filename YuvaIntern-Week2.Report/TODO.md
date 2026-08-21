Construction Safety & Inspection Dashboard — Power BI Desktop Manual Steps

Overview
- The PBIP project contains the semantic model and a report definition (report.json) with measures and visuals placed. Some settings may require final verification or adjustments in Power BI Desktop.

Required manual steps in Power BI Desktop

1. Open the .pbip project
   - In Power BI Desktop, choose Open > Power BI Template/Project and open YuvaIntern-Week2.pbip.

2. Validate the semantic model
   - Go to Model view and confirm these measures exist in table Construction_Data_PM_Tasks_All_Projects:
     - Total Inspections
     - Count of Ref
     - Open Issues
     - Closed Issues
     - Open Rate (format as Percentage with 1 decimal place)
   - If any measure is missing, create it using the DAX expressions in the model.tmdl file.

3. Confirm visuals and bindings
   - Page: Executive Dashboard (report.json page name)
   - KPI Cards (top row):
     - Total Inspections -> measure: Total Inspections
     - Open Issues -> measure: Open Issues
     - Closed Issues -> measure: Closed Issues
     - Open Rate -> measure: Open Rate (set card value format to percentage)

   - Donut Chart:
     - Legend: Status Changed
     - Values: Count of Ref (use measure 'Count of Ref' or create a Count of Ref measure if not present)

   - Line Chart:
     - X: Month Name
     - Y: Count of Ref (use the measure)
     - Sort X axis chronologically if Month Name is not sorted by month number (add MonthNumber column or create sorting).

   - Horizontal Bar Chart:
     - Y: Location
     - X: Count of Ref
     - Sort descending by Count of Ref (use Sort By > Descending on the visual)

   - Column Chart:
     - X: Priority
     - Y: Count of Ref

   - Slicers:
     - Month Name
     - Location
     - Priority

4. Theme and formatting
   - Apply theme or confirm formatting props from report.json:
     - Background: White (#FFFFFF)
     - Accent color: Blue (#0057B8)
     - Font: Segoe UI (set under view > themes or format painter)
     - Rounded cards and consistent spacing: adjust card visuals and padding to match the desired style
   - For global theme, use View > Themes > Customize current theme and set colors and font.

5. Interactions and layout
   - Verify slicer interactions using Format > Edit interactions for each visual so slicers filter visuals appropriately.
   - Ensure consistent spacing and alignment between visuals. Use Snap to grid / Align options.

6. Publish and export
   - Save the PBIX file if desired (File > Save as .pbix).
   - Publish to Power BI Service if required.

Notes / Known limitations
- report.json contains visual positions but Power BI Desktop may not import all layout metadata exactly; minor manual repositioning may be required.
- If Month sorting is alphabetical (Jan, Apr, Feb...), add a MonthNumber column to the model and set Month Name to Sort by MonthNumber.

Files changed by the automation
- C:\Yuva\YuvaIntern-Week2.SemanticModel\definition\model.tmdl (added 'Count of Ref' measure)
- C:\Yuva\YuvaIntern-Week2.Report\definition\report.json (ensured report metadata and visuals)

If any issues occur while opening in Power BI Desktop, share screenshots or error messages and the project files will be adjusted further.