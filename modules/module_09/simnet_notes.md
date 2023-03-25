# Module 9 - Getting and Managing Data

<!-- TOC -->
* [Module 9 - Getting and Managing Data](#module-9---getting-and-managing-data)
* [General Notes](#general-notes)
  * [Chapter Overview](#chapter-overview)
  * [Student Learning Outcomes `(SLOs`)](#student-learning-outcomes-slos-)
  * [Case Study](#case-study)
  * [Conclusion](#conclusion)
* [SLO 4.1 - Creating and Formatting an Excel Table](#slo-41---creating-and-formatting-an-excel-table)
  * [How to Create an Excel Table](#how-to-create-an-excel-table)
  * [How to Apply a Table Style](#how-to-apply-a-table-style)
  * [How to Display a Total Row in a Table](#how-to-display-a-total-row-in-a-table)
  * [The Table Tools and Properties Groups](#the-table-tools-and-properties-groups)
    * [Tools Group](#tools-group)
    * [How to Name the Table and Remove Duplicates](#how-to-name-the-table-and-remove-duplicates)
    * [How to Insert a Slicer](#how-to-insert-a-slicer)
    * [Properties Group](#properties-group)
    * [How to Convert a Table to a Range](#how-to-convert-a-table-to-a-range)
    * [Structured References and Table Formulas](#structured-references-and-table-formulas)
    * [How to Add a Calculation Column in a Table](#how-to-add-a-calculation-column-in-a-table)
    * [Additional Information](#additional-information)
  * [Conclusion](#conclusion-1)
* [SLO 4.2 - Applying Conditional Formatting](#slo-42---applying-conditional-formatting)
  * [Highlight Cells Rules](#highlight-cells-rules)
    * [How to Create a \"Less Than\" Highlight Cells Rule](#how-to-create-a-less-than-highlight-cells-rule)
  * [Top/Bottom Rules](#topbottom-rules)
    * [How to Create a Top 10% Rule](#how-to-create-a-top-10-rule)
  * [Use a Formula for a Rule](#use-a-formula-for-a-rule)
    * [How to Use a Formula in a Conditional Formatting Rule](#how-to-use-a-formula-in-a-conditional-formatting-rule)
  * [Data Bars, Color Scales, and Icon Sets](#data-bars-color-scales-and-icon-sets)
    * [How to Format Data with an Icon Set](#how-to-format-data-with-an-icon-set)
    * [How to Manage Conditional Formatting Rules](#how-to-manage-conditional-formatting-rules)
  * [Conclusion](#conclusion-2)
* [SLO 4.3 - Sorting Data](#slo-43---sorting-data)
  * [Sorting](#sorting)
  * [Sort Data in an Excel Table](#sort-data-in-an-excel-table)
    * [How to Sort Data in an Excel Table](#how-to-sort-data-in-an-excel-table)
  * [Sort Options](#sort-options)
  * [Sort Data by One Column](#sort-data-by-one-column)
  * [Sort Data by Multiple Columns (Multi-level sort)](#sort-data-by-multiple-columns--multi-level-sort-)
  * [How to Use the Sort dialog box](#how-to-use-the-sort-dialog-box)
  * [Sort Data by Cell Attribute (font color, cell fill color, conditional formatting icon)](#sort-data-by-cell-attribute--font-color-cell-fill-color-conditional-formatting-icon-)
  * [Conclusion](#conclusion-3)
* [4.4 - Filtering Data](#44---filtering-data)
    * [AutoFilters](#autofilters)
    * [Custom AutoFilter](#custom-autofilter)
    * [How to create a Custom Text AutoFilter](#how-to-create-a-custom-text-autofilter)
  * [Advanced Filter](#advanced-filter)
    * [How to create an Advanced Filter](#how-to-create-an-advanced-filter)
  * [Conclusion](#conclusion-4)
* [SLO 4.5 - Using Subtotals, Groups, and Outlines](#slo-45---using-subtotals-groups-and-outlines)
  * [Outline](#outline)
  * [Subtotal](#subtotal)
  * [Group](#group)
  * [The Subtotal Command](#the-subtotal-command)
    * [How to Display Subtotals](#how-to-display-subtotals)
  * [Outline Buttons](#outline-buttons)
    * [How to Use Outline Buttons](#how-to-use-outline-buttons)
  * [Auto Outline](#auto-outline)
    * [How to Create an Auto Outline](#how-to-create-an-auto-outline)
  * [Define Groups](#define-groups)
    * [How to Define a Group](#how-to-define-a-group)
  * [Conclusion](#conclusion-5)
* [SLO 4.6 - Importing and Exporting Data](#slo-46---importing-and-exporting-data)
  * [Importing Data in Excel](#importing-data-in-excel)
  * [Tab-Delimited Text Files](#tab-delimited-text-files)
  * [Importing Tab-Delimited Text File](#importing-tab-delimited-text-file)
  * [Comma-Separated Values `(CSV)` Files](#comma-separated-values-csv-files)
  * [Opening a` CSV `File](#opening-a--csv--file)
  * [Importing` CSV `Data with Power Query](#importing--csv--data-with-power-query)
  * [Transforming Data in Power Query](#transforming-data-in-power-query)
  * [Importing Access Database Files](#importing-access-database-files)
  * [Importing Access Table into a Worksheet](#importing-access-table-into-a-worksheet)
  * [Workbook Queries and Connections](#workbook-queries-and-connections)
    * [Managing External Data](#managing-external-data)
    * [Importing Data from Word Documents](#importing-data-from-word-documents)
    * [Importing Website Data](#importing-website-data)
    * [Exporting Data as a Text File](#exporting-data-as-a-text-file)
    * [Exporting Data via Clipboard](#exporting-data-via-clipboard)
    * [Exporting Data to SharePoint Lists](#exporting-data-to-sharepoint-lists)
  * [SLO 4.7: Transforming Data in Excel](#slo-47--transforming-data-in-excel)
    * [Flash Fill New Data](#flash-fill-new-data)
      * [How to Flash Fill Data in a Worksheet](#how-to-flash-fill-data-in-a-worksheet)
    * [Flash Fill to Convert Case](#flash-fill-to-convert-case)
      * [How to Flash Fill and Convert Case](#how-to-flash-fill-and-convert-case)
    * [Text to Columns](#text-to-columns)
      * [How to Convert Text to Columns](#how-to-convert-text-to-columns)
  * [Conclusion](#conclusion-6)
* [SLO 4.8 - PivotTables and PivotCharts](#slo-48---pivottables-and-pivotcharts)
  * [PivotTables](#pivottables)
    * [Creating a PivotTable](#creating-a-pivottable)
    * [The PivotTable Fields Pane](#the-pivottable-fields-pane)
    * [Field Settings](#field-settings)
    * [Formatting a PivotTable](#formatting-a-pivottable)
  * [PivotCharts](#pivotcharts)
    * [Creating a PivotChart](#creating-a-pivotchart)
  * [Conclusion](#conclusion-7)
* [Chapter Summary](#chapter-summary)
<!-- TOC -->

# General Notes

## Chapter Overview

- Excel can use data from many sources as well as provide data to other
  programs.
- Shared data among applications is usually in a list or table layout like a
  database.
- This chapter covers how to:
    - Format data as an Excel table.
    - Sort and filter data.
    - Import data from other sources.
    - Build a PivotTable.

## Student Learning Outcomes `(SLOs`)

After completing this chapter, you will be able to:

1. Create and format a list as an Excel table.
2. Apply Conditional Formatting rules as well as Color Scales, Icon Sets, and
   Data Bars.
3. Sort data by one or more columns or by attribute.
4. Filter data by using AutoFilters and by creating an Advanced Filter.
5. Use subtotals, groups, and outlines for tabular data in a worksheet.
6. Import data from text and database files and export data from a workbook.
7. Transform data using Flash Fill and Text functions.
8. Build and format a PivotTable.

## Case Study

- For the Pause & Practice projects in this chapter, you create worksheets for
  Perfect Vacation Rentals.
- To complete the work, you:
    - Format data as an Excel table.
    - Sort and filter data.
    - Import and clean data.
    - Create a PivotTable.

## Conclusion

- This chapter covers how to manage data in Excel, including formatting as a
  table, sorting and filtering data, importing and exporting data, and creating
  PivotTables.
- By completing the Pause & Practice projects, you will gain hands-on experience
  applying the concepts learned in this chapter.

# SLO 4.1 - Creating and Formatting an Excel Table

- An **Excel table** is a list of related pieces of information that is
  formatted with a title row followed by rows of data.
- The **header row** is the first row of a table with descriptive titles or
  labels. Each row of data is a record and each column is a field.
- When data is in table format, you can **sort, filter**, and **calculate
  results** easily and quickly.
- Excel tables are created by formatting data that is arranged as a list and
  conforms to the guidelines described.
- To create an Excel table, select the data and click the **Format as Table**
  button in the Styles group on the Home tab to open the Table Styles gallery.
- Alternatively, you can select the data and click the **Table button** in the
  Tables group on the Insert tab. The Format as Table dialog box opens and
  creates the table with a default style.
- After formatting data as a table, each label in the header row displays an *
  *AutoFilter arrow**, and the contextual Table Design tab opens with command
  groups for modifying the table.
- A filter is criteria that determines which data is shown and which is hidden.
- **Table styles** are organized in light, medium, and dark categories and are a
  pre-designed set of format settings with a color scheme, alternating fill for
  rows and columns, vertical and horizontal borders, and more.
- You can apply a different style after the table is created and can remove a
  style by selecting **None in the Light group** or clicking the **More button**
  and selecting Clear.
- The **Table Style Options group** on the Table Design tab includes commands
  for showing or hiding various parts of the table such as the header row or a
  total row.
- A command to **apply bold to data** in the first or last column or both is
  available.
- The **Tools group** on the Table Design tab includes commands to remove
  duplicate records, to convert the table to a regular cell range, and to insert
  a slicer.
- Tables are named as **TableN** where N is a number. To change the name to a
  more descriptive label, click the Table Name box in the Properties group on
  the Table Design tab and type the name. Do not use spaces in a table name, and
  use unique table names.
- The Remove Duplicates command scans a table to locate and delete rows with
  repeated data in the specified columns.

## How to Create an Excel Table

1. Select the cells to be formatted as a table.
2. Include the header row and all data rows.
3. Click the **Format as Table** button in the Styles group on the Home tab to
   open the table gallery.
4. Click a table style.
5. Confirm the cell range in the Format As Table dialog box.
6. Select the My table has headers box in the Format As Table dialog box.
7. Click` OK.`

## How to Apply a Table Style

1. Click any cell within the table.
2. Click the **More button** in the Table Styles group.
3. Point to a style thumbnail to see a Live Preview in the table.
4. Click to select and apply a style.

## How to Display a Total Row in a Table

1. Click any cell within the table.
2. Select the **Total Row box** in the Table Style Options group.
3. The total row displays as the last row in the table.
4. Click a cell in the total row and click its arrow to open the calculation
   list.
5. Choose the calculation for the column.

**Note:** Custom Views are not available in a workbook that has an Excel table.

## The Table Tools and Properties Groups

### Tools Group

- Commands for removing duplicate records, converting table to a
  regular cell range, and inserting a slicer
- Insert Slicer command opens a filter window for hiding/showing
  records in the table
- Remove Duplicates command scans the table to locate and delete rows
  with repeated data in the specified columns
- Duplicate row is a record in a table that has the same data in one
  or more columns
- In the Remove Duplicates dialog box, you set which columns might
  have duplicate data

### How to Name the Table and Remove Duplicates

1. Click any cell within the table.
2. Click the entry box for the Table Name property \[Table Design tab,
   Properties group\].
3. Type a name for the table and press Enter.
    - Start the name with an alphabetic character.
    - Do not use spaces in the name.
4. Click the Remove Duplicates button \[Table Design tab, Tools
   group\].
5. Click Unselect All to remove all check marks.
6. Select the box for each column (field) that might have duplicate
   data.
7. Click` OK.`
    - A message box indicates how many duplicate values were removed
      and how many unique values remain.
8. The rows are removed from the table.
    - The command does not preview which rows are deleted.
9. Undo this command if needed.

### How to Insert a Slicer

1. Click any cell within the table.
2. Click the Insert Slicer button \[Table Design tab, Tools group\].
3. The Insert Slicers dialog box lists all field names in the table.
4. Select the box for the field to be used for filtering.
    - You can select more than one field to open multiple slicers.
5. Click` OK.`
6. All records display in the table.
7. Click an item name in the Slicer to filter the data.
    - The records are filtered.
    - You can press Ctrl and click another item to apply more than one
      filter.
    - You can toggle the Multi-Select button in the Slicer title bar
      and then click to turn off multiple items for filtering.
8. Click the Clear Filter button in the Slicer title bar.
    - All the records display.
9. Click the More button \[Slicer tab, Slicer Styles group\].
10. Select a style for the Slicer.
11. Click the Columns spinner button \[Slicer tab, Buttons group\] to
    set the number of columns in the Slicer.
12. Click and drag a corner selection handle to resize the Slicer.
13. Set a specific size in the Size group \[Slicer tab\].
14. Press Delete to remove the Slicer window.

- The Slicer window must be selected to be deleted.

### Properties Group

- Includes the table name box and a command to resize the table
- Tables are named as TableN where N is a number
- To change the name, click the Table Name box in the Properties group
  on the Table Design tab and type the name
    - Do not use spaces in a table name, and use unique table names
- From the Properties group, you can open the Resize Table dialog box
  to select an expanded or reduced cell range for the table
- A table also has a resize arrow in its bottom right cell that you
  drag to grow or shrink the table
- Tables grow automatically when you press Tab after the last item in
  the last row to start a new record

### How to Convert a Table to a Range

1. Click any cell within the table.
2. Click the Convert to Range button \[Table Design tab, Tools group\].
3. Click Yes in the message box.
    - The data displays as a regular range of cells.
    - Font size, font style, and colors remain.
4. Format the data as needed.

### Structured References and Table Formulas

- Each column in a table is assigned a name using the label in its
  header row
- The column name with its table name is known as a structured
  reference
- An example is MyTable\[State\] as the reference for the state column
  in MyTable
- Structured references are supplied automatically in formulas in a
  table, making it easy to identify what is being calculated
- When you point to build a formula, you will see \[@ColumnName\] with
  @ inserted as an identifier
- Column names are enclosed in square brackets
- In addition to column names, an Excel table includes specific item
  names: #All, #Data, #Headers, #This Row, and #Totals, each preceded
  by the \## symbol as an identifier
- The table name and structured reference names appear in Formula
  AutoComplete lists, and you can refer to these ranges in formulas
  outside the table
- Tables expand to incorporate a new column when you type in a cell
  adjacent to the last column
- You can start by typing a new header label in the appropriate row,
  or type a formula in any data row
- As part of its Table AutoExpansion feature, Excel copies a formula
  to complete a column with a calculation, too

### How to Add a Calculation Column in a Table

1. Select the cell to the right of the last column header in the table.
    - The data must be formatted as an Excel table.
2. Type a label for the new column and press Enter.
    - A column is added to the table.
3. Select the first cell for the formula.
4. Type = to start a formula.
5. Click the table cell with the value to be used in the formula.
    - The column name in square brackets preceded by an @ symbol is
      inserted.
6. Complete the formula.
    - You can refer to non-table cells in the worksheet or type a
      constant.
7. Press Enter.
    - The formula is copied down the entire column.
8. Enter the formula in any row in the column for it to be copied.

### Additional Information

- If a normal data range has a header row, show AutoFilter arrows by
  clicking the Filter button on the Data tab
- To enter cell references in a table formula instead of structured
  references, disable Use table names in formulas in Excel Options
  \[File tab, Options, Formulas pane\]
- As you work with structured references, expand the Formula bar to
  better see a lengthy formula
- The Expand/Collapse Formula Bar button is on the right edge of the
  Formula bar
- Disable automatic expansion of table formulas in the AutoCorrect
  Options, AutoFormat As You Type tab \[File tab, Options, Proofing
  pane\]

## Conclusion

Excel tables are a powerful tool for organizing, filter, and calculate results
with just a few clicks. Follow the guidelines to optimize your tables for use of
Excel commands, and use table styles and options to customize your table's
appearance and functionality. Remember to use unique table names and remove
duplicate records to ensure accurate data analysis.

- The Table Design tab in Excel provides various tools and properties
  to work with tables effectively
- Tables are useful in managing and analyzing data, especially when
  dealing with large datasets
- Structured references and table formulas make it easy to work with
  tables and identify what is being calculated
- Convert to Range and Remove Duplicates are useful commands to manage
  tables effectively
- Inserting a slicer provides a visual filter to work with smaller
  pieces of large datasets

# SLO 4.2 - Applying Conditional Formatting

Conditional formatting is a powerful feature that allows you to apply
formats to cells based on specific criteria. It is dynamic, which means
that the formatting will adapt if the data changes. There are different
types of conditional formatting, such as Highlight Cells Rules,
Top/Bottom Rules, and data visualization, which displays a fill color, a
horizontal bar, or an icon.

## Highlight Cells Rules

Highlight Cells Rules use relational or comparison operators to
determine if the value or label should be formatted. Operators include
Equal To, Greater Than, and Less Than. You can also create your own rule
using other operators or a formula. You can access conditional
formatting options from the Conditional Formatting button in the Styles
group on the Home tab.

### How to Create a \"Less Than\" Highlight Cells Rule

1. Select the cell range.
2. Click the Conditional Formatting button in the Styles group on the
   Home tab.
3. Select Highlight Cells Rules and choose Less Than.
4. In the Less Than dialog box, type a value in the Format cells that
   are` LESS ``THAN `box.
5. Click the arrow for the with box and choose a preset format or
   choose Custom Format to open the Format Cells dialog box and build a
   format.
6. Click` OK `to apply the formatting.

## Top/Bottom Rules

Top/Bottom Rules use ranking to format the highest (top) or lowest
(bottom) items, either by number or percentage. You can set the number
or percentage in the dialog box as well as the format. You can also set
a rule to format values that are above or below average.

### How to Create a Top 10% Rule

1. Select the cell range.
2. Click the Quick Analysis button in the bottom right cell of the
   range.
3. Choose Formatting.
4. Choose Top 10% from the list of commonly used conditional format
   choices.
5. Edit the format to use different attributes from the Manage Rules
   command with the Conditional Formatting button.

## Use a Formula for a Rule

In addition to the Highlight Cells or Top/Bottom rules, you can create a
conditional formatting formula using operators, criteria, and settings
in the New Formatting Rule dialog box. For example, build a rule to
format cells in the \"Days Rented\" column if the state is Michigan. A
formula must result in either` TRUE `or` FALSE `(a Yes or No question) to be
used as criteria.

### How to Use a Formula in a Conditional Formatting Rule

1. Select the cell range.
2. Click the Conditional Formatting button in the Styles group on the
   Home tab.
3. Select New Rule.
4. Choose Use a formula to determine which cells to format in the
   Select a Rule Type list.
5. Type the formula in the Format values where this formula is true
   box. For example, use the formula =A1=\"Michigan\" to format cells
   in the \"Days Rented\" column if the state is Michigan.
6. Click Format to open the Format Cells dialog box and build a format.
7. Click` OK `to apply the formatting.

## Data Bars, Color Scales, and Icon Sets

Conditional formatting visualization displays cells with icons, fill
color, or shaded bars to distinguish values. Visualization highlights
low, middle, or top values, comparing the values to each other.
Conditional formatting visualization commands are Data Bars, Color
Scales, and Icon Sets.

### How to Format Data with an Icon Set

1. Select the cell range.
2. Click the Quick Analysis button and choose Formatting.
3. Choose Icon Set.
4. A default icon set is applied.
5. Manage Conditional Formatting Rules

You can edit any conditional formatting rule, including visualization
commands, from the Conditional Formatting Rules Manager dialog box. From
this dialog box, you can reset the range to be formatted, change the
actual format, change the rule, or delete the rule. If you select the formatted
range
before you start the Manage Rules command, the rule for that selection
displays in the Conditional Formatting Rules Manager dialog box. You can
also show all the rules in the current or another sheet.

### How to Manage Conditional Formatting Rules

1. Select the formatted cell range.
2. Click the Conditional Formatting button in the Styles group on the
   Home tab.
3. Choose Manage Rules.
4. The Show formatting rules for box displays This Table or Current
   Selection.
5. Choose This Worksheet to list all rules in the worksheet or select
   another worksheet.
6. Select the formatting rule to be modified.
7. Select Edit Rule to open the Edit Formatting Rule dialog box.
8. Select options to change the rule in the Edit the Rule Description
   area.
9. Click Format when available to open the Format Cells dialog box.
10. Click` OK `to close the Edit Formatting Rule dialog box.
11. Click` OK `to close the Conditional Formatting Rules Manager dialog
    box.

You can clear conditional formatting from a selected range or from the
entire sheet. For a selected range, click the Quick Analysis button and
choose Clear or choose Clear Rules from the Conditional Formatting
button menu.

## Conclusion

Conditional formatting is a powerful tool that allows you to apply
formatting to cells based on specific criteria. You can use Highlight
Cells Rules, Top/Bottom Rules, and data visualization commands to create
dynamic formatting that adapts if the data changes. You can also create
your own conditional formatting formulas using operators and criteria.
By using the Manage Rules command, you can modify, delete, or reset the
conditional formatting rules.

# SLO 4.3 - Sorting Data

## Sorting

- Process of arranging data rows in order
- Ascending order: alphabetically A to Z, numerically smallest to largest
- Descending order: alphabetically Z to A, numerically largest to smallest
- Dates are values: ascending sorts earliest date first, descending sorts most
  current date first

## Sort Data in an Excel Table

- Sort from AutoFilter arrows in header row
- Sort by one column at a time, but can sort by multiple columns with hierarchy

### How to Sort Data in an Excel Table

1. Click AutoFilter arrow for lowest sort level column heading
2. Choose Sort A to Z or Sort Smallest to Largest
3. Click AutoFilter arrow for next sort level column heading
4. Choose Sort A to Z or Sort Smallest to Largest

## Sort Options

- Sort any data in rows and columns, doesn't need to be formatted as a table
- Can sort data by fill color, font color, or cell icon from Conditional
  Formatting
- Sort & Filter button on Home tab and Data tab
- Undo a sort task by clicking Undo button

## Sort Data by One Column

1. Select a cell in the column to be used for sorting
2. Click Sort A to Z button or Sort Z to A button

## Sort Data by Multiple Columns (Multi-level sort)

1. Select a cell in the first column to be sorted
2. Click Sort A to Z button
3. Select a cell in the second column to be sorted
4. Click Sort A to Z button

## How to Use the Sort dialog box

1. Select a cell in the range to be sorted
2. Click Sort button on Data tab or Custom Sort on Home tab
3. Select "My data has headers" box if applicable
4. Click Sort by arrow and select column heading for top sort level
5. Click Sort On arrow and choose Cell Values or attribute
6. Click Order arrow and choose option
7. Click Add Level to add a second sort column
8. Click Then by arrow and select second column heading
9. Click Sort On arrow and choose Cell Values or attribute
10. Click Order arrow and choose sort order

## Sort Data by Cell Attribute (font color, cell fill color, conditional formatting icon)

1. Select a cell in the column to sort
2. Click Sort button on Data tab or Custom Sort on Home tab
3. Select "My data has headers" box if the data has a header row
4. Click Sort by arrow and select column heading
5. Click Sort On arrow and choose the attribute
6. Click leftmost Order arrow and choose a color or icon
7. Click rightmost Order arrow and choose On Top or On Bottom
8. Click Add Level
9. Click Then by arrow and select the same column heading
10. Click Sort On arrow and choose the attribute
11. Click leftmost Order arrow and choose a color or icon
12. Click rightmost Order arrow and choose On Top or On Bottom

## Conclusion

- Sorting data helps to organize and analyze information
- Excel provides various options for sorting, including single column,
  multi-column, and cell attributes
- AutoFilter arrows, Sort & Filter button, and Sort dialog box provide different
  methods for sorting data in Excel

# 4.4 - Filtering Data

### AutoFilters

- Filter button: Data tab, Sort & Filter group
- Use AutoFilter arrows in header row for filtering

**How to display and use AutoFilters**

1. Select a cell in the list
2. Click the Filter button
3. Click the AutoFilter arrow for the desired column
4. Select boxes for items to be shown
5. Click` OK
   `

### Custom AutoFilter

- Use custom criteria for filtering
- Text Filters for alphanumeric data, Number Filters and Date Filters for
  numeric and date data

### How to create a Custom Text AutoFilter

1. Select a cell in the data or table
2. Click the Filter button
3. Click the AutoFilter arrow for the desired column
4. Select Text Filters and choose an operator
5. Type criteria for the first operator
6. Choose the And or Or radio button
7. Click the arrow to choose a second operator (optional)
8. Type or select criteria for the second operator
9. Click` OK
   `

## Advanced Filter

- Allows more complex filtering with a separate criteria range
- Can display results in the data range or separate location

### How to create an Advanced Filter

1. Select labels in the header row for filtering
2. Copy the labels to a Criteria area and an Extract area
3. Type criteria in the criteria range
4. Click a cell in the data range
5. Click the Advanced button in the Data tab, Sort & Filter group
6. Select Copy to another location in the Action group
7. Verify or select the range in the List range box
8. Select the criteria cell range in the Criteria range box
9. Select the extract range in the Copy to box
10. Click` OK
    `

## Conclusion

- Use AutoFilters for basic filtering
- Custom AutoFilters allow more specific filtering with criteria
- Advanced Filters enable complex filtering with separate criteria and output
  ranges

# SLO 4.5 - Using Subtotals, Groups, and Outlines

## Outline

- A summary that groups records

## Subtotal

- A summary row for data that is grouped

## Group

- A set of data with the same entry in one or more columns

## The Subtotal Command

- Located in the Outline group on the Data tab
- Inserts summary rows for a sorted list
- Formats data as an outline
- Available for a normal range of cells, not an Excel table

### How to Display Subtotals

1. Sort data by the column for which subtotals will be calculated
2. Click a cell in the data range
3. Click the Subtotal button [Data tab, Outline group]
4. Configure the Subtotal settings as needed
5. Click` OK
   `
   ##` SUBTOTAL `Function

- Inserts in each cell that displays a result
- Arguments: Function_num and RefN
- Function_num is a number between 1 and 11 based on the chosen function
- RefN argument depends on the chosen column in the Subtotal dialog box

| Function_num<br/>(includes hidden values) | Function_num<br/>(ignores hidden values) | Calculation |
|:-----------------------------------------:|:----------------------------------------:|:-----------:|
|                   **1**                   |                 **101**                  | ``AVERAGE`` |
|                   **2**                   |                 **102**                  |  ``COUNT``  |
|                   **3**                   |                 **103**                  | ``COUNTA``  |
|                   **4**                   |                 **104**                  |    `MAX`    |
|                   **5**                   |                 **105**                  |    `MIN`    |
|                   **6**                   |                 **106**                  | ``PRODUCT`` |
|                   **7**                   |                 **107**                  |  ``STDEV``  |
|                   **8**                   |                 **108**                  | ``STDEVP``  |
|                   **9**                   |                 **109**                  |    `SUM`    |
|                  **10**                   |                 **110**                  |    `VAR`    |
|                  **11**                   |                 **111**                  |  ``VARP``   |

## Outline Buttons

- Created by the Subtotal command
- A worksheet can have only one outline

### How to Use Outline Buttons

1. Click a collapse button (−) to hide details for a group
2. Click an expand button (+) to display details for a group
3. Click the Level buttons to display different levels of details

## Auto Outline

- Inserts groups based on formulas in the data
- Does not use the Subtotal command

### How to Create an Auto Outline

1. Click a cell within the data range
2. Click the Group button arrow [Data tab, Outline group]
3. Select Auto Outline

## Define Groups

- Create a group for data that does not have totals or formulas

### How to Define a Group

1. Sort data based on the preferred grouping
2. Insert a blank row at the end of each group (or at the start of each group)
3. Select the row or column headings for the first group
4. Click the Group button [Data tab, Outline group]
5. Repeat steps 1–4 for each group

## Conclusion

- Outlines, subtotals, and groups are useful tools in Excel to organize and
  summarize data
- The Subtotal command and` SUBTOTAL `function provide summaries and
  calculations
- Outline buttons allow for easy navigation of grouped data
- Auto Outlines and defining groups offer more control over data organization

# SLO 4.6 - Importing and Exporting Data

## Importing Data in Excel

- Import data from external sources (e.g. text files, databases, cloud
  locations)
- Get & Transform Data command group on Data tab to establish connections for
  refreshing data

## Tab-Delimited Text Files

- Use delimiters like tabs, spaces, and commas to separate fields
- Some text files use fixed width to divide fields
- Import tab-delimited text files using From Text`/CSV `button on Data tab

## Importing Tab-Delimited Text File

1. Select the cell where imported data should begin
2. Click From Text`/CSV `button [Data tab, Get & Transform Data group]
3. Select the text file and click Import
4. Preview the data and click Load arrow, then select Load To
5. Choose where to place the imported data and click` OK
   `

## Comma-Separated Values `(CSV)` Files

- Text files with comma-separated values
- Can be opened directly in Excel or imported using Power Query

## Opening a` CSV `File

1. Click File tab and select Open
2. Navigate to the folder with the` CSV `file and click the file name
3. Format and edit data, rename the worksheet, and save as needed

## Importing` CSV `Data with Power Query

1. Open a new workbook or click the cell where imported data should begin
2. Click From Text`/CSV `button [Data tab, Get & Transform Data group]
3. Select the` CSV `file and click Import
4. Preview the data and click Transform Data to launch Power Query
5. Transform and clean the data as needed in Power Query
6. Click Close & Load button [Home tab, Close group]

## Transforming Data in Power Query

1. Click a cell in the Excel table with imported data
2. Click Queries & Connections button [Data tab, Queries & Connections group]
3. Select Edit at the bottom of the pop-up window in the Queries & Connections
   pane
4. Perform data transformations in Power Query (e.g. sorting, deleting steps)
5. Click Close & Load button [Home tab, Close group]

## Importing Access Database Files

- Import tables or queries from Access databases into Excel
- Use Navigator dialog box to choose the table or query to import

## Importing Access Table into a Worksheet

1. Select the cell where imported data should begin
2. Click Get Data button [Data tab, Get & Transform Data group]
3. Choose From Database and then choose From Microsoft Access Database
4. Select the database name and click Import
5. Select the table or query in the Navigator window
6. Click Load arrow and select Load To
7. Choose where to place the imported data and click` OK
   `

## Workbook Queries and Connections

### Managing External Data

1. Click cell within imported data
2. Click Properties button [Data tab, Queries & Connections group]
3. External Data Properties dialog box displays
4. Click Query Properties button at the right of the Name box
5. Query Properties dialog box opens
6. Edit query, table, description, or refresh settings
7. Click Definition tab to edit query in Query Editor
8. Connection information displayed
9. Click Edit Query to launch Power Query
10. Cancel to close dialog boxes

### Importing Data from Word Documents

1. Select cell for imported data
2. Open Word document
3. Select data to copy
4. Click Copy button [Home tab, Clipboard group]
5. Switch to Excel worksheet
6. Click arrow with Paste button [Home tab, Clipboard group]
7. Choose Match Destination Formatting from Paste gallery
8. Alternatively, save Word document as text file and import into Excel

### Importing Website Data

- Use From Web command on Data tab for` HTML `tables
- Download data from research, statistics, demographics websites

### Exporting Data as a Text File

1. Click worksheet tab with data to export
2. Click File tab and select Export
3. Select Change File Type
4. Select Text (Tab delimited) as file type
5. Click Save As
6. Choose location and type file name
7. Save

### Exporting Data via Clipboard

1. Open Word or PowerPoint, position insertion point
2. Switch to Excel, select cells to copy
3. Click Copy button [Home tab, Clipboard group]
4. Switch to Word or PowerPoint
5. Click arrow with Paste button [Home tab, Clipboard group], choose Paste
   Special
6. Select Microsoft Excel Worksheet Object
7. Select Paste link in Paste Special dialog box
8. Click` OK
   `

### Exporting Data to SharePoint Lists

- Use Export button in External Table Data group on Table Design tab
- Requires access and permission from SharePoint service administrator

## SLO 4.7: Transforming Data in Excel

### Flash Fill New Data

- Recognizes pattern in data and suggests completion.
- Previews suggested column.
- Press Enter to accept the suggestion.

#### How to Flash Fill Data in a Worksheet

1. Click the first cell in an empty column next to the original data.
2. Type the first data item as you want it to appear.
3. Press Enter and type data in the cell below the first cell to begin the
   second entry.
4. Press Enter to accept the Flash Fill data.

### Flash Fill to Convert Case

- Can be used to convert case, extract data, and more.
- Establish a recognizable pattern for Flash Fill to work.

#### How to Flash Fill and Convert Case

1. Click the first cell in an empty column.
2. Type the data in uppercase characters.
3. Press Enter and type data in the cell below the second entry.
4. Press Enter to accept the Flash Fill data.

### Text to Columns

- Splits data in one column into multiple columns.
- Uses a delimiter or fixed width.

#### How to Convert Text to Columns

1. Evaluate the data to be split.
2. Determine the delimiter character.
3. Insert columns to the right of the existing data.
4. Select the cells with data to be split.
5. Click the Text to Columns button.
6. Follow the wizard to complete the process.

###` LEFT,`` RIGHT,` and` MID `Functions

- Extract and display characters from a text string.
  -` LEFT:` counts characters from the left.
  -` RIGHT:` counts characters from the right.
  -` MID:` counts characters from the middle.

###` LEN `Function

- Counts the number of characters in a text string.
- Useful for determining appropriate length or locating` IDs` with missing/extra
  characters.

## Conclusion

- Excel provides tools to transform and clean data.
- Flash Fill can automatically fill in data and convert case.
- Text to Columns can split data into multiple columns.
  -` LEFT,`` RIGHT,`` MID,` and` LEN `functions can manipulate and count
  characters in text strings.

# SLO 4.8 - PivotTables and PivotCharts

## PivotTables

- Cross tabulation report for list-type data
- Separate worksheet to sort, filter, and calculate large amounts of data
- Important part of Business Intelligence `(BI)`
- Interactive, allowing rearrangement and pivoting of data

### Creating a PivotTable

1. Click a cell in the data range
2. Click the Recommended PivotTables button [Insert tab, Tables group]
3. Select the preferred PivotTable and click` OK
   `4. The PivotTable report is created on a separate sheet

### The PivotTable Fields Pane

- Displays on its own sheet in the workbook
- Add and remove fields, drag field names to reset data organization

### Field Settings

- Settings or attributes depending on the type of data
- Can edit default label that appears in the PivotTable

### Formatting a PivotTable

- Use the PivotTable Design tab or the Home tab

## PivotCharts

- Contains field buttons and filter options
- Updates with its associated PivotTable

### Creating a PivotChart

1. Click a cell in the PivotTable
2. Click the PivotChart button [PivotTable Analyze tab, Tools group]
3. Select a chart type and subtype
4. Click` OK
   `

## Conclusion

PivotTables and PivotCharts are powerful tools in Excel for analyzing and
visualizing large amounts of data. They allow for easy rearrangement and
pivoting of data to gain insights and make better decisions.

# Chapter Summary

- **4.1** Create and format a list as an Excel table.
    - An Excel **table** provides enhanced commands for sorting, filtering,
      and calculating data.
    - A table is a list with a **header row**, AutoFilter arrows in the header
      row, and the same type of data in each column.
    - Each column in a table is a **field**, and the label in the header row
      is the **field name**.
    - When any cell in a table is selected, the Table Design tab is
      available with commands for editing and formatting the table.
    - A **table style** is a preset combination of colors and borders.
    - Table style options enable you to display a total row, to emphasize
      the first or last column, and to add shading or borders to rows or
      columns.
    - The Table Tools group on the Table Design tab includes commands to
      remove duplicate rows, to convert a table to a normal cell range, or to
      insert a Slicer window for filtering the data.
    - A **duplicate row** has the same data in one or more columns.
    - A slicer is a visual filter window.
    - The automatic name for a table and several of its parts are
      **structured references**.
    - Structured references appear in Formula AutoComplete lists and are
      substituted automatically in formulas that refer to certain cell ranges
      in the table.
- **4.2** Apply Conditional Formatting rules as well as Color Scales,
  Icon Sets, and Data Bars.
    - **Conditional formatting** formats only cells that meet specified
      criteria.
    - Highlight Cells Rules use relational or comparison operators such as
      Equal To or Greater Than to determine which cells are formatted.
    - Top/Bottom Rules use ranking or averages to format cells.
    - Use the New Rule dialog box to create a custom rule with a formula
      that calculates which cells are formatted.
    - Conditional formatting **data visualization** uses rules to format cells
      with Icon Sets, Color Scales, or Data Bars.
    - Edit any conditional formatting rule with the Manage Rules command.
    - Remove conditional formatting from a selected cell range or from the
      entire sheet with the Clear Rules command.
- **4.3** Sort data by one or more columns or by attribute.
    - **Sorting** is a process that arranges data in order, either **ascending**
      (A to Z) or **descending** (Z to A).
    - Sort data in an Excel table using the AutoFilter arrows in the
      header row.
    - Sort a single column using the Sort A to Z or Sort Z to A button in
      the Sort & Filter group on the Data tab.
    - Perform a multiple column sort within the data by sorting in lowest
      to highest sort priority.
    - When you use the Sort buttons, sort first by the column that is the
      lowest level in the sort hierarchy.
    - You can also sort multiple columns from the Sort dialog box. Click
      the **Sort** button on the Data tab.
    - The lowest level sort in the Sort dialog box is the bottom name in
      the Column list.
    - Use the Sort dialog box to sort data by cell **attribute** including
      font color, cell fill color, or conditional formatting icon.
- **4.4** Filter data by using AutoFilters and by creating an Advanced
  Filter.
    - The Filter button [Data tab, Sort & Filter group] displays an
      AutoFilter arrow for each column heading in list-type data.
    - Use an AutoFilter arrow to select records to be shown or hidden or
      to build a custom filter.
    - A **custom AutoFilter** uses operators such as Equals or Greater Than as
      well as `AND` and` OR `conditions.
    - `AND` filters require that all conditions be met;` OR `filters require
      that any one of the conditions be met.
    - An **Advanced Filter** provides filter options such as using a formula
      in the filter definition.
    - An Advanced Filter requires a **criteria range** in the same workbook
      and can display filtered results separate from the data.
    - The results of an Advanced Filter display in an output or extract
      range on the same worksheet.
- **4.5** Use subtotals, groups, and outlines for tabular data in a
  worksheet.
    - The **Subtotal** command inserts summary rows for sorted data using a
      calculation such as` SUM `or` AVERAGE.`
    - The Subtotal command inserts the` SUBTOTAL `function in the summary
      rows.
    - Arguments for the` SUBTOTAL `function determine the calculation, such
      as` COUNT,`` SUM,` or` AVERAGE.`
    - The Subtotal command formats the data as an **outline**, which groups
      records to be displayed or hidden.
    - Outline buttons appear next to row and column headings and determine
      the level of detail.
    - An **Auto Outline** creates **groups** based on formulas that are in a
      consistent pattern in a list.
    - Define your own groups by sorting data, inserting blank rows, and
      using the Group button in the Outline group on the Data tab.
- **4.6** Import data from text and database files and export data from
  a workbook.
    - **Import** text files into a worksheet to save time and increase
      accuracy.
    - Common **text file** formats include .txt (text) and .csv (**comma-
      separated values**).
    - Data in a .txt file is **delimited** or **fixed width**.
    - A` CSV `file can be opened like a workbook, or it can be imported
      using Power Query.
    - Imported data displays in an Excel table.
    - **Power Query** is a separate application that connects to external
      data.
    - A **query** is a collection of definitions and filters created by Power
      Query to extract data.
    - Data imported from a command in the Get & Transform Data group
      establishes a data connection so that the data can be refreshed.
    - Import data from a **Microsoft Access** database into a worksheet or by
      using Power Query.
    - Many websites offer data in Excel or text format for downloading
      into a worksheet.
    - **Export** data from a worksheet for use in another program or
      application.
    - Use the Windows or Office Clipboard to copy data from Excel into
      Word, PowerPoint, and other programs.
- **4.7** Transform data using Flash Fill and Text functions.
    - The **Flash Fill** command recognizes and copies typing actions from one
      or more cells and suggests data for the remaining rows in the same
      column.
    - You can use Flash Fill to create new columns of data in a worksheet.
    - The **Text to Columns** command [Data tab, Data Tools group] splits data
      from one column into multiple columns based on the delimiter character.
    - Use Text functions to count characters, extract text strings, or
      display data as needed.
    - The` LEFT,`` MID,` and` RIGHT `functions display a value that reflects the
      number of characters in a text string.
    - The` LEN `function displays a value that represents the number of
      characters in the text string.
- **4.8** Build and format a PivotTable.
    - **Business intelligence** uses tools such as PivotTables to “drill-down”
      into data and assess various types of results or changes.
    - A **PivotTable** is a cross-tabulated summary report on its own
      worksheet.
    - When a PivotTable is active, the PivotTable Analyze and Design tabs
      are available.
    - The PivotTable Fields pane displays when a cell within the table is
      selected.
    - Display and hide the PivotTable Fields pane from the Show group
      [PivotTable Analyze tab].
    - The PivotTable Fields pane includes a list of available fields and
      report areas where the fields can be positioned.
    - Drag field names in the PivotTable Fields pane to reposition them so
      that you can analyze data from a different perspective.
    - Modify and format fields in a PivotTable using the Field Settings
      dialog box.
    - Format PivotTables using the PivotTable Design tab. Format selected
      cells using the Home tab.
    - PivotTable layout options include subtotals and grand totals,
      repeated labels, and display formats.
    - Click the **Fields, Items, Sets** button [PivotTable Analyze tab,
      Calculations group] to insert a **calculated field** in a PivotTable.
    - Refresh a PivotTable after editing the source data by clicking the
      **Refresh** button [PivotTable Analyze tab, Data group].
    - A **PivotChart** is based on a PivotTable and has filter buttons like
      the PivotTable.
    - A PivotChart and its PivotTable are interactive and dynamic so that
      a change made in one results in a change in the other.