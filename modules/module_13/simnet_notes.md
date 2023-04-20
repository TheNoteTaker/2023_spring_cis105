# Module 14 - Creating a Database and Tables

<!-- TOC -->

* [Module 14 - Creating a Database and Tables](#module-14---creating-a-database-and-tables)
* [General Notes](#general-notes)

<!-- TOC -->

# General Notes

## Overview

- Microsoft Access is the leading PC-based database management system software
  in both the personal and business market.
- Access helps organize and manage personal, business, and educational data in a
  structure known as a database.
- This chapter covers the basics of working with a relational database and using
  the Access database management system.

## Student Learning Outcomes (SLOs)

- **SLO 1.1:** Explains data organization concepts, uses the Access Navigation
  Pane, and identifies objects.
- **SLO 1.2:** Creates a database; sets database properties; understands
  database object naming conventions and data types; creates and manipulates
  database objects; and opens, closes, backs up, and recovers a database.
- **SLO 1.3:** Creates a table in Datasheet view; edits the default primary key;
  adds a new field; edits field and table properties; saves a table; enters
  records into a table; and understands edit and navigation modes.
- **SLO 1.4:** Imports data records from Excel.
- **SLO 1.5:** Navigates table records in Datasheet view; customizes a datasheet
  by modifying field order, row height, column width, field alignment, font
  style and size, row colors, and gridlines; hides and unhides fields; and adds
  total rows.
- **SLO 1.6:** Searches, sorts, and filters records in a table.

## Conclusion

- This chapter provides an overview of Microsoft Access, its benefits, and its
  basic features for creating and managing a database.
- Readers can learn how to create a database and tables, import data from Excel,
  navigate and customize datasheets, and search, sort, and filter records in a
  table.
- The Pause & Practice projects allow readers to apply and practice key skills,
  with the first project orienting them to using Access and the remaining
  projects helping them build a database for Central Sierra Insurance.

# SLO 1.1: Understanding Database Concepts

## Introduction

- Databases are used to manage data in organizations.
- Database management system (DBMS) is software that enables creating, managing,
  sorting, retrieving, querying and reporting related data in a database.

## Organize Data

- Access is a relational DBMS that organizes data into related tables.
- Tables store data about one type or grouping of information.
- Fields describe one aspect of a business object or activity, and a record is a
  collection of related data fields.
- A database is a collection of related tables.
- Each table is connected to another through common fields.

## Access Interface

- The interface is how you interact with Access to create, modify, delete, and
  use different objects of your database application.
- Access interface has a ribbon that contains six primary tabs and contextual
  tabs that become available when you are working on different tasks.
- The major objects in Access are tables, forms, reports, and queries.
- A database must contain at least one table.

## Major Access Objects

- Tables store data records.
- Forms provide an interface to view, add, update, and delete data in a table.
- Queries find data in the database and enable performing actions such as
  updating or deleting records.
- Reports enable viewing and printing the data in the database in a formatted
  and professional way.

## Objects in Access

- Access has two additional objects: Macros and Modules to add functionality to
  forms and reports in the database.
- You can add a button to a form and attach specific actions to each event, like
  opening a form when the button is clicked.

## Using the Navigation Pane

- The Navigation Pane displays all objects in the database application.
- The Shutter Bar Open/Close Button opens and closes the Navigation Pane.
- Use the expand/collapse arrow to change the display of groups.
- You can customize the Navigation Pane by using the Navigation menu.
- The top half of the Navigation menu enables you to select a category while the
  bottom half enables you to filter by a specific group within that category.
- The width of the Navigation Pane can be adjusted to see the complete text of
  longer object names.

## Using Datasheet and Design View of a Table

- You can open tables directly from the Navigation Pane in Datasheet or Design
  view.
- Datasheet view is used to enter, modify, delete, or view the data records.
- Design view enables you to build or modify the basic structure or
  functionality of an object.
- Switch between views of an open table using the View button in the Home tab,
  Views group.
- Datasheet view displays the name of the table at the top of the datasheet.
  Each row represents a record in the table, and each column represents a field.
- The record selector is the gray cell on the left of each row. The active cell
  is the field where you enter or edit data. The last row in a table is the
  append row identified with an asterisk.
- The horizontal and vertical scroll sliders automatically display if the Access
  window is not wide enough or tall enough to display all of the rows and
  columns.

_Note:_ The Append row allows adding new records to a table, and the
rose-colored border surrounds the active cell.

# SLO 1.2 - Creating and Managing a Database

## 1. Creating a New Database

- Access extension: `.accdb`
- Create using a template or a blank database

### 1.1 Templates

- Predesigned databases containing pre-built objects
- Installed with Access and available at Office.com
- Can be customized for specific needs

#### How to Create a New Database Using a Template:

1. Open Access and go to the Start page
2. Click the New button to view sample templates
3. Select desired template
4. Type a file name for your database in the File Name box
5. Click the Browse button to choose a location to save the database, then click
   OK
6. Click the Create button
7. The new database opens, and you may need to click Get Started and Enable
   Content

### 1.2 Database Properties

- Properties include creation date and author
- View and edit properties via the Database Properties dialog box in the
  Backstage view

#### How to View and Edit Database Properties:

1. Click the File tab and then the Info button
2. Click the "View and edit database properties" link
3. Click the Summary tab
4. View and edit properties as desired, then click OK

## 2. Create a Blank Database

- Create from scratch when no suitable template exists
- Access automatically creates one new table

### How to Create a New Blank Database:

1. Open Access and go to the Start page
2. Select Blank database
3. Type a file name in the File Name box
4. Click the Browse button to choose a location to save the database, then click
   OK
5. Click the Create button

### 2.1 Access Naming Rules

- Use meaningful names for objects, fields, and controls
- Avoid spaces; use capital letters for compound words
- Common prefixes: `tbl` (table), `frm` (form), `rpt` (report), `qry` (query)

### 2.2 Data Types

- Each field must be assigned a specific data type
- Data types determine properties and usage

### 2.3 Create and Save Database Objects

- Create objects using the Create tab
- Save objects to keep a permanent copy

#### How to Create a New Object:

1. Click the Create tab on the Ribbon
2. Click the object button of the object you wish to create
3. Save the object and change the default name if needed

## Saving and Managing Database Objects

### Save a Copy of an Object

- Click Save As button on File tab
- Select Save Object As in File Types area
- Access suggests default name "Copy of" + original object name
- Type desired name in Save 'Object name' to box
- Select type of object to create
- Click OK
- Press F12 to open Save As dialog box
- Use Copy and Paste to save a copy using Navigation Pane

### Open a Database Object

- Open object directly from Navigation Pane
- Double-click object name to open in default view
- Right-click object name and select desired view from context menu
- Press Enter to open object using its default view

### Rename a Database Object

- Object must be closed
- Right-click object and select Rename from context menu
- Type new name and press Enter

### Delete a Database Object

- Object must be closed
- Select object in Navigation Pane and click Delete button
- Click Yes to confirm deletion

## Closing, Opening, and Managing Databases

### Close a Database

- Click Close button on File tab
- Access remains open

### Open a Database

- Click Open button on File tab
- Click Browse button and select database file
- Click Open button or double-click file name

### Back Up a Database

- Click File tab, then Save As button
- Select Back Up Database in Advanced grouping
- Click Save As button
- Select desired location for backup file
- Change default name if desired
- Click Save button

### Recover Database Objects from a Backup

- Open original database file
- Delete or rename damaged object
- Click New Data Source button, select From Database, then Access
- Import tables, queries, forms, reports, macros, and modules into current
  database
- Select objects to import and click OK
- Click Close button and verify successful recovery

# SLO 1.3 - Creating and Using a Table in Datasheet View

## Importance

- This section teaches you how to use Datasheet view to quickly create a table
  and enter data records in a new blank database in Access.

## Creating a New Table in Datasheet View

- Databases can contain multiple tables, and a new table can be added to a
  database from the Create tab.
- Tables created using the Table button automatically open in Datasheet view.

### How to Create a New Table in Datasheet View

1. Click the Create tab on the Ribbon.
2. Click the Table button in the Tables group.
3. A new table opens in Datasheet view.
4. The Ribbon updates, and the Table Fields and Table contextual tabs become
   available, with the Table Fields tab selected.

## Editing the Default Primary Key

- A primary key is a field that contains a unique value for each record.
- When a new table is created in Datasheet view, Access automatically includes
  one field, named ID, and designates it as the primary key.
- The default data type for this field is AutoNumber, and its value is
  automatically completed by Access after you add data into the other fields and
  the record is added into the table.

### How to Edit the Default Primary Key

1. Double-click the cell containing the ID field name (column header) in your
   table to select it.
2. Type the new name for your field and press the down arrow key.
3. Change the default data type of AutoNumber as desired.
4. Click the Data Type drop-down arrow to display the list of data types.
5. Select the desired Data Type property.

## Adding New Fields

- You must create a field for each of the different pieces of information you
  want to keep in your table.
- In Datasheet view, you can add new fields to a table by either entering the
  field names or entering the data values directly.

### How to Add New Fields by Entering Field Names

1. Click the Click to Add column header.
2. The drop-down list of data types displays.
3. Select the appropriate data type.
4. Type the desired field name in the column header.
5. Press the Enter key.

### How to Add New Fields Using a Data Type Button

1. Click the appropriate data type button in the Add & Delete group on the Table
   Fields tab.
2. Type the field name in the column header.
3. Press the Enter key.

### How to Add New Fields Using the More Fields Button

1. Click the More Fields button to display the drop-down list.
2. Select the appropriate data type.
3. Type the field name in the column header.
4. Press the Enter key.

### How to Add New Fields by Entering Data

1. Click the cell in the first Click to Add column.
2. Type the data value into the cell.
3. Click the Click to Add column header to add the field.
4. Access assigns a default field name, such as Field1, and selects that column.

## Additional Tips

- You need to verify that Access assigned the correct data type to each field.
- You should also change the default field name to be a descriptive field name.
- Fields added using the Click to Add option always add to the right of the last
  field in your table.
- Occasionally, you may wish to add a field between existing fields.

### How to Add New Fields between Existing Fields

1. Put your pointer on the field name of the column to the right of where you
   wish to insert a new field.
2. Click to select that column.
3. Right-click to open the context menu.
4. Select Insert Field.
5. The new field is inserted to the left of the selected field.
6. It is assigned a default field name, such as Field1.
7. Double-click the field name to change the name.

## How to Delete a Field

- To delete a field, hover the pointer over the field name, and click to select
  the column.
- Click the Delete button [Table Fields tab, Add & Delete group] to delete the
  field.
- Access will prompt to confirm the deletion, click Yes to delete the field.

Alternatively, you can right-click to open the context menu and select Delete
Field.

Note that the Undo button cannot undo deletions of table fields.

## Editing Field Properties in Datasheet View

- In Datasheet view, you can edit some field properties.
- To edit field properties, select the cell containing the field name (column
  header) you wish to change.
- Click the Name & Caption button [Table Fields tab, Properties group].
- In the Enter Field Properties dialog box, type a new value in the Name
  property.
- Type a value in the Caption and Description properties as desired.
- Click OK to close the dialog box.
- Click the Data Type option drop-down
  arrow [Table Fields tab, Formatting group].
- Select the desired Data Type property.
- Type a new value in the Field Size
  property [Table Fields tab, Properties group].

## Adding a Table Description

- You can set the Description property of a table to provide a more informative
  summary of its contents.
- Right-click the table name in the Navigation Pane and select Table Properties
  from the context menu.
- In the Table Properties dialog box, type a description in the Description
  property.
- Click OK to save the changes.

Note that the table properties can be displayed in the Navigation Pane by
selecting View By > Details from the Title bar context menu.

## Saving a Table

- To save a new table, click the Save button [File tab] and enter the desired
  name in the Table Name box.
- Alternatively, open the Save As dialog box by clicking the Save button on the
  Quick Access toolbar or by pressing Ctrl+S.

To save changes to an existing table, click the Save button [File tab], click
the Save button on the Quick Access toolbar, press Ctrl+S, or right-click the
tab with the table name and select Save from the context menu.

## Closing a Table

- To close a table, click the X (Close button) in the table tab, press Ctrl+W or
  Ctrl+F4, or right-click the tab with the table name and select Close from the
  context menu.

## Opening a Table

- To open an existing table in Datasheet view, click the Shutter Bar Open/Close
  Button or press F11 to open the Navigation Pane.
- Double-click the table name to open it in Datasheet view.
- Alternatively, select the table in the Navigation Pane and press Enter, or
  right-click the table in the Navigation Pane and select Open from the context
  menu.
- To open an existing table in Design view, select the table in the Navigation
  Pane and press Ctrl+Enter, or right-click the table in the Navigation Pane and
  select Design View from the context menu.

## Renaming a Table

- Close the table if it is open.
- Right-click the table name in the Navigation Pane and select Rename from the
  context menu.
- Type a new name and press Enter.

Note that renaming a table may cause problems for other objects that depend on
the table as their source of data.

## Adding, Editing, and Deleting Records

- In Datasheet view, add a record by clicking the first empty cell in the append
  row, and then entering the data values for the record.
- Use the Tab key to move to the next field, and press Enter when completed.
- To edit a field in a record, click to select the field, edit the data, and
  press Enter to save the changes.
- To delete a record from a table, click the record selector

When you are working in Datasheet view, you may not always notice that you are
switching between Edit and Navigation modes. Access automatically switches
between these modes depending on the actions you take. To switch between Edit
and Navigation modes manually, press the F2 key. On some laptops, you may need
to press the Fn key and the F2 key together to access the F2 key function.

In Edit mode, you can add new characters to the right of the insertion point in
a field. In Navigation mode, the insertion point is hidden, and an entire field
is selected. If you type in Navigation mode, your entry replaces the existing
value.

Table 1-5 describes the arrow key functions in the different modes:

| Key         | Navigation Mode Action                           | Edit Mode Action                                                                                                                                                                             |
|:------------|:-------------------------------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Up arrow    | Selects the current field in the previous record | Switches to Navigation mode and selects the current field in the previous record                                                                                                             |
| Down arrow  | Selects the current field in the next record     | Switches to Navigation mode and selects the current field in the next record                                                                                                                 |
| Left arrow  | Selects the previous field in the current record | Moves the insertion point left one character. If the insertion point is to the left of the first character, switches to Navigation mode and selects the previous field in the current record |
| Right arrow | Selects the next field in the current record     | Moves the insertion point right one character. If the insertion point is to the right of the last character, switches to Navigation mode and selects the next field in the current record    |

When you are in Navigation mode, you can select multiple cells by holding down
the Shift key and using the arrow keys to move the selection. You can also
select entire columns or rows by clicking the column or row selector,
respectively.

# SLO 1.4 - Importing Data Records from Excel

- Importing data records from an Excel file is a convenient way to add data to a
  table in Access.
- Before importing the file, you need to ensure that the Excel file is formatted
  correctly.
- It is important to close the Access table before starting the import process
  to avoid being prompted to close it during the import process.

## How to Import Data from Excel

1. Click the New Data Source button located in the External Data tab, Import &
   Link group.
2. Select From File, and then select Excel.
3. The Get External Data – Excel Spreadsheet dialog box launches.
4. Click the Browse button to locate and select the Excel file that contains the
   records you want to import.
5. Select the Append a copy of the records to the table radio button.
6. Click the drop-down arrow in the table name box and select the desired table.
7. Click OK to launch the Import Spreadsheet Wizard.
8. The data records should display in the Wizard window.
9. Click the Next button to advance to the next page of the Wizard.
10. Access confirms the name of the table to append the records.
11. Click Finish.
12. The Wizard displays the Save Import Steps screen.
13. If you are frequently going to repeat an import process, you can save the
    steps to perform the action more quickly.
14. Click the Close button.
15. Open the table to verify that the records were successfully added into the
    table.

## Conclusion:

- Importing data records from Excel can save time and effort, especially when
  dealing with a large amount of data.
- It is important to format the Excel file correctly and close the Access table
  before starting the import process.
- The Import Spreadsheet Wizard guides the user through the process, and the
  option to save import steps can be useful for frequently repeated imports.

# SLO 1.5 - Exploring Datasheet View

## Purpose of Datasheet View
- Add, modify, delete, or view data records
- Move among the fields and records
- Adjust the display of data within a table
- Hide fields and display totals in Datasheet view

## Navigation Buttons
- Provide additional methods to move among table records
- Located in the bottom-left corner of a table in Datasheet view
- Contains navigation buttons, filter status message, and a search box
- Navigation buttons functionality explained in Table 1-9

## Refine the Datasheet Layout
- Adjust the default display settings in Datasheet view several ways
- Settings change both the view of the table on the screen and the printed contents of a table
- Modify the field order
- Select a display font face and font size
- Modify the row height
- Modify the column width

### Modify the Field Order
- Move a single column or a group of adjacent columns
- Place your pointer over the field name of the column you wish to move
- Click to select the column
- Click, hold, and drag the pointer to move the column to the new location

### Select a Display Font Face and Font Size
- Access provides many standard text formatting options available in other Microsoft Office programs
- Font face specifies the shape of the text
- Font size specifies the size of text, measured in points (pt.)
- Font style options include bolding, italicizing, and underlining
- Changes to font face, size, or style are applied to the entire datasheet

### Modify the Row Height
- Adjust the row height to change the amount of space that displays between rows
- When you change the row height in Access, all of the rows change to the new height
- Place your pointer on the border of a record selector of a row in your table
- Click, hold, and drag the resize arrow to increase or decrease the height of the rows

### Modify the Column Width
- Adjust the column width to change the amount of text that displays
- Set each field to a desired width
- Select a set of adjacent columns and set the same width to all of them in one operation
- Place your pointer on the right border of the field name of the column you wish to change
- Click, hold, and drag the resize arrow to increase or decrease the width of the field(s)

## Conclusion 1
- Datasheet view is primarily used to add, modify, delete, or view data records and provides many options to adjust the display of data within a table.
- Navigation buttons and the Navigation bar allow for easy movement among table records.
- Refining the Datasheet layout can improve the readability and presentation of data in the table.

## How to Modify the Column Width of All Fields Using AutoFit

1. Click on the Select All button in the upper-left corner of the datasheet to select the entire datasheet. All of the rows and columns are highlighted in blue.
2. Place the pointer on the right border of any of the field names.
3. Double-click the resize arrow to have all the fields adjust to their widest entry or to the width of the field caption if that is larger than the contents.

### Another way to modify the column width
- Right-click the field name to open the context menu and select Field Width to open the Column Width dialog box. Enter the desired width and click OK or click Best Fit to automatically adjust the size to the widest entry.

## Modify the Field Alignment

- The default field alignment is determined by the data type, or position relative to the edges of each field.
- Change the default alignment by selecting the field and then clicking the desired alignment button in the Text Formatting group [Home tab].
- Alignment changes to a field are applied across all the database records, not just one cell.

## Display Gridlines

- Gridlines help visually frame the rows and columns in a datasheet with a border.
- To change the gridline settings, click the Gridlines button in the Text Formatting group [Home tab] to display the gallery of gridline options and then select the desired setting.

### Gridline Display Options
- Horizontal and vertical gridlines in a gray color.

## Display Alternate Row Colors

- Alternate row colors help you easily scan the records and fields in your datasheet.
- To change the alternate row color settings, click the Alternate Row Color button in the Text Formatting group [Home tab] to display the gallery of color options. Select the desired color.

## Use the Datasheet Formatting Dialog Box

- The Datasheet Formatting dialog box enables you to change several other formatting options that are not available through the Home tab. These options include the Cell Effect, Background Color, and Gridline Color.

### How to Make Multiple Formatting Changes with the Datasheet Formatting Dialog Box
1. Click the Text Formatting launcher [Home tab, Text Formatting group] to open the Datasheet Formatting dialog box.
2. Select the desired settings of the different options.
3. Click OK when finished.

## Hide and Unhide Table Fields

- If a table has many fields, it may not be possible to see all of them at the same time. You can hide data to optimize your screen space. Access does not provide any visual indications that table columns have been hidden.

### How to Hide a Table Field
- Select the column to be hidden.
- Click the More button [Home tab, Records group] and then select Hide Fields.

### How to Unhide Table Fields
- Click the More button [Home tab, Records group] and then select Unhide Fields to open the Unhide Columns dialog box.
- Checked fields display in the table, and unchecked fields are hidden.
- Click all of the field check boxes that you wish to unhide.
- Click Close.

## Add Total Rows

- At times you may wish to have Access calculate totals based on the data in a table. You can summarize data in a table with a Total row.

### How to Add a Total Row
- Click the Totals button [Home tab, Records group].
- A Total row is added below the append row.
- Click the Total row cell of the field you wish to summarize, click the drop-down arrow in that cell, and select the desired function.
- The available functions vary depending on the data type of the field.
- The table updates and the calculated value displays in the Total row cell.
- You can add totals to multiple fields in the table

## Save Changes

- If you make changes to the design settings of the datasheet layout, you must save the table for those changes to be permanent.
- Select one of the previously discussed options for saving the design changes to your table.

## Conclusion 2

- Microsoft Access provides several formatting options for modifying the datasheet layout to make it more readable and user-friendly.
- Users can modify the column width of all fields using AutoFit, modify the field alignment, display gridlines and alternate row colors, use the Datasheet Formatting dialog box, hide and unhide table fields, and add total rows.
- The changes made to the design settings must be saved for them to be permanent.