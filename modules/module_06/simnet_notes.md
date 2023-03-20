# Excel - Chapter 1 - Creating and Editing Workbooks

<!-- TOC -->
* [Excel - Chapter 1 - Creating and Editing Workbooks](#excel---chapter-1---creating-and-editing-workbooks)
* [SLO 1.1: Creating, Saving, and Opening a Workbook](#slo-11--creating-saving-and-opening-a-workbook)
* [Create a New Workbook](#create-a-new-workbook)
* [Save and Close a Workbook](#save-and-close-a-workbook)
* [Open a Workbook](#open-a-workbook)
* [Save a Workbook with a Different File Name](#save-a-workbook-with-a-different-file-name)
* [Workbook File Formats](#workbook-file-formats)
* [SLO 1.2 - Entering and Editing Data](#slo-12---entering-and-editing-data)
* [Worksheet Overview](#worksheet-overview)
* [Display Options](#display-options)
* [Entering Data](#entering-data)
  * [How to Enter Data into a Workbook](#how-to-enter-data-into-a-workbook)
* [Editing Cell Contents](#editing-cell-contents)
  * [How to Edit Cell Contents](#how-to-edit-cell-contents)
* [Clearing Cell Contents](#clearing-cell-contents)
  * [How to Clear Cell Contents](#how-to-clear-cell-contents)
* [Aligning and Indenting Text](#aligning-and-indenting-text)
  * [How to Align and Indent Text](#how-to-align-and-indent-text)
* [Selecting Cells](#selecting-cells)
* [Using the Fill Handle](#using-the-fill-handle)
  * [How to Use the Fill Handle to Create a Series](#how-to-use-the-fill-handle-to-create-a-series)
* [AutoComplete and Pick From Drop-down List](#autocomplete-and-pick-from-drop-down-list)
  * [How to Use AutoComplete and Pick From Drop-down List](#how-to-use-autocomplete-and-pick-from-drop-down-list)
* [Cut, Copy, and Paste Cell Contents](#cut-copy-and-paste-cell-contents)
  * [Cut Cell Contents](#cut-cell-contents)
    * [Move or Cut Cell Contents](#move-or-cut-cell-contents)
    * [MORE INFO](#more-info)
  * [Copy Cell Contents](#copy-cell-contents)
    * [MORE INFO](#more-info-1)
  * [Paste Options](#paste-options)
    * [Paste Options Table](#paste-options-table)
    * [How to Use an Operation with Paste Special Command](#how-to-use-an-operation-with-paste-special-command)
* [SLO 1.3 - Using Functions](#slo-13---using-functions)
  * [The AutoSum Button](#the-autosum-button)
  * [How to Use the`SUM`Function](#how-to-use-the-sum-function)
  * [Function Syntax](#function-syntax)
  * [Copy a Function](#copy-a-function)
  * [Edit an Argument](#edit-an-argument)
  * [How to Resize an Argument Range](#how-to-resize-an-argument-range)
    * [`AVERAGE`Function](#average-function)
  * [How to Use the`AVERAGE`Function](#how-to-use-the-average-function)
    * [`MAX`and`MIN`Functions](#max-and-min-functions)
  * [How to Use the`MAX`Function with Nonadjacent Cells](#how-to-use-the-max-function-with-nonadjacent-cells)
* [SLO 1.4 - Formatting Cell Data](#slo-14---formatting-cell-data)
* [Formatting Cell Data](#formatting-cell-data)
  * [Font Face, Size, Style, and Color](#font-face-size-style-and-color)
    * [HOW TO: Customize Font, Style, Font Size, and Font Color](#how-to--customize-font-style-font-size-and-font-color)
  * [The Format Painter](#the-format-painter)
    * [HOW TO: Use the Format Painter](#how-to--use-the-format-painter)
  * [Number Formats](#number-formats)
    * [HOW TO: Format Numbers](#how-to--format-numbers)
* [Custom Borders and Fill](#custom-borders-and-fill)
  * [Add Custom Borders](#add-custom-borders)
  * [Use Fill Choices](#use-fill-choices)
* [Merge Cells](#merge-cells)
  * [Merge Cells Horizontally](#merge-cells-horizontally)
  * [Merge Cells Vertically](#merge-cells-vertically)
* [Alignment Group](#alignment-group)
  * [Fine-tuning Label Appearance](#fine-tuning-label-appearance)
  * [Change Orientation](#change-orientation)
  * [Center Across Selection](#center-across-selection)
* [Wrap Text](#wrap-text)
  * [Wrap Text in a Cell](#wrap-text-in-a-cell)
* [Cell Styles](#cell-styles)
  * [Apply Cell Styles](#apply-cell-styles)
  * [What is a Cell Style?](#what-is-a-cell-style)
* [Workbook Themes](#workbook-themes)
  * [Change the Workbook Theme](#change-the-workbook-theme)
  * [What is a Workbook Theme?](#what-is-a-workbook-theme)
  * [Apply a Custom Theme](#apply-a-custom-theme)
* [SLO 1.5 - Modifying Columns, Rows, and Sheets](#slo-15---modifying-columns-rows-and-sheets)
  * [Adjust Column Width and Row Height](#adjust-column-width-and-row-height)
  * [AutoFit Columns and Rows](#autofit-columns-and-rows)
* [Insert and Delete Columns and Rows](#insert-and-delete-columns-and-rows)
  * [How to: Insert a Column or a Row](#how-to--insert-a-column-or-a-row)
  * [Another Way: Delete Columns or Rows](#another-way--delete-columns-or-rows)
* [The Insert and Delete Dialog Boxes](#the-insert-and-delete-dialog-boxes)
* [Hide and Unhide Columns and Rows](#hide-and-unhide-columns-and-rows)
* [Insert and Delete Worksheets](#insert-and-delete-worksheets)
* [Rename Worksheets and Change Tab Color](#rename-worksheets-and-change-tab-color)
* [Move and Copy Worksheets](#move-and-copy-worksheets)
* [SLO 1.6 - Navigating in a Workbook](#slo-16---navigating-in-a-workbook)
  * [Common Navigation Commands](#common-navigation-commands)
    * [MORE INFO](#more-info-2)
  * [The Name Box](#the-name-box)
    * [HOW TO: Navigate Using the Name Box](#how-to--navigate-using-the-name-box)
  * [The Go To Command](#the-go-to-command)
    * [ANOTHER WAY](#another-way)
      * [HOW TO: Navigate in a Workbook Using the Go To Command](#how-to--navigate-in-a-workbook-using-the-go-to-command)
    * [The Find Command](#the-find-command)
    * [Wildcard Characters](#wildcard-characters)
      * [ANOTHER WAY](#another-way-1)
      * [HOW TO: Search for Data Using the Find Command](#how-to--search-for-data-using-the-find-command)
* [SLO 1.7 - Modifying Screen Appearance of a Workbook](#slo-17---modifying-screen-appearance-of-a-workbook)
  * [Workbook Views](#workbook-views)
    * [How to Switch Workbook Views Using the Status Bar](#how-to-switch-workbook-views-using-the-status-bar)
  * [Zoom Options](#zoom-options)
  * [Freeze Panes](#freeze-panes)
    * [How to Freeze and Unfreeze Panes](#how-to-freeze-and-unfreeze-panes)
  * [Split a Worksheet into Panes](#split-a-worksheet-into-panes)
    * [How to Split a Worksheet](#how-to-split-a-worksheet)
  * [Hide or Unhide Worksheets](#hide-or-unhide-worksheets)
    * [How to Hide and Unhide Worksheets](#how-to-hide-and-unhide-worksheets)
  * [Switch Windows Command](#switch-windows-command)
    * [How to Switch Windows Using the Ribbon](#how-to-switch-windows-using-the-ribbon)
  * [View Multiple Worksheets](#view-multiple-worksheets)
    * [How to View Two Worksheets at the Same Time](#how-to-view-two-worksheets-at-the-same-time)
* [SLO 1.8 - Managing Page, Print, and Document Settings](#slo-18---managing-page-print-and-document-settings)
  * [The Page Setup Dialog Box](#the-page-setup-dialog-box)
    * [Page Tab](#page-tab)
    * [Margins Tab](#margins-tab)
    * [Header/Footer Tab](#headerfooter-tab)
    * [Sheet Tab](#sheet-tab)
  * [Margins, Page Orientation, and Paper Size](#margins-page-orientation-and-paper-size)
  * [Headers and Footers](#headers-and-footers)
    * [Preset Headers and Footers](#preset-headers-and-footers)
  * [Page Breaks](#page-breaks)
* [Chapter Summary](#chapter-summary)
<!-- TOC -->

# SLO 1.1: Creating, Saving, and Opening a Workbook

# Create a New Workbook

By default, a new workbook includes one worksheet. Here are the ways to create a
new workbook:

1. Launch Excel and press Esc to leave the Start page and open a blank workbook.
2. Click Blank workbook from the Excel Start page to create a new blank
   workbook.
3. Click the File tab from an open Excel window, and select **New** in the left
   pane in the Backstage view.

# Save and Close a Workbook

To save a new workbook:

1. Click the **File** tab to display the Backstage view.
2. Select Save As in the left pane to display the Save As area.
3. Select the location to save your workbook.
4. Type the file name in the file name box.
5. Click **Save**.

To re-save a workbook with the same file name, press **Ctrl+S** or click the
Save
button on the title bar.

To close a workbook, click the File tab and choose Close or press **Ctrl+W** or
**Ctrl+F4**.

# Open a Workbook

To open a workbook:

1. Click the File tab to open the Backstage view.
2. Click Open to display the Open area.
3. Select the location and folder where the workbook is stored.
4. Locate and click the workbook name in the pane on the right.
5. Click Open.

# Save a Workbook with a Different File Name

To save a workbook with a different file name:

1. Click the File tab to open the Backstage view.
2. Click Save As to display the Save As area.
3. Navigate to and select the folder name.
4. Type or edit the workbook name in the file name area.
5. Click Save to save the file and return to the workbook.

# Workbook File Formats

Excel workbooks are saved as .xlsx files, and you can save a workbook in other
formats for ease in sharing data. The table below lists common formats for
saving an Excel workbook.

| Type of Document             | File Name Extension | Uses of This Format                                                                                                                    |
|------------------------------|---------------------|----------------------------------------------------------------------------------------------------------------------------------------|
| Excel Workbook               | .xlsx               | Excel workbook compatible with versions 2010 and later.                                                                                |
| Excel Macro-Enabled Workbook | .xlsm               | Excel workbook with embedded macros.                                                                                                   |
| Excel 97-2003 Workbook       | .xls                | Excel workbook compatible with older versions of Microsoft Excel.                                                                      |
| Excel Template               | .xltx               | Model or sample Excel workbook stored in the Custom Office Templates folder.                                                           |
| Excel Macro-Enabled Template | .xltm               | Model or sample Excel workbook with embedded macros stored in the Custom Office Templates folder.                                      |
| Portable Document Format     | .pdf                | An image of the workbook for viewing that can be opened with reader software.                                                          |
| Text (tab-delimited)         | .txt                | Data only with columns separated by a tab character. File can be opened by many applications. Windows and Mac formats are available.   |
| Comma Separated Values       | .csv                | Data only with columns separated by a comma. File can be opened in Excel and many applications. Windows and Mac formats are available. |
| OpenDocument Spreadsheet     | .ods                | Workbook for the Open Office suite as well as Google Docs.                                                                             |
| Web Page                     | .htm, .html         | Excel workbook formatted for posting on a website that includes data, graphics, and linked objects.                                    |

# SLO 1.2 - Entering and Editing Data

# Worksheet Overview

- A worksheet consists of columns and rows
- Column labels are letters and row labels are numbers
- A cell is the intersection of a column and a row and has a cell reference
- A range is a rectangular group of cells

# Display Options

- Use the View tab to change gridlines, headings, Formula bar, and ruler display
  options

# Entering Data

- Data can be labels, values, or formulas
- Use apostrophe (**'**) to enter numbers as labels
- Use Formula bar to view entire entry when a label is longer than the cell

## How to Enter Data into a Workbook

1. Select the cell and type the data
2. Press Enter, Tab, directional arrow key, Ctrl+Enter, or click Enter button in
   Formula bar to complete entry and activate cell

# Editing Cell Contents

1. To edit data as you type, use Backspace or Delete keys
2. Double-click cell or press F2 to start Edit mode
3. Edit data in the cell or in the Formula bar and press Enter or completion key

## How to Edit Cell Contents

1. Double-click the cell to be edited or press F2
2. Click to position the insertion point in the cell or Formula bar
3. Edit the data and press Enter or completion key

# Clearing Cell Contents

1. To replace existing data, select the cell and type the new data
2. To delete data from a cell, select the cell and press Delete or click Clear
   button in Editing group
3. Use Clear All, Clear Formats, or Clear Contents to remove different aspects
   of a cell's content

## How to Clear Cell Contents

1. Select the cell or cells
2. Press Delete on the keyboard or use Clear button in Editing group

# Aligning and Indenting Text

1. Use Alignment group to change vertical and horizontal alignment of data in a
   cell
2. Use Increase and Decrease Indent buttons to offset data from the left edge of
   a cell

## How to Align and Indent Text

1. Select the cell or cells
2. Click horizontal text alignment button for left, center, or right alignment
3. Click vertical text alignment button for top, middle, or bottom alignment
4. Click indent button to increase or decrease the indent

# Selecting Cells

1. Select a single cell by clicking it
2. Select a range by clicking the first cell and dragging the pointer in any
   direction to select adjacent cells
3. Use Table 1-2 to learn additional selection methods

# Using the Fill Handle

1. Use the Fill Handle to create a series or copy data
2. Drag the Fill Handle across the cell range for the series or to copy data
3. Release the pointer and use Auto Fill Options button to complete the series

## How to Use the Fill Handle to Create a Series

1. Type the first item in the series and press Enter
2. Type two or three entries to identify a custom series
3. Select the cell with the entry and select all cells with data that identify
   the pattern
4. Point to the Fill Handle to display the Fill pointer and drag it through the
   last cell for the series
5. Release the pointer and use the Auto Fill Options button to choose an option

# AutoComplete and Pick From Drop-down List

1. AutoComplete displays a suggested label in a column when you type a character
   or two that matches a label already in the column
2. Pick From Drop-down List is a sorted list of labels in a column
3. Use these commands to complete new data or edit an existing entry

## How to Use AutoComplete and Pick From Drop-down List

1. Type the first character in a label
2. Type a second and third character when many labels begin with the same letter
3. Press Enter to accept the Auto

# Cut, Copy, and Paste Cell Contents

## Cut Cell Contents

### Move or Cut Cell Contents

Use one of three methods to cut data:

- Drag and drop to quickly move cells within a visible range on the worksheet.
    1. Select the cell(s) to be moved.
    2. Point to a border of the selection to display a move pointer.
    3. Click and drag the selection to the new location.
    4. Release the pointer to paste the data.
- Click the Cut button [Home tab, Clipboard group].
    1. Select the cell or range to be moved.
    2. A moving border surrounds the source cell(s).
    3. Select the destination cell for data that was cut.
    4. Select the top-left cell in a destination range.
    5. Click the Paste button [Home tab, Clipboard group].
- Right-click the destination cell and select Insert Cut Cells to avoid
  overwriting existing data.

### MORE INFO

The Office Clipboard stores up to 24 cut or copied items from Office
applications. Items are available for pasting in any application. Click the
Clipboard launcher to open the Clipboard task pane.

## Copy Cell Contents

1. Select the cell(s) to be copied.
2. Point to a border of the selection to display the move pointer.
3. Press Ctrl to display the copy pointer.
4. Drag the selection to the desired location.
5. Release the pointer to place the copy.
6. Continue to press Ctrl and drag to the next location for another copy.
7. Release the pointer and then release Ctrl to complete the task.

### MORE INFO

Copying data places a duplicate on the Clipboard so that it can be pasted in
another location. When you use a Ribbon, keyboard, or context menu command, you
can paste the data multiple times.

## Paste Options

When you paste data, you can choose how it is copied in the new location. To see
the options, click the lower half of the Paste
button [Home tab, Clipboard group] to open the Paste Options gallery. The
gallery also displays when you right-click a destination cell.

### Paste Options Table

The table includes Paste command icons and the related tasks.

| Paste Option                  | Description                                                                                                                 |
|:------------------------------|:----------------------------------------------------------------------------------------------------------------------------|
| **Paste Group**               |                                                                                                                             |
| Paste                         | Copies all data and formatting; this is the default option.                                                                 | 
| Formulas                      | Copies all data and formulas, but no formatting.                                                                            | 
| Formulas & Number Formatting  | Copies all data and formulas with formatting.                                                                               | 
| Keep Source Formatting        | Copies all data and formatting; same as Paste.                                                                              | 
| Match Destination Formatting  | Copies data and applies formatting used in the destination.                                                                 | 
| No Borders                    | Copies all data, formulas, and formatting except for borders.                                                               | 
| Keep Source Column Widths     | Copies all data, formulas, and formatting and sets destination column widths to the same width as the source columns.       | 
| Transpose                     | Copies data and formatting and reverses data orientation so that rows are pasted as columns and columns are pasted as rows. | 
| Merge Conditional Formatting  | Copies data and Conditional Formatting rules. This option is available only if the copied data has conditional formatting.  | 
| **Paste Values Group**        |                                                                                                                             |
| Values                        | Copies only formula results (not the formulas) without formatting.                                                          | 
| Values & Number Formatting    | Copies only formula results (not the formulas) with formatting.                                                             |
| Values & Source Formatting    | Copies only formula results (not the formulas) with all source formatting.                                                  |
| **Other Paste Options Group** |                                                                                                                             |
| Formatting                    | Copies only formatting, no data.                                                                                            |
| Paste Link                    | Pastes a 3-D reference to data or a formula.                                                                                |
| Picture                       | Pastes a picture object of the data. Data is static, but the object can be formatted like any object.                       |
| Linked Picture                |                                                                                                                             |

### How to Use an Operation with Paste Special Command

1. Enter or select values for the mathematical operation in the destination
   range.
    - The value(s) will be used to calculate a result with the source data.
2. Select the source data.
    - The number of cells to be copied must equal the number of cells in the
      destination range.
3. Click the Copy button [Home tab, Clipboard group] or press Ctrl+C.
    - A moving border surrounds the cell(s).
4. Select the destination cell or range.
    - You can click the first cell in a range.
5. Click the bottom half of the Paste button [Home tab, Clipboard group].
    - The Paste Options gallery displays.
6. Choose Paste Special (Figure 1-19).
7. Select the radio button for the operation to be performed and click OK.
    - The arithmetic executes in the destination cell and replaces the original
      cell(s).

# SLO 1.3 - Using Functions

- A formula calculates a result for numeric data, while a function is a built-in
  formula.
- The terms "formula" and "function" are used interchangeably.

## The AutoSum Button

- The AutoSum button is available on the Home tab [Editing group] and on the
  Formulas tab [Function Library group].
- It displays the `SUM` function by default, but can also be used to choose
  `AVERAGE`, Count Numbers, `MAX`, or `MIN`.
- To use the`SUM`function:
    1. Click the cell where you want to show the total.
    2. Click the AutoSum button [Home tab, Editing group].
    3. Excel inserts the function `=SUM()` with a suggested range of cells to be
       added between the parentheses.
    4. If the range is correct, press **Enter**, **Ctrl+Enter**, or click the
       Enter button in the Formula bar to complete the function.
    5. If the suggested range is not correct, select a different range or choose
       cells individually and then press a completion key.
    6. The sum displays in the cell, and the function is visible in the Formula
       bar.
- Another way to enter the`SUM`function is by typing `=SUM()` in a cell by
  pressing **Alt** and the **+**/**=** key on the top keyboard row.

## How to Use the`SUM`Function

1. Click the cell for the total.
2. Click an empty cell at the bottom of a column of values or at the end of a
   row
   of values.
3. Click the AutoSum button [Home tab, Editing group].
4. A proposed range is adjacent and above or to the left of the formula cell.
5. The Argument ScreenTip below the cell displays the elements needed.
6. Press Enter to accept the range and complete the function.
7. Drag the pointer to select a different range before pressing Enter.
8. Alternatively, click the cell for the total and click the AutoSum
   button [Home tab, Editing group].
9. An adjacent cell range is selected.
10. Click the first cell to be included and type , (a comma).
11. Click the next cell to be added and type , (a comma).
12. Separate nonadjacent cells or ranges with a comma.
13. Cells or ranges are color coded to the argument in the formula.
14. Press Enter to complete the function.
15. More information:
    - Click the empty cell adjacent to a column or row of values and
      double-click the AutoSum button to insert the`SUM`function.

## Function Syntax

- An Excel function has syntax, the required elements and the order of those
  elements for the function to work.
- Every function starts with the equals sign (=) followed by the name of the
  function and a set of parentheses.
- Within the parentheses, you enter the argument(s).
- An argument is the cell reference or value required to complete the function.
- A function can have a single argument or multiple arguments.
- In `=SUM(B5:B9)`, the argument is the range `B5:B9`.
- Multiple arguments are separated by commas.
- The`SUM`function has one required argument, number1, and others are optional.
- The proper syntax for a`SUM`function is: `=SUM(
  number1, [number2], [number3], ...)`
- The`SUM`function ignores empty cells or cells that include labels.

## Copy a Function

- When data follows a pattern, you enter a function once and copy it to similar
  columns or rows.
- A copied formula adjusts to its location in the worksheet.
- To copy the function:
    1. Click the cell with the function to be copied.
    2. Point to the Fill Handle in the lower-right corner.
    3. Click and drag the Fill pointer across the cells where the function
       should be pasted.
    4. Release the pointer.
    5. The results display, the function is copied, and the AutoFill Options
       button appears.
- Alternatively,
    1. Click the cell with the function to be copied and press **Ctrl+C** to
       copy.
    2. Select a range of cells as needed.
    3. Click the destination cell and press **Ctrl+V** to paste.
    4. Click the first cell when the destination is a range.
    5. Press ESC to turn off the moving border.

## Edit an Argument

- Functions and formulas automatically recalculate when you edit values used in
  the calculation.
- Automatic calculation ensures that your results update as often as needed when
  you make changes to your work.
- To change a cell or range argument in a function:
    1. Use Edit mode to enter a different cell reference in the Formula bar or
       in the cell.
    2. You can often use the Range Finder to drag or select a new range.
    3. The Range Finder is an Excel feature that highlights and color codes
       formula cells as you enter or edit a formula or function.
    4. Drag a Range Finder handle to select a different argument range,
       expanding or shrinking the number of included cells.
- Another way to edit an argument is to double-click the cell with the function
  and edit the argument in the cell.

## How to Resize an Argument Range

1. Double-click the cell with the function to start Edit mode.
2. Drag the selection handle up to shrink the argument range.
3. Drag the selection handle down to expand the argument range.
4. For ease in resizing a range, leave a blank row above or a blank column to
   the
   left of the function cell.
5. Press Enter to complete the change.
6. Click the Enter button in the Formula bar.

### `AVERAGE`Function

1. The `AVERAGE` function calculates the arithmetic mean of a range of cells.
2. The mean is determined by adding all the values and dividing the result by
   the
   number of values.
3. The `AVERAGE` function ignores empty cells or cells with labels.
4. When calculating an average, do not include total values in the range to be
   averaged.
5. The `AVERAGE` function has one required argument, number1, and others are
   optional.
6. The proper syntax for an`AVERAGE`function is: `=AVERAGE(
   number1, [number2], [number3], ...)`

## How to Use the`AVERAGE`Function

1. Click the cell for the function.
2. Click the AutoSum button arrow [Home tab, Editing group] and choose Average.
3. The function and an assumed argument range display in the Formula bar and in
   the cell.
4. The Argument ScreenTip below the cell displays the syntax.
5. The next argument displays in bold in the ScreenTip.
6. Select a different range or type cell addresses separated by commas as
   needed.
7. Press Enter or any completion key.

### `MAX`and`MIN`Functions

1. The`MAX`function finds the largest value in a range, and the`MIN`function
   finds the smallest value.
2. Each function has one argument, numberN.
3. The argument can be a cell or a range, up to 255 cells (or ranges).
4. There must be at least one cell or range argument.
5. The proper syntax for`MAX`and`MIN`functions is: `=MAX(
   number1, [number2], [number3], ...)` and `=MIN(
   number1, [number2], [number3], ...)`

## How to Use the`MAX`Function with Nonadjacent Cells

1. Click the cell for the function.
2. Click the AutoSum button arrow [Home tab, Editing group] and choose Max.
3. The function displays in the cell and in the Formula bar.

# SLO 1.4 - Formatting Cell Data

# Formatting Cell Data

- A new workbook uses a default theme named Office which is a collection of
  fonts, colors, and special effects.
- You can change the theme as well as individual font attributes, or you can add
  fill and borders to highlight or emphasize data.

## Font Face, Size, Style, and Color

- A font is a type design for alphanumeric characters, punctuation, and keyboard
  symbols.
- Font size specifies the size of the character, measured in points (pt.).
- Font style refers to the thickness or angle of the characters and includes
  settings such as bold, underline, or italic.
- Font color is the hue of the characters in the cell.
- Apply any of these font attributes to a single cell, a range of cells, or
  selected characters in a cell.

### HOW TO: Customize Font, Style, Font Size, and Font Color

1. Select the cell or range to be formatted.
2. Click the Font drop-down list [Home tab, Font group].
3. Point to a font name.
4. Click the Font Size drop-down list [Home tab, Font group] and select a size.
5. Type a custom size in the Font Size area.
6. Click the Increase Font Size or Decrease Font Size
   button [Home tab, Font group] to adjust the font size to the next or previous
   size in the Font Size drop-down list.
7. Click the Bold, Italic, or Underline button [Home tab, Font group] to apply a
   style.
8. Click the Font Color drop-down list [Home tab, Font group] and select a
   color.

## The Format Painter

- The Format Painter copies formatting attributes and styles from one cell to
  another cell or range.

### HOW TO: Use the Format Painter

1. Select a cell that contains the formatting you want to copy.
2. Click the Format Painter button [Home tab, Clipboard group].
3. Double-click the Format Painter button to lock it for painting formats to
   multiple areas in the worksheet.
4. Click the cell or drag across the range to be formatted.

## Number Formats

- Format values in your work with currency symbols, decimal points, commas,
  percent signs, and more so that data are quickly recognized and understood.

### HOW TO: Format Numbers

1. Select the cell or range of values.
2. Click the Number Format drop-down list to choose a format.
3. Click a specific button [Home tab, Number group] to apply its format.
4. Press Ctrl+1 to open the Format Cells dialog box and click the Number tab to
   select a format.
5. Click the Increase Decimal or Decrease Decimal button to control the number
   of decimal places.

# Custom Borders and Fill

## Add Custom Borders

1. Select the cell or range.
2. Open the Format Cells dialog box:
    - Click the Font launcher [Home tab, Font group].
    - Press Ctrl+1.
    - Right-click a cell in the range and select Format Cells.
3. Click the Border tab.
4. To remove all borders, click None.
5. Choose a color from the Color drop-down list.
6. Select the color, the line style, and then the location.
7. Select a line style in the Style area.
8. Click Outline in the Presets area to apply an outside border.
9. Build a custom border by clicking the desired icons in the Border area or by
   clicking the desired position in the Preview area.
10. Click the Fill tab.
11. Select a color tile in the Background Color area.
12. Select a color from the Pattern Color list and select a pattern from the
    Pattern Style list.
13. Click Fill Effects to apply a gradient. Choose two colors, a shading style,
    and a variant.
14. Click OK to close each dialog box.

## Use Fill Choices

- Use fill choices carefully because they affect the readability of your data.
- From the Fill tab, you can design patterned or gradient fills.
- A pattern fill uses crosshatches, dots, or stripes.
- A gradient is a blend of two or more colors.

# Merge Cells

## Merge Cells Horizontally

1. Select the cells to be merged and centered.
2. Include the label and blank cells to the right to span the width of your
   data.
3. Click the Merge & Center button [Home tab, Alignment group].
    - The data is horizontally centered across the selected cell range.
4. Choose Merge Across to combine the cells and left or right align the label.

## Merge Cells Vertically

1. Select cells in a column to be merged.
2. Include the label and blank cells below to span the desired height.
3. Click the Merge & Center button [Home tab, Alignment group].
    - The data is centered and aligned at the bottom of the merged cell.
4. Change the text orientation and other alignment choices to better position a
   label.

# Alignment Group

## Fine-tuning Label Appearance

- The Alignment group on the Home tab has three commands: Merge & Center
  commands, Wrap Text, and Orientation choices.
- The Merge & Center command combines two or more cells into one cell.
- Use this command to center a main label over multiple columns or across
  multiple rows.
- The Wrap Text command displays a label on multiple lines within the cell.
    - You control where the label splits by inserting a manual line break.
- The Orientation button provides a way to add visual impact to labels by
  displaying them at an angle.

## Change Orientation

1. Select the cell or range to be formatted.
2. Click the Orientation button [Home tab, Alignment group].
3. A list of common angles displays.
4. Click the name of an orientation option.
    - The data is reformatted.
5. Row height increases to accommodate the labels.
6. Select the cell or range with angled labels.
7. Click the Orientation button [Home tab, Alignment group] and choose Format
   Cell Alignment.
8. The Format Cells dialog box opens to the Alignment tab.
9. The angle setting displays as degrees and by the red triangle in the preview.
10. Drag the red triangle to 0 degrees to return to the default orientation.
    - You can type 0 or use the spinner arrows.
11. Row height decreases as needed.

## Center Across Selection

- The Center Across Selection command horizontally centers multiple rows across
  multiple columns.
- When you have two or three rows of labels to center over the same range:
    1. Select the cell range that includes the data to be centered.
    2. Select all rows with data to be centered.
    3. Select columns to identify the range over which the labels should be
       centered.
    4. Click the Alignment launcher [Home tab, Alignment group].
    5. Click the Horizontal drop-down arrow and choose Center Across Selection.
        - Each label is centered across the selected range.

# Wrap Text

## Wrap Text in a Cell

1. Type the first word or two in the label.
2. Press Alt+Enter to insert a manual line break.
    - The cell is still active.
3. Finish typing the label and press Enter to finish.
    - The label displays on multiple lines in the cell.
4. The first line of the label displays in the Formula bar.
5. The Wrap Text command has been activated.
6. Select the cell with the label and click the Expand Formula Bar arrow at the
   right edge of the Formula bar.
    - The Formula bar expands to display two lines.
    - The Formula bar displays a vertical scroll bar for labels that are entered
      on more than three lines.
7. Alternatively, type the label without line breaks.
8. Select the cell and click the Wrap Text button in the Alignment
   group [Home tab].
    - The label displays on multiple lines based on the column width.
    - The Wrap Text button is a toggle to turn the command on and off.
9. After using the Wrap Text command, adjust the column width or row height to
   better display the label.

# Cell Styles

## Apply Cell Styles

1. Select the cell or range.
2. Click the More button in the Cell Styles gallery.
3. Click the Cell Styles button [Home tab, Styles group] to open the gallery.
4. Point to a style name to see a Live Preview in the worksheet.
5. Click a style name to apply it.

## What is a Cell Style?

- A cell style is a set of formatting elements that includes font style, size,
  color, alignment, borders, and fill, as well as number formats.
- When you apply a cell style, the style format overwrites formatting already
  applied.
- After a cell style is applied, however, you can change any of the attributes.
- The Cell Styles gallery is in the Styles group on the Home tab.

# Workbook Themes

## Change the Workbook Theme

1. Click the Theme button [Page Layout tab, Themes group].
2. The Themes gallery opens.
3. Point to a theme name.
4. Live Preview displays the worksheet with new format settings.
5. Click a theme icon to apply a different theme.
6. Display the Font button list [Home tab, Font group] to see the theme font
   names at the top of the list.
7. Display the Font Color and Fill Color button galleries [Home tab, Font group]
   to see new theme colors.

## What is a Workbook Theme?

- A workbook theme is a professionally designed set of fonts, colors, and
  effects.
- When you change the workbook theme, data are reformatted with the new theme
  fonts, colors, and effects.
- Change the theme to quickly restyle a worksheet without having to individually
  edit cell formats.
- A theme includes a font name for headings and one for body text.
- You can see the names of Theme Fonts at the top of the Font list.
- Theme Colors are identified in the color palettes for the Fill and Font Color
  buttons [Home tab, Font group].
- Cell styles use font and color settings from the theme.
- The Themes gallery lists built-in themes, and additional themes are available
  online.
- You can also create and save your own theme.

## Apply a Custom Theme

1. Click the Theme button [Page Layout tab, Themes group].
2. The Themes gallery opens.
3. Click the Browse for Themes link at the bottom of the gallery.
4. The Microsoft Office website opens in your default browser.
5. Select a theme and click Download.
6. Follow the instructions to download and install the theme.
7. Return to Excel and click the Themes button.
8. Click the Custom heading.
9. Click the new custom theme to apply it to the workbook.

# SLO 1.5 - Modifying Columns, Rows, and Sheets

- A worksheet includes a default number of rows and columns in the default width
  and height for the workbook theme. Over 1 million rows and more than 16,000
  columns are available. You can modify the width of a column, insert or delete
  columns or rows, change the height of a row, and more.
- Unused rows and columns are ignored when you print a worksheet.

## Adjust Column Width and Row Height

- The default column width in a worksheet using the Office theme is 8.43
  characters. Change column width to any value between 0 and 255 characters. The
  default row height is 15 points.
- You can change column width or row height by dragging the border between
  column or row headings, by using the context menu, or by selecting a command
  from the Format button list [Home tab, Cells group].
- To change the column width or row height:
    1. Click a cell in the column to be adjusted.
    2. Point to the border between the column heading and the heading on its
       right.
    3. The resize pointer is a double-pointed arrow.
    4. Drag the pointer to the right to widen the column. Data that was
       truncated displays when the column is wide enough.
    5. Drag the pointer to the left to make the column narrower. Data may be
       truncated when the column is not wide enough. Values display as a series
       of #### symbols when the column is not wide enough.
    6. Click a row heading.
    7. Point to the border between the row heading and the row heading below.
    8. Point to any border within multiple selected rows.
    9. Drag the pointer up or down to set a new height. Data may be cut off if
       the font is too large for the row height.
- Alternatively:
    1. Select a cell in the row or column to be adjusted.
    2. Select a cell in each of several rows or columns to be adjusted.
    3. Click the Format button [Home tab, Cells group] and choose Row Height or
       Column Width.
    4. The Row Height or Column Width dialog box opens.
    5. Enter a height in points or a column width in characters and click OK.
    6. The row height adjusts for the rows in which you selected cells.
    7. The column width adjusts for the columns in which you selected cells.

## AutoFit Columns and Rows

- The AutoFit feature resizes column width or row height to fit the width or
  height of the longest or tallest entry. AutoFit a column by double-clicking
  the right border of the column heading. To AutoFit a row, double-click the
  bottom border of the row heading. You can AutoFit multiple columns or rows by
  first selecting them and double-clicking any border within the selection.
- To AutoFit a column width or row height:
    1. Point to the border to the right of the column heading or below the row
       heading.
    2. Drag to select more than one column or row.
    3. Point to any border in a selection.
    4. Double-click the resize arrow.
    5. All selected columns and rows adjust to the longest or tallest entry.

# Insert and Delete Columns and Rows

You can insert or delete rows or columns in a worksheet. When you insert or
delete rows or columns, Excel moves existing data to make room for new data or
fills the gap left by deleted data. Functions or formulas are automatically
updated to include an inserted row or column, and they reflect deleted rows or
columns in the argument or range.

## How to: Insert a Column or a Row

1. Right-click the row heading below the row where a new row should appear, or
   right-click the column heading to the right of where a new column should
   appear.
2. Choose "Insert" from the context menu.
    - A new row displays above the selected row. Existing rows move down.
    - An inserted column displays to the left of the selected column and
      existing columns move right.
    - To insert multiple columns or rows, select the number of columns or rows
      that you want to insert.
    - Use any "Insert" command to insert the new columns or rows.

Alternatively, you can:

1. Select a cell in the column to the right of where a new column is to be
   inserted or a cell in the row below where a new row should display.
2. Click the arrow on the "Insert" button in the "Cells" group on the "Home"
   tab.
3. Select "Insert Sheet Columns" or "Insert Sheet Rows".
    - The entire column or row is inserted.
    - The row height adjusts for the rows in which you selected cells.
    - The column width adjusts for the columns in which you selected cells.

## Another Way: Delete Columns or Rows

1. Right-click the row heading for the row to be deleted, or right-click a
   column heading as needed.
2. Choose "Delete" from the context menu.
    - The entire row (or column) is deleted.
    - Remaining rows shift up; remaining columns shift left.
    - Data are deleted and the remaining columns and rows shift to the left or
      up. Most functions or formulas are updated if you delete a row or column
      that is within the argument range.

Alternatively, you can:

1. Select a cell in the column or row to be deleted.
2. Click the arrow on the "Delete" button in the "Cells" group on the "Home"
   tab.
3. Select "Delete Sheet Columns" or "Delete Sheet Rows".
    - The entire column or row is deleted.

# The Insert and Delete Dialog Boxes

1. To insert cells, open the Insert dialog box from the Cells group on the
   Format tab.
2. Select the option to insert cells.
3. Existing cells shift down or right to make room for the inserted cells.
4. To delete cells, open the Delete dialog box from the Cells group on the
   Format tab.
5. Select the option to delete cells.
6. Existing cells shift left or up to fill the gap.
7. You can also right-click a single cell and choose Insert or Delete from the
   context menu to open the dialog box.

# Hide and Unhide Columns and Rows

1. Click the row or column heading of the row or column to be hidden.
2. Right-click one of the selected row or column headings to open the context
   menu.
3. Select Hide to hide the entire row or column.
4. Formula references to hidden cells are maintained.
5. To display hidden columns or rows, drag across column or row headings that
   surround hidden elements.
6. Drag from the column to the left of a hidden column to one column to the
   right of a hidden column.
7. Drag from the row above hidden rows to one row below hidden rows.
8. Right-click one of the selected row or column headings to open the context
   menu.
9. Select Unhide to show the hidden columns or rows.
10. Alternatively, select the column or row headings, click the Format button in
    the Home tab, choose Hide & Unhide, and select the desired command.

# Insert and Delete Worksheets

1. To insert a worksheet, click the New Sheet button to the right of the
   worksheet tabs, or select Insert Sheet from the Insert button in the Home
   tab, Cells group.
2. Alternatively, right-click an existing sheet tab and select Insert to open
   the Insert dialog box.
3. Select Worksheet and click OK.
4. The new sheet is inserted to the left or right of the active sheet and named
   SheetN, where N is the next available number.
5. To delete the active worksheet, right-click the sheet tab and select Delete
   or select Delete Sheet from the Delete button in the Home tab, Cells group.
6. To move to the next worksheet in the tab order, press Ctrl+Page Down. To move
   to the previous sheet, press Ctrl+Page Up.

# Rename Worksheets and Change Tab Color

1. Double-click the worksheet tab or right-click and select Rename to open the
   Rename Sheet dialog box.
2. Type the new name on the tab and press Enter.
3. Right-click the worksheet tab and select Tab Color from the context menu to
   open the palette. Alternatively, select Tab Color from the Format button in
   the Home tab, Cells group.
4. Select a color.

# Move and Copy Worksheets

1. To move a worksheet within the same workbook, simply drag the tab to the
   desired location.
2. To move or copy a worksheet to another workbook, right-click the sheet tab
   and select Move or Copy or select Move or Copy Sheet from the Format button
   in the Home tab, Cells group.
3. In the Move or Copy dialog box, choose a sheet name in the Before sheet list.
4. To copy the sheet into a different workbook, choose its name.
5. Choose (move to end) to place the cut or copied sheet as the last tab.
6. Select the Create a copy box as needed.
7. Click OK to close the dialog box.
8. To copy a worksheet within the same workbook, press Ctrl+Alt and drag the
   pointer to the desired location.
9. Release the pointer first and then release Ctrl+Alt to copy the sheet. A
   copied worksheet within the same workbook has the same name as the original
   followed by (2).
10. To move or copy multiple sheets, select sheet names using Ctrl or Shift. The
    word [Group] displays in the title bar.
11. After you complete the move or copy command, right-click any tab in the
    group and choose Ungroup.

# SLO 1.6 - Navigating in a Workbook

- Navigating and scrolling in a worksheet changes the active cell or repositions
  what you see without moving the insertion point.
- Common keyboard navigation commands move the active cell.
- Rolling the mouse wheel or dragging a scroll bar or a scroll arrow adjusts
  what you see on screen but does not change the active cell.

## Common Navigation Commands

This table lists common navigation commands in a workbook. It lists the command
and the result of each command.

| Command                          | Result                                                                           |
|:---------------------------------|:---------------------------------------------------------------------------------|
| Up, Down, Left, and Right arrows | Moves the active cell up, down, left, or right one cell.                         |
| Tab                              | Moves the active cell one cell to the right.                                     |
| Shfit+Tab                        | Moves the active cell one cell to the left.                                      |
| Page Down                        | Moves the active cell down one screen.                                           |
| Page Up                          | Moves the active cell up one screen.                                             |
| Alt+Page Down                    | Moves the active cell one screen to the right.                                   |
| Alt+Page Up                      | Moves the active cell one screen to the left.                                    |
| Ctrl+Page Down                   | Moves to the next sheet in the workbook.                                         |
| Ctrl+Page Up                     | Moves to the previous sheet in the workbook.                                     |
| Ctrl+Home                        | Moves the active cell to cell A1.                                                |
| Ctrl+End                         | Moves the active cell to the last cell with data (bottom row, rightmost column). |
| End+arrow                        | Moves the active cell to the last cell with data in the direction of the arrow.  |

### MORE INFO

The number of rows or columns moved by navigation commands depends on the number
of rows visible on your screen.

## The Name Box

- Use the Name box at the left edge of the Formula bar to move the insertion
  point to a specific cell or range of cells.
- When a workbook has named ranges, you can click a name to navigate to that
  range.
- Range names are discussed in SLO 2.3: Using Mixed, Relative, and 3D
  References.

### HOW TO: Navigate Using the Name Box

1. Click the Name box.
2. The active cell address displays and is highlighted.
3. Type a single cell address (A1, C10) or a cell range (A1:B5) (Figure 1-59).
4. You can type cell addresses using lowercase letters (a1, a1:b5).
5. Press Enter.
6. The cell or range is selected.
7. Click the Name box arrow.
8. Existing named ranges display in alphabetical order.
9. Click a name in the list.
10. The cell range is selected.

## The Go To Command

- The Go To command [Find & Select button, Home tab, Editing group] opens the Go
  To dialog box.
- Enter a cell reference or range name or select from recently used names and
  cell references.

### ANOTHER WAY

- Press F5 or Ctrl+G to open the Go To dialog box.

#### HOW TO: Navigate in a Workbook Using the Go To Command

1. Select cell A1 to start at the beginning of a worksheet.
2. Press F5 (FN+F5) to open the Go To dialog box.
3. Click the Reference box and type a new cell reference.
4. Type a single cell address (a1, c10).
5. Type a cell range (a1:b5 or c4:c18).
6. Select a reference from the list (Figure 1-60).
7. Click OK.
8. The cell(s) are selected in the worksheet.

### The Find Command

- The Find command locates occurrences of a character string in a workbook.
- A character string is any alphanumeric combination, such as �per� or �13.�
- You can use wildcard characters, restrict the search to the current sheet,
  search by row or column, match case, or look for formats.

### Wildcard Characters

This table lists wildcard characters, their results, and examples.

| Character  | Result and Example                                                                                                |
|:----------:|:------------------------------------------------------------------------------------------------------------------|
|    `?`     | Finds any single character. s?t finds three-character strings such �sit� or �s2t� but not �salt.�                 |
|   `???`    | Finds any three characters. ???rd finds five-character strings such as �third� or �123rd� but not �keyboard.�     |
|    `*`     | Finds any number of characters. m*n finds all character strings that begin with �m� and end with �n.�             |
| `*street*` | Finds all addresses that include the word �street.�                                                               |
|    `~?`    | Finds a question mark in the data.                                                                                |
|    `~*`    | Finds an asterisk in the data.                                                                                    |
| `*date~?`  | Finds character strings that include �date� preceded by any number of characters but followed by a question mark. |
|  `~*See*`  | Finds character strings that display �*See� followed by any number of characters.                                 |

#### ANOTHER WAY

- Press **Ctrl+F** to open the Find and Replace dialog box.

#### HOW TO: Search for Data Using the Find Command

1. Select cell A1 to start at the beginning of the worksheet.
2. Click the Find & Select button [Home tab, Editing group] and select Find.
3. The dialog box includes the Replace tab.
4. Click the Options >> button to expand the Find and Replace dialog box.
5. Click the Find what box and type the search string.
6. Type as many characters as needed to identify the data.
7. Use wildcard characters to shorten typed strings (* and ?).
8. Click the Within arrow to choose the current sheet or the entire workbook.
9. Click the Search arrow to choose By Rows or By Columns.
10. Your choice depends on the data layout.
11. For small files, you may not see any difference in speed.
12. Click the Look In arrow.
13. Choose Formulas to locate any occurrence of a character string or value.
14. Choose Values to ignore character strings or numbers that are part of a
    formula.
15. Select the Match case box if appropriate.
16. The option locates only occurrences that match capitalization as entered in
    the Find what box.
17. Select the Match entire cell contents box as needed.
18. Do not use this option when searching for partial data in a cell.
19. Click Find Next.
20. The cell with the first occurrence of the Find what string is highlighted.
21. Click Find All to display a list of all occurrences of the Find what
    string (Figure 1-61).
22. Resize the dialog box as needed.
23. Click an item in the list to navigate to the cell.
24. Edit the data in the worksheet as needed.
25. Close the dialog box.

# SLO 1.7 - Modifying Screen Appearance of a Workbook

## Workbook Views

- Access workbook views from the `View` tab on the Ribbon.
- Available views include `Normal`, `Page Break Preview`, and `Page Layout`.
- Create custom views or work in `Sheet View`.
- Use the buttons on the right side of the status bar to quickly switch between
  Normal view, Page Layout view, and Page Break Preview.
- Use the Ribbon Display Options button (on the right of the workbook title bar)
  to hide the Ribbon for maximum screen space in any workbook view.

### How to Switch Workbook Views Using the Status Bar

1. Click the `View` tab.
2. The `Normal` button is activated in the Workbook Views group.
3. Click the `Page Layout` button in the Status bar.
4. This view displays the header, the footer, and margin areas. You can edit
   data in Page Layout view.
5. Click the `Page Break Preview` button in the Status bar. When data does not
   fit on a single page, a dashed line indicates where the second page starts.
   You can edit data in Page Break Preview.
6. Click the `Normal` button in the Status bar.

## Zoom Options

- Change the sheet�s magnification to see more of the data at once or to
  scrutinize content more carefully using the Zoom group on the View tab or the
  Zoom controls in the Status bar.
- The Status bar controls include a `Zoom In` button, a `Zoom Out` button, and
  the `Zoom slider`.
- To change zoom levels:
    1. Click the `Zoom In` button in the Status bar. The magnification increases
       in 10% increments with each click.
    2. Click the `Zoom Out` button in the Status bar. The magnification
       decreases in 10% increments with each click.
    3. Drag the `Zoom slider` in the middle of the horizontal line to the left
       or right to set a value. The magnification percentage displays to the
       right of the line.
    4. Click the `Zoom` button on the View tab, Zoom group. The Zoom dialog box
       includes preset values. Alternatively, click the magnification value in
       the Status bar to open the Zoom dialog box.
    5. Choose a magnification level by selecting a radio button in the Zoom
       dialog box.
    6. Select `Fit Selection` for a selected range to fill the screen.
       Choose `Custom` to enter any magnification value.
    7. Click `OK`.
- The `Zoom` group on the View tab has `100%` and `Zoom to Selection` buttons.

## Freeze Panes

- In a worksheet with many rows and columns of information, you may not see all
  the data at once.
- Use the Freeze Panes command to lock rows or columns in view so that you can
  position data for easy review.
- The worksheet divides into two or four panes (sections) with a thin black
  border to indicate the divisions.
- The active cell becomes the top-left corner of the moving pane.

### How to Freeze and Unfreeze Panes

1. Select a cell in the worksheet.
2. Click a cell one row below the last row to be frozen.
3. Click a cell one column to the right of the last column to be locked.
4. Click the `Freeze Panes` button on the View tab, Window group.
5. Select `Freeze Top Row` to lock row 1 in view, regardless of the active cell.
6. Choose `Freeze First Column` to keep the first column in view, regardless of
   the active cell.
7. Select `Freeze Panes`. All rows above and all columns to the left of the
   active cell are locked in position. A thin black border identifies frozen
   rows and columns.
8. Scroll the data in any direction.
9. Click the `Freeze Panes` button on the View tab, Window group.
10. Select `Unfreeze Panes`. Pane borders are removed, and all data is visible.

## Split a Worksheet into Panes

- Use the Split command to divide a worksheet into two or four display panes.
- Each pane displays the sheet, but you can arrange each pane to show different
  data.
- A splitter bar is a light gray bar that spans the window, either horizontally,
  vertically, or both.
- You can drag a splitter bar to resize the panes.

### How to Split a Worksheet

1. Select the cell for the location of a split.
2. To split the worksheet into two horizontal panes, click a cell in column A (
   except cell A1).
3. To split the worksheet into two vertical panes, click a cell in row 1 (but
   not cell A1).
4. To split the worksheet into four panes, click a cell.
5. Click the `Split` button on the View tab, Window group.
6. Drag a splitter bar to resize a pane.
7. Move one pane with a two-pointed arrow pointer. To move all panes at the same
   time, use a four-pointed arrow.
8. Click the `Split` button on the View tab, Window group to remove the split.
9. To remove a window split, double-click a splitter bar.

## Hide or Unhide Worksheets

- A hidden worksheet is a worksheet whose tab does not display.
- Use the Format button on the Home tab or right-click the sheet tab to hide a
  worksheet.
- You cannot hide a worksheet if it is the only sheet in the workbook.

### How to Hide and Unhide Worksheets

1. Right-click the sheet tab to hide.
2. Press `Ctrl` to select multiple sheets, then right-click any one of the
   selected tabs.
3. Select `Hide`. The worksheet tab is hidden.
4. Right-click any worksheet tab.
5. Select `Unhide`. The Unhide dialog box opens.
6. Select the tab name to unhide and click `OK`.

## Switch Windows Command

- The Switch Windows command changes or cycles through open workbooks.
- You can switch windows from the Ribbon, from the Windows taskbar, or with the
  keyboard shortcut `Ctrl+F6`.
- When you click the `Switch Windows` button on the View tab, Window group, a
  menu lists the names of open workbooks.
- To view two workbooks side by side, use the `View Side by Side` command.
- For side-by-side viewing, you can scroll the windows together or separately.
- `Synchronous Scrolling` moves both workbooks in the same direction at the same
  time.

### How to Switch Windows Using the Ribbon

1. Open two or more workbooks.
2. Click the `Switch Windows` button on the View tab, Window group.
3. The names of open workbooks display.
4. Click the name of the workbook to view.
5. Press `Ctrl+F6` to switch to the other open workbook.

## View Multiple Worksheets

- The New Window command opens a new separate window with the same workbook.
- The new window displays the workbook name in the title bar followed by a dash
  and a number.
- The View Side by Side command allows you to position different parts of the
  workbook to compare or monitor changes as you work.
- Use the Arrange All command to create your preferred screen layout.

### How to View Two Worksheets at the Same Time

1. Click the `View` tab and click the `New Window` button on the `Window` group.
2. A second window opens for the same workbook and is active. Both windows are
   maximized.
3. Click the `View Side by Side` button on the `View` tab, Window group.
4. Both windows display the same sheet. The second window displays 2 after the
   file name in the window title bar.
5. The `View Side by Side` button is unavailable when only one window is open.
6. Click the `Synchronous Scrolling` button on the `View` tab, Window group to
   turn off synchronous scrolling as needed. The button toggles the command on
   and off.
7. Arrange and display the data in each window as needed.
8. Click the `Maximize` button in either title bar. Both windows of the same
   workbook are still open.
9. Press `Ctrl+F6` to switch to the other window.
10. Click the `Close` button (X) in the title bar to close the current window.
    The other window displays.
11. Click the `Maximize` button if the window is at a tiled size.

# SLO 1.8 - Managing Page, Print, and Document Settings

## The Page Setup Dialog Box

- Use the Page Setup dialog box to control how a worksheet prints and access
  headers and footers commands.
- To open the Page Setup dialog box:
    - Click the Page Setup launcher on the Page Layout tab, Page Setup group.
    - OR, click the launcher in the Scale to Fit or the Sheet Options group on
      the Page Layout tab.
- The dialog box includes tabs for Page, Margins, Header/Footer, and Sheet.
- Table 1-7 summarizes the available settings on each tab.

### Page Tab

- Set the orientation to Portrait or Landscape.
- Use Scaling to shrink or enlarge the printed worksheet to the paper size.
- Choose a paper size or set print quality.
- Define the first page number.

### Margins Tab

- Adjust top, bottom, left, and right worksheet margins.
- Set header and footer top and bottom margins.
- Center on page horizontally and/or vertically.

### Header/Footer Tab

- Choose a preset header or footer layout.
- Create a custom header or footer.
- Specify first and other page headers.
- Set header and footers to use the same font size and scale as the worksheet
  and to align with sheet margins.

### Sheet Tab

- Identify a print area other than the entire worksheet.
- Identify print titles to repeat on each page.
- Print gridlines, row and column headings, comments, and error messages.
- Apply changes to multiple worksheets at the same time by selecting the
  worksheet tabs first to group the worksheets.

## Margins, Page Orientation, and Paper Size

- The Margins button [Page Layout tab, Page Setup group] lists Normal, Wide, and
  Narrow for margin settings as well as Custom Margins.
- To set margins:
    1. Click the Margins button.
    2. Select Custom Margins to open the Page Setup dialog box.
    3. Type each margin setting as desired or use the spinner buttons to change
       the margin value.
    4. Select the Horizontally box in the Center on page area to center the data
       between the left and right margins.
    5. You can also vertically center the data on the page between the top and
       bottom margins.
    6. Click the Page tab in the Page Setup dialog box.
    7. Select the radio button for Portrait or Landscape.
    8. Click the Paper size arrow and select a paper size.
    9. Click OK to close the dialog box.
- To format multiple worksheets:
    - Group the worksheet tabs by clicking the first tab name, holding Shift,
      and clicking the last tab name to select all sheets between those two
      names. OR, right-click any sheet tab and choose Select All Sheets to group
      all sheets in the workbook.

## Headers and Footers

- A header is information that prints at the top of each page. A footer is data
  that prints at the bottom of each page. Use a header or a footer to display
  identifying information such as a company name or logo, page numbers, the
  date, or the file name.
- Each header or footer has left, middle, and right sections. Information in the
  left section is left aligned, data in the middle section is centered, and
  material in the right section is right aligned.
- Predefined headers and footers include one, two, or three elements, separated
  by commas. One element prints in the center section, two elements print in the
  center and right sections, and three elements use all three sections.
- Headers and footers are not visible in Normal view, but they can be viewed in
  Page Layout view or Print Preview.
- To insert a custom header and footer using the Ribbon:
    1. Select the worksheet.
    2. Press Ctrl and click another tab name to apply the same header or footer
       to multiple sheets.
    3. Click the Insert tab and click the Header & Footer button [Text group].
    4. Click a header section.
    5. Click an element name in the Header & Footer Elements
       group [Header & Footer tab].
    6. A code displays with an ampersand (&) and the element name is enclosed in
       square brackets.
    7. Click a different header section and type your own label.
    8. Select and delete codes and text in a section.
    9. Click the Go To Footer button [Header & Footer tab, Navigation group].
    10. Click a footer section.
    11. Click an element name in the Header & Footer Elements
        group [Header & Footer tab] or type your own label.
    12. You can combine your own labels with predefined elements.
    13. Click any worksheet cell.
    14. The header and footer appear as they will print in Page Layout view.
    15. The Header & Footer tab no longer displays.
    16. Click the Normal button in the Status bar.
    17. The header and footer are not visible in Normal view.

### Preset Headers and Footers

To insert preset headers or footers on multiple sheets:

1. Select a worksheet tab.
2. Press Ctrl and click a second tab.
3. Two worksheets are grouped, and "Group" appears in the title bar after the
   file name.
4. Right-click any tab and click Select All Sheets to add a header or footer to
   all sheets at once.
5. Click the Page Setup launcher [Page Layout tab, Page Setup group].
6. Click the Header/Footer tab.
7. Click the arrow for the Header or the Footer text box.
8. Choose (none) to remove a header or footer, or choose a predefined header or
   footer.
9. Click Custom Header or Custom Footer to open the Header or Footer dialog box.
10. Click a section and then click a button to insert an element.
11. Point to a button to see its ScreenTip.
12. The text or code appears in the section.
13. Click OK to close the Header or Footer dialog box.
14. The custom header or footer displays in the preview area.
15. Click OK to close the Page Setup dialog box.
16. Click the File tab and select Print while the worksheets are grouped.
17. Click the Next Page button at the bottom of the print preview area in the
    Backstage
18. Both sheets display the same header or footer.
19. Return to the worksheet.
20. Right-click a sheet tab and choose Ungroup Sheets.

To remove headers and footers from multiple worksheets, select the sheet tabs
so that they are grouped, then:

1. In Page Layout view, click the header or footer section and delete labels
   and codes.
2. From the Page Setup dialog box, simply select (none) from the Header or
   Footer drop-down list.

## Page Breaks

- A page break is a printer code that starts a new page. When worksheet data
  spans more than one printed page, automatic page breaks are inserted based on
  the paper size, the margins, and scaling. Automatic page breaks adjust as you
  add or delete data.
- You can insert manual page breaks, too. Manual page breaks do not adjust as
  you edit worksheet data.
- In Page Break Preview:
    - An automatic page break displays as a dotted or dashed blue line.
    - A manual page break appears as a solid blue line.
- To insert a manual page break:
    1. Select the cell for the location of a page break.
    2. Select a cell in column A below the row where you want to start a new
       page.
    3. Select a cell in row 1 to the right of the column where you want to start
       a new page.
    4. Select a cell below and to the right of where new pages should start.
    5. Click the Page Layout tab and then click the Breaks
       button [Page Setup group].
    6. Select Insert Page Break.
    7. A manual page break displays as a solid blue line. Existing automatic
       page breaks may adjust.
- A manual page break displays as a thin, solid line in the worksheet in Normal
  view.

# Chapter Summary

- **1.1** Create, save, and open an Excel workbook (p. E1-3).
    - A new Excel workbook has one worksheet and is named BookN, with
      numbers assigned in order throughout a work session.
    - Save a workbook with a descriptive name that identifies the contents
      and purpose.
    - Create a new workbook from the Excel Start page or from the New
      command on the File tab.
    - A workbook opened or copied from an online source opens in **Protected
      View**.
    - An Excel workbook is an **.xlsx** file, but it can be saved in other
      formats for easy sharing of data.
- **1.2** Enter and edit data in a worksheet (p. E1-8).
    - Data is entered in a cell which is the intersection of a column and
      a row. Each cell has an address or reference.
    - Data is recognized as a label or a value. Labels are not used in
      calculations. A label is left-aligned in the cell; a value is 
      right-aligned.
    - A **formula** is a calculation that displays a value as its result in
      the worksheet.
    - Press **F2 (FN+F2)** or double-click a cell to start Edit mode. After
      making a change, press **Enter** to complete the edit.
    - Horizontal alignment options for cell data include Align Left,
      Center, and Align Right. Data can also be indented from the cell border.
      Vertical alignment choices are Top Align, Middle Align, and Bottom
      Align.
    - A group or selection of cells is a **range**. Use the pointer and
      keyboard shortcuts to select a range.
    - Use the **Fill Handle** to create a series that follows a pattern. The
      Fill Handle copies data when no pattern exists.
    - **AutoComplete** supplies a suggestion for a column entry that begins
      with the same character as a label already in the column.
    - **Pick From Drop-down List** displays a sorted list box with labels in
      the column when you right-click a cell.
    - Cut, copy, and paste commands include drag and drop, as well as
      regular Windows **Cut**, **Copy**, and **Paste** buttons on the Home tab,
      in context menus, and with keyboard shortcuts.
    - The **Paste Special** dialog box includes an option to paste data using
      an arithmetic operation.
- **1.3** Use functions to build a simple formula (p. E1-22).
    - A **function** is a built-in formula. Results appear in the worksheet
      cell, but the function displays in the Formula bar.
    - The **AutoSum** button [Home tab, Editing group] inserts the SUM
      function in a cell to add the values in a selected range.
    - Click the **AutoSum** button arrow to use _AVERAGE_, _MAX_, or _MIN_, 
      as well as Count Numbers.
    - Each function has **syntax** which is its required parts in the required
      order. A function begins with an equals sign (=), followed by the name
      of the function. After the name, the function **argument** displays in
      parentheses.
    - An argument is one or more cell references.
    - An argument range displays cell addresses separated by a colon
      (B1:B5) for adjacent cells or by commas for nonadjacent cells (B1, B4,
      B8).
    - When you copy a function, Excel adjusts cell references to match
      their locations in the worksheet.
    - Change a function argument by entering new cell references in the
      Formula bar or by resizing the cell range in the worksheet.
- **1.4** Format cell data with font attributes, borders, fill, cell
  styles, and themes (p. E1-27).
    - Apply font attributes from the Font group on the Ribbon or from the
      Format Cells dialog box.
    - Font attributes include the **font** name or face, the **font size**, font
      styles such as bold, italic, and underline, and the **font color**.
    - The **Format Painter** copies formatting from one cell to another.
    - Number formats include decimals, commas, currency symbols, and
      percent signs. They also determine how negative values appear in the
      worksheet.
    - Apply number formats from the Ribbon or the Format Cells dialog box.
    - Add borders or fill color to cells for easy identification,
      clarification, and emphasis.
    - The **Merge & Center** commands combine two or more cells into one and
      centers the data within that new cell.
    - Display labels at an angle using the **Orientation** command.
    - Use the **Center Across Selection** command to center multiple rows of
      data across the same range.
    - Wrap a label to multiple lines in a cell by pressing **Alt+Enter** to
      start a new line after the word where a new line should begin.
    - A **cell style** is a preset collection of font face, font size,
      alignment, color, borders, and fill color and is based on the theme.
    - A **theme** is a collection of fonts, colors, and special effects for a
      workbook.
    - Use themes and cell styles to quickly format a worksheet with
      consistent, professionally designed elements.
- **1.5** Modify columns, rows, and sheets in a workbook (p. E1-37).
    - The default column width and row height depend on the font size
      defined by the workbook theme.
    - Change the width of a column or the height of a row by dragging the
      border between the headings.
    - **AutoFit** a column or row by double-clicking the border between the
      headings.
    - The **Format** button [Home tab, Cells group] includes commands to
      change **row height** and **column width**.
    - Insert or delete a column or a row by right-clicking the column or
      row heading and choosing Insert or Delete.
    - The **Insert** and **Delete** buttons [Home tab, Cells group] include
      commands for inserting and deleting rows and columns.
    - The Insert and Delete dialog boxes have options to insert or delete
      cells by shifting data in a worksheet.
    - Hide one or multiple rows or columns; they are still available for
      use in a formula.
    - The number of worksheets in a workbook is limited only by the amount
      of computer memory.
    - Click the **New Sheet** button next to the worksheet tabs to insert a
      new sheet in a workbook.
    - The **Insert** and **Delete** buttons [Home tab, Cells group] include
      options to insert and delete worksheets.
    - Double-click a worksheet tab, type a name, and press **Enter** to rename
      a sheet.
    - Apply a **tab color** to distinguish a particular sheet tab.
    - Move a worksheet tab to another location in the list of tabs by
      dragging it to that location.
    - Create an exact duplicate of a worksheet in the same or another
      workbook with the **Move or Copy** command on the context menu.
- **1.6** Navigate in a workbook (p. E1-49).
    - Use basic keyboard directional arrows and the mouse to position the
      pointer.
    - Excel has specific navigation commands as well as common Windows
      movement commands.
    - Click the **Name** box, type a cell address or range, and press **Enter** 
      to move the active cell(s) to that location.
    - Use the **Go To** command to go to a single cell or a range of cells.
    - The **Find** command searches for labels or values in a workbook.
- **1.7** Modify screen appearance of a workbook by adjusting zoom size,
  changing views, and freezing panes (p. E1-52).
    - View a workbook in three ways: **Normal**, **Page Layout**, and 
      **Page Break Preview**.
    - Switch views using the Status bar buttons or from the Workbook Views
      group on the View tab.
    - Adjust the zoom size to determine how much data is visible at once.
      Zoom controls are available in the Status bar and from the View tab.
    - Use the **Freeze Panes** command to lock selected rows or columns on
      screen to scroll data while keeping important data in view.
    - Use the **Split** command to divide the screen into multiple sections to
      see different areas of a large worksheet at once.
    - Hide a worksheet from view and unintended editing.
    - When multiple workbooks are open, the **Switch Windows** command lists
      the open workbook names.
    - Switch among open workbooks using the Windows taskbar or the
      keyboard shortcut **Ctrl+F6**.
    - The **New Window** command displays the same workbook in a second
      window. Use this command with the **View Side by Side** command to see
      different areas or sheets of a workbook at the same time.
- **1.8** Manage page setup options, print settings, and document
  properties (p. E1-58).
    - Adjust most page settings from the Page Setup dialog box.
    - The Page Layout command tab includes the **Margins** and **Orientation**
      buttons for quick access to these settings.
    - Add **headers** and **footers** to print at the top and bottom of each 
      page.
    - Group worksheets and apply page setup choices to multiple sheets at
      the same time.
    - Excel inserts automatic page breaks when data require more than one
      printed page.
    - Insert manual **page breaks** to define where a new page starts and
      override automatic breaks.
    - Print **gridlines**, row and column headings, or **print titles** from
      commands on the Page Layout tab or in the Page Setup dialog box.
    - You can scale a worksheet to fit a printed size or define a **print
      area** other than the entire sheet.
    - From the **Print** command in the Backstage view, adjust margins and
      column settings, change the orientation, or print the entire workbook.
    - A **document property** is **metadata** stored with the file.
    - Several properties are supplied by Excel and cannot be edited. Other
      document properties are added or edited by the user.
    - Several properties are visible and can be edited in the Info area in
      the Backstage view.
    - Open the Properties dialog box to build or review all properties.