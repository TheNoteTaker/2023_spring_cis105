# Module 7 - Working with Formulas and Functions

<!-- TOC -->
* [Module 7 - Working with Formulas and Functions](#module-7---working-with-formulas-and-functions)
* [General Notes](#general-notes)
* [SLO 2.1 - Build and Edit Basic Formulas](#slo-21---build-and-edit-basic-formulas)
  * [Building and Editing Formulas](#building-and-editing-formulas)
    * [Type Arithmetic Operators](#type-arithmetic-operators)
    * [Type a Formula](#type-a-formula)
    * [Point and Click to Build a Formula](#point-and-click-to-build-a-formula)
    * [Edit a Formula](#edit-a-formula)
  * [Conclusion](#conclusion)
* [SLO 2.2 - Set Mathematical Order of Operations in a Formula](#slo-22---set-mathematical-order-of-operations-in-a-formula)
* [Excel - Chapter 2 - Working with Formulas and Functions](#excel---chapter-2---working-with-formulas-and-functions)
  * [Setting Mathematical Order of Operations](#setting-mathematical-order-of-operations)
    * [Mathematical Order of Precedence](#mathematical-order-of-precedence)
    * [Useful Tips and Tricks](#useful-tips-and-tricks)
    * [How to Use Multiple Operators in a Formula](#how-to-use-multiple-operators-in-a-formula)
  * [Conclusion](#conclusion-1)
* [SLO 2.3 - Use Absolute, Mixed, Relative, and 3D References in a Formula](#slo-23---use-absolute-mixed-relative-and-3d-references-in-a-formula)
  * [Table 2-3: Cell Reference Types](#table-2-3--cell-reference-types)
  * [Copy a Formula with a Relative Reference](#copy-a-formula-with-a-relative-reference)
    * [HOW TO:    Copy a Formula with a Relative Reference](#how-to--copy-a-formula-with-a-relative-reference)
  * [Build and Copy a Formula with an Absolute Reference](#build-and-copy-a-formula-with-an-absolute-reference)
    * [HOW TO:    Build and Copy a Formula with an Absolute Reference](#how-to--build-and-copy-a-formula-with-an-absolute-reference)
  * [Build and Copy a Formula with a Mixed Reference](#build-and-copy-a-formula-with-a-mixed-reference)
    * [HOW TO: Build and Copy a Formula with a Mixed Reference](#how-to--build-and-copy-a-formula-with-a-mixed-reference)
  * [Create a Formula with a 3D Reference](#create-a-formula-with-a-3d-reference)
    * [MORE INFO](#more-info)
    * [HOW TO:    Create a Formula with a 3D Reference](#how-to--create-a-formula-with-a-3d-reference)
  * [Range Names and Formula AutoComplete](#range-names-and-formula-autocomplete)
    * [Naming a Range](#naming-a-range)
    * [Using Range Names in a Formula](#using-range-names-in-a-formula)
    * [Defining and Scoping a Range Name](#defining-and-scoping-a-range-name)
    * [Navigation Using Range Names](#navigation-using-range-names)
      * [Additional Information](#additional-information)
* [SLO 2.4 - Use Formula Auditing Tools in a Worksheet](#slo-24---use-formula-auditing-tools-in-a-worksheet)
* [Excel - Chapter 2 - Working with Formulas and Functions](#excel---chapter-2---working-with-formulas-and-functions-1)
  * [Formula Auditing Tools](#formula-auditing-tools)
  * [Using Formula Auditing Tools](#using-formula-auditing-tools)
    * [Trace an Error](#trace-an-error)
    * [Formula Auditing Group](#formula-auditing-group)
    * [Formula Auditing Buttons](#formula-auditing-buttons)
    * [Trace Precedents and Dependents](#trace-precedents-and-dependents)
    * [How to Use Formula Auditing Tools](#how-to-use-formula-auditing-tools)
  * [Formula Correction Message Window](#formula-correction-message-window)
  * [Circular Reference](#circular-reference)
  * [Additional Information](#additional-information-1)
  * [Conclusion](#conclusion-2)
* [SLO 2.5 - Work With Statistical and Date & Time Functions](#slo-25---work-with-statistical-and-date--time-functions)
  * [Excel Functions](#excel-functions)
    * [Insert Function Dialog Box](#insert-function-dialog-box)
    * [Function Arguments Dialog Box](#function-arguments-dialog-box)
  * [How to Build a Function in the Function Arguments Dialog Box](#how-to-build-a-function-in-the-function-arguments-dialog-box)
  * [The MEDIAN and MODE Functions](#the-median-and-mode-functions)
  * [How to Use the MEDIAN Function](#how-to-use-the-median-function)
  * [MODE Functions](#mode-functions)
  * [How to Use the MODE.SNGL Function](#how-to-use-the-modesngl-function)
  * [Count Functions](#count-functions)
    * [MORE INFO](#more-info-1)
    * [HOW TO: Use the COUNT Function](#how-to--use-the-count-function)
    * [COUNTA and COUNTBLANK Functions](#counta-and-countblank-functions)
    * [HOW TO: Use the `COUNTA` and `COUNTBLANK` Functions](#how-to--use-the-counta-and-countblank-functions)
  * [AutoCalculate](#autocalculate)
  * [Calculation Options](#calculation-options)
    * [How to View AutoCalculate Results](#how-to-view-autocalculate-results)
  * [TODAY and NOW Functions](#today-and-now-functions)
    * [More Info](#more-info-2)
    * [How to Insert TODAY with Formula AutoComplete](#how-to-insert-today-with-formula-autocomplete)
* [SLO 2.6 - Use Functions From the Financial, Logical, and Lookup & Reference Categories](#slo-26---use-functions-from-the-financial-logical-and-lookup--reference-categories)
* [SLO 2.6: Using Financial, Logical, and Lookup Functions](#slo-26--using-financial-logical-and-lookup-functions)
  * [Financial Functions](#financial-functions)
    * [PMT Function](#pmt-function)
    * [Example](#example)
  * [IF Function](#if-function)
    * [How to Build an IF Function](#how-to-build-an-if-function)
    * [Lookup Functions](#lookup-functions)
    * [How to Use the XLOOKUP Function](#how-to-use-the-xlookup-function)
  * [VLOOKUP and HLOOKUP Functions](#vlookup-and-hlookup-functions)
    * [How to Display Results Using VLOOKUP](#how-to-display-results-using-vlookup)
    * [How to Use the HLOOKUP Function](#how-to-use-the-hlookup-function)
* [SLO 2.7 - Work With Text Functions](#slo-27---work-with-text-functions)
* [Excel - Chapter 2 - Working with Formulas and Functions](#excel---chapter-2---working-with-formulas-and-functions-2)
  * [Working with Text Functions](#working-with-text-functions)
  * [The TEXTJOIN Function](#the-textjoin-function)
    * [HOW TO:    Use the TEXTJOIN Function](#how-to--use-the-textjoin-function)
  * [The CONCAT Function](#the-concat-function)
    * [ANOTHER WAY](#another-way)
    * [HOW TO: Use the CONCAT Function](#how-to--use-the-concat-function)
    * [Additional Information](#additional-information-2)
  * [Conclusion](#conclusion-3)
* [SLO 2.8 - Build functions From The Math & Trig Category](#slo-28---build-functions-from-the-math--trig-category)
  * [ROUND Function](#round-function)
    * [HOW TO: Type the ROUND Function](#how-to--type-the-round-function)
  * [SUMIF Function](#sumif-function)
    * [HOW TO: Build a SUMIF Function](#how-to--build-a-sumif-function)
  * [SUMPRODUCT Function](#sumproduct-function)
    * [HOW TO: Build a SUMPRODUCT Function](#how-to--build-a-sumproduct-function)
  * [Conclusion](#conclusion-4)
* [Chapter Summary](#chapter-summary)
<!-- TOC -->

# General Notes

# SLO 2.1 - Build and Edit Basic Formulas

## Building and Editing Formulas

- An Excel formula is an expression or statement that uses arithmetic operations
  to perform a calculation.
- Basic arithmetic operations are addition, subtraction, multiplication, and
  division.
- The formula is entered in a cell and refers to other cells or ranges in the
  worksheet or workbook.
- A function is a built-in formula, which begins with an equals sign (=) similar
  to a formula.
- After the equals sign, you enter the address of the first cell, followed by
  the arithmetic operator, followed by the next cell in the calculation.
- Basic arithmetic operators are listed in Table 2-1.

### Type Arithmetic Operators

- Type arithmetic operators using the numeric keypad or the symbols on the
  number keys at the top of the keyboard.
- Table 2-1: Arithmetic Operators

| Character | Operation      |
|:---------:|:---------------|
|   **+**   | Addition       |
|   **-**   | Subtraction    |
|   __*__   | Multiplication |
|   **/**   | Division       |

### Type a Formula

- For a simple formula, click the cell, type the equals sign, type the first
  cell address, type the operator, type the next cell address, and so on.
- Press any completion key when the formula is finished.
- Formulas are not case-sensitive, so you can type cell addresses in lowercase
  letters.
- As you type a cell address after the equals sign, Formula AutoComplete
  displays a list of functions and range names that match the character that you
  typed.
- If you press Esc before completing a formula, it is canceled, and nothing is
  entered in the cell.

### Point and Click to Build a Formula

- It is often easier to point to and click the cell to enter its address in a
  formula.
- This method also guards against your typing the wrong address.
- You will not see a Formula AutoComplete list, but the Range Finder will
  highlight and color code the cells.

### Edit a Formula

- Edit a formula by double-clicking its cell or by clicking the Formula bar
  while the cell is active to start Edit mode.
- The Range Finder highlights and color codes the formula cell and the argument
  cells.
- Add, remove, or change cells and operators to build a new formula in the cell
  or in the Formula bar.

## Conclusion

In summary, formulas are expressions or statements that use arithmetic
operations to perform a calculation. Excel provides basic arithmetic operators,
such as addition, subtraction, multiplication, and division. You can type a
formula using the keyboard or point and click to build a formula. Additionally,
you can edit a formula by double-clicking its cell or by clicking the Formula
bar while the cell is active to start Edit mode.

# SLO 2.2 - Set Mathematical Order of Operations in a Formula

# Excel - Chapter 2 - Working with Formulas and Functions

## Setting Mathematical Order of Operations

- A formula can have multiple operators, such as both addition and
  multiplication.
- Excel follows mathematical order of operations, which is the sequence of
  arithmetic calculations.
- The order in which arithmetic in a formula is carried out depends on the
  operator as well as left-to-right order.
- The basic sequence is left to right, but multiplication and division are done
  before addition and subtraction.
- Use parentheses to control the order of operations.
- In a formula with multiple operators, references enclosed in parentheses are
  calculated first.
- Table 2-2 shows the mathematical order of precedence.

### Mathematical Order of Precedence

The table explains how arithmetic is carried out when a formula has multiple
operators.

| Operator |   Operation    | Order of Precedence |
|:--------:|:--------------:|:-------------------:|
|  **()**  |  Parentheses   |        First        |
|  **^**   |    Exponent    |       Second        |
|  __*__   | Multiplication |        Third        |
|  **/**   |    Division    |        Third        |
|  **+**   |    Addition    |       Fourth        |
|  **−**   |  Subtraction   |       Fourth        |

- If a value with an exponent exists, it is calculated next.
- When two operators have the same precedence, Excel calculates them from left
  to right.

### Useful Tips and Tricks

- Use the following acronym to recall the arithmetic order of operations: Please
  Excuse My Dear Aunt Sally.
- An exponent raises a number to a power. Use the caret symbol (^) to build a
  formula with an exponent.
- Because addition and subtraction are equal in priority, Excel calculates from
  left to right.
- Without parentheses, Excel calculates the multiplication first and then
  follows left to right order.

### How to Use Multiple Operators in a Formula

1. Type `=` to start the formula.
2. The calculation in parentheses can occur anywhere in the formula.
3. Type `(` (left parenthesis) to start the calculation that should have
   priority.
4. Click the first cell for the calculation to be enclosed in parentheses.
5. Type the arithmetic operator.
6. Click the next cell for the calculation to be enclosed in parentheses.
7. Type `)` (right parenthesis) to end the calculation with priority.
8. You must have matching parentheses.
9. Complete the formula.
10. Press Enter.

## Conclusion

- Understanding the order of operations in Excel formulas is crucial to getting
  accurate results.
- Use parentheses to control the order of operations in formulas with multiple
  operators.
- Use the acronym Please Excuse My Dear Aunt Sally to recall the order of
  operations.
- Remember to use the caret symbol (`^`) to build a formula with an exponent.

# SLO 2.3 - Use Absolute, Mixed, Relative, and 3D References in a Formula

- Using Absolute, Mixed, Relative, and 3D References
- Cell references in a formula are identified in different ways. How a cell
  reference is identified affects how the formula adjusts when copied.

## Table 2-3: Cell Reference Types

The table shows a cell reference, identifies its type, and explains what happens
when the formula is copied.

| Cell Reference | Reference Type  | Behavior When Copied                                    |
|:--------------:|:---------------:|:--------------------------------------------------------|
|      `B2`      |  **Relative**   | Cell address updates to new location(s).                |
|     `$B$2`     |  **Absolute**   | Cell address does not change.                           |
|     `$B2`      |    **Mixed**    | Column does not change; row updates to new location(s). |
|     `B$2`      |    **Mixed**    | Row does not change; column updates to new location(s). |
|  `Sheet2!B2`   | **3D Relative** | Cell address updates to new location(s).                |

## Copy a Formula with a Relative Reference

- When a formula with relative references is copied, the copied formulas use
  cell addresses that reflect their locations in the worksheet.

### HOW TO:    Copy a Formula with a Relative Reference

1. Select the cell with the formula.
2. Select multiple adjacent cells as needed.
3. Press Ctrl+C to copy the formula(s) or click the Copy
   button [Home tab, Clipboard group].
4. You can also click and drag the Fill pointer if the copies are adjacent to
   the original formula(s).
5. Select the destination cell or range.
6. Click the first cell in a destination range when pasting multiple cells.
7. When you use the Fill pointer, the cells are copied as soon as you release
   the pointer.
8. Press Ctrl+V to paste the formula.
9. The pasted formula reflects its location on the worksheet.
10. Ignore the Paste Options button.
11. Press Esc to cancel the Paste command if a moving border displays.

## Build and Copy a Formula with an Absolute Reference

- An absolute reference in a formula maintains that address when copied.
- To change a cell address to an absolute reference, press the F4 (FN+F4)
  function key while keying or editing the formula.

### HOW TO:    Build and Copy a Formula with an Absolute Reference

1. Click the cell for the formula.
2. Type an equals sign (=) to begin the formula.
3. Click the cell to be referenced in the formula.
4. Press F4 (FN+F4).
5. The reference becomes absolute (Figure 2-8).
6. Complete the formula and press Enter.
7. Select the cell with the formula.
8. Press Ctrl+C to copy the formula(s) or drag the Fill pointer if copies are
   adjacent to the original formula.
9. Use the Copy button on the Home tab if you prefer.
10. Select the destination cell and press Ctrl+V to paste the formula.
11. The pasted formula uses the same cell reference as the original.
12. Use the Paste button on the Home tab if you prefer.
13. Ignore the Paste Options button.

## Build and Copy a Formula with a Mixed Reference

- A mixed cell reference has one relative and one absolute reference in the cell
  address. Part of the copied formula is unchanged, and part is updated. A mixed
  cell reference has one dollar sign.

### HOW TO: Build and Copy a Formula with a Mixed Reference

1. Click the cell for the formula.
2. Type an equals sign (**=**) to begin the formula.
3. Click the cell to be referenced in the formula.
4. Press **F4 (FN+F4)**.
5. The reference is absolute.
6. Press **F4 (FN+F4)**
7. Complete the formula and press Enter.
8. Select the formula cell.
9. Press Ctrl+C to copy the formula(s).
10. You can use the Copy button on the Home tab.
11. Drag the Fill pointer if copies are adjacent to the original formula.
12. Alternatively, select the destination cell or range.
13. Press Ctrl+V to paste the formula.
    * The pasted formula does not change the absolute part of the reference but
      updates the relative part.
    * Use the Paste button on the Home tab if you prefer.

## Create a Formula with a 3D Reference

An important feature of an Excel workbook is its ability to use data from any
sheet in a formula. This type of formula is called a 3D reference because it
uses more than one sheet or surface to calculate a result. A 3D reference
includes the name of the worksheet and can be absolute, mixed, or relative.

Table 2-4 illustrates sample 3D references.

| Cell on Sheet1 | Formula on Sheet1       | Results                                                                      |
|----------------|-------------------------|------------------------------------------------------------------------------|
| B4             | =B3+Sheet2!B3+Sheet3!B3 | Adds the values in cell B3 on Sheet1, Sheet2, and Sheet3.                    |
| B7             | =SUM(B4:B6)+Sheet2!A12  | Adds the values in cells B4:B6 on Sheet1 to the value in cell A12 on Sheet2. |
| C8             | =C7*Sheet2!$D$4         | Multiplies the value in cell C7 by the value in cell D4 on Sheet2.           |
| D10            | =Sheet2!$B$2/Sheet3!B2  | Divides the value in cell B2 on Sheet2 by the value in cell B2 on Sheet3.    |

### MORE INFO

An external reference formula refers to a cell in another workbook. It can be
absolute, mixed, or relative. An external reference includes the name of the
workbook in square brackets as in [OurBook] SheetName!B$2.

### HOW TO:    Create a Formula with a 3D Reference

1. Click the cell for the formula and type =.
2. Click the first cell for the formula.
    - The 3D reference can occur anywhere in the formula.
3. Type the operator.
4. Select the sheet tab with the cell to be used in the formula.
5. Click the cell required in the formula.
6. Press F4 (FN+F4) to make the reference absolute or mixed as needed (Figure
   2-10).
7. Type the next operator and complete the formula.
    - The next cell can be on the same or another sheet.
8. Press Enter or any completion key.

## Range Names and Formula AutoComplete

A range name is a label assigned to a single cell or a group of cells that can
be used instead of cell references in a formula. This can be helpful when you
want to use descriptive names to help interpret formulas. In addition, named
cell ranges appear in the Formula AutoComplete list for ease in building
formulas.

### Naming a Range

To name a range, you can select the cell or range to be named and click on the
Name Box to the left of the Formula bar. Then, you can type the name and press
Enter. You can also use the Defined Names group on the Formulas tab to create,
edit, and delete range names. To assign a name to existing formulas that refer
to the cells identified in the range name, you can use the Apply Names command.

When naming a cell range, it is important to follow these basic rules:

- Use a descriptive name.
- Start with a letter or underscore.
- Do not include spaces.
- Avoid using cell references like "A1" in the name.

You can use row and column labels to create range names using the Create from
Selection button on the Formulas tab.

### Using Range Names in a Formula

To use a named range in a formula, type the first character of the name to
display the Formula AutoComplete list. Select the range name in the list and
press Tab or double-click the name to enter it in the formula. Range names are
automatically substituted in formulas when you select the named cell range in
the worksheet while building a formula.

### Defining and Scoping a Range Name

To define and scope a range name, select the cell range to be named and click
the Define Name button in the Defined Names group. Type the range name in the
Name Box and choose the name of a worksheet as desired from the Scope drop-down
list. You can use a range name multiple times in the workbook by scoping names
to a worksheet.

### Navigation Using Range Names

In large worksheets, using named ranges quickly moves the insertion point to
where you want to work. You can use the Name Box and the Go To command to
navigate to a particular range.

#### Additional Information

- Long range names only partially display in the Name Box list.
- The Use in Formula button in the Defined Names group displays a list of range
  names for insertion in a formula.
- The Paste Name dialog box displays a list of current range names with cell
  references as documentation.

Overall, using range names can help to make formulas easier to read and
interpret, and can simplify the process of navigating through large worksheets.

---

**Key Takeaways:**

- Range names are labels assigned to cells or groups of cells that can be used
  instead of cell references in a formula.
- Naming a cell range requires using a descriptive name that follows specific
  rules.
- Using named ranges can simplify the process of navigating through large
  worksheets.
- Defining and scoping a range name can make it possible to use a range name
  multiple times in a workbook.

# SLO 2.4 - Use Formula Auditing Tools in a Worksheet

# Excel - Chapter 2 - Working with Formulas and Functions

## Formula Auditing Tools

- Formula auditing is the process of reviewing formulas for accuracy in Excel.
- Excel automatically audits formulas as you enter them and when you open a
  workbook based on its own error checking rules.
- Excel recognizes specific types of errors and flags them. For example, when
  you set a SUM range, adjacent cells not included in the argument are usually
  flagged as a potential error.
- When Excel finds an error in a formula, it displays a small green error
  triangle in the top-left corner of the cell.
- The Trace Error button has a drop-down list with options for dealing with the
  error.

## Using Formula Auditing Tools

### Trace an Error

1. Click the cell with the triangle error.
2. Point to the Trace Error button to display a ScreenTip.
3. Click the Trace Error drop-down list.
4. Choose an option for the error.

### Formula Auditing Group

- The Formula Auditing group on the Formulas tab includes tools to check
  formulas for logic, consistency, and accuracy.
- These commands enable you to examine what contributes to a formula and to
  analyze identified or potential errors.
- The tools are helpful, but they do require that you correct the problem.

### Formula Auditing Buttons

The table displays the button names and explains what each button does.

| Button           | Description                                                                                                 |
|------------------|-------------------------------------------------------------------------------------------------------------|
| Trace Precedents | Displays lines with arrows to identify all cells referenced in the formula in the active cell               |
| Trace Dependents | Displays lines with arrows to all cells that use the active cell directly or indirectly in a formula        |
| Remove Arrows    | Removes all lines and arrows from the Trace Precedents or Trace Dependents buttons                          |
| Show Formulas    | Displays formulas in the cells                                                                              |
| Error Checking   | Checks data against the error rules in Excel Options                                                        |
| Evaluate Formula | Steps through each part of a formula and displays an outcome for each part so that an error can be isolated |
| Watch Window     | Opens a floating window that displays selected cells and values for monitoring                              |

### Trace Precedents and Dependents

- A precedent is a cell that contributes to the formula results. Excel displays
  lines from the formula cell to each precedent cell for an easy way to audit
  your formula.
- A dependent is a cell that is affected by the active cell. Excel displays
  lines from the active cell to each cell that depends on the value in that
  cell.
- If a precedent or dependent cell is located on a different worksheet, a black
  line and arrow point to a small worksheet icon, but this does not identify the
  error cell address.

### How to Use Formula Auditing Tools

1. Click the Show Formulas button [Formulas tab, Formula Auditing group].
2. Each formula displays in its cell.
3. Show or hide formulas by pressing Ctrl+` (the grave accent next to the 1 key
   at the top of the keyboard).
4. Select the formula cell.
5. Only a formula cell has a precedent.
6. Click the Trace Precedents button [Formulas tab, Formula Auditing group].
7. Lines and arrows identify cells that contribute to the formula results.
8. You can trace precedents and dependents with formulas shown or hidden.
9. Click a cell that may be referenced in a formula.
10. Trace dependents for cells with or without formulas.
11. Click the Trace Dependents button [Formulas tab, Formula Auditing group].
12. Lines and arrows trace to cells that are affected by the active cell.
13. Trace both precedents and dependents as needed.
14. Click the Remove Arrows button [Formulas tab, Formula Auditing group].
15. Clicking the button removes all arrows.
16. Click the Remove Arrows drop-down list to specify the type of arrow to
    remove.
17. Click the Show Formulas button [Formulas tab, Formula Auditing group].
18. Formulas are hidden.
19. Show or hide formulas by pressing Ctrl+`.

## Formula Correction Message Window

- Some errors are noted before you complete the entry.
- When you press Enter, Excel displays a message box with a suggested
  correction.
- For this type of error, review the suggestion and select Yes if it is valid.
- If you click No, an alert box paraphrases the error.
- On occasion, Excel may find an error to which it cannot offer a suggested
  solution, and you need to solve the problem on your own.

## Circular Reference

- A circular reference error occurs when a formula includes the cell address of
  the formula.
- For example, if the formula in cell B10 is =B5+B10, the reference is circular.
- When you try to complete such a formula, a message box opens, but Excel does
  not correct the error or prevent you from keeping it in the worksheet with
  inaccurate results.
- The Status bar displays the location of circular references when they exist.

## Additional Information

- Use the Go to Special dialog
  box [Find & Select button, Home tab, Editing group] to highlight cells with
  formula errors, blank cells, and cells with precedents and dependents. This
  command helps identify potential problems in a workbook.

## Conclusion

- Formula auditing tools are essential for identifying and correcting errors in
  Excel.
- These tools can help you check formulas for logic, consistency, and accuracy,
  trace precedents and dependents, and remove errors in a formula.
- Additionally, Excel's error checking rules and correction message windows help
  you ensure that your formulas are accurate.
- By using these tools, you can save time and ensure that your data is accurate
  and reliable.

# SLO 2.5 - Work With Statistical and Date & Time Functions

## Excel Functions

- Excel functions are arranged into categories on the Formulas command tab in
  the Function Library group.
- The group includes the AutoSum button and an Insert Function button which
  opens the Insert Function dialog box.
- The Insert Function dialog box allows you to search for a function by name,
  category, or a brief description of what you want to do.
- Use a formula name for a specific function, otherwise it is just a
  mathematical formula and not a function.

### Insert Function Dialog Box

- You can insert a function from any category using the Insert Function dialog
  box.
- Display the dialog box by clicking the Insert Function button on the Formulas
  tab or the Insert Function button to the left edge of the Formula bar.
- Choose a function by clicking on a category, selecting the function name and
  clicking OK.

### Function Arguments Dialog Box

- The Function Arguments dialog box opens when you choose a function name from a
  category on the Formulas tab or when you use the Insert Function dialog box.
- It shows each argument with an entry box and an explanation.
- You can select cells or type directly in the entry box.
- Shift+F3 opens the Insert Function dialog box.

## How to Build a Function in the Function Arguments Dialog Box

1. Click the cell for the function.
2. Click the Insert Function button on the left edge of the Formula bar.
3. Choose the function name and click OK.
4. The Function Arguments dialog box opens for the selected function.
5. Move the dialog box as needed to see your data.
6. Click the first argument entry box.
    - Required argument names are bold.
    - Optional argument names are not bold.
7. Click a cell or drag to select a range.
    - The cell or range address displays in the entry box.
    - The dialog box collapses when you select cells.
8. Click the second argument entry box.
    - Some functions require more than one argument.
9. Click a cell or drag to select a range for the entry box.
10. Click OK when the arguments are complete.

## The MEDIAN and MODE Functions

- The Statistical function category includes calculations that determine
  averages, counts, percentiles, and ranks.
- Central tendency measures the average, the middle, or the most common.
- The `MEDIAN` and the `MODE` functions determine two central tendency
  statistics
  for a range of values.
- The `MEDIAN` function calculates the middle value in a range of cells.
- The `MEDIAN` function ignores labels and empty cells.
- The `MEDIAN` function has one argument, numberN.
- A number argument can be a range such as `D5:D19`.
- You can look for the median among up to 255 numbers (or ranges).
- The proper syntax for a `MEDIAN` function is: `=MEDIAN(
  number1, [number2], [number3], …)`
- The number1 argument is required which means you must have at least one value
  or cell range.
- Multiple arguments are separated by commas.
- Table 2-6 illustrates how multiple number arguments can be used in statistical
  functions.

|     MEDIAN Functions      | NumberN Arguments Used              | Description                      |
|:-------------------------:|:------------------------------------|:---------------------------------|
|     `MEDIAN(D5:D19`)      | Number1                             | One cell range                   |
| `MEDIAN(D5:D19, E5:E19)`  | Number1, Number2                    | Two cell ranges                  |
| `MEDIAN(D5:D19, E5, F12)` | Number1, Number2, Number3           | One range and two separate cells |
| `MEDIAN(D5, D7, D9, D11`) | Number1, Number2, Number3, Number 4 | Four nonadjacent cells           |

## How to Use the MEDIAN Function

1. Click the cell for the function.
2. Click the More Functions button [Formulas tab, Function Library group] and
   point to Statistical.
3. Scroll the list to find and click on MEDIAN.
4. The Function Arguments dialog box opens with an assumed number1 argument
   range.
5. Select the cell range as needed.

## MODE Functions

- A MODE function calculates the one value that appears most often in a range,
  or you can choose to find the two or three most repeated values.
- The two MODE functions are MODE.SNGL and MODE.MULT.
- Both functions have one argument, numberN, but you can look for the most
  common value in up to 255 numbers (or ranges).
- The proper syntax for a MODE function is:
    - `=MODE.SNGL(number1, [number2], [number3], …)`
    - `=MODE.MULT(number1, [number2], [number3], …)`

## How to Use the MODE.SNGL Function

1. Click the cell for the function.
2. Click the Insert Function button in the Formula bar.
3. Type mode in the Search for a function entry box.
4. The Insert Function dialog box will open, listing function names that match
   the search string.
5. Select `MODE.SNGL` from the list and click OK.
6. The Function Arguments dialog box will open with an assumed number1 argument
   range.
7. Select the cell range as needed.
8. You can also type cell addresses rather than selecting cells.
9. When the selected cells have been named, the range name is substituted in the
   formula.
10. Press Enter or click OK to return the result.

It is worth noting that the `MODE.SNGL` function returns the most frequently
occurring value in a set of values. If there are multiple modes with the
same frequency, only the first one is returned. If there are no modes, the
function will return the #N/A error value.

## Count Functions

Count functions tally the number of items in a range. The COUNT function has one
argument, valueN, and can include up to 255 cells or ranges. The proper syntax
for a COUNT function is:

```
=COUNT(value1, [value2], [value3], …)
```

- COUNT (Count Numbers): Counts the cells that contain values; ignores labels.
  Example syntax: `=COUNT(A1:A15)`.
- COUNTA: Counts the cells that contain any data type. Example
  syntax: `=COUNTA(A1:A15)`.
- COUNTBLANK: Counts empty cells. Example syntax: `=COUNTBLANK(A1:A15)`.
- COUNTIF: Counts the cells that meet the criteria argument. Example
  syntax: `=COUNTIF(A1:A15, “Services”)`.
- COUNTIFS: Counts the cells in one or more criteria ranges that meet respective
  criteria arguments. Example syntax: `=COUNTIFS(B2:B5,“=A”,C2:C5,“=21”)`.

### MORE INFO

The Database function category includes DCOUNT and DCOUNTA functions. This
category is available from the Insert Function dialog box.

### HOW TO: Use the COUNT Function

1. Click the cell for the function.
2. Click the arrow beside the AutoSum
   button [Formulas tab, Function Library group].
3. Select Count Numbers.
4. The COUNT function is inserted with an assumed range.
5. The function appears in the cell and in the Formula bar (Figure 2-33).
6. Select the cell range.
7. Type cell addresses rather than selecting cells.
8. Press Enter or any completion key.

If you use COUNT with a range of labels, the result is zero (0).

### COUNTA and COUNTBLANK Functions

From the Statistical group, choose `COUNTA` to count cells with any type of
data.
It counts nonblank cells. `COUNTA` has one argument, ValueN, similar to the
NumberN argument. You can count cells across up to 255 cells or ranges. Here is
the proper syntax for the `COUNTA` function:

```
=COUNTA(value1, [value2], [value3], …)
```

The `COUNTBLANK` function counts cells that are empty in a single range of
adjacent cells. Cells with values, labels, dates, or formulas are not counted.
In many rows of data, you might use COUNTBLANK to determine how many cells are
missing data. Here is the proper syntax for the COUNTBLANK function:

```
=COUNTBLANK(range)
```

### HOW TO: Use the `COUNTA` and `COUNTBLANK` Functions

1. Click the cell for the function.
2. Click the More Functions button [Formulas tab, Function Library group] and
   point to Statistical.
3. Scroll the list and click `COUNTA`.
4. The `COUNTA` Function Arguments dialog box opens.
5. Verify or click the Value1 entry box.
6. Click the first cell to be evaluated.
7. Select the first range as needed.
8. Click the Value2 entry box.
9. Select or type each cell address or range to be counted.
10. Click OK when the ValueN arguments are complete (Figure 2-34).
11. Click the cell for the function.
12. Type =cou to display the Formula AutoComplete list.
13. Double-click `COUNTBLANK` in the list to insert the function.
14. You can also click the function name once and press Tab to insert it.
15. Select the cells for range argument (Figure 2-35).
16. Press Enter.
17. The closing right parenthesis is automatically supplied.

## AutoCalculate

- AutoCalculate enables you to see certain statistical results for selected
  cells.
- The AutoCalculate area is located on the right side of the Status bar.
- The calculations available for automatic display are Average, Count, Numerical
  Count, Maximum, Minimum, and Sum.
- Review or change what you see for your data by right-clicking the Status bar
  to open the Customize Status Bar menu.

## Calculation Options

- Automatic calculation means that formulas are computed and recalculated after
  each entry or edit.
- For extremely large worksheets with millions of calculations, you might see
  lags in speed as the whole sheet is recalculated for every change you make.
- You can set manual calculation to delay recalculation until your work is
  complete.
- The Calculation group on the Formulas tab includes command buttons that enable
  you to fine-tune automatic calculation as needed.

### How to View AutoCalculate Results

1. Select the cell range or individual cells in the worksheet.
2. View the results on the Status bar.
3. Selected calculation results appear on the Status bar when appropriate; no
   SUM displays when you select labels.
4. Right-click the right side of the Status bar to open the Customize Status Bar
   menu.
5. If a function name has a check mark to its left, it is active.
6. Select a calculation name to activate it and show it on the Status bar.
7. Select a calculation name with a check mark to hide it from the Status bar.

## TODAY and NOW Functions

- The Date & Time function category includes calculations for date and time
  arithmetic, for converting dates and times to values, and for controlling how
  dates and times display.
- Excel treats a date as a serial number, a unique value assigned to each date.
- The TODAY function inserts the current date in the cell and updates each time
  the workbook opens. It uses the computer’s clock, and the syntax is =TODAY().
- The NOW function uses the syntax =NOW() and has no arguments.
- These two functions are volatile which means that the result depends on the
  current date, time, and computer.
- You can format the results of either of these functions to display both the
  date and the time.
- The Format Cells dialog box has a Date category and a Time category with many
  preset styles.
- In addition, two date formats are available from the Number Format drop-down
  list [Home tab, Number group]; they are Short Date and Long Date.
- The Number tab in the Format Cells dialog box includes a Locale setting for
  selecting the location used for date and time formats.

### More Info

- Display the Excel Options dialog box [File tab] and select the Advanced pane
  to set a 1904 date system which matches the date system in the Macintosh
  operating system.
- Select TODAY or NOW from the Date & Time function category. Because they are
  widely used and have no arguments, you may prefer to type either with the help
  of Formula AutoComplete.

### How to Insert TODAY with Formula AutoComplete

If you want to quickly insert the current date into a cell, you can use the
Formula AutoComplete feature. Here are the steps to insert TODAY using Formula
AutoComplete:

1. Click on the cell where you want to insert the date.
2. Type the equals sign (`=`) followed by the letters "to".
3. A list of functions that start with the letters "to" will appear. Scroll down
   until you find "TODAY".
4. Double-click on "TODAY" or select it and press "Tab" to insert the function
   name and open the parentheses.
5. Press "Enter" to insert the function and the current date will appear in the
   cell.

It's important to note that the TODAY function is volatile, which means that the
result of the function depends on the current date, time, and computer. If you
want to update the date every time the worksheet is opened, you can use the
TODAY function without any arguments.

You can also format the date and time by using the Format Cells dialog box. This
dialog box allows you to choose from a variety of date and time formats, as well
as set the location used for date and time formats.

# SLO 2.6 - Use Functions From the Financial, Logical, and Lookup & Reference Categories

# SLO 2.6: Using Financial, Logical, and Lookup Functions

## Financial Functions

- Functions in the Financial category calculate loan payments, interest earned,
  return on investment, etc.
- PMT function: calculates a constant payment amount over a specified period of
  time at a stated interest rate.
- PMT function has five arguments: rate, nper, pv, fv (optional), type
  (optional).
- Proper syntax for PMT formula: `=PMT(rate, nper, pv, [fv], [type])`

### PMT Function

- Calculates monthly payment for a selected payback time.
- Three required arguments: **rate**, **nper**, **pv**.
    - **Rate:** interest rate divided by 12 to determine a monthly payment.
    - **Nper:** number of years times 12, to determine the total number of
      payments.
    - **Pv:** amount of money borrowed.
    - **Fv** and type are optional.
    - **Type** argument indicates when the payment occurs.

### Example

- PMT function in cell `B7`: `=PMT(B6/12,B5*12,B4,,1)`
- Arguments separated by commas.
- `B6/12`: interest rate divided by 12 to determine monthly payment.
- `B5*12`: number of years times 12, to determine total number of payments.
- `B4`: amount of money borrowed.
- Type argument set to 1 to indicate payment occurs at beginning of period.

## IF Function

- Returns custom results based on the condition or statement tested.
- Syntax: =IF(logical_test, value_if_true, value_if_false)
- Three arguments: logical_test, value_if_true, and value_if_false.
- Use comparison or logical operators in the arguments.
- Comparison operators include =, <>, >, >=, <, and <=.

### How to Build an IF Function

1. Click the cell for the function.
2. Click the Logical button [Formulas tab, Function Library group].
3. Select IF in the list.
4. The Function Arguments dialog box for IF opens.
5. Click the Logical_test box.
6. Click the cell that will be evaluated.
7. Type a comparison operator immediately after the cell address.
8. Type a value or click a cell to complete the expression to be tested.
9. Click the Value_if_true box.
10. Type the label or click a cell that contains data that should display if the
    logical_test is true.
11. Click the Value_if_false box.
12. Type the label or click a cell that contains data that should display if the
    logical_test is false.

### Lookup Functions

- A LOOKUP function displays a piece of data from an existing list in the
  workbook.
- XLOOKUP is the latest addition to the LOOKUP family of functions and is
  available in current subscription versions of Excel.
- Syntax: =XLOOKUP(lookup_value, lookup_array,
  return_array, [if_not_found], [match_mode], [search_mode])
- Three required and three optional arguments.
- Match_mode and Search_mode arguments control how the search is performed.
- A binary search uses an algorithm to speed up the process.

|    Argument     | How Search and Results                                                     |
|:---------------:|:---------------------------------------------------------------------------|
| **Match_Mode**  |                                                                            |
|        0        | Exact match. If not found, display #N/A. This is the default.              |
|       -1        | Exact match. If not found, display the next smaller result.                |
|        1        | Exact match. If not found, display the next larger result.                 |
|        2        | Wildcard match using *, ?, ~                                               |
| **Search_Mode** |                                                                            |
|        1        | Start searching at first item. This is the default.                        |
|       -1        | Start searching at last item.                                              |
|        2        | Binary search that requires look_up array to be sorted in ascending order  |
|       -2        | Binary search that requires look_up array to be sorted in descending order |

### How to Use the XLOOKUP Function

1. Click the cell for the function.
2. Click the Lookup & Reference button [Formulas tab, Function Library group]
   and select XLOOKUP.
3. The Function Arguments dialog box opens.
4. Click the Lookup_value box.
5. Click the cell to be matched or located.
6. Click the Lookup_array box.
7. Select the cell range with the lookup data.
8. Click the Return_array box.
9. Select the cell range that contains the data to be displayed.
10. Three arguments are required.

## VLOOKUP and HLOOKUP Functions

- All Excel versions include VLOOKUP (vertical) and HLOOKUP (horizontal).
- VLOOKUP requires a lookup range that is organized in columns, and HLOOKUP uses
  a range that is arranged in rows.
- VLOOKUP function syntax: =VLOOKUP(lookup_value, table_array,
  col_index_num, [range_lookup]).
- Three required and one optional argument.
- Use false in the Range_lookup argument box to find an exact match.
- HLOOKUP function syntax: =HLOOKUP(lookup_value, table_array,
  row_index_num, [range_lookup]).
- The HLOOKUP function has a row_index_num argument instead of the col_index_num
  argument.

### How to Display Results Using VLOOKUP

1. Click the cell for the function.
2. Click the Lookup & Reference button [Formulas tab, Function Library group]
   and select VLOOKUP.
3. The Function Arguments dialog box opens.
4. Click the Lookup_value box.
5. Click the cell to be matched.
6. Click the Table_array box.
7. Select the cell range with the lookup data.
8. Press F4 (FN+F4) to make the reference absolute if cell addresses are used.
9. Click the Col_index_num box.
10. Type the column number, counting from the left in the table array, that
    contains the data to be displayed for the result.
11. Type false in the Range_lookup argument box to find an exact match.
12. Click OK.

### How to Use the HLOOKUP Function

1. Click the cell for the function.
2. Click the Lookup & Reference button [Formulas tab, Function Library group]
   and select HLOOKUP.
3. The Function Arguments dialog box opens.
4. Click the Lookup_value box.
5. Click the cell to be matched.
6. Click the Table_array box.
7. Select the cell range with the lookup data.
8. Press F4 (FN+F4) to make the reference absolute when cell addresses are used.
9. Click the Row_index_num box.
10. Type the row number, counting the first row of the table_array as 1.
11. Leave the Range_lookup argument box empty if the first row of data in the
    table array is arranged alphabetically from left to right.
12. Type false for the argument if you are not sure the data is sorted.
13. Click OK.

# SLO 2.7 - Work With Text Functions

# Excel - Chapter 2 - Working with Formulas and Functions

## Working with Text Functions

- Text functions split, join, or convert labels and values.
- With a Text function, you can format a value as a label to control alignment,
  display labels in uppercase or lowercase letters, and more.

## The TEXTJOIN Function

- The TEXTJOIN function combines strings of text, values, or characters.
- A common use of TEXTJOIN might be to display city and state in one cell or to
  display first and last name in a single cell.
- The TEXTJOIN function has three required arguments and one optional argument.
  The syntax for a TEXTJOIN function is:
  =TEXTJOIN(delimiter, ignore_empty, text1, [text2], …)
- The delimiter argument is the character used to separate the textN arguments.
  This character is inserted after each textN argument except the last one. You
  can have up to 252 textN arguments but only one delimiter. A text argument is
  a cell reference or data that you type.
- For most TEXTJOIN formulas, the delimiter is a space, but the delimiter can
  include multiple characters. For example, if you want to display “City, State”
  in a single cell, the delimiter is a comma and space combination.
- TEXTJOIN results are often placed in a separate column or row, and the
  original data can be hidden when it is located on the same sheet. The source
  data must be accessible so that the function works.

### HOW TO:    Use the TEXTJOIN Function

1. Click the cell for the result.
2. Click the Text button [Formulas tab, Function Library group] and choose
   TEXTJOIN.
3. Click the Delimiter box and press the Spacebar.
4. A space character displays with quotation marks in the dialog box and the
   Formula bar.
5. Type other combinations of characters as needed.
6. Click the Ignore_empty box.
7. Leave this box empty for the default TRUE option.
8. A blank cell is ignored in the result so that a blank is not joined to or
   included with result data.
9. Click the Text1 argument box.
10. Select the first cell with data to be joined.
11. Alternatively, type a text string, a value, or a punctuation character.
12. Click the Text2 box and select the cell or type data to be joined (Figure
    2-58).
13. You can use a combination of cell references and typed data among the TextN
    arguments.
14. Complete TextN argument boxes as needed.
15. The same delimiter is used after each textN argument.
16. Click OK.

## The CONCAT Function

- Excel has another Text function CONCAT that joins or combines data strings.
- The CONCAT function does not have a delimiter argument which enables you to
  determine how each two textN arguments are joined. You might use a comma and a
  space between the first two arguments but a period and a space between the
  next two.
- The CONCAT function was originally named CONCATENATE which means to link or
  join.
- The syntax for a CONCAT function is:
  =CONCAT(text1, [text2], ...)
- The CONCAT function has one required argument textN, where N is a number from
  1 to 255. The text argument can be cell references or any character strings
  that you type.

### ANOTHER WAY

- Join data in a formula by using the ampersand (&) as an operator. The formula
  =A1&B1 returns the same result as =CONCAT(A1,B1).
- When you concatenate text strings, you need to plan for spacing between
  characters to separate words or values. Like TEXTJOIN results, CONCAT results
  are usually displayed in a separate column or row to preserve the source data.

### HOW TO: Use the CONCAT Function

1. Click the cell for the result.
2. Click the Text button [Formulas tab, Function Library group] and choose
   CONCAT.
3. Click the Text1 argument box.
4. Select the first cell with data as the Text1 argument.
5. You can type a text string rather than selecting a cell.
6. Click the Text2 box and enter the next string to be concatenated.
7. Type the character(s) that should follow the first argument.
8. Press the Spacebar or type punctuation as needed.
9. Click the Text3 box and select the cell or type the next data (Figure 2-59).
10. You must include an argument for each occurrence of a space or punctuation
    character in your desired result.
11. Complete TextN argument boxes as needed.
12. Click OK.

### Additional Information

- The TEXTJOIN and CONCAT functions are useful for combining data strings.
- The TEXTJOIN function allows for the insertion of a delimiter between each
  textN argument.
- The CONCAT function does not include a delimiter argument and instead allows
  the user to concatenate two or more text strings together.
- Both functions can be used in a formula by using the ampersand (&) as an
  operator.
- When using either function, spacing between characters should be planned for
  to separate words or values.
- It is common practice to display the results of either function in a separate
  column or row to preserve the source data.

## Conclusion

- Excel's TEXTJOIN and CONCAT functions are powerful tools for combining data
  strings.
- TEXTJOIN allows for the insertion of a delimiter between each textN argument
  and CONCAT allows for the concatenation of two or more text strings.
- Both functions can be used in a formula by using the ampersand (&) as an
  operator, and spacing between characters should be planned for to separate
  words or values.
- When using either function, it is recommended to display the results in a
  separate column or row to preserve the source data.

# SLO 2.8 - Build functions From The Math & Trig Category

- Math & Trig function category includes SUM function and others for:
    - Absolute value of data
    - Cosine of an angle
    - Power of a number
- `ROUND`, `SUMIF`, and `SUMPRODUCT` are discussed in this section

## ROUND Function

- Rounding adjusts a value to display a specified number of decimal places or a
  whole number
- ROUND function has two required arguments: number and num_digits
- When a value used in rounding is 5 or greater, results round up to the next
  digit
- Increase Decimal and Decrease Decimal buttons follow arithmetic rounding
  principles for displaying values in the cell, as do the Accounting Number and
  Currency formats
- Excel has `ROUNDDOWN` function that always rounds a number down or toward
  zero. The `ROUNDUP` function rounds a value up, away from zero.

### HOW TO: Type the ROUND Function

1. Click the cell for the function
2. Type `=rou` and press Tab
3. Select the cell to be rounded
4. Type a value as the number argument
5. Type , (a comma) to separate the arguments
6. Type a number to set the number of decimal places
7. Press Enter

## SUMIF Function

- `SUMIF` function combines a sum with an If statement to limit which data are
  included in the total based on criteria
- In a `SUMIF` function, cell values are included in the sum only if they match
  the criteria
- `SUMIF` function has two required arguments and one optional argument

### HOW TO: Build a SUMIF Function

1. Click the cell for the function
2. Click the Math & Trig button [Formulas tab, Function Library group] and
   select `SUMIF`
3. Click the Range box
4. Select the cell range to be compared against the criteria
5. Press F3 (FN+F3) to paste a range name rather than selecting the range
6. Click the Criteria box
7. Type text to be compared or a value to be matched
8. Click the Sum_range box
9. Select the cells to be added
10. If the cells to be added are the same as those in the range argument, leave
    the entry blank
11. Click OK

## SUMPRODUCT Function

- `SUMPRODUCT` function calculates the sum of the product of several ranges
- It multiplies the cells identified in its array arguments and then totals
  those individual products
- An array is a collection of values in a row, a column, or both, basically a
  range
- The `SUMPRODUCT` function has the following
  syntax: `=SUMPRODUCT(array1, array2, [arrayN])`
- It is possible to use more than two arrays so that you could have an array3
  argument or even more

### HOW TO: Build a SUMPRODUCT Function

1. Click the cell for the function
2. Click the Math & Trig button [Formulas tab, Function Library group] and
   select `SUMPRODUCT`
3. Click the Array1 box
4. Select the first cell range to be multiplied
5. Click the Array2 box
6. Select the next cell range to be multiplied
7. Use as many arrays as required for the calculation
8. Click OK

## Conclusion

- Excel's Math & Trig functions include `SUM`, `ROUND`, `SUMIF`,
  and `SUMPRODUCT`
- `ROUND` function rounds values based on specified number of decimal places
- `SUMIF` function combines a sum with an If statement to limit which data are
  included in the total based on criteria
- `SUMPRODUCT` function calculates the sum of the product of several ranges

# Chapter Summary

- **2.1** Build and edit basic formulas (p. E2-106).
    - A formula is a calculation that uses arithmetic operators, worksheet
      cells, and constant values. Basic arithmetic operations are addition,
      subtraction, multiplication, and division.
    - Type a formula in the cell or point and click to select cells.
    - When you type a formula, **Formula AutoComplete** displays suggestions
      for completing the formula.
    - Formulas are edited in the Formula bar or in the cell to change cell
      addresses, use a different operator, or add cells to the calculation.
- **2.2** Set mathematical order of operations in a formula (p. E2-109).
    - Excel follows mathematics rules for the order in which operations
      are carried out when a formula has more than one operator.
    - Control the sequence of calculations by placing operations that
      should be done first within parentheses.
    - Use the following acronym to help remember the order of arithmetic
      operations: **P**lease **E**xcuse **M**y **D**ear **A**unt **S**ally
      (**P**arentheses, **E**xponentiation, **M**ultiplication, **D**ivision,
      **A**ddition, **S**ubtraction).
- **2.3** Use absolute, mixed, relative, and 3D references in a formula
    - A **relative cell reference** in a formula is the cell address which,
      when copied, updates to the address of the copy.
    - An **absolute cell reference** is the cell address with dollar signs, as
      in $A$5. This reference does not change when the formula is copied.
    - A mixed cell reference contains one relative and one absolute
      address, as in $B5 or B$5. When copied, the absolute part of the
      reference does not change.
    - A **3D cell reference** is a cell in another worksheet in the same
      workbook. It includes the name of the worksheet followed by an
      exclamation point, as in Inventory!B2.
    - Name a single cell or a group of cells with a defined **range name**.
    - Use range names in formulas instead of cell addresses.
    - Formula AutoComplete displays range names so that you can paste them
      in a formula.
    - A named range is an absolute reference.
    - Navigate from the Name Box to named ranges in the workbook.
- **2.4** Use formula auditing tools in a worksheet (p. E2-120).
    - Excel highlights certain types of formula errors as you work, but
      you still need to review errors and make corrections.
    - Excel automatically error-checks formulas based on its internal
      rules. A potential error is marked in the upper-left corner of the cell
      with a small triangle.
    - **Formula auditing** tools include several commands to aid your review
      of workbook formulas and functions.
    - The Formula Auditing group on the Formulas tab includes the Trace
      Precedents and Trace Dependents buttons.
    - A **circular reference** is an error that occurs when a formula includes
      the address of the formula.
    - Occasionally errors are noted as you press a completion key and can
      be quickly corrected by accepting the suggested correction in the
      message window.
- **2.5** Work with Statistical and Date & Time functions
    - Use the Insert Function dialog box to search for a function from a
      category or by a description of your task.
    - The Function Arguments dialog box details required and optional
      arguments and helps you complete a function according to its syntax.
    - The `MEDIAN` function calculates the value that represents the precise
      middle of a range.
    - The MODE functions find the most common value (`MODE.SNGL`) or values
      (`MODE.MULT`) in a range.
    - The `COUNT` function tallies the number of cells in a range but only
      includes cells with values.
    - Use `COUNTA` to count any type of data and `COUNTBLANK` to count empty
      cells.
    - The **AutoCalculate** feature displays numerical results such as Sum,
      Average, and Count on the Status bar for selected cells.
    - Excel defaults to automatic calculation so that all formulas are
      recalculated as soon as you make an edit.
    - The Date & Time category includes a TODAY function and a NOW
      function that display the current date and time.
- **2.6** Use functions from the Financial, Logical, and Lookup &
  Reference categories
    - The PMT function from the Financial category calculates a constant
      payment amount for a loan.
    - The Logical function `IF` evaluates a statement or condition and
      displays a particular result when the statement is true and another
      result when the condition is false.
    - A Lookup function displays data from a list located in another part
      of the workbook.
    - `XLOOKUP` searches for a value in one range and displays the match
      from another range.
    - An **array** is a collection of values, which usually maps to a range in
      Excel.
    - All versions of Excel include `VLOOKUP` (vertical) and `HLOOKUP`
      (horizontal) functions.
- **2.7** Work with Text functions
    - Text functions are used to manage strings of text, values, or
      characters.
    - The `TEXTJOIN` function joins a series of labels or other data using a
      single delimiter between each item.
    - A **delimiter** is a character that separates data items.
    - `CONCAT` joins or combines data strings with specified separators for
      each item.
    - When you use a Text function to join data, you must keep the source
      data in an accessible location so that the function can execute as
      expected.
- **2.8** Build functions from the Math & Trig category
    - The `ROUND` function adjusts a value up or down based on the number of
      decimal places.
    - The `SUMIF` function includes cells in a total only if they meet a set
      **criteria** or condition.
    - The `SUMPRODUCT` function multiplies corresponding cells from an array
      and then totals the results of each multiplication.