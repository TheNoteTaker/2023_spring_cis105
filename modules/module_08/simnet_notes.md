# Module 8 - Creating and Editing Charts

# General Notes

## Chapter Overview

- Excel can be used to graph or chart data in addition to building formulas and
  functions in a worksheet.
- After selecting values and labels, a professional looking chart can be created
  with a few clicks of the mouse.
- This chapter covers the basics of creating, editing, and formatting Excel
  charts.

# SLO 3.1 - Creating an Excel Chart

## Introduction

- An Excel chart visually represents numeric data in a worksheet.
- Charts help identify trends, make comparisons, and recognize patterns in the
  data.
- Charts are dynamic and linked to the data, so when values change, the chart is
  automatically updated.

## Creating a Chart Object

- A chart object is a distinct item surrounded by a square border that contains
  chart elements such as titles, axes, and gridlines that can be edited.
- Source data are the cells that contain values and labels to be graphed.
- It is important to select appropriate data for a chart because it is possible
  to build charts that illustrate nothing of any consequence.
- Data for a chart is best arranged with labels and values in adjacent cells and
  no empty rows.
- The Quick Analysis button appears in the lower-right corner of the selection
  when you select contiguous source data.
- Excel includes different types of charts and can recommend the best type based
  on your selected data.

### Excel Chart Types

- Excel has many chart types and most types have subtypes or variations.
- The most common chart types are column or bar, line, and pie.

|         Chart Type         | Purpose                                                                                                                                                   | Data Example                                                                                               | Categories (Labels)         | Values (Numbers)                    |
|:--------------------------:|:----------------------------------------------------------------------------------------------------------------------------------------------------------|:-----------------------------------------------------------------------------------------------------------|:----------------------------|:------------------------------------|
|         **Column**         | Illustrates data changes over a period of time or shows comparisons among items                                                                           | Monthly sales data for three automobile models                                                             | Horizontal axis             | Vertical axis                       |
|          **Bar**           | Displays comparisons among individual items or values at a specific period of time                                                                        | Number of pairs sold for eight shoe styles                                                                 | Verical axis                | Horizontal axis                     |
|          **Pie**           | Uses one data series to display each value as a percentage of the whole                                                                                   | Expenses or revenue by department for one quarter                                                          | NA                          | One data series shown by slice size |
|          **Line**          | Displays trends in data over time, emphasizing the rate of change                                                                                         | Number of weekly website views over a ten week period                                                      | Horizontal axis             | Vertical axis                       |
|          **Area**          | Displays the magnitude of change over time and shows the rate of change                                                                                   | Yearly consumption of apples, bananas, and pears over a ten year period                                    | Horizontal axis             | Vertical axis                       |
| **XY (Scatter) or Bubble** | Displays relationships among numeric values in two or more data series; these charts do not have a category                                               | Number of times a patient visits a doctor, amount billed to insurance, and cost billed to patient          | Horizontal axis (value 1-x) | Vertical axis (value 2-y)           |
|          **Map**           | Compares and plots values across geographical regions such as countries, states, counties, or postal codes                                                | Tax revenue from all counties in a state or province                                                       | Either                      | Either                              |
|         **Stock**          | Displays three series of data to show fluctuations in stock prices from high to low to close                                                              | Opening, closing, and high price of Microsoft stock each day for 30 days                                   | Horizontal axis             | Vertical axis                       |
|        **Surface**         | Displays results for combinations of two sets of values on a surface                                                                                      | Windchill factor based on wind speed and outdoor temperature                                               | NA                          | NA                                  |
|         **Radar**          | Displays the frequency of multiple data series relative to a center point. There is an axis for each category                                             | Style comfort and vlaue ratings for three snow boot styles                                                 | NA                          | NA                                  |
|        **TreeMap**         | Displays a hierarchical view of data with different sized and colored rectangles and sub-rectangles to compare the sizes of groups                        | Soda sales by product name, continent, country, and city                                                   | NA                          | NA                                  |
|        **Sunburst**        | Displays a hierarchical view of data with concentric rings. The top hierarchy is the inner ring and each outer ring is related to its adjacent inner ring | Tablet sales by screen size, continent, country, and cities                                                | NA                          | NA                                  |
|       **Histogram**        | Column-style chart that shows frequencies within a distribution                                                                                           | Number of students in each of five grade categories for an exam                                            | Horizontal axis             | Vertical axis                       |
|     **Box & Whisker**      | Displays the distribution of data with minimum, mean, maximum, and outlier values                                                                         | Sales prices of homes in five suburbs during a three week period                                           | Horizontal axis             | Vertical axis                       |
|       **Waterfall**        | Plots each element in a running total and displays negative and positive effects of each on the totals                                                    | Banking or savings account register                                                                        | Horizontal axis             | Vertical axis                       |
|         **Funnel**         | Plots values that progressively decrease from one stage or process to the next                                                                            | Job applications recieved, applications selected, online interviews, team interviews                       | Verical axis                | Horizontal axis                     |
|         **Combo**          | Uses two types of charts to graph values that are widely different                                                                                        | Line chart for number of monthly website visits and a column chart for monthly sales of a selected product | Either                      | Either                              |

## Size and Position a Chart Object

- A chart object is active and surrounded by eight selection handles when
  selected.
- To move a chart object, point to its border and click.
- To resize a chart object, drag a corner handle.
- Chart Design and Format tabs are available on the Ribbon when a chart is
  active or selected.

## Creating a Chart Sheet

- A chart sheet is an Excel chart that displays on its own sheet in the
  workbook.
- A chart sheet does not have rows, columns, and cells, but the chart is linked
  to its data on the source worksheet.
- Excel uses default sheet names Chart1, Chart2, and so on, but you can type a
  descriptive name in the Move Chart dialog box.
- Source data for a chart need not be contiguous.

### Quick Summary of Steps:

1. Select the data you want to include in the chart.
2. Click the Quick Analysis button in the lower-right corner of the selection.
3. Select the Charts command and choose the recommended chart type.
4. Size and position the chart object as desired.
5. Optionally, move the chart object to its own sheet with the Move Chart button
   on the Chart Design tab.

### Conclusion

Creating an Excel chart is an essential tool for identifying trends, making
comparisons, and recognizing patterns in data. By selecting appropriate data and
using the recommended chart types, you can create effective and meaningful
charts that support your work. Remember to size and position the chart object,
and you can also create a chart sheet to display the chart on its own sheet in
the workbook.

# SLO 3.2 - Designing Charts

- Excel provides tools to enhance chart appearance for originality, readability,
  and appeal.
- Chart Design tab includes commands for:

1. selecting a chart layout
2. choosing a chart style
3. switching the data
4. changing the chart type

- A newly created chart has a default layout, color, and style.

## Apply a Quick Layout

- A chart layout is a set of elements and their location in the chart.
- Elements are individual parts of a chart, such as a main title, a legend, and
  axes titles.
- Quick Layout button [Chart Design tab, Chart Layouts group] opens a gallery of
  predefined layouts.
- Live Preview redraws the chart as you point to an option in the gallery.
- Chart layout adds an element such as a title, which displays a generic label
  like "Chart Title."
- You can add or remove individual elements, as well as edit placeholder text.

### Figure 3-6 Quick Layout gallery and Live Preview

![Quick Layout Gallery and Live Preview](https://i.imgur.com/0B2Q7Rw.png)

## Apply a Chart Style

- A chart style is a preset combination of colors and effects for a chart, its
  background, and its elements.
- Chart styles available for a chart are based on the current workbook theme.
- Changing the theme updates the chart style colors, and the chart reflects the
  new color palette.
- Chart styles are available in the Chart Styles group on the Chart Design tab.
- You can preview the effects of a chart style by pointing to each style in the
  gallery.

## Change Chart Colors

- The workbook theme and the chart style form the basis for a chart’s color
  scheme.
- The Chart Styles group [Chart Design tab] includes a Change Colors button with
  optional color palettes.
- Palettes are divided into Colorful and Monochromatic groups.

## Switch Row and Column Data

- Excel plots the data based on the rows and columns you select in the worksheet
  and the chart type.
- In a column chart, the x-axis is the category along the bottom of the chart,
  and the y-axis along the left represents the data series (the values).
- You can change how the data are plotted.

## Change the Chart Type

- When a chart does not depict what you intended, or you need a variety of chart
  types for a project, change the chart type.
- Changing the type assumes that the source data is the same and that it can be
  graphed in the new chart.
- Do not change a column chart with three data series into a pie chart because a
  pie chart has only one series.
- The Change Chart Type button is on the Chart Design tab in the Type group.
- For a selected chart object or sheet, you can also right-click and choose
  Change Chart Type from the context menu.
- The Change Chart Type dialog box includes Recommended Charts and All Charts
  tabs like the Insert Chart dialog box.

## Filter Source Data

- A chart displays all the categories and all the data series for the selected
  cells.
- You can filter or refine that data by hiding categories or series.
- A filter is a requirement or condition that identifies which data is shown and
  which data is hidden.
- Chart filters do not change the underlying cell range for a chart, but they
  enable you to focus on particular data.
- Filter chart data from the Chart Filters quick button or in the Select Data
  Source dialog box [Chart Design tab, Data group].

## Editing Chart Source Data

### Adding or Removing Cells

- The source data of a chart is made up of its cells.
- To edit the source data, add, remove, or change cells.
- To add a data series to a column chart, add another column to the data range.
- To remove a data series, delete the corresponding cells.
- To change the data range of a chart, select the chart and drag the sizing
  arrow in a corner of the highlighted cell range to expand or shrink the range.

### Editing Chart Source Data for a Chart Sheet

- Click the Select Data button on the Chart Design tab [Data group].
- The Select Data Source dialog box opens.
- You can use the dialog box for a chart sheet or a chart object.
- Select the new cell range in the worksheet.
- If the cell ranges are not adjacent, select the first one, type a comma, and
  select the next range.
- Type cell addresses in the Chart data range box rather than selecting cells.
- Click OK.
- The chart sheet is active and redrawn with the edited source data range.

### Resizing the Chart Source Data Range

- Select the chart object.
- The source cells are highlighted.
- Point to the corner of the highlighted cell range to display the resize
  pointer.
- You can resize the range from any corner.
- Drag to resize the cell range.
- The chart sheet is redrawn with the new range.

## Printing Charts

### Printing a Chart with Data

- Deselect the chart object by clicking a worksheet cell.
- Size and position the chart and scale the sheet to fit a printed page.
- Use regular Page Setup commands to complete the print task, such as choosing
  the orientation or inserting headers or footers.

### Printing a Chart Object

- Select the chart object.
- Use regular Print and Page Setup options.
- A selected chart object, by default, prints scaled to fit a landscape page.

### Printing a Chart Sheet

- Click the chart sheet tab.
- Click the File tab and select Print.
- Choose print settings as needed.
- Change the margins to grow or shrink the chart.
- Add a header or footer as desired.
- Click Print.
- The chart sheet prints.

# SLO 3.3 - Managing Chart Elements

## Chart Elements

- A chart element is a separate object that can be added, removed, formatted,
  sized, and positioned as you design a chart.
- Chart layout and style affect which elements are initially displayed.

|            Element             | Description                                                                                                                                                                                        |
|:------------------------------:|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|            **Axis**            | Horizontal or vertical boundary that identifies what is plotted.                                                                                                                                   |
|         **Axis title**         | Optional description for the categories or values.                                                                                                                                                 |
|         **Chart area**         | Background for the chart; can be filled with a color, gradient, or pattern.                                                                                                                        |
|        **Chart floor**         | Base or bottom for a 3-D chart.                                                                                                                                                                    |
|        **Chart title**         | Optional description or name for the chart.                                                                                                                                                        |
|         **Chart wall**         | Vertical background for a 3-D chart.                                                                                                                                                               |
|         **Data label**         | Optional element that displays values with the marker for each data series.                                                                                                                        |
|        **Data marker**         | Element that represents individual values. The marker is a bar, a column, a slice, or a point on a line.                                                                                           |
|         **Data point**         | A single value or piece of data from a data series.                                                                                                                                                |
|        **Data series**         | Group of related values that are in the same column or row and translate into the columns, lines, pie slices, and other markers.                                                                   |
|         **Error bars**         | Vertical or horizontal bars that represent a potential error range or degree of uncertainty for each value.                                                                                        |
|          **Gridline**          | Horizontal or vertical line that extends across the plot area to help identify values.                                                                                                             |
| **Horizontal (category) axis** | Describes what is shown in the chart and is created from row or column labels. In a bar chart, the category axis is the vertical axis; the category axis is the horizontal axis in a column chart. |
|           **Legend**           | Element that explains symbols, textures, or colors used to differentiate data series.                                                                                                              |
|         **Plot area**          | Rectangular area bounded by the horizontal and vertical axes.                                                                                                                                      |
|         **Tick mark**          | Small line or marker on an axis to guide in reading values                                                                                                                                         |
|         **Trendline**          | Line or curve that displays averages in the data and can be used to forecast future averages.                                                                                                      |
|   **Vertical (value) axis**    | Shows the numbers on the chart. In a bar chart, the vertical axis is along the bottom; in a column chart, the vertical axis is along the side.                                                     |

## Chart and Axes Titles

- A main chart title is positioned above the chart, but within the chart area.
- An axis title displays along the bottom or left or right boundary of a chart.
- Axes titles may not be necessary if the chart graphs the data well.
- ScreenTip names or describes chart elements when you point to them.
- The Chart Elements drop-down list on the chart Format tab allows you to select
  an element.

### Add Chart and Axis Titles

1. Click the chart object or the chart sheet tab.
2. Click the Chart Elements quick button in the top-right corner of the chart.
3. Select the Chart Title box.
4. Click the Chart Title arrow in the Chart Elements pane to select a position
   for the title.
5. Select the Chart Title placeholder text and type a title to replace the
   placeholder text.
6. Click the chart border to deselect the title object.
7. Click the Chart Elements quick button and select the Axis Titles box.
8. Select the Axis Title placeholder text and type a title to replace the
   placeholder text.
9. Click the chart border to deselect the axis title object.

### Remove Chart Elements

- Removing chart elements, especially axes titles, creates more room for the
  chart data.
- To remove a chart element, select it and press Delete or click the Chart
  Elements quick button and deselect the box for that element.
- The Chart Elements box on the chart Format tab confirms which element is
  selected.

## Formatting Chart Elements

- Most chart layouts include placeholders for a chart title and axes titles, or
  you can add the elements.
- You must edit the placeholder text to personalize your chart.
- Text within a chart object can be formatted with font attributes from the Font
  and Alignment groups when the element is selected.
- To apply a format to a portion of the text, click to place an insertion point
  inside the element and select the characters to be changed.
- Double-clicking a chart element opens the Format pane with format and design
  choices for that element.

## Additional Information

- Clicking the Add Chart Element button [Chart Design tab, Chart Layouts group]
  shows a chart element.
- Clicking the Chart Elements quick button in the top-right corner of a chart
  displays a pane with a list of available elements for the chart type.
- Pressing Enter while typing in a title text box inserts a new line in the
  title.
- Pressing Backspace removes the new line.
- Right-clicking a chart element and choosing Delete also removes the element.

## Data Labels

- Data labels display the number or percentage represented by a column, bar, pie
  slice, or other marker on the chart.
- They can be used to display a precise value that a marker cannot show due to
  the scale used on the value axis.
- Data labels should be used sparingly when the chart has few data series to
  avoid clutter.
- To add data labels, select the chart object or sheet tab, click the Chart
  Elements quick button or the same button on the Chart Design tab in the Chart
  Layouts group, select the Data Labels box, and choose a location for the
  labels.
- Live Preview shows data labels for each data point in the data series.
- The position choices for the labels are listed, and Live Preview updates as
  you point to each option.

## Data Table

- A chart data table is a columnar display of the values for each data series in
  a chart, located just below the chart.
- It can supply valuable information when readers do not have access to the
  source data for a chart.
- To include a data table below your chart, use the Add Chart Element button on
  the Chart Design tab in the Chart Layouts group or the Chart Elements quick
  button.
- A data table can be shown with or without legend markers, a mini legend to the
  left of the table.

## Error Bars

- Error bars illustrate a degree of variability or uncertainty in the data and
  are relevant in scientific or statistical charts.
- They can be designed to show a standard 5% error, a percentage, or a standard
  deviation.
- Error bars can be displayed for area, bar, column, line, and XY charts.
- To add them to your chart, use the Add Chart Element button on the Chart
  Design tab or the Chart Elements quick button.
- A related Format task pane provides options for customizing the error bars.

## Gridlines and the Legend

- Gridlines are vertical or horizontal lines that guide in interpreting the
  values by extending across the plot area of a chart.
- Major and minor vertical and horizontal gridlines can be displayed along with
  tick marks on the axes, which are part of the axis formatting.
- Gridlines are not available for chart types that do not have axes, such as pie
  charts.
- To display gridlines, use the Add Chart Element button on the Chart Design tab
  or the Chart Elements quick button.
- A legend is a list of colors, symbols, or textures used to differentiate the
  data series.
- It can be displayed at the top, bottom, right, or left of the plot area.
- To display a legend, use the Chart Elements quick button and select the Legend
  box in the Chart Elements pane, and choose a location for the legend.
- Live Preview updates as you point to each option.

## Trendlines

- A trendline plots patterns using a moving average for the data and assumes a
  time element in the data.
- It is a straight or curved line that can extend to predict future averages.
- A basic linear trendline is appropriate for values that tend to increase or
  decrease as time passes.
- Not all charts are suitable for trendlines, such as a pie or a treemap chart.
- To add a trendline, use the Chart Elements quick button and select the
  Trendline box in the Chart Elements pane, and choose a type of trendline.
- Alternatively, right-click a data series marker and choose Add Trendline to
  open the Format Trendline task pane and build a trendline.
- The Trendline Format task pane provides options for customizing the trendline,
  including logarithmic and other statistical concepts.

# SLO 3.4 - Formatting Chart Elements

## Overview

- Chart type, layout, and style contribute to the overall appearance of your
  chart.
- Each chart element can be formatted, including fill or outline color, and
  custom format options are specific to the element.
- The Format Task Pane provides custom commands for each chart element,
  including shape, fill, and color options.

## The Format Task Pane

- The Format Task Pane consolidates shape, fill, and color options for each
  chart element.
- Open the Format Pane for a chart object by double-clicking or right-clicking
  the element and choosing Format [Element Name] from the menu.
- The Format Pane opens to the right of the workbook window and automatically
  updates to reflect the selected element.
- For each options group, a Format Task Pane has at least two buttons, including
  Fill & Line, Effects, and Options.

### How to Use the Format Task Pane to Change Fill

1. Double-click the chart element to open the Format Task Pane for the element
   on the right.
2. Choose a chart object that can be formatted with a fill color such as the
   chart area, the plot area, or a data series in a column or bar chart.
3. Confirm that the intended chart element is selected in the Format Task Pane.
4. Click the name of the Options group at the top of the task pane.
5. Click the Fill & Line button in the Format Task Pane.
6. Click Fill to expand the group.
7. Click the Color button to open the gallery and choose the desired color.
8. Select another element in the chart to see its Format Task Pane.

### How to Use the Format Task Pane to Format Data Labels

1. Click the Chart Elements drop-down list (Format tab, Current Selection
   group).
2. Choose Series [Name] Data Labels in the list.
3. All data labels are selected.
4. Click the Format Selection button (Format tab, Current Selection group).
5. The Format Data Labels task pane displays.
6. Choose Label Options at the top of the task pane.
7. Expand a command group to display format options.
8. Select format options for the labels.

## Apply a Shape Style

- Shape styles are predesigned sets of borders, fill colors, and effects.
- Shape styles display in a gallery and are grouped as Theme Styles or Presets.
- Apply a shape style by selecting the chart element, clicking the More button
  in the Shape Styles group, and choosing an icon to apply.

## Apply Shape Fill, Outline, and Effects

- Fill color, outline width and color, and special effects are available
  separately, and your choices override shape style settings already applied.
- Apply a gradient fill to a chart element by selecting the element, clicking
  the Shape Fill button, and choosing Gradient from the gallery.
- Apply an outline to a chart element by selecting the element, clicking the
  Shape Outline button, and choosing a color and width from the gallery.

## Conclusion

In conclusion, formatting chart elements can greatly enhance the visual appeal
of your charts. The Format Task Pane provides a variety of custom commands for
each chart element, including shape, fill, and color options. Shape styles and
special effects can be applied to create a predesigned set of borders, fill
colors, and effects. Careful use of fill, borders, and effects is essential to
avoid overly formatted charts.

# SLO 3.5 - Enhancing a Chart

## Adding Pictures, Icons, and WordArt to a Chart

- Pictures can be inserted as separate design objects or used as fill for a
  chart element.
- Use WordArt for a distinctive look on a chart title.
- Drawing a shape such as an arrow or a lightning bolt can highlight a
  particular data point.

### Using Icons as Shape Fill

- Icons can be used to fill the chart area, the plot area, or a marker.
- Icons can be selected from the available Microsoft group or an image file from
  your computer, external media, or online source.
- Always verify permission to use an image in your work.
- The size and design of the icon or picture is important to avoid making the
  chart appear busy.
- You may be able to edit attributes of an icon or picture in Excel, but this
  depends on the original format.
- To insert an icon or illustration from the Illustrations group:
    1. Insert the image in the worksheet apart from the chart.
    2. Design and size it as desired and copy it to the Clipboard.
    3. From the Format task pane, insert the image from the Clipboard into the
       chart
       element.
    4. After inserting the image, the task pane displays Stretch and Stack
       options:
    5. Stretch enlarges and elongates a single copy of an icon to fill the
       element.
    6. The Stack choice sizes and repeats the image to fill the area.
    7. Use the Stack and Scale setting for precise matching of the number of
       images
       used per unit of value.

### Inserting Shapes in a Chart

- The Shapes gallery is located in the Insert Shapes group on the chart Format
  tab.
- Shapes can be used to highlight or emphasize a point in the chart.
- They can be drawn anywhere on the chart but are usually placed on the
  background or the plot area so they do not obscure any data.
- Once drawn, a shape is an object.
- When selected, the Shape Format tab is available with commands for altering
  the appearance of the shape as well as the Format Shape task pane.

### Using WordArt in a Chart

- WordArt is text with preset font style, fill, and effects.
- In a chart, you select an element and style it with WordArt.
- In a worksheet, you draw a WordArt text box and position it.
- WordArt is best used for main chart titles because smaller elements may be
  unreadable.
- You can apply a WordArt style from the chart Format tab, WordArt Styles group,
  and format a WordArt style by individually changing its fill, outline, and
  effects.

## Other Chart Enhancements

### The Selection Pane

- Objects, including charts, are placed on invisible layers in a worksheet.
- The Selection Pane command [Format tab, Arrange group] displays the names of
  objects such as charts, icons, images, and shapes.
- The Selection pane enables you to hide, display, or rearrange objects in your
  workbook.

### Alt Text for a Chart

- Alt text is a brief description or explanation of an object.
- It is alternative text for a visual.
- Alt text is read aloud by a screen reader such as Windows Narrator, to aid
  users with visual impairment.
- You can increase accessibility of your charts by adding Alt Text that further
  explains the data.

---

## Conclusion:

In conclusion, enhancing a chart with pictures, icons, WordArt, shapes, and Alt
text can make it more visually appealing and accessible to users. While doing
so, it is important to consider the size and design of the added elements to
avoid making the chart appear busy. Additionally, using the Selection Pane can
help in organizing the different elements of the chart.

# SLO 3.6 - Creating and Editing Charts

### Building Pie, Line, and Combo Charts

- A pie chart represents one set of related values and shows the proportion of
  each value to the total.
- Line charts illustrate how values change over a period of time.
- A combo chart combines at least two sets of values and plots each data series
  with a different chart type.
- A combo chart can also have two value axes, one on the left and one on the
  right.

### Create a 3-D Pie Chart

- A pie chart graphs one data series and illustrates how each number relates to
  the whole.
- Be cautious about the number of categories, because a pie chart with hundreds
  of slices is difficult to interpret and does not depict the relationship among
  the values.
- From the Insert Pie or Doughnut Chart button on the Insert tab, you can select
  from a gallery of 2-D or 3-D pie types as well as a doughnut shape.
- You can also select options for a pie chart with a bar or another pie chart.
- As you point to each subtype, Live Preview previews the chart.

#### HOW TO: Create a 3-D Pie Chart Sheet

1. Select the cells that include the values and the category labels.
2. The labels and values need not be adjacent.
3. Pie charts do not have axes.
4. Click the Insert Pie or Doughnut Chart button [Insert tab, Charts group].
5. Choose 3-D Pie.
6. A pie chart object displays in the worksheet (Figure 3-42).
7. A 2-D pie includes subtypes for a pie chart with a second pie chart or a pie
   chart with a bar chart.
8. Click the Move Chart button [Chart Design tab, Location group].
9. Click the New sheet button.
10. Type a name for the chart sheet and click OK.

### Pie Chart Elements and Options

- In a pie chart, the data series is represented by the whole pie. Each data
  point is a slice of the pie.
- You can format the data series as a whole or as individual slices.
- A pie chart can display a legend and a title.
- Use data labels in place of a legend because they show the same information.
- Custom commands for a pie chart are the angle of the first slice and the
  percent of explosion.

#### HOW TO: Rotate a Pie Chart and Explode a Slice

1. Double-click the pie shape.
2. The Format Data Series task pane opens.
3. Another way to open the task pane is to right-click the pie shape and choose
   Format Data Series.
4. Click the Series Options button in the pane.
5. The first slice starts at the top of the chart or 0° (zero degrees).
6. Drag the vertical slider to set the Angle of first slice.
7. The first slice arcs to the right.
8. The percentage is shown in the entry box.
9. You can also type a percentage or use the spinner arrows to set a value.
10. Click the data point (slice) to explode.
11. The Format Data Point task pane opens.
12. Drag the vertical slider to set the Pie Explosion (Figure 3-43).
13. The larger the percentage, the farther the slice moves from the whole.
14. The percentage is shown in the entry box.
15. Type a percentage or use the spinner arrows to set a value.

### Create a Line Chart

- A line chart graphs one or more data series and shows how the values progress
  over a span of time.
- Line charts include a horizontal axis (x-axis) and a vertical axis (y-axis).
- The categories are usually along the horizontal axis and depict equal
  intervals such as every month or every year for the time period.
- Line charts in the Charts group include area charts.
- An area chart can be described as a line chart with the space below the line
  filled in.
- It can better emphasize the magnitude of the changes (Figure 3-44).

#### HOW TO: Create a Line Chart

1. Select the cells that include values and category labels. The labels and
   values need not be adjacent.
2. A line chart with too many data series is difficult to interpret.
3. Click the Insert Line or Area Chart button [Insert tab, Charts group].
4. Subtypes include a ScreenTip with suggestions for how to use the subtype.
5. Choose Line with Markers from the 2-D group.
6. A line chart object displays in the worksheet (Figure 3-45).

### Line Chart Elements and Options

- Each data series displays as a line, and each data point is the actual value.
- If there are not too many categories along the horizontal axis, it is a good
  idea to include line markers to clearly show where the value is plotted.
- A marker in a line chart is a shape distinct from the line such as a square or
  a triangle.
- A line chart can display a legend and a title.
- Custom commands include the width of the line, the line style (solid, dashed,
  dotted), the end point, and the transparency.
- Line width is the thickness of the line, measured in points like fonts.
- You can also choose from several built-in marker symbols and select marker
  color and size.

#### HOW TO: Format Data Series in a Line Chart

1. Double-click the line.
2. The Format Data Series task pane opens.
3. Alternatively, click the Chart Element drop-down
   list [Format tab, Current Selection group] and select the Series “Name.” Then
   click the Format Selection button.
4. Click the Fill & Line button in the pane.
5. Line and Marker options are available in this group.
6. Expand the Line group if it is collapsed.
7. Click the Color arrow and choose a line color.
8. Document theme colors, standard colors, and custom colors are available.
9. Click the Width spinners arrow to adjust the thickness of the line (Figure
   3-46).
10. Each click adjusts the size in .25-point increments.
11. You can click the entry box and type a custom width, too.
12. Live Preview shows your changes as you make them.
13. Click Marker in the pane to expand the marker options.
14. The command subgroups are Marker Options, Fill, and Border.
15. Click Marker Options to expand the group.
16. Select the Built-in radio button and click the Type drop-down list.
17. Built-in markers include several common shapes.
18. Click a shape for the marker.
19. Click the Size spinners arrow to shrink or grow the marker.
20. Live Preview shows your changes as you make them.
21. Expand the Fill group if it is collapsed.
22. Click the Color arrow and choose a color for the marker (Figure 3-47).
23. Choose the same color as the line as needed.
24. Expand the Border group as desired.
25. Markers generally do not have borders.

### Building Pie, Line, and Combo Charts

- Pie chart represents one set of related values and shows the proportion of
  each value to the total.
- Line charts illustrate how values change over a period of time.
- A combo chart combines at least two sets of values and plots each data series
  with a different chart type.
- A combo chart can have two value axes, one on the left and one on the right.

### Create a Combo Chart

- A combo chart includes at least two chart types such as a line chart and a
  column chart.
- Perfect Vacation Rentals can use a combo chart to compare rental days at one
  property to total rental days per month.
- From the Insert Combo Chart button in the Charts group on the Insert tab as
  well as from the All Charts tab in the Insert Chart dialog box, you can select
  from a gallery of combo chart subtypes.
- To create a combo chart:
    1. Select the cell ranges for the chart.
    2. Click the Insert Combo Chart button [Insert tab, Charts group].
    3. Point to a subtype to see a description and a Live Preview.
    4. Select the chart subtype.
    5. A combo chart object displays in the worksheet.
    6. Move the chart to its own sheet as desired.
    7. Edit the data series as needed.
    8. Decide which set of values should display on the secondary axis.

### Combo Chart Elements and Options

- A combo chart has at least two data series, each graphed in its own chart
  type.
- In a combo chart, you can display values on two vertical axes.
- The axis on the left is the primary axis; the one on the right is secondary.
- For both axes, Excel builds the number scale based on the data.
- Select the data series to be plotted on the secondary axis and choose the
  option from the Series Options command group in the Format task pane.
- You can also define a secondary axis in the Insert Chart or the Change Chart
  Type dialog boxes.
- A combo chart has the same elements and commands as a regular column or line
  chart.
- Apply chart styles and layouts, as well as shape fill, outline, effects,
  pictures, shapes, and WordArt.

### HOW TO: Display a Secondary Axis on a Combination Chart

1. Select the chart object or the chart sheet tab.
2. Select the data series to be shown on the secondary axis.
3. Select the line or any column in a line-column combo chart.
4. You can also choose the data series from the Chart Elements drop-down
   list [Format tab, Current Selection group].
5. Click the Format Selection button [Format tab, Current Selection group].
6. Click the Series Options button in the task pane.
7. Click the Secondary Axis radio button.
8. Additional options for the axis depend on the chart type.
9. Close the Format Data Series task pane.

# SLO 3.7 - Creating and Editing Charts

- Specialty chart types visualize data in ways that previously required
  engineering, mapping, or other expert software.
- The Charts group on the Insert tab includes buttons for Hierarchy, Statistic,
  and Map charts.
- To create a map chart, you must have high-detail geographical data and access
  to the Bing Map service.

## Creating Specialty Charts

- Specialty charts have characteristics of standard Excel charts as well as
  their own requirements or limitations.
- Examples of specialty charts are Hierarchy, Statistic, and Map charts.

## Create a Sunburst Chart

- A sunburst chart is a hierarchy chart that illustrates the relationship among
  categories and subcategories of data.
- The Insert Hierarchy Chart button has two subtypes, Sunburst and TreeMap.
- A sunburst chart has one data series and no axes. It can display a title, data
  labels, and a legend.
- Data should be arranged in hierarchical order that depicts what you intend.
- Perfect Vacation Rentals might keep a record of fees collected from additional
  services by state, service, and month.
- The inner ring is the top level of the hierarchy, the state. The second level
  is the actual service (Car Rental, Housekeeping, etc.), and the outside ring
  is the month.
- From the chart, you can see that, in Indiana, a visiting chef contributed more
  to revenue than the other services.

## Create a Map Chart

- A map chart is a one-dimensional chart that displays high-level geographic
  details.
- High-level geography means continents, countries, states, or provinces.
- A map chart illustrates a set of values for a set of locations.
- A map chart cannot plot cities or streets (low-level geography).
- Interactive detailed maps can be built using 3D Map and geography data types
  to connect to Microsoft Bing.
- Use common geographical names for labels such as “Country,” “State,” or
  “Province.”
- Expert users suggest that you format the data as an Excel table.
- A map chart does not have axes.
- A map chart can display data labels and a title.

## Create a Waterfall Chart

- A waterfall chart is a financial chart that displays a moving total for
  positive and negative values.
- A waterfall chart graphs how each expense or outlay affects the account.
- The Insert Waterfall, Funnel, Stock, Surface, or Radar Chart button is in the
  Chart group on the Insert tab.
- A waterfall chart plots one data series and has a legend that clarifies the
  increase, decrease, and total colors.
- A waterfall chart can display axes, titles, gridlines, and data labels.

## More Info

- A TreeMap chart illustrates hierarchical data using rectangles of various
  sizes and colors to represent the divisions and subdivisions.
- Detailed maps can be built using 3D Map and geography data types to connect to
  Microsoft Bing.

## How to's

### Create a Sunburst Chart

1. Select the source data with labels and values.
2. The leftmost column in the hierarchy displays as the innermost ring.
3. Click the Insert Hierarchy Chart button and select Sunburst from the subtypes
   menu.
4. A chart object displays in the worksheet.
5. Click the Move Chart button and select New sheet.
6. Type a name for the chart sheet and click OK.
7. The Chart Design and Format tabs are available.
8. Click the sunburst shape to select the circle.
9. Point to a name in the inner ring and click to select the data point.
10. Open the Format Data Point task pane and format the data point as desired.
11. Point to a name in the second ring from the center and click to select the
    data point.
12. Open the Format Data Point task pane and format the data point as desired.
13. Format each ring as needed.
14. Click outside the chart background to deselect the chart.

### Create a Map Chart

1. Prepare the data for the chart.
2. Use common geographical names for labels such as “Country,” “State,” or
   “Province.”
3. Expert users suggest that you format the data as an Excel table.
4. Select the cell range with labels or click a cell in the table.
5. Click the Insert Map Chart button and select Filled Map to insert the chart
   object.
6. The chart plots the values in the first column to the right of the
   identifying geography data if you clicked a cell in an Excel table.
7. Size and position the chart object.
8. Move the chart to its own sheet as desired.
9. A map chart does not have axes.
10. A map chart can display data labels and a title.
11. Data series format options are specific to the map chart.
12. Click a worksheet cell to deselect the chart.
13. Save the chart with the workbook as usual.

### Create a Waterfall Chart Sheet

1. Select the source data with labels and values.
2. Click the Insert Waterfall, Funnel, Stock, Surface, or Radar Chart button and
   select Waterfall.
3. A chart object is placed in the worksheet.
4. Click the Move Chart button and select New sheet.
5. Type a name for the chart sheet and click OK.
6. The Chart Design and Format tabs are available.
7. A waterfall chart plots one data series and has a legend that clarifies the
   increase, decrease, and total colors.
8. A waterfall chart can display axes, titles, gridlines, and data labels.

### Format a Waterfall Chart

1. Click the Chart Elements quick button.
2. Only elements available for the chart type are listed.
3. Verify or select the Chart Title, Legend, and Data Labels boxes.
4. Click the Legend object in the chart.
5. Your first click selects the legend and its entries.
6. Click the Increase entry in the Legend object.
7. Click the Format Selection button to open the Format Legend Entry task pane.
8. Click the Fill & Line button in the pane and expand the Fill group.
9. Select the radio button for the type of Fill as needed.
10. Click the Color button and choose a color from the palette.
11. The color applies to all data points that are positive values.
12. Click the Decrease entry in the Legend object.
13. Click the Color button and choose a color from the palette.
14. The color applies to data points that are negative values.
15. Click the Total entry in the Legend object and select its color.
16. Close the Format Legend Entry task pane.
17. Click the ending marker to select only that column.
18. Click the Format Selection button to open the Format Data Point task pane.
19. Click the Series Options button in the task pane and expand the Series
    Options group.
20. Select the Set as total box.
21. The marker rests on the bottom axis and is no longer floating.
22. Close the task pane and click outside the chart to deselect it.

## Conclusion

Specialty chart types such as hierarchy, statistic, map, sunburst, and waterfall
charts can be created using Excel's chart tools. To create a map chart, you need
high-detail geographical data and access to the Bing Map service. A sunburst
chart is a hierarchy chart that illustrates the relationship among categories
and subcategories of data. A waterfall chart is a financial chart that displays
a moving total for positive and negative values. These chart types have their
own requirements or limitations, but they can be useful for visualizing data in

# SLO 3.8: Inserting and Formatting Sparklines

Sparklines are miniature charts that can be inserted in a cell or cell range to
illustrate trends and patterns without adding a separate chart object or sheet.
Here are some key points related to inserting and formatting sparklines:

## Inserting Sparklines

- Three types of sparklines are available: Line, Column, and Win/Loss.
- The Sparklines commands are located on the Insert command tab.
- To insert column sparklines in a worksheet:
    1. Select the range of data to be graphed.
    2. Click the Column Sparkline button [Insert tab, Sparklines group].
    3. Select or edit the data range as needed.
    4. Click the Location Range box and select the cell or range for the
       sparklines.
    5. Click OK.

## Formatting Sparklines

- When sparklines are selected, the Sparkline tab opens and includes options for
  changing the appearance of the sparklines.
- The Style gallery options change the color scheme for the entire sparkline
  set.
- The Marker Color command enables you to choose a different color for
  identified values. You can select marker colors to highlight the high, low,
  first, or last value as well as negative values.
- To format sparklines:
    1. Click a cell with a sparkline.
    2. Click the Sparkline tab if it is not active.
    3. Click the More button [Sparkline tab, Style group] to display a gallery
       of styles.
    4. Choose a sparkline style.
    5. Select the High Point box in the Show group to apply a default color
       based on the sparkline style to the highest value in the sparkline for
       each row.
    6. Click the Sparkline Color button [Sparkline tab, Style group] to choose a
       color for the sparkline group.
    7. Click the Marker Color button [Sparkline Tools Design tab, Style group]
       to choose a color for highlighted marker values.

## Sparkline Design Tools

- The Sparkline tab also includes options for changing the type of sparkline,
  such as making a column sparkline into a line sparkline.
- When you define a location range for sparklines, they are embedded as a group.
  If you click any cell in the sparkline range, the entire group is selected.
  Subsequent commands affect all sparklines in the group.
- From the Group section on the Sparkline tab, you can ungroup the sparklines to
  select and format each sparkline cell individually.

## Clearing Sparklines

- You can remove sparklines from a worksheet with the Clear command in the Group
  group on the Sparkline tab.
- After sparklines are cleared, you may also need to delete the column where
  they were located or reset row heights and column widths.

More Info: Ungroup sparklines to create individual sparklines in each cell.

---

## Conclusion

Sparklines are a useful tool for illustrating trends and patterns in a compact
way. To insert and format sparklines, you can follow the steps listed above,
including selecting the type of sparkline, choosing a color scheme, and
highlighting specific marker values. Additionally, the Sparkline tab includes
options for changing the type of sparkline and ungrouping sparkline groups to
select and format each sparkline cell individually. Finally, if you need to
clear sparklines, you can use the Clear command and adjust row heights and
column widths as needed.

# Chapter Summary

- **3.1** Create Excel chart objects and chart sheets.
- A **chart** is a visual representation of worksheet data.
- A **chart object** is an item or element in a worksheet.
- A **chart sheet** is an Excel chart on its own tab in the workbook.
- The cells with values and labels used to build a chart are its
  **source data**.
- Chart objects and sheets are linked to their source data and contain
  chart elements such as data labels or a chart title.
- Commonly used chart types are Column, Line, Pie, and Bar, and Excel
  can build statistical, financial, geographical, and scientific charts.
- Size and position a chart object in a worksheet, or move it to its
  own sheet using the **Move Chart** button in the Location group on the Chart
  Design tab.
- The **Quick Analysis** tool includes a command group for charts, and it
  appears when the selected source data is contiguous.
- **3.2** Design charts by changing the layout, style, colors, and type.
    - A **chart layout** is a set of elements and their locations in a chart.
    - The Quick Layout button [Chart Design tab, Chart Layouts group]
      includes predefined layouts for the current chart type.
    - A **chart style** is a predefined combination of colors and effects for
      chart elements.
    - Chart styles are based on the current workbook theme.
    - The **Change Colors** command [Chart Design tab, Chart Styles group]
      provides color palettes for customizing a chart.
    - Excel plots data based on the number of rows and columns selected
      and the chart type, but row and column data can be switched.
    - Use the **Change Chart Type** button on the Chart Design tab to change
      selected chart types into another type.
    - **Filter** chart data to hide and display values or categories without
      changing the source data.
    - Edit source data to use a different cell range or to add or remove a
      data series.
    - Print a chart object with its worksheet data or on its own sheet.
    - A chart sheet prints on its own page in landscape orientation.
- **3.3** Manage chart elements including titles, data labels,
  gridlines, and trendlines.
    - Each **chart element** is a separate object in the chart.
    - Chart elements include chart and axes titles, data labels, **legends**,
      **gridlines**, and more.
    - Chart elements depend on the chart type; a pie chart, for example,
      does not have axes.
    - Click the **Chart Elements** quick button to display or remove
      individual elements.
    - Chart title elements display with placeholder text when first added
      to a chart.
- **3.4** Format chart elements with shape styles, fill, outlines, and
  special effects.
    - A chart element has a Format task pane that includes fill, outline,
      and effects commands as well as specific options for the element.
    - A **shape style** is a predesigned set of fill colors, borders, and
      effects.
    - Apply **shape fill, outline**, and **effects** to a chart element from the
      chart Format tab.
    - Some formats can also be applied to a selected chart element from
      the Home tab.
- **3.5** Enhance a chart with icons, shapes, WordArt, and Alt Text.
    - For chart shapes that have a fill color, you can use an icon or
      image as fill.
    - Icons and images used as fill can be from your own or online
      sources, but not all images work well as fill.
    - Shapes are predefined outline drawings available from the Insert
      Shapes group on the chart Format tab or from the Illustrations group on
      the Insert tab.
    - Place shapes on a chart to highlight or draw attention to a
      particular element.
    - **WordArt** is a preset design, often applied to a chart title.
    - Charts, icons, shapes, and other objects are placed on invisible
      layers in the worksheet.
    - Use the Selection pane to rearrange the order of the layers so that
      icons or images display in front of or behind each other.
    - Add **Alt Text** to a chart to briefly describe its contents and purpose
      for improved accessibility.
    - Alt Text is read aloud by a **screen reader** such as Windows Narrator.
- **3.6** Build pie charts, line charts, and combo charts.
    - A **pie chart** has one data series and shows each data point as a slice
      of the pie.
    - A pie chart does not have axes, but it does have options to rotate
      or explode slices.
    - A **line chart** illustrates how values change over a period of time.
    - Line charts display the categories along the horizontal axis and the
      values along the vertical axis.
    - A **combo chart** uses at least two chart types to highlight, compare,
      or contrast differences in data or values.
    - Format a combo chart to show a secondary axis when values are widely
      different or use different scales.
- **3.7** Create specialty charts.
    - Specialty charts include hierarchy, statistical, funnel, scatter
      charts, and more.
    - A **sunburst chart** has one data series that is grouped in a
      **hierarchy**.
    - A sunburst chart illustrates the relationship among the hierarchies
      in a pie-like chart with concentric rings.
    - A **map chart** is a simple display of high-level geographic data such
      as countries, states, or provinces.
    - A map chart has one data series and no axes.
    - Format a map chart to show only the region of interest and choose
      your own colors to distinguish the values.
    - A **waterfall chart** depicts a running total for positive and negative
      values.
    - A waterfall chart is used for financial or other data that shows
      inflows and outlays of resources.
    - A waterfall chart resembles a column chart and has category and
      value axes.
    - Place a specialty chart as an object or on a separate sheet.
- **3.8** Insert and format sparklines in a worksheet.
    - A sparkline is a miniature chart in a cell or range of cells in the
      worksheet.
    - Three **sparkline** types are available: **Line, Column**, and
      **Win/Loss**.
    - When inserted in a range of cells, sparklines are grouped and can be
      ungrouped.
    - The Sparkline tab is visible when a sparkline or its group is
      selected and includes formatting options such as setting the color and
      identifying high and low values.