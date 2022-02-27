# Excel For Data Science and Analysis

In our day-to-day tasks, we use excel in performing almost all tasks whether it is collecting data, storing data, analysing data, or data cleaning. It is the primary knowledge that should be expected from a person expertizing in the field of Data Science. It is so simple to use and to store the results to provide a basic statistics of the given data.

## Power Queries

Power Query is a business intelligence tool available in Excel that allows you to import data from many different sources and then clean, transform and reshape your data as needed. It allows you to set up a query once and then reuse it with a simple refresh. It's also pretty powerful.

***Step 1-*** *Importing the files from the folder on your PC (files can be imported from Web or a Database as well).*

![Img1](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img1.PNG)

***Step 2-*** *Browse the folder where the excel files are stored.*

![Img2](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img2.PNG)

***Step 3-*** *Click on OK after selecting the folder path to the CSV files.*

![Img3](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img3.PNG)

***Step 4-*** *A window will appear showing the data. Click on Combine & Edit to combine all the excel files in that folder.* 

![Img4](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img4.PNG)

*A new window appears with all the data of the combined CSV files.*

![Img5](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img5.PNG)

### Performing functions on an excel file is known as Data Cleaning.

***Step 5-*** *Performing operationas on the Data can be performed. To remove a column from the CSV file simply select the column from the header. Right click and and Remove.*

![Img6](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img6.PNG)

***Step 6-*** *Merging columns is done by selecting the two columns and and Right click. Go to the Merge Columns.*

![Img7](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img7.PNG)

*Choose a separator. It could be anything you need to use as a delimiter for separation between the two column. In case of a certain separator not available select "custom" from the drop down menu and choose the selector of your choice. Select the "New Column name" and press OK.*

![Img8](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img8.PNG)

*The two merged columns appears to be something like this.*

![Img9](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img9.PNG)

***Step 6-*** *One column can be splitted into two. To perform this task simply choose the column and Right Click. In the Home tab selct "Split Column" and choose "By delimiter". This will split the column by the provided delimiter.* 

![Img10](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img10.PNG)

*This operation gives the following result after it is performed.*

![Img11](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img11.PNG)

***Step 7-*** *The data type of a column can also be changed. Select the column and Right Click. Choose the data type for the given column.*

![Img12](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img12.PNG)

***Step 8-*** *Once the Data Cleaning has been performed, go to the Home tab and click on Close & Load.*

![Img13](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img13.PNG)

*A new Excel Sheet after Data Cleaning is visible. Save this file for later purpose. The later added files to the same folder performs the Data Cleaning automatically.*

![Img14](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img14.PNG)

## Formulas 

Apart from using Power Queries, there are wide range of formulas to perform operations in the data of the Excel File. Just navigate to the Formulas and click "Insert Function".

![Img15](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img15.jpg)

1. CONCATENATE
<br>Combine the values of several cells into one
<br>Formula: Combine the values of several cells into one.

![Img16](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img16.jpg)

2. VLOOKUP
<br>The formula allows you to look up data that is arranged in vertical columns.
<br>Formula: =VLOOKUP(LOOKUP_VALUE,TABLE_ARRAY, COL_INDEX_NUM, [RANGE_LOOKUP])

![Img17](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img17.jpg)

3. LEN
<br>To get the number of characters in a given cell.
<br>Formula: =LEN(SELECT CELL)

![Img18](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img18.jpg)

4. SUMIF
<br>It adds up the values in cells which meet a selected number.
<br>Formula: =SUMIF(RANGE,CRITERIA,[sum_range])

![Img19](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img19.jpg)

5. DAYS/NETWORKDAYS
<br>To determine the number of days between two calendar dates.
<br>Formula: =NETWORKDAYS(SELECT CELL, SELECT CELL,[numberofholidays])

![Img20](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img20.jpg)

<br>Formula: =NETWORKDAYS(SELECT CELL, SELECT CELL,[numberofholidays])

![Img21](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img21.jpg)

6. SUBSTITUTE
<br>Replacing cells in bulks 
<br>Formula: =SUBSTITUTE(A1,"p","s")

![Img22](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img22.jpg)

7. MINIF/MAXIF
<br>Minimum of a set of values, and match on criteria.
<br>Formula: =MIN(IF(RANGE1,CRITERIA1,RANGE2))

![Img23](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img23.jpg)

<br>Maximum of a set of values, and match on criteria.
<br>Formula: =MAX(IF(RANGE1,CRITERIA1,RANGE2))

![Img24](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img24.jpg)

8. COUNTIFS
<br>Counts the numbers how many times a value appears based on one criteria.
<br>Formula: =COUNTIFS(RANGE,CRITERIA)

![Img25](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img25.jpg)

9. LEFT/RIGHT
<br>will return the “x” number of characters from the beginning of the cell.
<br>Formula: =LEFT(SELECT CELL,NUMBER)

![Img26](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img26.jpg)

<br>will return the “x” number of characters from the end of the cell.
<br>Formula: =RIGHT(SELECT CELL,NUMBER)

![Img27](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img27.jpg)

## Pivot Tables

### Creating Pivot Table

Select the table or the rows and columns. Go to the insert tab and click on the PivotTable. 

![Img28](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img28.jpg)

The range of the table is selected. Now choose a new worksheet or the existing worksheet to work with the Pivot Tables. Press OK.

![Img29](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img29.jpg)

The extreme right section of the new worksheet created contains all the column names of the table. On the buttom are four fields for the formatting of the Pivot Table according to your choice.

![Img30](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img30.jpg)

Dragging UnitCost to the VALUES section will give the sum of the total unit costs present in the table.

![Img31](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img31.jpg)

Dragging the items to the ROWS with show all the items in a row. Choosing both items to the rows and UnitCost to the VALUES with show the items in a row along with sum of the total of the each item and the grand total.

![Img32](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img32.jpg)

The Pivot Table for the 5 fields appears to be like this.

![Img33](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img33.jpg)

A chart of this new worksheet can be formed. Go to the insert tab and click on Recommended Charts. Choose the type of chart you want to form. 

![Img44](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img34.jpg)

![Img45](https://github.com/ananyachibber21/Excel-For-Data-Science-and-Analysis/blob/main/Screenshots/Img35.jpg)
