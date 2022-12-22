# Excel-Excel
Excel-Excel organizes a group of consistently formatted excel workbooks into a single workbook. Each set of data from each of the consistently formatted workbooks will correspond to a single row at the destination.

# How it works 
 
## Startup Page

Currently when the app opens there a single button with a "+". In the future, there will also be options to open saved configuration. For now, this button will move the user to the next page.

![StartPage](https://github.com/yzjnxsantiago/Excel-Excel/blob/main/Images/StartPage.png)

## File Setup Page

The next page will be used to obtain the source directory and destination excel spreadsheet. The source directory is the folder where a set of consistently formatted excel workbooks (e.g. Applications) but with different information. All the information will be organized onto a destination spreadsheet.

![FileSetup](https://github.com/yzjnxsantiago/Excel-Excel/blob/main/Images/FileSetup.png)

## Sheet Setup Page

The sheet setup page is used to select which sheets will be used to transfer data. When moving to the Cell to Column page, the checked sheets will be the currently selected set of sheets.

![SheetSetup](https://github.com/yzjnxsantiago/Excel-Excel/blob/main/Images/SheetSelection.png)

## Cell to Column Page

The cell to column page is used to select the  cells of the currently selected set of sheets and configure the destination column each of these cells. To do this a cell must be created by typing the name of the cell in the entry box and clicking the button "Confirm Cell". After this a cell will appear in the label frame to the right of this where the user can drag and drop to the columns. Clicking Next Sheets will finish the set.

![CellSelect](https://github.com/yzjnxsantiago/Excel-Excel/blob/main/Images/CellSelection.png)

## Finish

Once the user clicks finish, the program will move to a loading page which uses threading to update the app (Loading., Loading.., Loading...) while also working on the task. Here is an example of the final result loaded onto the test destination workbook.

![FinalResult](https://github.com/yzjnxsantiago/Excel-Excel/blob/main/Images/FinalResult.png)

Note: Sheet Validation is still being developed. The purpose of this is so that if an application has information about which sheets are completed, the program will be able to use this information to only move the information from the completed sheets. The menu is also currently being developed.
