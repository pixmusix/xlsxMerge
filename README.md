# xlsxMerge
A simple user interface for merging two datasets into one dataset. *xlsxMerge* performs a fast Full Outer Joins on two sheets from an Excel workbook. Perfect for a small buisness!

## When would you use xlsxMerge

*xlsxMerge* is useful when you have two sets of data from different sources which you would like to combine into one Dataset in Excel.
- For Example: Merging "Supplier PriceList" into your "Inventory" based on a "SKU". 
- For Example: Merging "User Responses" into your "User Database" based on "UserID". 
- For Example: Merging "Google Maps Data" into "Local Street Directory" based on "PostCode" 

## Quick Start

- [Download Microsofts Visual Studio](https://visualstudio.microsoft.com/vs/)
- Open the solution in Visual Studio. (The file called "xlsxMerge.sln")
- [Set the project to release mode](https://learn.microsoft.com/en-us/visualstudio/debugger/how-to-set-debug-and-release-configurations?view=vs-2022)
- In the menubar, under *Build*, select *Build Solution* (ctl+sft+b)
- Locate xlsxMerge.exe in your projects director under ./bin/release/
- Run xlsxMerge.exe

## What is a Full Outer Join

![imb](https://upload.wikimedia.org/wikipedia/commons/thumb/3/3d/SQL_Join_-_05b_A_Full_Join_B.svg/330px-SQL_Join_-_05b_A_Full_Join_B.svg.png)

To oversimplify, for each cell in ColumnA, *Full Outer Joins* will search ColumnB for matching cells. 
If a match is found, a *Full Outer Join* merges the rows where the match was found into one big row.
If a row is not found, the row will be left in place and *Nulls* will fill the gaps.
Importantly for small buisness applications a Full Join is lossless, meaning that data that failed to match is preseved and added to the output sheet.

##### Left
|      | LeftKey | LeftValue |
|------|---------|-----------|
| Row1 | Foo     | 4         |
| Row2 | Bar     | 5         |

##### Right

|      | RightKey | RightValue |
|------|----------|------------|
| Row1 | Foo      | 15         |
| Row2 | Jip      | 12         |

##### Out

|                | LeftKey | LeftValue | RightKey | RightValue |
|----------------|---------|-----------|----------|------------|
| MatchedRow     | Foo     | 4         | Foo      | 15         |
| UnmatchedLeft  | Bar     | 5         | Null     | Null       |
| UnmatchedRight | Null    | Null      | Jip      | 12         |

## User Interface.

![ima](xlsxMergeUI.jpg)

#### Tutorial

1. With your data prepared in an excel spreadsheet, save and close the workbook.
2. Open *xlsxMerge* and click *Load*, then select and open your XLSX file.
3. In the dropdowns that appear, select which sheets you would like to merge (note, you can select the same sheet twice).
4. For both sheets, choose which column contains the primary keys for that data, i.e. the values you would like to match up. A key on the left will match with the first instance of the same key on the right. xlsxMerge is greedy and wants an exact match.
5. For both sheets, give which row you would like to start matching from (inclusive). e.g. If you have row headers in row 1 of your xlsx sheet, you may give row 2 as your starting row (ignoring the header row).
6. Confirm that you are happy with the result in the feedback window to the right.
7. Click *Save*. This will save your output in your local directory.

## Further Reading

- [I strongly suggest learning the =INDEX(MATCH)) Functions in vanilla excel](https://www.youtube.com/watch?v=F264FpBDX28)
- [See also this excellent python libray for reading and writing in excel](https://openpyxl.readthedocs.io/en/stable/)

## Dependencies

- [Microsoft.Office.Interop.Excel](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel?view=excel-pia)
