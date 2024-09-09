Excel Table Properties
1. Did not expect that _pk_ would not work with tables. (NewFunctions.script line 21)
2. Still looking for method to iterate over table value, specifically when starting with a record other than 1. (NewFunctions.script line 34)
3. Changing columnNames does not change the key values (NewFunctions.script line 50)

AdditionalWorkbookActions
1. When setting a value for tabColor, whether by color name or by number, the value returned is one less than the value set (e.g. white --> 254,254,254) (SetTabColors.script lines 9-19)
2. When returning a value for a tab that is not colored (e.g. a new worksheet), the value returned is (255,0,0). After deleting worksheet.tabColor, the value returned is (0,0,0). (ResetTabColors.script).
3. copyFrom does not work as expected. Placing the worksheet object in a variable and then using copyFrom:variable throws a "must be a worksheet object error."
4. Hyperlinks property incorrectly appends two additional backslashes for hyperlinks that reference file systems in UNC format. (ReturnHyperlinks.script line 13)
5. Placing workbook objects in a list (e.g. put [workbook(ResourcePath(filename1)),workbook(ResourcePath(filename2))] into myHyperlinksExcels) breaks the objects.
6. Unable to determine the syntax that will return the hyperlink of a single cell.
7. Returning the value of a cell which contains a hyperlink only returns "HTTPS." (ReturnHyperlinks.script line 11)
