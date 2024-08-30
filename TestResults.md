Excel Table Properties
1. Did not expect that _pk_ would not work with tables. (NewFunctions.script line 21)
2. Still looking for method to iterate over table value, specifically when starting with a record other than 1. (NewFunctions.script line 34)
3. Changing columnNames does not change the key values (NewFunctions.script line 50)

AdditionalWorkbookActions
1. When setting a value for tabColor, whether by color name or by number, the value returned is one less than the value set (e.g. white --> 254,254,254) (SetTabColors.script lines 9-19)
2. When returning a value for a tab that is not colored (e.g. a new worksheet), the value returned is (255,0,0). After deleting worksheet.tabColor, the value returned is (0,0,0). (ResetTabColors.script).
3. 
