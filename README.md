# Multi_select_drop_down_with_multiple_sheets_integration

In the following code, We need to create these major functionalities ->
  1. Create a multi-select drop-down list
  2. Use the comma-separated values present in the cells of a column, as source of the drop-down list
      (comma separated values in a cell will be converted to a list, and will be used as source for the drop-down list)
  3. Fetch these comma-separated values from other sheet, based on matching of an ID field in both the sheets.
      (i.e., Based on value in a column ,say H, in the drop-down-list sheet, we check the value matching with that in the
       source sheet's Id column, say, K, and then use the comma-separated values in the corresponding row to fill the drop-down list).

# ExcelMacros

## Files ->
  
  # worksheet_change_for_drop_down_list_sheet (Sheet 5) -> 
    - Source Sheet (Sheet 4) -> Source ID column (B) and Source Comma-separated values Column (H)
    - Target Sheet (Sheet 5) -> Target ID column (A) and Target Drop-down list Column (C)
    - The file serves the purpose of auto updating the lists and its values.
    - This file, performs mainly 2 functions ->
      * Update the drop-down list options of corresponding cell in the column C of Sheet 5; for any change in the target ID column of a cell in the column A.
      * Matches the text in the col A, with the values present in the Column B of the Sheet 4, and updates the drop-down list values of the column C.
    - Each drop-down list has an option of 'Clear' which removes all the values present in the corresponding cell.
    - It handles the case, where We can add new values in the column A and it automatically updates the corresponding column in the drop-down column C, based on the ID match.

  # worksheet_change_for_source_sheet (Sheet 4) -> 
    - Source Sheet (Sheet 4) -> Source ID column (B) and Source Comma-separated values Column (H)
    - Target Sheet (Sheet 5) -> Target ID column (A) and Target Drop-down list Column (C)
    - The file serves the purpose of auto updating the lists and its values.
    - This file, ->
      * Updates the drop-down-list values in Sheet 5 (column C), based on the changes in the comma-separated values in Sheet 4 (Column H).
    
  # split_values_to_list ->
    - This file contains the helper function, SplitValuesToList, which splits the string (having comma-separated values) passed as inputString and converts it 
        to a list of values; with an additional value of 'Clear'

  # update_drop_down_lists ->
    - This file works on the Existing values present in the Excel sheet.
    - While using Excel, We may have the case, where, We have data already present, and we need to add the functionality of multi-select dropdown to the sheet.
    - Running this macro file, updates/creates the drop-down list in the column B, based on the existing values present in the cells of column A.

# IMPORTANT NOTE -> 
  - **worksheet_change_for_drop_down_list_sheet** and **worksheet_change_for_source_sheet** files, are important for handling the updates that will be happening in the file columns.
  - **update_drop_down_lists** file,is important for handling the drop-down list for existing values in the columns.
  

## Steps to create a multi select drop down list without repeatition in a column, based on comma separated values present in another column ->
    eg -> 

          Sheet 4                                              Sheet 5
    Col B         Col H                                 Col A            Col C
     ABC          X,Y,Z                                 ABC              X,Y ( X,Y,Z,Clear - Drop-down list values)
     GHK          P,Q,R,S                               ABC              X ( X,Y,Z,Clear - Drop-down list values)
                                                        GHK              P,Q,R  ( P,Q,R,S,Clear - Drop-down list values)

    Steps ->
        Step 1: Press Alt + F11 to open the VBA Editor.
        Step 2: Go to Insert > Module to insert a new module.
        Step 3: Copy and paste the code of the "**split_values_to_list**" file into the module window.
        Step 4: Go to Insert > Module to insert a new module.
        Step 5: Copy and paste the code of the "**update_drop_down_lists**" file into the module window.
        Step 6: In the Project Explorer window, find your worksheet name under "Microsoft Excel Objects" (e.g., "Sheet12 (Hello World)").
        Step 7: Double-click on the worksheet name where, you want the drop-down list to be present, to open the code window for that worksheet.
        Step 8: Under the (General) tab, select "Worksheet" and under the (Declarations) tab, select "Change".
        Step 9: Copy and paste the code of the "**worksheet_change_for_drop_down_list_sheet**" file, into the window.
        Step 10: This file will serve on the Sheet on which, You want the drop-down list to be present
        Step 11: Double-click on the worksheet name where, you want the source data to be present, to open the code window for that worksheet.
        Step 12: Under the (General) tab, select "Worksheet" and under the (Declarations) tab, select "Change".
        Step 13: Copy and paste the code of the "**worksheet_change_for_source_sheet**" file, into the window.
        Step 14: This file will serve on the Sheet on which, You want the source data to be present
        Step 15: Run the Macro, update_drop_down_lists ("UpdateDropDownLists"), by pressing Alt + F8 and selecting the macro.
        
# IMPORTANT POINTS ->
  - In the code, I have used column 'B'(source Id) and 'H'(comma-separated values) as my Source and column 'A'(Target Id) and 'C'(Drop-down-list) as my Target Column.
  - Also, Do update the Sheet name, based on your sheet, I have used "Sheet4" for my Source Sheet and "Sheet5" for my Target Sheet.
  - You can find your sheet name in the VBA editor, in the Project Explorer window, (say, "Sheet1 (World)"), then you have to use "Sheet1" as your sheet name.
