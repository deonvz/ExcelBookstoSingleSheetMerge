Attribute VB_Name = "ExcelBookstoSingleSheetMerge"
Sub mergeFiles()
    'Merges all files in a folder to a main file.
    'Written by Deon van Zyl
    
    'Define variables:
    Dim numberOfFilesChosen, i As Integer
    Dim tempFileDialog As FileDialog
    Dim mainWorkbook, sourceWorkbook As Workbook
    Dim tempWorkSheet As Worksheet
    
    Set mainWorkbook = Application.ActiveWorkbook
    Set tempFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    'Allow the user to select multiple workbooks
    tempFileDialog.AllowMultiSelect = True
    
    numberOfFilesChosen = tempFileDialog.Show
    
    'Loop through all selected workbooks
    For i = 1 To tempFileDialog.SelectedItems.Count
        
        'Open each workbook
        Workbooks.Open tempFileDialog.SelectedItems(i)
        
        Set sourceWorkbook = ActiveWorkbook
        
        'Copy all workbooks data to the current sheet
        
        'change "A2" with cell reference of start point for every files here
        'for example "B3:IV" to merge all files start from columns B and rows 3
        'If you're files using more than IV column, change it to the latest column
        'Also change "A" column on "A65536" to the same column as start point
        Range("A2:IV" & Range("A65536").End(xlUp).Row).Copy
        ThisWorkbook.Worksheets(1).Activate
         
        'Do not change the following column. It's not the same column as above
        Range("A65536").End(xlUp).Offset(1, 0).PasteSpecial
        Application.CutCopyMode = False
        
        'Copy each worksheet to the end of the main workbook
        'For Each tempWorkSheet In sourceWorkbook.Worksheets
        '    tempWorkSheet.Copy after:=mainWorkbook.Sheets(mainWorkbook.Worksheets.Count)
        'Next tempWorkSheet
        
        'Close the source workbook
        sourceWorkbook.Close
    Next i
    
End Sub
