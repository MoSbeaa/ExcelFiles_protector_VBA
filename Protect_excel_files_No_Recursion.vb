' Declare a global variable to store the main folder path
Global mainFolder As Variant
Sub StartProtecting()
    Call getData
    
    Call gettingCells
    
    Call makeLoop
    
    ' Display a message box indicating that the operation is complete
    MsgBox "All Done"
End Sub
'It prompts the user to select a folder and then lists the names of
'all subfolders in the selected folder in column A of the active worksheet.
'If there are no subfolders in the selected folder, a message box is displayed to inform the user
Sub getData()
    ' Clear columns A to C
    Columns("A:C").Clear

    ' Declare variables
    Dim Cell As Range
    Dim folder As Variant
    Dim SubFolders As Variant
    Dim vArray As Variant

    ' Set the starting cell for output
    Set Cell = Range("A1")

    ' Use the FileDialog object to prompt the user to select a folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the main folder that contain all the subfolders"
        If .Show Then
            folder = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With

    ' Store the selected folder path in the global variable
    mainFolder = folder

    ' Use the Shell.Application object to get a collection of subfolders in the selected folder
    With CreateObject("Shell.Application")
        Set SubFolders = .Namespace(folder).Items
        SubFolders.Filter 32, "*"
    End With

    ' Check if there are any subfolders in the selected folder
    If SubFolders.Count = 0 Then
        MsgBox "There are No Subfolders in this Directory."
        Exit Sub
    End If

    ' Resize the vArray array to hold the names of the subfolders
    ReDim vArray(1 To SubFolders.Count, 1 To 1)

    ' Loop through the subfolders and store their names in the vArray array
    For n = 0 To SubFolders.Count - 1
        vArray(n + 1, 1) = SubFolders.Item(n).Name
    Next n

    ' Output the names of the subfolders to the worksheet starting at cell A1
    With Cell.Resize(n, 1)
        .NumberFormat = "0"
        .Value = vArray
    End With



End Sub
'It prompts the user to select an Excel file and then copies data from columns A and B of the selected file
'into columns B and C of the active worksheet in the current workbook.
'If there is a mismatch between the folder names in column A and column B,
'a message box is displayed to inform the user and columns A to C are cleared.
Sub gettingCells()
    ' Declare variables
    Dim fd As Office.FileDialog
    Dim strFile As String

    ' Create a FileDialog object to prompt the user to select an Excel file
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        ' Clear any existing filters and add a filter for Excel files
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx?", 1

        ' Set the dialog title and disallow multiple file selection
        .Title = "Choose The Passwords File"
        .AllowMultiSelect = False

        ' Show the dialog and get the selected file path
        If .Show = True Then
            strFile = .SelectedItems(1)
        Else
            ' If the user cancels the dialog, clear columns A to C and display a message box
            Columns("A:C").Clear
            MsgBox "Process Canceled"
            Exit Sub
        End If
    End With

    ' Open the selected Excel file and copy data from column B (passwords)
    Workbooks.Open Filename:=strFile
    ActiveSheet.Range("B2:B1000").Select
    Selection.Copy

    ' Paste the copied data into column C of the active sheet in the current workbook
    Windows(ThisWorkbook.Name).Activate
    ActiveSheet.Columns("C:C").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    ' Open the selected Excel file again and copy data from column A (folder names)
    Workbooks.Open Filename:=strFile
    ActiveSheet.Range("A2:A1000").Select
    Selection.Copy

    ' Paste the copied data into column B of the active sheet in the current workbook
    Windows(ThisWorkbook.Name).Activate
    ActiveSheet.Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    ' Close the opened Excel file without saving changes
    Workbooks.Open Filename:=strFile
    ActiveWindow.Close

    ' Check if the folder names in column B match those in column A
    For i = 1 To Rows.Count
        If Cells(i, 1).Value <> Cells(i, 2).Value Then
            ' If there is a mismatch, display a message box and clear columns A to C
            MsgBox "Either the folder / " & Cells(i, 2).Value & " / is not exist in the directory. Or the folder / " & Cells(i, 1).Value & " /is not listed in the excel file,. Please check again"
            Columns("A:C").Clear
            Exit Sub
        End If
    Next i

End Sub
' Subroutine to loop through rows in a table and call the addPassword subroutine for each row
Sub makeLoop()
    ' Activate the current workbook
    Windows(ThisWorkbook.Name).Activate

    ' Declare a Range variable to represent the table
    Dim table As Range

    ' Set the table range to include all cells from A1 to the last cell in the used range
    Set table = Range("A1", Range("A1").End(xlToRight).End(xlDown))

    ' Loop through each row in the table
    For Row = 1 To table.Rows.Count
        ' Check if the first cell in the row is not empty
        If table(Row, 1).Value <> "" Then
            ' Call the addPassword subroutine with arguments from columns A and C of the current row
            Call addPassword(Cells(Row, 1).Value, Cells(Row, 3).Value)
        End If
    Next Row

    ' Clear columns A to C
    Columns("A:C").Clear
End Sub
' Subroutine to add a password to Excel files in a specified folder and its subfolders
Sub addPassword(folderName As Variant, pswd As String)
    ' Declare variables
    Dim FSO As Object
    Dim folder As Object, subfolder As Object
    Dim wb As Object

    ' Create a FileSystemObject to work with the file system
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ' Build the path to the specified folder using the global mainFolder variable
    folderPath = mainFolder & "\" & folderName & "\"

    ' Get a reference to the specified folder
    Set folder = FSO.GetFolder(folderPath)

    ' Disable various Excel application settings to speed up processing
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
        .EnableEvents = False
        .AskToUpdateLinks = False
    End With

    ' Loop through each file in the specified folder
    For Each wb In folder.Files
        ' Check if the file is an Excel file
        If Right(wb.Name, 3) = "xls" Or Right(wb.Name, 4) = "xlsx" Or Right(wb.Name, 4) = "xlsm" Then
            ' Attempt to open the file with an incorrect password to check if it is already password protected
            On Error Resume Next
            Workbooks.Open wb, , , , "daafdsfafasfff"
            If Err.Number > 0 Then
                ' If an error occurs (i.e. the file is password protected), skip to the next file
                GoTo 25
            End If

            ' Open the file without a password and save it with the specified password
            Set masterWB = Workbooks.Open(wb)
            ActiveWorkbook.SaveAs Filename:=Application.ActiveWorkbook.FullName, password:=pswd
            ActiveWorkbook.Close True
        End If

        ' Label for skipping to the next file if an error occurs when opening a password protected file
25 Next

' Loop through each subfolder in the specified folder (1 level deep)
    For Each subfolder In folder.SubFolders
        ' Loop through each file in the subfolder
        For Each wb In subfolder.Files
            ' Check if the file is an Excel file and save it with the specified password if it is not already password protected (same as above)
            If Right(wb.Name, 3) = "xls" Or Right(wb.Name, 4) = "xlsx" Or Right(wb.Name, 4) = "xlsm" Then
                Set masterWB = Workbooks.Open(wb)
                ActiveWorkbook.SaveAs Filename:=Application.ActiveWorkbook.FullName, password:=pswd
                ActiveWorkbook.Close True
            End If

        Next
        
    ' Loop through each subfolder in the specified folder (2 level deep)
    For Each subFolder2 In subfolder.SubFolders
       For Each wb In subFolder2.Files
           If Right(wb.Name, 3) = "xls" Or Right(wb.Name, 4) = "xlsx" Or Right(wb.Name, 4) = "xlsm" Then
               Set masterWB = Workbooks.Open(wb)
               ActiveWorkbook.SaveAs Filename:=Application.ActiveWorkbook.FullName, password:=pswd
               ActiveWorkbook.Close True
           End If
        Next
        ' Loop through each subfolder in the specified folder (3 level deep)
        For Each subfolder3 In subFolder2.SubFolders
            For Each wb In subfolder3.Files
                If Right(wb.Name, 3) = "xls" Or Right(wb.Name, 4) = "xlsx" Or Right(wb.Name, 4) = "xlsm" Then
                    Set masterWB = Workbooks.Open(wb)
                    ActiveWorkbook.SaveAs Filename:=Application.ActiveWorkbook.FullName, password:=pswd
                    ActiveWorkbook.Close True
                End If
            Next
           ' Loop through each subfolder in the specified folder (4 level deep)
            For Each subfolder4 In subfolder3.SubFolders
               For Each wb In subfolder4.Files
                   If Right(wb.Name, 3) = "xls" Or Right(wb.Name, 4) = "xlsx" Or Right(wb.Name, 4) = "xlsm" Then
                       Set masterWB = Workbooks.Open(wb)
                       ActiveWorkbook.SaveAs Filename:=Application.ActiveWorkbook.FullName, password:=pswd
                       ActiveWorkbook.Close True
                   End If
               Next
               ' Loop through each subfolder in the specified folder (5 level deep)
                For Each subfolder5 In subfolder4.SubFolders
                   For Each wb In subfolder5.Files
                       If Right(wb.Name, 3) = "xls" Or Right(wb.Name, 4) = "xlsx" Or Right(wb.Name, 4) = "xlsm" Then
                           Set masterWB = Workbooks.Open(wb)
                           ActiveWorkbook.SaveAs Filename:=Application.ActiveWorkbook.FullName, password:=pswd
                           ActiveWorkbook.Close True
                       End If
                   Next
                   ' Loop through each subfolder in the specified folder (6 level deep)
                    For Each subfolder6 In subfolder5.SubFolders
                       For Each wb In subfolder6.Files
                           If Right(wb.Name, 3) = "xls" Or Right(wb.Name, 4) = "xlsx" Or Right(wb.Name, 4) = "xlsm" Then
                               Set masterWB = Workbooks.Open(wb)
                               ActiveWorkbook.SaveAs Filename:=Application.ActiveWorkbook.FullName, password:=pswd
                               ActiveWorkbook.Close True
                           End If
                       Next
                       ' Loop through each subfolder in the specified folder (7 level deep)
                        For Each subfolder7 In subfolder6.SubFolders
                           For Each wb In subfolder7.Files
                               If Right(wb.Name, 3) = "xls" Or Right(wb.Name, 4) = "xlsx" Or Right(wb.Name, 4) = "xlsm" Then
                                   Set masterWB = Workbooks.Open(wb)
                                   ActiveWorkbook.SaveAs Filename:=Application.ActiveWorkbook.FullName, password:=pswd
                                   ActiveWorkbook.Close True
                               End If
                           Next
                           ' Loop through each subfolder in the specified folder (8 level deep)
                            For Each subfolder8 In subfolder7.SubFolders
                               For Each wb In subfolder8.Files
                                   If Right(wb.Name, 3) = "xls" Or Right(wb.Name, 4) = "xlsx" Or Right(wb.Name, 4) = "xlsm" Then
                                       Set masterWB = Workbooks.Open(wb)
                                       ActiveWorkbook.SaveAs Filename:=Application.ActiveWorkbook.FullName, password:=pswd
                                       ActiveWorkbook.Close True
                                   End If
                               Next
                            
                            ' Loop through each subfolder in the specified folder (9 level deep)
                            'For Each subfolder9 In subfolder8.SubFolders
                            '   For Each wb In subfolder9.Files
                             '      If Right(wb.Name, 3) = "xls" Or Right(wb.Name, 4) = "xlsx" Or Right(wb.Name, 4) = "xlsm" Then
                              '         Set masterWB = Workbooks.Open(wb)
                               '        ActiveWorkbook.SaveAs Filename:=Application.ActiveWorkbook.FullName, password:=pswd
                                '       ActiveWorkbook.Close True
                                 '  End If
                              ' Next
                           ' Next
                            
                            
                            Next
                            
                        Next
                    
                    Next
                
                
                Next
            
        Next
            
            
     Next
        
        
Next
 
       
Next
        
        



End Sub
