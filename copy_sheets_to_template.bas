Sub CopySheetsToTemplate()

    Dim sourceFolder As String
    Dim outputFolder As String
    Dim templatePath As String
    Dim fileName As String

    Dim sourceWb As Workbook
    Dim templateWb As Workbook
    Dim sourceWs As Worksheet
    Dim newWs As Worksheet

    Dim outputFileName As String

    ' Select source folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Source Folder"

        If .Show <> -1 Then Exit Sub

        sourceFolder = .SelectedItems(1) & "\"
    End With

    ' Select output folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Output Folder"

        If .Show <> -1 Then Exit Sub

        outputFolder = .SelectedItems(1) & "\"
    End With

    ' Select template workbook
    With Application.FileDialog(msoFileDialogFilePicker)

        .Title = "Select Template Workbook"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xlsm"

        If .Show <> -1 Then Exit Sub

        templatePath = .SelectedItems(1)

    End With

    ' Process all Excel files in the source folder
    fileName = Dir(sourceFolder & "*.xlsx")

    Do While fileName <> ""

        Set sourceWb = Workbooks.Open(sourceFolder & fileName)
        Set templateWb = Workbooks.Open(templatePath)

        ' Copy each worksheet from the source workbook
        For Each sourceWs In sourceWb.Worksheets

            sourceWs.Copy After:=templateWb.Sheets(templateWb.Sheets.Count)

            Set newWs = templateWb.Sheets(templateWb.Sheets.Count)

            ' Preserve the worksheet name where possible
            On Error Resume Next
            newWs.Name = sourceWs.Name
            On Error GoTo 0

        Next sourceWs

        ' Create output filename
        outputFileName = outputFolder & _
                         Left(fileName, InStrRev(fileName, ".") - 1) & _
                         "_Output.xlsx"

        Application.DisplayAlerts = False

        templateWb.SaveAs outputFileName, _
                          FileFormat:=xlOpenXMLWorkbook

        Application.DisplayAlerts = True

        templateWb.Close SaveChanges:=False
        sourceWb.Close SaveChanges:=False

        fileName = Dir

    Loop

    MsgBox "Sheets copied successfully."

End Sub
