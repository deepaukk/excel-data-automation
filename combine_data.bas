Sub CombineData()

    Dim sourceFolder As String
    Dim outputFolder As String
    Dim templatePath As String
    Dim fileName As String
    Dim sourceWb As Workbook
    Dim templateWb As Workbook
    Dim sourceWs As Worksheet
    Dim templateWs As Worksheet
    Dim lastRow As Long
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
        Set sourceWs = sourceWb.Worksheets(1)

        ' Open a fresh copy of the template for each source file
        Set templateWb = Workbooks.Open(templatePath)
        Set templateWs = templateWb.Worksheets(1)

        ' Find the last row containing data
        lastRow = sourceWs.Cells(sourceWs.Rows.Count, "A").End(xlUp).Row

        ' Copy columns A:C from source workbook
        ' into columns C:E of the template
        If lastRow >= 1 Then

            templateWs.Range("C1").Resize(lastRow, 3).Value = _
                sourceWs.Range("A1:C" & lastRow).Value

        End If

        ' Create output filename
        outputFileName = outputFolder & _
                         Left(fileName, InStrRev(fileName, ".") - 1) & _
                         "_output.xlsx"

        ' Save output workbook
        Application.DisplayAlerts = False

        templateWb.SaveAs outputFileName, FileFormat:=xlOpenXMLWorkbook

        Application.DisplayAlerts = True

        templateWb.Close SaveChanges:=False
        sourceWb.Close SaveChanges:=False

        fileName = Dir

    Loop

    MsgBox "Data combination completed."

End Sub
