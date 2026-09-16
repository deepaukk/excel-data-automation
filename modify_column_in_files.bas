Sub ModifyColumnInFiles()

    Dim sourceFolder As String
    Dim destinationFolder As String
    Dim fileName As String
    Dim wb As Workbook
    Dim lastTwoDigits As String
    Dim fileBaseName As String

    ' Select source folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Source Folder"

        If .Show <> -1 Then Exit Sub

        sourceFolder = .SelectedItems(1) & "\"
    End With

    ' Select destination folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Destination Folder"

        If .Show <> -1 Then Exit Sub

        destinationFolder = .SelectedItems(1) & "\"
    End With

    ' Process all Excel files in the source folder
    fileName = Dir(sourceFolder & "*.xlsx")

    Do While fileName <> ""

        Set wb = Workbooks.Open(sourceFolder & fileName)

        ' Get the filename without the extension
        fileBaseName = Left(fileName, InStrRev(fileName, ".") - 1)

        ' Extract the last two characters from the filename
        lastTwoDigits = Right(fileBaseName, 2)

        ' Update the required range
        wb.ActiveSheet.Range("B2:B431").Value = "100100" & lastTwoDigits

        ' Save the processed workbook in the destination folder
        wb.SaveAs destinationFolder & fileName
        wb.Close SaveChanges:=False

        fileName = Dir

    Loop

    MsgBox "Processing completed."

End Sub
