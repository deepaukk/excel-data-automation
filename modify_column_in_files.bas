Sub ModifyColumnInFiles()

    Dim sourceFolder As String
    Dim destinationFolder As String
    Dim fileName As String
    Dim workbook As Workbook
    Dim lastTwoDigits As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Source Folder"

        If .Show = -1 Then
            sourceFolder = .SelectedItems(1) & "\"
        Else
            Exit Sub
        End If
    End With

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Destination Folder"

        If .Show = -1 Then
            destinationFolder = .SelectedItems(1) & "\"
        Else
            Exit Sub
        End If
    End With

    fileName = Dir(sourceFolder & "*.xlsx")

    Do While fileName <> ""

        Set workbook = Workbooks.Open(sourceFolder & fileName)

        lastTwoDigits = Right(Left(fileName, InStrRev(fileName, ".") - 1), 2)

        workbook.ActiveSheet.Range("B2:B431").Value = "100100" & lastTwoDigits

        workbook.SaveAs destinationFolder & fileName
        workbook.Close SaveChanges:=False

        fileName = Dir

    Loop

    MsgBox "Processing completed."

End Sub
