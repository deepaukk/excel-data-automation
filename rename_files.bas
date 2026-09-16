Sub RenameFiles()

    Dim folderPath As String
    Dim fileName As String
    Dim newFileName As String

    Dim ws As Worksheet
    Dim matchResult As Variant
    Dim rowNumber As Long

    ' Select folder containing the files
    With Application.FileDialog(msoFileDialogFolderPicker)

        .Title = "Select Folder"

        If .Show <> -1 Then Exit Sub

        folderPath = .SelectedItems(1) & "\"

    End With

    ' Use the active worksheet as the reference list
    Set ws = ActiveSheet

    ' Process files in the selected folder
    fileName = Dir(folderPath & "*.*")

    Do While fileName <> ""

        ' Find the current filename in column A
        matchResult = Application.Match(fileName, ws.Columns("A"), 0)

        If Not IsError(matchResult) Then

            rowNumber = CLng(matchResult)

            ' Get the new filename from column B
            newFileName = Trim(ws.Cells(rowNumber, "B").Value)

            If newFileName <> "" Then

                ' Rename the file
                Name folderPath & fileName _
                    As folderPath & newFileName

            End If

        End If

        fileName = Dir

    Loop

    MsgBox "File renaming completed."

End Sub
