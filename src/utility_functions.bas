Attribute VB_Name = "utility_functions"
Option Explicit


Private Function printPdfUsingAdobeSdk(ByVal filePath As String)
    Dim acroApp As acroApp
    Dim avDoc As AcroAVDoc
    Dim pdDoc As AcroPDDoc

    Set acroApp = New acroApp
    Set avDoc = New AcroAVDoc
    Set pdDoc = New AcroPDDoc

    Dim methodeReturn As Variant
    '  Dim avDoc As Object
    ' Set avDoc = CreateObject("AcroExch.AVDoc")
    
    methodeReturn = acroApp.Hide() ' this methode must call bfore call "Exit()" methode

    If avDoc.Open(filePath, "") Then

        Set pdDoc = avDoc.GetPDDoc()
        methodeReturn = avDoc.PrintPagesSilent(0, pdDoc.GetNumPages - 1, 2, 0, 0)
        
    End If
    
    methodeReturn = acroApp.CloseAllDocs() ' this methode must call bfore call "Exit()" methode
    
    methodeReturn = acroApp.Exit()
    
    Debug.Print "printing by acroApp"
    ' methodeReturn = acroApp.Show() ' if this methode call "Exit()" methode not work

End Function


Private Function returnSelectedFilesFullPathArr(ByVal initialPath As String) As Variant
    Dim fileDialog As Object
    Dim selectedFiles As Variant
    Dim i As Long
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "Select Files"
        .AllowMultiSelect = True
         .InitialFileName = initialPath
        If .Show = -1 Then
            ReDim selectedFiles(1 To .SelectedItems.Count)
            For i = 1 To .SelectedItems.Count
                selectedFiles(i) = .SelectedItems.Item(i)
            Next i
        End If
    End With

    returnSelectedFilesFullPathArr = selectedFiles
End Function
