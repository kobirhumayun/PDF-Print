Attribute VB_Name = "Main"
Option Explicit

 Sub PDF_Print()
Attribute PDF_Print.VB_ProcData.VB_Invoke_Func = "P\n14"
     
    Dim filePathArr As Variant
'    filePathArr = Application.Run("utility_functions.returnSelectedFilesFullPathArr", "G:\PDL Customs\Export LC, Import LC & UP\Import LC With Related Doc\YEAR-2023")  ' initial path
    filePathArr = Application.Run("utility_functions.returnSelectedFilesFullPathArr", "")  ' no initial path
    
    Dim filePath As Variant
    
    Dim i As Long
  
    For i = LBound(filePathArr) To UBound(filePathArr)
        
        filePath = filePathArr(i)
        
        Application.Run "utility_functions.printPdfUsingAdobeSdk", filePath ' print current at file
                
    Next i
           
    MsgBox "Printing Process Completed"
    
 End Sub
