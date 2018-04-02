Attribute VB_Name = "Module1"
Sub spike_excel_r_macro_01()
Attribute spike_excel_r_macro_01.VB_Description = "spike_excel_r_macro_01"
Attribute spike_excel_r_macro_01.VB_ProcData.VB_Invoke_Func = "j\n14"

'Mustafa Alpay: Excel_Macro_Loop_Subdirectories_To_Get_Data

Dim wb As Workbook
Dim myPath As String
Dim myFile As String
Dim myExtension As String
Dim FldrPicker As FileDialog

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

'retrieve source folder with file dialog
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder(s)"
      '.AllowMultiSelect = False
      .AllowMultiSelect = True
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'if canceled
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'source excel files
  myExtension = "*.xls*"
  

'loop subfolders
Dim fso As Object
Dim folder As Object
Dim subfolders As Object
Set fso = CreateObject("Scripting.FileSystemObject")

Set folder = fso.GetFolder(myPath)
Set subfolders = folder.subfolders
Dim dirName As String
Dim WrdArray() As String

For Each subfolder In subfolders

    dirName = subfolder.Name
    WrdArray() = Split(dirName, "_")
    MsgBox subfolder.Path & "," & WrdArray(0) & "-" & WrdArray(1)


    myPath = subfolder.Path & "\"

    'target source file path
    myFile = Dir(myPath & myExtension)
    
    'control: MsgBox myPath & ", " & myFile

    'Loop through each Excel file in folder
    Do While myFile <> ""
        'Set variable equal to opened workbook
          Set wb = Workbooks.Open(Filename:=myPath & myFile)
        
        'Ensure Workbook has opened before moving on to next line of code
          DoEvents
        
        'Change First Worksheet's Background Fill Blue
        'wb.Worksheets(1).Range("A1:Z1").Interior.Color = RGB(51, 98, 174)
          
        
        
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets.Add(After:= _
                 ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                 
        'sheet name
        ws.Name = WrdArray(0) & "-" & WrdArray(1)
        
        'predefined range
        ws.Range("A1:K19").Value = wb.Worksheets(7).Range("A1:K19").Value

        'save before closing workbook
          wb.Close SaveChanges:=True
          
        'double chehck if workbook has closed
          DoEvents
    
        'next file name
          myFile = Dir
    Loop
  
Next 'for each subfolders

'task complete message to the user
  MsgBox "Task Complete!"

ResetSettings:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
