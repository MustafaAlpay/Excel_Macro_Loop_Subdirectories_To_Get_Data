Attribute VB_Name = "Module1"
Sub spike_excel_r_macro_01()
Attribute spike_excel_r_macro_01.VB_Description = "spike_excel_r_macro_01"
Attribute spike_excel_r_macro_01.VB_ProcData.VB_Invoke_Func = "j\n14"
'//todo==>1-loop sub-multiple directories 2-name:directory name _ile explode edecek!!!


'
' spike_excel_r_macro_01 Macro
' spike_excel_r_macro_01
'

'https://www.thespreadsheetguru.com/the-code-vault/2014/4/23/loop-through-all-excel-files-in-a-given-folder
'Sub LoopAllExcelFilesInFolder()
'PURPOSE: To loop through all Excel files in a user specified folder and perform a set task on them
'SOURCE: www.TheSpreadsheetGuru.com

Dim wb As Workbook
Dim myPath As String
Dim myFile As String
Dim myExtension As String
Dim FldrPicker As FileDialog

'Optimize Macro Speed
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual

'Retrieve Target Folder Path From User
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder(s)"
      '.AllowMultiSelect = False
      .AllowMultiSelect = True
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.xls*"
  'myExtension = "\*.xls*"
  

'loop subfolders


Dim fso As Object
Dim folder As Object
Dim subfolders As Object
Set fso = CreateObject("Scripting.FileSystemObject")

'Set folder = fso.GetFolder("C:\SAP\")
Set folder = fso.GetFolder(myPath)
Set subfolders = folder.subfolders
Dim dirName As String
Dim WrdArray() As String

For Each subfolder In subfolders
    'MsgBox subfolder.Name
    dirName = subfolder.Name
    WrdArray() = Split(dirName, "_")
    MsgBox subfolder.Path & "," & WrdArray(0) & "-" & WrdArray(1)


    myPath = subfolder.Path & "\"

    'Target Path with Ending Extention
    myFile = Dir(myPath & myExtension)
    
    MsgBox myPath & ", " & myFile
    
    'myFile = Dir(subfolder & myExtension)
    'myFile = Dir(subfolder.Path & myExtension)

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
        ws.Name = WrdArray(0) & "-" & WrdArray(1) 'wb.Name 'wb.Worksheets(2).Name
        
        'predefined range
        ws.Range("A1:K19").Value = wb.Worksheets(7).Range("A1:K19").Value
        
        
        
        
        
        'Save and Close Workbook
          wb.Close SaveChanges:=True
          
        'Ensure Workbook has closed before moving on to next line of code
          DoEvents
    
        'Get next file name
          myFile = Dir
    Loop
  
Next 'For Each subfolder In subfolders

'Message Box when tasks are completed
  MsgBox "Task Complete!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True


'
'
'
End Sub
