# Source code

## Module 1
```
#If VBA7 Then
    Declare PtrSafe Function SetCurrentDirectoryA Lib "KERNEL32" (ByVal lpPathName As String) As Long
#Else
    Declare Function SetCurrentDirectoryA Lib "KERNEL32" (ByVal lpPathName As String) As Long
#End If


Sub Import_previous_report()

'disable all the alerts
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False

'select the file sorce, may exit the program here
SetCurrentDirectoryA "\\maxhealth.com\filecabinet\Headquarters\Departments\AccountsPayable\Private\General\Audit_Team\Concur\Daily Expense Report Assignment"
PrevFile = Application.GetOpenFilename
If PrevFile = False Then
    End
End If

'open and copy content from the previous report
Range("F3").Value = PrevFile
Workbooks.Open (PrevFile)
Range("A1").CurrentRegion.Copy

'paste content to current template
Windows("Mary Template.xlsm").Activate
Sheets("Step1").Activate
Range("A6").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
Application.CutCopyMode = False 'clean the clipboard
Workbooks(2).Close 'close the previous report
Sheets("Step2").Activate

End Sub
```

## Module 2
```
#If VBA7 Then
    Declare PtrSafe Function SetCurrentDirectoryA Lib "KERNEL32" (ByVal lpPathName As String) As Long
#Else
    Declare Function SetCurrentDirectoryA Lib "KERNEL32" (ByVal lpPathName As String) As Long
#End If


Sub Import_todays_export()

'disable all the alerts
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False

'select the file sorce, may exit the program here
SetCurrentDirectoryA "\\maxhealth.com\filecabinet\Headquarters\Departments\AccountsPayable\Private\General\Audit_Team\Concur\Daily Expense Report Assignment"
PrevFile = Application.GetOpenFilename
If PrevFile = False Then
    End
End If

'open and copy content from the todays export
Range("F3").Value = PrevFile
Workbooks.Open (PrevFile)
Range("A1").CurrentRegion.Copy

'paste content to current template
Windows("Mary Template.xlsm").Activate
Sheets("Step2").Activate
Range("A6").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
Application.CutCopyMode = False 'clean the clipboard
Workbooks(2).Close 'close the todays export

'check for row number
Dim lastrow As Long
lastrow = Range("A7").End(xlDown).Row

'vlookup autofill
Range("L6").Value = "Assign to"
Range("M6").Value = "Assign date"
Range("L7:L" & lastrow).FormulaR1C1 = "=VLOOKUP(RC[-11],Step1!C[-11]:C[-9],3,FALSE)"
Range("M7:M" & lastrow).FormulaR1C1 = "=VLOOKUP(RC[-12],Step1!C[-12]:C[-10],2,FALSE)"

'paste to next sheet
Range("A6").CurrentRegion.Copy
Sheets("Step3").Activate
Range("A6").PasteSpecial Paste:=xlPasteValuesAndNumberFormats

End Sub
```

## Module 3
```
Sub Check_assignment()

'disable all the alerts
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False

Dim lastrow As Long
lastrow = Range("A7").End(xlDown).Row

'replace #N/A with blank
Worksheets("Step3").Columns("L").Replace What:="#N/A", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True
Worksheets("Step3").Columns("M").Replace What:="#N/A", Replacement:="", SearchOrder:=xlByColumns, MatchCase:=True

Range("M7:M" & lastrow).NumberFormat = "m/d/yyyy"
Range("N6").Value = "Assigned?"

'set drop down list and colors
Dim i As Long
i = 7
For i = 7 To lastrow
    If Cells(i, "L").Value = "" Then
        Cells(i, "L").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Formula1:="Debra Wilson, Franklin Lhotsky, Jacqueline Harrison, Kisha Terry, Marlo Hilton, Nancy Zhang, Peggy Elias, Ronda Kromah, Stephanie Jones Boone, Scott Markel, Mary Keke, Sandra Esders"
        Cells(i, "L").Interior.Color = RGB(255, 235, 156)
        Cells(i, "N").Value = "No"
    Else
        Cells(i, "L").Interior.Color = RGB(198, 239, 206)
        Cells(i, "N").Value = "Yes"
    End If
Next i

End Sub
```

## Module 4
```
#If VBA7 Then
    Declare PtrSafe Function SetCurrentDirectoryA Lib "KERNEL32" (ByVal lpPathName As String) As Long
#Else
    Declare Function SetCurrentDirectoryA Lib "KERNEL32" (ByVal lpPathName As String) As Long
#End If

Sub Format_report()

'disable all the alerts
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False

Dim lastrow As Long
lastrow = Range("A6").End(xlDown).Row

Worksheets("Step3").Range("A6:A" & lastrow).Copy Destination:=Worksheets("Step4").Range("A1")
Worksheets("Step3").Range("M6:M" & lastrow).Copy Destination:=Worksheets("Step4").Range("B1")
Worksheets("Step3").Range("L6:L" & lastrow).Copy Destination:=Worksheets("Step4").Range("C1")
Worksheets("Step3").Range("B6:K6" & lastrow).Copy Destination:=Worksheets("Step4").Range("D1")

Application.CutCopyMode = False 'clean the clipboard

Sheets("Step4").Activate
Range("C:C").Validation.Delete
Range("C:C").Interior.Color = xlNone

End Sub
```

# Module 5
```
#If VBA7 Then
    Declare PtrSafe Function SetCurrentDirectoryA Lib "KERNEL32" (ByVal lpPathName As String) As Long
#Else
    Declare Function SetCurrentDirectoryA Lib "KERNEL32" (ByVal lpPathName As String) As Long
#End If

Sub Export_report()

'disable all the alerts
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False

'ask for report saving folder
Dim fileExplorer As FileDialog
Set fileExplorer = Application.FileDialog(msoFileDialogFolderPicker)
fileExplorer.AllowMultiSelect = False 'To allow or disable to multi select
With fileExplorer
    If .Show = -1 Then 'Any folder is selected
        folderpath = .SelectedItems.Item(1)
    Else ' else dialog is cancelled
        End
    End If
End With

'name the report file
NewFile = folderpath & "\" & Format(Date, "mmddyy") & " CONCUR.xlsx"
Range("O5").Value = NewFile

'write the report file
Range("A1").CurrentRegion.Copy
Sheets("Export").Activate
Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
ActiveSheet.Copy
ActiveWorkbook.SaveAs Filename:=NewFile
ActiveWorkbook.Close

MsgBox "Report saved! Please rename the file if today is not the date of the report."

End Sub
```
