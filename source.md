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
