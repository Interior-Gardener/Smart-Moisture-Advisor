Option Explicit
Dim wordApp, reportDoc, taskDir, fso, args
Set fso = CreateObject("Scripting.FileSystemObject")
Set args = WScript.Arguments
taskDir = args(0)
Set wordApp = CreateObject("Word.Application")
wordApp.Visible = False
wordApp.DisplayAlerts = 0
WScript.Echo "Word created via VBScript"
Set reportDoc = wordApp.Documents.Open(taskDir & "\SMA_Major_Project_Report_10_Pages.docx",False,True,False)
WScript.Echo "Report opened"
WScript.Echo "Page count: " & reportDoc.ComputeStatistics(2)
reportDoc.ExportAsFixedFormat taskDir & "\SMA_Major_Project_Report_10_Pages.pdf", 17
WScript.Echo "PDF exported"
reportDoc.Close 0
wordApp.Quit
