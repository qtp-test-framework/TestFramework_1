'========== Declare and initialize the variables ==========
Dim sTestCaseFolder, strScriptPath, sQTPResultsPath, sQTPResultsPathOrig
 
sBatchSheetPath = "D:\Mohit\QTP\QTP_Excel\Batch.xls"
sTestCaseFolder = "C:\Users\Mohit-Ankit\Documents\Unified Functional Testing\"
sQTPResultsPathOrig = "D:\Mohit\QTP\"
 
'========== Create an object to access QTP's objects, methods and properties ==========
Set qtpApp = CreateObject("QuickTest.Application")
 
'Open QTP if it is already not open
If qtpApp.launched <> True Then
 qtpApp.Launch
End If
 
qtpApp.Visible = True
 
'========== Set the run options for QTP ==========
qtpApp.Options.Run.ImageCaptureForTestResults = "OnError"
qtpApp.Options.Run.RunMode = "Fast"
 
'Set ViewResults property to false. This is because if you run many test cases in batch, you would not want QTP to open a separate result window for each of them
qtpApp.Options.Run.ViewResults = False
 
' ========== Read test cases from batch excel sheet ==========
Set xl_Batch = CreateObject("Excel.Application")

Set wb_Batch = xl_Batch.WorkBooks.Open(sBatchSheetPath)
 
'Loop through all the Rows
'1st row is header and the test case list starts from 2nd row. So, For loop is started from 2nd row
For iR = 2 to 1000
	
 'Get the value from the Execute column
 If xl_Batch.Cells(iR, 1).Value = "Yes" Then
 
 'Get Test Case Name
 sTestCaseName = xl_Batch.Cells(iR, 2).Value
	 
	 strScriptPath = ""
 'Get the location where the test case is stored
 strScriptPath = sTestCaseFolder & sTestCaseName
	
 'Open the Test Case in Read-Only mode
 qtpApp.Open strScriptPath, True
 WScript.Sleep 2000
 
 'Create an object of type QTP Test
 Set qtpTest = qtpApp.Test
 
 'Instruct QTP to perform next step when error occurs
 qtpTest.Settings.Run.OnError = "NextStep"
 
 'Create the Run Results Options object
 Set qtpResult = CreateObject("QuickTest.RunResultsOptions")
 
 'Set the results location. This result refers to the QTP result
 sQTPResultsPath = sQTPResultsPathOrig
 sQTPResultsPath = sQTPResultsPath & sTestCaseName
 qtpResult.ResultsLocation = sQTPResultsPath
 
 'Run the test. The result will automatically be stored in the location set by you
 WScript.Sleep 2000
 qtpTest.Run qtpResult
 
 'MsgBox "Ending Loop"
 
 ElseIf xl_Batch.Cells(iR, 1).Value = "No" Then
 'Do nothing. You don't have to execute the test cases marked as No
 
 ElseIf xl_Batch.Cells(iR, 1).Value = "" Then
 'Blank value means that the list of test cases has finished
 'So you can exit the for loop
 Exit For
  'MsgBox "After Loop"
 
 End If
 
Next