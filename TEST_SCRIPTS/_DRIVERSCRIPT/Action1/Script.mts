﻿
'Driver Script

On Error Resume Next

If QCUtil.IsConnected Then
	Msgbox "Please Disconnect ALM Connection from UFT"
	ExitTest
End If


Environment("ProjectName") = "UFTBaseFramework"  'Need to change the project name for different Projects
TestPath =  Environment("TestDir")  'It will give the path upto test Action
Length_TestPath = Len(TestPath)
StrartPosition_ProjectName = Instr(TestPath, Environment("ProjectName") )
Environment("ProjectParentFolder") = Mid(TestPath,1,StrartPosition_ProjectName-2)
Environment("ProjectFolder") = Environment("ProjectParentFolder") & "\" & Environment("ProjectName")


LoadFunctionLibrary Environment("ProjectFolder")&"\FUNCTION_LIBRARY\Environment_Functions.qfl"	
LoadFunctionLibrary Environment("ProjectFolder")&"\FUNCTION_LIBRARY\FW_Functions.qfl"	
LoadFunctionLibrary Environment("ProjectFolder")&"\FUNCTION_LIBRARY\CM_Functions.qfl"
LoadFunctionLibrary Environment("ProjectFolder")&"\FUNCTION_LIBRARY\CSO_Functions.qfl"
LoadFunctionLibrary Environment("ProjectFolder")&"\FUNCTION_LIBRARY\Variables_Initialization.qfl"

Call fn_EnvironmentVariables()

Call fn_CloseBrowsers("ALL")

Call fn_KillExcel()
wait 1

Set objFSO = CreateObject("Scripting.FileSystemObject")	
If objFSO.FileExists(Environment("DriverSheetFolderPath") & "\DriverSheet_HTMLReport.mht") Then
	objFSO.DeleteFile(Environment("DriverSheetFolderPath") & "\DriverSheet_HTMLReport.mht")
End If


'Reading from Test Cases sheet
Set oExcel_TestCases = CreateObject("Excel.Application")
Set oWB_TestCases = oExcel_TestCases.Workbooks.Open(Environment("DriverSheetFolderPath") & "\" & "DriverSheet.xlsx")
Set oSheet_TestCases = oWB_TestCases.WorkSheets(1)
oExcel_TestCases.Visible = true
intTestCasesCount = oSheet_TestCases.UsedRange.Rows.Count

For intSheetRowNo = 2 To intTestCasesCount
	'Status
	oSheet_TestCases.Cells(intSheetRowNo,5).Font.colorindex = 1
	oSheet_TestCases.Cells(intSheetRowNo,5) = "No Run"	
	'Execution Time
	oSheet_TestCases.Cells(intSheetRowNo,6) = ""	
	'Result Path
	oSheet_TestCases.Cells(intSheetRowNo,7) = ""
Next

oWB_TestCases.Save
Wait 1
oWB_TestCases.Close
Wait 1
oExcel_TestCases.Quit
Wait 1


'Importing TestCases sheet into UFT Test
Datatable.ImportSheet Environment("DriverSheetFolderPath") & "\" & "DriverSheet.xlsx",1,1

intRowCount = Datatable.GetSheet(1).GetRowCount

For RowNo = 1 to intRowCount

    Datatable.GetSheet(1).SetCurrentRow(RowNo)
    
    If UCase(Trim(datatable.Value("Execution_Required"))) = "YES" Then

    	Environment("TestScenarioName") = Datatable.Value("Test_Name")        
	Environment("TestCaseName") = Datatable.Value("Action_Name")
	
	strTestPath = Environment("TestScriptsFolderPath") & "\" & Environment("TestScenarioName")
	
	'Executing TestCases one by one
	LoadAndRunAction strTestPath, Environment("TestCaseName")

    End If
 
Next

'Generating Main HTML report for Driver sheet
Call fn_DriverSheet_HTMLReport()


'Sending Email if required
If Trim(Ucase(Environment("SendStatusMail_DriversheetReport"))) = "YES" Then

	strBrowserName = Environment("Browser") 
	strURL = "https://mail.excelacom.in"
	strUsername = "balakrishna.m"
	strPassword = "59a5587d8301afd8baa6e0daea16330ff24feb07a8efebcbb361"
	strTo = "balakrishna.m@excelacom.in"	
'	strCC = "Century_Automation@excelacom.in"
	strSubject = "Automation Execution Status - Main Result - " & Date
	strBody = "Note : This is Autogenerated mail by Automation. So please dont give the reply for this mail. PFA for the Test Results. "
	strAttachmentPath = Environment("DriverSheetFolderPath") & "\MainResults.mht"
	
	Call fn_SendStatusMail(strBrowserName, strURL, strUsername, strPassword, strTo, strCC, strSubject, strBody, strAttachmentPath)
	
End If

On Error Goto 0


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


