'******************************************************************************************************************************			
'														SCRIPT  HEADER 			
'******************************************************************************************************************************
'	Test Name	:	
'	Action(s)		:	
'	Description	:	
'	DATE		:	
'	Author		:	
'******************************************************************************************************************************

Environment("envProjectName") = "UFTBaseFramework"	'Need to change the project name for different Projects
TestPath =  Environment("TestDir")
Length_TestPath = Len(TestPath)
StrartPosition_ProjectName = Instr(TestPath, Environment("envProjectName") )
Environment("envProjectParentFolder") = Mid(TestPath,1,StrartPosition_ProjectName-2)
Environment("envProjectFolder") = Environment("envProjectParentFolder") & "\" & Environment("envProjectName")

Set QCConnection = QCUtil.QCConnection    
If QCUtil.IsConnected Then
	LoadFunctionLibrary "[QualityCenter\Resources] Resources\Subject\Active\Century\SIP\FUNCTION_LIBRARY\" & "Environment_Functions"
	LoadFunctionLibrary "[QualityCenter\Resources] Resources\Subject\Active\Century\SIP\FUNCTION_LIBRARY\" & "FW_Functions"
ElseIf Not QCUtil.IsConnected Then
	LoadFunctionLibrary Environment("envProjectFolder") & "\FUNCTION_LIBRARY\Environment_Functions.qfl"	
	LoadFunctionLibrary Environment("envProjectFolder") & "\FUNCTION_LIBRARY\FW_Functions.qfl"	
End If


Call fn_StartTestCaseExecution() 	'Mandatory for all Actions


'CM

'Launching URL
Call fn_CM_LaunchURL()

'CM Login

strUsername = fn_GetTestdata("CM_Username",2) 
strPassword = fn_GetTestdata("CM_Password",2) 
strDomain = fn_GetTestdata("CM_Domain",2) 

If Environment("envFunStatus")= 1 Then
	Call fn_CM_Login(strUsername, strPassword, strDomain)
End If

If Environment("envFunStatus") = 1 Then
	Call fn_CM_Signout()
End If


Call fn_WriteTestCaseReport() 	'Mandatory for all Actions

'********************************************************************************************

