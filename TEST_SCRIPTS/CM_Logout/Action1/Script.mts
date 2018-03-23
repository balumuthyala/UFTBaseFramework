'******************************************************************************************************************************			
'														SCRIPT  HEADER 			
'******************************************************************************************************************************
'	Test Name	:	
'	Action(s)		:	
'	Description	:	
'	DATE		:	
'	Author		:	
'******************************************************************************************************************************

Environment("ProjectName") = "UFTBaseFramework"	'Need to change the project name for different Projects
TestPath =  Environment("TestDir")  'It will give the path upto test Action
Length_TestPath = Len(TestPath)
StrartPosition_ProjectName = Instr(TestPath, Environment("ProjectName") )
Environment("ProjectParentFolder") = Mid(TestPath,1,StrartPosition_ProjectName-2)
Environment("ProjectFolder") = Environment("ProjectParentFolder") & "\" & Environment("ProjectName")


'Set QCConnection = QCUtil.QCConnection    
If QCUtil.IsConnected Then
	LoadFunctionLibrary "[QualityCenter\Resources] Resources\Subject\Active\Century\SIP\FUNCTION_LIBRARY\" & "Environment_Functions"
	LoadFunctionLibrary "[QualityCenter\Resources] Resources\Subject\Active\Century\SIP\FUNCTION_LIBRARY\" & "FW_Functions"
ElseIf Not QCUtil.IsConnected Then
	LoadFunctionLibrary Environment("ProjectFolder")&"\FUNCTION_LIBRARY\Environment_Functions.qfl"	
	LoadFunctionLibrary Environment("ProjectFolder")&"\FUNCTION_LIBRARY\FW_Functions.qfl"	
End If


Call fn_StartTestCaseExecution() 	'Mandatory for all Actions


'CM

'Launching URL
Call fn_CM_LaunchURL()

'CM Login

strUsername = fn_GetTestdata("CM_Username",2) 
strPassword = fn_GetTestdata("CM_Password",2) 
strDomain = fn_GetTestdata("CM_Domain",2) 

Call fn_CM_Login(strUsername, strPassword, strDomain)

Call fn_CM_Signout()


Call fn_WriteTestCaseReport() 	'Mandatory for all Actions

'********************************************************************************************

