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

If Environment("envFunStatus")= 1 Then
	Call fn_CM_Login(strUsername, strPassword, strDomain)
End If

'Creating new customer 
If Environment("envFunStatus") = 1 Then
	Call fn_CM_CustomerCreation(strCustomerName, strCustomerID)
End If

'Create Service Account
If Environment("envFunStatus") = 1 Then
	Call fn_CM_ServiceAccountCreation(strServiceAccountID)
'	Create Contact
	If Environment("envFunStatus") = 1 Then
		strContactType ="Account Primary"
		Call fn_CM_ContactCreation(strContactType, strContactID)	
	End If
End If

'Create Billing Account
If Environment("envFunStatus") = 1 Then
	Call fn_CM_BillingAccountCreation(strServiceAccountID, strBillingAccountID)
End If

'Create Service Address
If Environment("envFunStatus") = 1 Then
	strAddressFormat="On-net"
	Call fn_CM_AddressCreation(strAddressFormat, strSitename, strAddressID)
	Environment("envSiteName") = strSiteName
'	Create Contact
	If Environment("envFunStatus") = 1 Then		
		strContactType="Site Technical Contact"
		Call fn_CM_ContactCreation(strContactType, strContactID)	
	End If
	
End If
	
'Create Site Billing
If Environment("envFunStatus") = 1 Then		
	Call fn_CM_SiteBillingCreation(strAddressID, strBillingAccountID,strSiteName,strAddressID2)	
End If


'Service Page
If Environment("envFunStatus") = 1 Then
	Call fn_CM_ServiceSelection()
End If

'Feature Page
If Environment("envFunStatus") = 1 Then
	Call fn_CM_FeatureSelection()
End If

'Process Page
If Environment("envFunStatus") = 1 Then
	Call fn_CM_Process(strSiteName,strSitename2,strSitename3)
End If


'Order Summary Page
If Environment("envFunStatus") = 1 Then
	Call fn_CM_OrderSummary(strServiceRequestID)
End If

If Environment("envFunStatus") = 1 Then
	Call fn_CM_Signout()
End If


'CSO 
'Environment("envFunStatus")= 1
'strServiceRequestID="2767790"
'Environment("envServiceRequestID") = strServiceRequestID


'Launching URL
If Environment("envFunStatus") = 1 Then
	Call fn_CSO_LaunchURL()
End If

'CSO Login
If Environment("envFunStatus")= 1 Then
	strUsername = fn_GetTestdata("CSO_Username",2) 
	strPassword = fn_GetTestdata("CSO_Password",2) 
	strDomain = fn_GetTestdata("CSO_Domain",2) 
	Call fn_CSO_Login(strUsername, strPassword, strDomain)
End If

'strServiceRequestID="2767251"
'Environment("envServiceRequestID") = strServiceRequestID

'Search SR
If Environment("envFunStatus")= 1 Then
	Call fn_CSO_SearchSR_PSWorkist(strServiceRequestID)
End If


'Provisioning
If Environment("envFunStatus") = 1 Then
	
	strActivityType_Expected="new"

	Call fn_CSO_Provisioning_PSWorklist("Conduct Site Survey(Coax),Conduct Coax Survey,Obtain Site Agreement(s),Build House Account,Obtain Site Permits(Coax),Obtain Coax Permits,Complete Site Build(Coax),Complete Coax Build,Build Update Local Biller Account")	
	
	intNewconnect="1"
	
	SupplementType = fn_GetTestdata("SupplementType",2)
	
	Call fn_CSO_Provisioning_PSWorklist("Design SIP Service,Schedule Trunk Install")
	
	Call fn_CSO_VerifyTaskStatus("Gain Toll-free RespOrg", "Deferred/Inprogress")
	
	Call fn_CSO_VerifyTaskStatus("Provision E911", "Deferred/Inprogress")

	Call fn_CSO_Provisioning_PSWorklist("Pre Provision Switch")

	Call fn_CSO_VerifyTaskStatus("Provision Toll free", "Deferred/Inprogress")
	
	Call fn_CSO_VerifyTaskStatus("CIC Toll-free", "Deferred/Inprogress")
	
	Call fn_CSO_Provisioning_PSWorklist("Install Trunk CPE,Update Cramer,Cutover Trunk,Provision DA/DL")
	
	Call fn_CSO_VerifyTaskStatus("Provision DA/DL", "Deferred/Inprogress/Completed/Fallout")
		
	Call fn_CSO_VerifyTaskStatus("Assign TN", "Deferred/Inprogress")

	Call fn_CSO_VerifyTaskStatus("Mediation-Manage Account", "Deferred/Inprogress")
	
	Call fn_CSO_VerifyTaskStatus("Mediation-Manage Device", "Deferred/Inprogress")
	
	Call fn_CSO_Provisioning_PSWorklist("Start Trunk Billing,Validate Billing")
	
End If

'Search SR
If Environment("envFunStatus")= 1 Then
	Call fn_CSO_SearchSR_PSWorkist(strServiceRequestID)
End If

'CSO Signout
If Environment("envFunStatus")= 1 Then
	Call fn_CSO_Signout()
End If

'Results
Call fn_WriteTestCaseResults(strTestCaseName)





