'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Asset Revaluation_p4_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 19th March
'.................Modified By :
'.................Modified Date/Details :
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	DataRowSet= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Asset Revaluation_p4_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Fixed Assets Sales_p3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call SetTextbox("Company Code","ANBZ-BUKRS","",DT_ABAW_0100_COMPANY_CODE,False)
Call SetTextbox("Asset","ANBZ-ANLN1","",DT_ABAW_0100_ASSET,False)
Call SetTextbox("Document Date","ANEK-BLDAT","",Replace(DT_ABAW_0100_DOCUMENT_DATE,"/","."),False)
Call SetTextbox("Posting Date","ANEK-BUDAT","",Replace(DT_ABAW_0100_POSTING_DATE,"/","."),False)

Call SetTextbox("Posting period","ANBZ-PERID","",DT_ABAW_0100_POSTING_PERIOD,False)
Call SetTextbox("Transaction Type","ANBZ-BWASL","",DT_ABAW_0100_TRANSACTION_TYPE,False)
Call TakeScreenShot()
Call PressEnter()  
If SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiEdit("guicomponenttype:=32","attachedtext:=Document Date","name:=ANEK-BLDAT","Index:=0").Exist(5)=True Then 
Call PressEnter()  
End If
Call TakeScreenShot()
Call SetTextbox("Amount posted","ANBZ-DMBTR","",DT_ABAW_0160_REVALODEP_CY,False)
Call SetTextbox("Asset Val. Date","ANBZ-BZDAT","",Replace(DT_ASSET_VAL_DATE,"/","."),False)

Call TakeScreenShot()
Call PressEnter()  
''wait 5
''Call ClickButton("Execute   \(Enter\)",True)
''wait 5
''Call ClickButton("Execute   \(Enter\)",True)
''wait 5
''Call ClickButton("Execute   \(Enter\)",True)
''wait 5
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call GetStatusBar("text","DT_ABAW_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_ABAW_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)


Call LogOff()
Call FinalStatus ()




