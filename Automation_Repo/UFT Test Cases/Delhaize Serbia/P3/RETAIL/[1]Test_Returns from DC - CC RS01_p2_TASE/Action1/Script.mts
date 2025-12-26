
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Returns from DC - CC RS01_p2_TASE
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


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


gstrTestCaseName = "Test_Returns from DC - CC RS01_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\DS\RETAIL\DT_Returns from DC - CC RS01_p2_TASE.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode MIRO----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()    


Call SetTextbox("Document Number","S_EBELN-LOW","",DT_DOCUMENT_NUMBER,False)
Call SetTextbox("Purchasing Organization","S_EKORG-LOW","",DT_ME9F_1000_PURCHASING_ORGANIZATION,False)
Call SetTextbox("Purchasing Group","S_EKGRP-LOW","",DT_ME9F_1000_PURCHASING_GROUP,False)
Call SetTextbox("Application","P_KAPPL","",DT_ME9F_1000_APPLICATION,False)
Call SetTextbox("Processing Status","P_VSTAT","",DT_ME9F_1000_PROCESSING_STATUS,False)
Call PressEnter()     
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyifGuiLabelExists(DT_ME9F_0120_CHECK_TEXT_OF_ZNEU)
Call TakeScreenShot

Call SelectCheckboxNoLabel("1","ON",False)
Call ClickButton("Display Message   \(Shift\+F8\)",False)
Call TakeScreenShot

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

