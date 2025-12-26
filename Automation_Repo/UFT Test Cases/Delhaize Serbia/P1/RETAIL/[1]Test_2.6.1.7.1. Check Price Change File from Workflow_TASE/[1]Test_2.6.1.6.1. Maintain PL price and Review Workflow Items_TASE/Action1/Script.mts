
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.6.1.6.1. Maintain PL price and Review Workflow Items
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


'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.6.1.6.1. Maintain PL price and Review Workflow Items
'.................Author : TCS 	   :Raushan
'................ Creation Date    :20th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_2.6.1.6.1. Maintain PL price and Review Workflow Items"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.4.1.2.5. Listing at DC.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode ZMDPC_PL_BL_PRICE_CM ----------------------------
'Enter the transaction code
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

'Enter the details
'Call SetCombo("P_ACTION","Workflow")
Call SetComboByKey("P_ACTION","2")
Call SetTextbox("Sales Organization","P_VKORG","","SS01",False)
Call SetTextbox("Distribution Channel","P_VTWEG","","01",False)
Call SetTextbox("Purchasing Group","S_EKGRP-LOW","",DT_ZMDPC_PL_BL_PRICE_CM_1000_PURCHASING_GROUP,False)
Call SetTextbox("Sales Price Condition Type","P_SALES","","VKPO",False)
Call SetTextbox("Price List","P_PLTYP","",DT_ZMDPC_PL_BL_PRICE_CM_1000_PRICE_LIST,False)
Call SetTextbox("Pricing Date","P_PRSDT","",ConvertDate(DT_ZMDPC_PL_BL_PRICE_CM_1000_PRICING_DATE),False)

'Capture the screenshot
Call TakeScreenShot()


'Click on Program;Execute in Background in Menu bar
Call SelectMenuBar("Program;Execute in Background")

'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Output Device","PRI_PARAMS-PDEST","",DT_ZMDPC_PL_BL_PRICE_CM_0100_OUTPUT_DEVICE,False)

'Capture the screenshot
Call TakeScreenShot()

'Click on Continue
Call ClickButton("Continue   \(Shift\+F1\)",True)
Wait(2)
Call TakeScreenShot()

'Click on Continue
Call ClickButton("Continue   \(Enter\)",True)
Wait(2)
Call TakeScreenShot()

'Click on Immediate
Call ClickButton("Immediate",True)
Wait(2)
Call TakeScreenShot()

'Click on save
Call ClickButton("Save   \(Ctrl\+S\)",True)
Wait(2)
Call TakeScreenShot()

'Verify the Status bar Mesage Type
Call VerifyStatusBarMessageType("S")

'Verify the status bar message
Call VerifyStatusBar(DT_ZMDPC_PL_BL_PRICE_CM_1000_CHECK_TEXT_OF_STATUSBAR)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

