
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Direct Re- Export without border crossing_p2
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
'.................Test Script Name :Test_Direct Re- Export without border crossing_p2
'.................Author : TCS 	   :Raushan
'................ Creation Date    :20th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Direct Re- Export without border crossing_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Direct Re- Export without border crossing_p2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode ME21N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
Wait(1)
Call TakeScreenShot()

'Enter the Organisation details
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)    
Call PressEnter()
Call TakeScreenShot()


'Click on Switch On Document Overview
Call ClickButtonIfExist("Switch On Document Overview   \(F8\)",False)
wait(2)

SAPGuiSession("Session").SAPGuiWindow("Create Purchase Order").SAPGuiToolbar("ToolBarControl").PressContextButton "SELECT"
SAPGuiSession("Session").SAPGuiWindow("Create Purchase Order").SAPGuiToolbar("ToolBarControl").SelectMenuItem "Purchase Requisitions"


'Enter the Purchase Requisition Number
Call SetTextbox("Purchase Requisition Number","SP\$00026-LOW","",DT_ME21N_1000_PURCHASE_REQUISITION_NUMBER,FALSE)
Call TakeScreenShot()

'Click on Execute
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)


'Click on GUI Tree
Call ActivateItemGuiTree(0,DT_ME21N_1000_PURCHASE_REQUISITION_NUMBER,DT_ME21N_1000_PURCHASE_REQUISITION_NUMBER)
Wait(2)
'Click on Continue
Call ClickButtonIfExist("Cancel",True)
Wait(2)

'Call ClickContextButtonToolBar("shell\[0\]",1)
SAPGuiSession("Session").SAPGuiWindow("Create Purchase Order").SAPGuiToolbar("ToolBarControl").PressButton "COPY"
Wait(5)
Call TakeScreenShot()


'Save the Details
Call ClickButton("Save   \(Ctrl\+S\)",False) 
Wait(2)

'Verify the message type
Call VerifyStatusBarMessageType("S")

'Validate If Reexport PO Retail Order is generated
Call GetStatusBar("item2","DT_REEXPORT_PO_NUM_OUTPUT")
VerifyStatusBar("Reexport PO Retail created under the number "&DT_REEXPORT_PO_NUM_OUTPUT)


'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

 @@ hightlight id_;_1_;_script infofile_;_ZIP::ssf1.xml_;_

