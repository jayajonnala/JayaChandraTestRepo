'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Distribution of consumables_p1_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 25th Feb
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

gstrTestCaseName = "Test_Distribution of consumables_p1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Distribution of consumables_p1_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 



'''----------------------Tcode ME21N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()


Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
wait(1)
Call TakeScreenShot()
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)

Call ClickButtonIfExist("Expand Items Ctrl\+F3",False)
wait(1)
Call TakeScreenShot()
Call SetTableData("SAPLMEGUITC_1211","A","1","","",DT_ME21N_1211_TABLECELL_A_0,False)
Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False)
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False)
Call SetTextbox("Supplying Site","MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_SUPPLYING_SITE,False)
Call TakeScreenShot()

Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
wait(1)
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False) 
Call TakeScreenShot()

'Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",Replace((DT_DELIVERY_DATE),"/","."),False)
Call PressEnter()
Call ClickButtonIfExist("Continue",True)
Call ClickButtonIfExist("Continue",True)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
Call TakeScreenShot()

Call SetTextbox("Cost Center","COBL-KOSTL","",DT_ME21N_1101_COST_CENTER,False)
'Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",Replace((DT_DELIVERY_DATE),"/","."),False)
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","",ConvertDate(DT_DELIVERY_DATE),False)

'Call PressEnter()
Call TakeScreenShot()

Call SelectMenuBar("Purchase Order;Save")
Wait(1)
Call ClickButtonIfExist("Continue",True)
'Call ClickButtonIfExist("Continue",True)
Wait(1)
Call ClickButtonIfExist("Save",True)
Wait(2)
Call TakeScreenShot()
Call GetStatusBar("item2","DT_PO_NUMBER_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC2)
Call VerifyStatusBar(DT_ME21N_0014_CHECK_TEXT_OF_STATUSBAR)

'----------------------------Additional Code To process in vl10b for outbound delivery doc creation_need for P2 script--------------
Call SetTcode(DT_TCODE) 
Call PressEnter()
Call TakeScreenShot()

Call SelectTab("TABSTRIP_ORDER_CRITERIA","Purchase Orders",False)
Call SetTextbox("Shipping Point/Receiving Pt","ST_VSTEL-LOW","","",False)
Call SetTextbox("Deliv\. Creation Date","ST_LEDAT-LOW","","",False)
Call SetTextbox("to","ST_LEDAT-HIGH","","",False)
Call SetTextbox("Purchasing Document","ST_EBELN-LOW","",DT_PO_NUMBER,False)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)

Wait(1)
Call TakeScreenShot()
Call SelectRowGuiGridbyRowNo("","",1,False)
Call TakeScreenShot()
Call ClickButton("Create Delivery in Background   \(Shift\+F7\)",False)
Wait(2)
Call TakeScreenShot()
'Call VerifyStatusBarMessageType("S")
'______________________________________________________________________________________________________________________
Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()

