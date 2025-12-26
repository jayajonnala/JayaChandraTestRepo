
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Regular purchasing in SW31 dry goods_p1
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
'.................Test Script Name : Test_Regular purchasing in SW31 dry goods_p1
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 19th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Regular purchasing in SW31 dry goods_p1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Regular purchasing in SW31 dry goods_p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode ME21N----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()  
Call TakeScreenShot() 


Call SetComboByKey("MEPO_TOPLINE-BSART",DT_ME21N_1105_MEPO_TOPLINEBSART)
'Call SetTextbox("Vendor","MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)
Call SetTextboxNoLabel("MEPO_TOPLINE-SUPERFIELD","",DT_ME21N_1105_VENDOR,False)
Call TakeScreenShot()
Call PressEnter()
' ClickButtonIfExist(tooltipOrButtonName, blnIsItPopup)


Call ClickButtonIfExist("Expand Items Ctrl+F3",False)
Call SetTextbox("Purch. Org.","MEPO1222-EKORG","",DT_ME21N_1221_PURCH_ORG,False) 
Call SetTextbox("Purch. Group","MEPO1222-EKGRP","",DT_ME21N_1221_PURCH_GROUP,False)     
Call PressEnter() 
Call SetTextbox("Company Code","MEPO1222-BUKRS","",DT_ME21N_1221_COMPANY_CODE,False)    
Call PressEnter()
Call TakeScreenShot()

Call SetTableData("SAPLMEGUITC_1211","Article","1","","",DT_ME21N_1211_TABLECELL_ARTICLE_0,False) 
Call SetTableData("SAPLMEGUITC_1211","PO Quantity","1","","",DT_ME21N_1211_TABLECELL_PO_QUANTITY_0,False)
Call SetTableData("SAPLMEGUITC_1211","Plnt","1","","",DT_ME21N_1211_TABLECELL_SITE_0,False)
Call SetTableData("SAPLMEGUITC_1211","Deliv. Date","1","","","",False) 

Call TakeScreenShot()
Call PressEnter()
Call PressEnter()

Call SelectTab("ITEM_DETAIL","Delivery",False)
Call TakeScreenShot()

Call VerifyCheckBoxValue("MEPO1313-WEORA","OFF")
Call TakeScreenShot()

Call SelectTab("ITEM_DETAIL","Conditions",False)
Call TakeScreenShot()

'''''Call SetTableData("SAPLV69ATCTRL_KONDITIONEN","Amount","1","","",1,False)
'''''Call PressEnter()

Call SelectMenuBar("Purchase Order;Save")
Wait(2)
Call TakeScreenShot()
Call ClickButtonIfExist("Save",True)
Call TakeScreenShot()

'Call ClickButtonIfExist("Hold",True)

Call GetStatusBar("item2","DT_PO_NUMBER_OUTPUT")
Call VerifyStatusBar("Standard PO Retail created under the number " & DT_PO_NUMBER_OUTPUT)

'----------------------Tcode ME29N----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_ME21N_0014_OKCD) 
Call PressEnter()
Call ClickButtonIfExist("Cancel   (F12)",True)
Call TakeScreenShot()
Call CheckTCodeScreen(DT_ME21N_0014_OKCD)


Call ClickButtonIfExist("Other Purchase Order   \(Shift\+F5\)",False)
wait(1)
Call TakeScreenShot()
Call SetTextboxPopupIfExist("MEPO_SELECT-EBELN","Pur\. Order",DT_PO_NUMBER_OUTPUT)
Call TakeScreenShot()
Call ClickButtonIfExist("Other Document   \(Enter\)",True)
wait(2)
Call TakeScreenShot()

Call ClickButtonIfExist("Expand Header Ctrl\+F2",True)
wait(2)
Call TakeScreenShot()

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

