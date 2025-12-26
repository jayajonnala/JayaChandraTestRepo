
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Distribution process to DS stores - SW41 (Fresh meat) - p2
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
'.................Test Script Name :Test_Distribution process to DS stores - SW41 (Fresh meat) - p2
'.................Author : TCS 	   :Raushan
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Distribution process to DS stores - SW41 (Fresh meat) - p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Distribution process to DS stores - SW41 (Fresh meat) - p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode ME23N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()


Call ClickButtonIfExist("Other Purchase Order   \(Shift\+F5\)",False)
wait(2)
Call SetTextboxPopupIfExist("MEPO_SELECT-EBELN","Pur\. Order",DT_ME23N_0003_PUR_ORDER)
Call PressEnter()
Wait(2)
Call TakeScreenShot()

Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
wait(2)

'Navigate to Purchase Order History Tab
Call SelectTab("ITEM_DETAIL","Purchase Order History",False)
Wait(1)
Call TakeScreenShot()

'----------------------Tcode VL02N----------------------------

Call SetTcode(DT_ME23N_0014_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ME23N_0014_OKCD)
Call TakeScreenShot()

'Enter the Outbound Delivery No
Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_ME23N_4004_OUTBOUND_DELIVERY,FALSE)
Call TakeScreenShot()
Call PressEnter()
Wait(3)
Call TakeScreenShot()

'Enter the Document Date
Call SetTextbox("Document Date","LIKP-BLDAT","",ConvertDate(DT_ME23N_1502_DOCUMENT_DATE),FALSE)
Call TakeScreenShot()
'
Call SelectCellGuiTable("SAPMV50ATC_LIPS_OVER","Docu. Batch No.","Itm",10,False)
Call TakeScreenShot()
Wait(1)
Call SendKey("{F4}")
Wait(3)
Call SendKey("{F2}")
Call ClickButtonIfExist("Copy   \(Enter\)",True)

Call SelectCellGuiTable("SAPMV50ATC_LIPS_OVER","Docu. Batch No.","Itm",20,False)
Call TakeScreenShot()
Wait(1)
Call SendKey("{F4}")
Wait(3)
Call SendKey("{F2}")
Call ClickButtonIfExist("Copy   \(Enter\)",True)

Call SelectCellGuiTable("SAPMV50ATC_LIPS_OVER","Docu. Batch No.","Itm",30,False)
Call TakeScreenShot()
Wait(1)
Call SendKey("{F4}")
Wait(3)
Call SendKey("{F2}")
Call ClickButtonIfExist("Copy   \(Enter\)",True)

Call SelectCellGuiTable("SAPMV50ATC_LIPS_OVER","Docu. Batch No.","Itm",40,False)
Call TakeScreenShot()
Wait(1)
Call SendKey("{F4}")
Wait(3)
Call SendKey("{F2}")
Call TakeScreenShot()
Call ClickButtonIfExist("Copy   \(Enter\)",True)


Call SelectCellGuiTable("SAPMV50ATC_LIPS_OVER","Docu. Batch No.","Itm",10,False)
Call TakeScreenShot()
Wait(1)
Call SetTableData("SAPMV50ATC_LIPS_OVER","Docu. Batch No.",1,"Itm",10,"10031439",False)
Call TakeScreenShot()

Call SelectCellGuiTable("SAPMV50ATC_LIPS_OVER","Docu. Batch No.","Itm",20,False)
Call TakeScreenShot()
Wait(1)
Call SetTableData("SAPMV50ATC_LIPS_OVER","Docu. Batch No.",2,"Itm",20,"10032689",False)
Call TakeScreenShot()

Call SelectCellGuiTable("SAPMV50ATC_LIPS_OVER","Docu. Batch No.","Itm",30,False)
Call TakeScreenShot()
Wait(1)
Call SetTableData("SAPMV50ATC_LIPS_OVER","Docu. Batch No.",3,"Itm",30,"10031439",False)
Call TakeScreenShot()

Call SelectCellGuiTable("SAPMV50ATC_LIPS_OVER","Docu. Batch No.","Itm",40,False)
Call TakeScreenShot()
Wait(1)
Call SetTableData("SAPMV50ATC_LIPS_OVER","Docu. Batch No.",4,"Itm",30,"EVG90800",False)
Call TakeScreenShot()


'Click on Post Goods Issue
Call ClickButton("Post Goods Issue   \(Shift\+F8\)",False) 


'Verify Status Bar message
Call VerifyStatusBar(DT_ME23N_4004_CHECK_TEXT_OF_STATUSBAR)

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************



