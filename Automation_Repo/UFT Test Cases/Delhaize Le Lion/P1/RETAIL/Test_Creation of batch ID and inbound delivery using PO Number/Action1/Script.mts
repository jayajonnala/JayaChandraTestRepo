
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Creation of batch ID and inbound delivery using PO Number
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



gstrTestCaseName = "TC_03_Test_Creation of batch ID and inbound delivery using PO Number_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_MD_ABI067_001 Create ZCXT Retail Customer Local or Foreign_TASE.xls"
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''
'
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''----------------------Tcode VL31N----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call SetTextbox("Purchase order","LV50C-BSTNR", "", DT_PO_NUMBER, False)
Call SetTextbox("Delivery date","RV50A-LFDAT_LA", "", ConvertDate(DT_DELIV_DATE), False)
Call TakeScreenShot()
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call SelectRowGuiTableByRow("SAPMV50ATC_LIPS_OVER_INB", 1, False)
Call TakeScreenShot()
Call ClickButton("Create batches",False)
Wait 5
Call TakeScreenShot()
Call ClickButtonifExist("OK   \(F5\)",False)
'Call TakeScreenShot()
wait 5
Call ClickButtonifExist("SPOP-OPTION_CAN",False)
Call GetTableCellData("SAPMV50ATC_LIPS_OVER_INB", "Batch", 1, "", "", "DT_BATCH_ID_OUTPUT", False)
Call GetTableCellData("SAPMV50ATC_LIPS_OVER_INB", "Article", 1, "", "", "DT_ARTICLE_OUTPUT", False)
Call GetTableCellData("SAPMV50ATC_LIPS_OVER_INB", "Description", 1, "", "", "DT_DESCRIPTION_OUTPUT", False)
Call  VerifyTableCellContent(1, "Delivery quantity", "SAPMV50ATC_LIPS_OVER_INB", DT_DELIVERY_QUNAITTY)

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Call GetStatusBar("item2", "DT_INBOUND_DELIVERY_OUTPUT")
Call VerifyStatusBar(Lcase("Inbound Delivery DLL "&DT_INBOUND_DELIVERY_OUTPUT&" has been saved"))

Call LogOff()
Call FinalStatus()
