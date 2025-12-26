
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Reverting Goods Receipt using the PO Number and Verification of the net result of material_TASE
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

gstrTestCaseName = "TC_10_Test_Reverting Goods Receipt and Verifying the net result of material_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_MD_ABI067_001 Create ZCXT Retail Customer Local or Foreign_TASE.xls"



'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''----------------------Tcode MIGO----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call SetCombo("GODYNPRO-ACTION", "Goods Receipt")
Call SetCombo("GODYNPRO-REFDOC", "Purchase Order")
Call PressEnter() 
Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER", 0, DT_PO_NUMMBER, False)
Call TakeScreenShot()

Call SetTextboxNoLabel("GODEFAULT_TV-BWART", 0, DT_GR_PO_REVERSAL, False)
Call TakeScreenShot()
Call ClickButton("Execute",False)
Call TakeScreenShot()

Call SetTextbox("Delivery Note","GOHEAD-LFSNR", 0, DT_DELIVERY_NOTE, False)
Call SetTextbox("HeaderText","GOHEAD-BKTXT", 0, DT_HEADER_TEXT, False)
Call TakeScreenShot()

Call ClickButtonifExist("Open detail data",False)
Call TakeScreenShot()

Call SelectTab("TS_GOITEM", "Where", False)
Call TakeScreenShot()

Call SetTextbox("Storage Location","GOITEM-LGOBE", 0, DT_STORAGE_LOCATION, False)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE", 0, "ON", False)
Call TakeScreenShot()

Call SelectTab("TS_GOITEM", "Batch", False)
Call TakeScreenShot()

Call VerifyTextBoxContent("Batch","GOITEM-CHARG", 0, DT_BATCH_NUMBER, False)

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot()

Call GetStatusBar("Item1", "DT_ARTICLE_NUMBER")

Call  ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

''----------------------Tcode MB51----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE_OCC1)
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call SetTextboxNoLabel("WERKS-LOW", 0, DT_SITE, False)
Call SetTextboxNoLabel("BUDAT-LOW", 0, ConvertDate(DT_POSTING_DATE), False)
Call SetTextboxNoLabel("USNAM-LOW", 0, DT_USER, False)
Call SetTextboxNoLabel("BUKRS-LOW", 0, "", False)
Call SetTextboxNoLabel("MBLNR-LOW", 0, "", False)

Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

Call GetNumberOfRows("shell", 0, "DT_NUMBER_OF_ROWS")
Call VerifyGridCellContent("", DT_NUMBER_OF_ROWS, "Quantity", 0, DT_NET_QUNAITY)

Call  ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus()
