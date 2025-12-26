

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Direct Export to foreign customers_p2
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

gstrTestCaseName = "Test_Direct Export to foreign customers_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\P3\DS\RETAIL\DT_Direct Export to foreign customers_p2_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
''----------------------VL03N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 

Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL03N_4004_OUTBOUND_DELIVERY,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SelectMenuBar("Extras;Delivery Output;Header")
Call TakeScreenShot
Call VerifyTableCellContent(1,"Output Type", "SAPDV70ATC_NAST3", DT_VL03N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)

'-----------------------VL02N-------------------------

Call SetTcode(DT_VL03N_4004_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextBox("Outbound Delivery","LIKP-VBELN",0,DT_VL03N_4004_OUTBOUND_DELIVERY_OCC1,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Post Goods Issue   \(Shift\+F8\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call TakeScreenShot

Call GetStatusBar("item3","DT_ARTICLE_DOCUMENT_NUM_OUTPUT")
Wait(2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyStatusBar("Outb. del.(Affiilia) "&DT_VL03N_4004_OUTBOUND_DELIVERY_OCC1&" saved, article document "&DT_VL02N_0200_GRIDCELL_0_ART_DOC_NUMBER&" created")

'--------------------------VL03N--------------------------
Call SetTcode(DT_VL03N_4004_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextBox("Outbound Delivery","LIKP-VBELN",0,DT_VL03N_4004_OUTBOUND_DELIVERY_OCC1,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Document Flow   \(F7\)",False)
Call TakeScreenShot
wait 2
Call ActivateNodeGuiTree(0, "#1;#1;#1")
Call TakeScreenShot
Call GetGridContent("GD.*", 0, "DOCNUM", 1, "<NA>", "<NA>", "DT_VL03N_0100_OUTPUT")
Call ClickButton("Back   \(F3\)",False)

'--------------------------MIGO-----------------------------
Call SetTcode(DT_VL03N_1000_OKCD)     
Call PressEnter() 

Call SetCombo("GODYNPRO-ACTION", "Display")
'''Call SetComboByKey("GODYNPRO-REFDOC", DT_VL03N_0010_GODYNPROACTION_OCC1)
Call PressEnter() 
wait 2
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextBoxNoLabel("GODYNPRO-MAT_DOC","0",DT_VL03N_2010_GODYNPROMAT_DOC,False)
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR","",DT_VL03N_2010_GODYNPRODOC_YEAR,False)
Call PressEnter() 
wait 2
Call TakeScreenShot
Call PressEnter() 
Call SelectTab("TS_GOHEAD", "Doc. info", False)
Call ClickButton("FI Documents",False)
Call TakeScreenShot
Call GetGridContent("Documents in Accounting",0,"Document Number",1,"Object type text", "Accounting document","DT_DOCNR_OUTPUT")

Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)
Call ClickButtonIfExist("Display Document   \(F2\)",True)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "Profit Center", 0, DT_VL03N_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)


Call LogOff()
Call FinalStatus ()





''''''''SAPGuiSession("Session").SAPGuiWindow("Outb. del.(Affiilia) 915146811_2").SAPGuiMenubar("mbar").Select "Extras;Delivery Output;Header"


'*********************************************End Of Script*********************************************************************

