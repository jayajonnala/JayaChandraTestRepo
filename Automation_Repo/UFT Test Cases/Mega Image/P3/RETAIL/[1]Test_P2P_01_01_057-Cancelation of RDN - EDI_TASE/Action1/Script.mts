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


'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_P2P_01_01_057-Cancelation of RDN - EDI_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 4th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_P2P_01_01_057-Cancelation of RDN - EDI_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_P2P_01_01_057-Cancelation of RDN - EDI_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)
'
''''Login'''

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'
''''----------------------Tcode MB01----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Date","MKPF-BLDAT","",ConvertDate(DT_MB01_0200_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","MKPF-BUDAT","",ConvertDate(DT_MB01_0200_POSTING_DATE),False)
Call SetTextbox("Movement Type","RM07M-BWARTWE","",DT_MB01_0200_MOVEMENT_TYPE,False)
Call SetTextbox("Site","RM07M-WERKS","",DT_MB01_0200_SITE,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()   
Call PressEnter() ' 
'Capture the screenshot
Call TakeScreenShot()

'Call SetTextbox("Vendor","MSEGK-LIFNR","",DT_MB01_0241_VENDOR,False)
Call SetTextboxNoLabel("MSEGK-LIFNR","",DT_MB01_0241_VENDOR,False)


Call SetTextbox("Article","MSEG-MATNR","0",DT_MB01_0241_ARTICLE,False)
Call SetTextbox("Article","MSEG-MATNR","1",DT_MB01_0241_ARTICLE_OCC1,False)

Call SetTextbox("Quantity","MSEG-ERFMG","0",DT_MB01_0241_QUANTITY,False)
Call SetTextbox("Quantity","MSEG-ERFMG","1",DT_MB01_0241_QUANTITY_OCC1,False)

Call SetTextbox("SLoc","MSEG-LGORT","0",DT_MB01_0210_STOR_LOCATION,False)
Call SetTextbox("SLoc","MSEG-LGORT","1",DT_MB01_0210_STOR_LOCATION_OCC1,False)

'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

'Post the Article No
Call ClickButton("Post   \(Ctrl\+S\)",False)
Wait(2)
Call ClickButtonIfExist("Save",True) 
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
'save it to data sheet
Call GetStatusBar("item1","DT_RDN_DOC_NUMBER_OUTPUT")
'reload data sheet
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'verify statusbar with datasheet feed
Call VerifyStatusBar("Document "& DT_RDN_DOC_NUMBER_OUTPUT &" posted")


''----------------------Tcode MB03----------------------------
''Enter the Tcode
Call SetTcode(DT_MB01_0200_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MB01_0200_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Article Doc\.","RM07M-MBLNR","",DT_RDN_DOC_NUMBER_OUTPUT,False)
Call SetTextbox("Art\. Doc\. Year","RM07M-MJAHR","",DT_MB01_0460_ART_DOC_YEAR,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
Wait(2) 
Call SendKey("{DOWN}")
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Choose   \(F2\)",False)

'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Output   \(Shift\+F2\)",False)

Call VerifyTableCellContentByRef("SAPDV70ATC_NAST3","Medium","EDI","Status",DT_MB01_0100_CHECK_SCREENTOP_OF_TABLECELL_STATUS_1)
Call VerifyTableCellContentByRef("SAPDV70ATC_NAST3","Medium","EDI","Output Type",DT_MB01_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_1)

Call FindRowNumber("SAPDV70ATC_NAST3","Medium","EDI","EDI_ROW_OUTPUT")
Call SelectRowGuiTableByRow("SAPDV70ATC_NAST3",EDI_ROW_OUTPUT,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Processing log   \(Ctrl\+F2\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call GetLabelContentByRefLabel("Message text","0","-96","DT_MB01_0120_CHECK_TEXT_OF_IDOC_0000000555562336_SENT_TO_SAP_OUTPUT",True)
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Cancel   \(F12\)",True)

''----------------------Tcode WE05----------------------------
'Enter the Tcode
Call SetTcode(DT_MB01_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MB01_0100_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("IDoc Number","DOCNUM-LOW","",Mid(DT_MB01_0120_CHECK_TEXT_OF_IDOC_0000000555562336_SENT_TO_SAP_OUTPUT,7,16),False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)

'Capture the screenshot
Call TakeScreenShot()
Call VerifyTextBoxContent("Current Status","EDIDC-STATUS","",DT_MB01_0100_CHECK_TEXT_OF_CURRENT_STATUS,False)
'
'''----------------------Tcode MIGO----------------------------
'Enter the Tcode
Call SetTcode(DT_MB01_0100_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MB01_0100_OKCD_OCC1)
'Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-ACTION",DT_MB01_0010_GODYNPROACTION)
'Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_RDN_DOC_NUMBER_OUTPUT,False)
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,DT_MB01_2010_GODYNPRODOC_YEAR,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()
Call PressEnter()

'item ok
Call ClickButtonIfExist("Open detail data",False)
Call SelectCheckbox("GODYNPRO-DETAIL_TAKE",0,"ON",False)
Call SetTableDataNoRef("SAPLMIGOTV_GOITEM","OK",1,"ON",False)
Call SetTableDataNoRef("SAPLMIGOTV_GOITEM","OK",2,"ON",False)

'Post the Article No
Call ClickButton("Post   \(Ctrl\+S\)",False)
Wait(2)
Call ClickButtonIfExist("Save",True) 
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
'save it to data sheet
Call GetStatusBar("item1","DT_CANCELLATION_DOC_NO_OUTPUT")
'reload data sheet
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'verify statusbar with datasheet feed
Call VerifyStatusBar("Article document "& DT_CANCELLATION_DOC_NO_OUTPUT &" posted")

'''_____________________________________________________________________________________________________

Call SetComboByKey("GODYNPRO-ACTION",DT_MB01_0010_GODYNPROACTION_OCC1)
'Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_CANCELLATION_DOC_NO_OUTPUT,False)
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,DT_MB01_2010_GODYNPRODOC_YEAR_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()
Call PressEnter()

Call SelectTab("TS_GOITEM",DT_MB01_0300_WHERE,False)
Call VerifyTextBoxContent("Movement type","GOITEM-BWART","",DT_MB01_0325_CHECK_TEXT_OF_MOVEMENT_TYPE,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS_GOITEM",DT_MB01_0300_OUTPUT,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Display outputs",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()


Call VerifyTableCellContentByRef("SAPDV70ATC_NAST3","Medium","EDI","Status",DT_MB01_0100_CHECK_TOOLTIP_OF_TABLECELL_STATUS_0)
Call VerifyTableCellContentByRef("SAPDV70ATC_NAST3","Medium","EDI","Output Type",DT_MB01_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)



Call FindRowNumber("SAPDV70ATC_NAST3","Medium","EDI","EDI_ROW_OUTPUT1")
Call SelectRowGuiTableByRow("SAPDV70ATC_NAST3",EDI_ROW_OUTPUT1,False)
'Call SelectRowGuiTableByRow("SAPDV70ATC_NAST3",DT_MB01_0100_SAPDV70ATC_NAST3_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Processing log   \(Ctrl\+F2\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call GetLabelContentByRefLabel("Message text","0","-96","DT_MB01_0120_CHECK_TEXT_OF_IDOC_0000000555562337_SENT_TO_SAP_SYSTEM_OUTPUT",True)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Cancel   \(F12\)",True)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


