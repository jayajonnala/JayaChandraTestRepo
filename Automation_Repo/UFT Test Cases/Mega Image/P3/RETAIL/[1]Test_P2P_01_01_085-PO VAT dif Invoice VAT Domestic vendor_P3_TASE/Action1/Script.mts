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
'.................Test Script Name : Test_P2P_01_01_085-PO VAT dif Invoice VAT Domestic vendor_P3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 3rd June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_P2P_01_01_085 vendor_P3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_P2P_01_01_085-PO VAT  dif Invoice VAT  Domestic vendor_P3_TASE.xls"
'''''----------------------Login----------------------------

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'''----------------------Tcode MB90----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Sort order","PM_NSORT","",DT_MB90_1000_SORT_ORDER,False)
Call SetTextbox("Processing Mode","PM_VERMO","",DT_MB90_1000_PROCESSING_MODE,False)
Call SetTextbox("Article Doc\. Year","PM_MJAHR","",DT_MB90_1000_ARTICLE_DOC_YEAR,False)
Call SetTextbox("Article Document","RG_MBLNR-LOW","",DT_MB90_1000_ARTICLE_DOCUMENT,False)
Call SelectCheckbox("PM_GRID",0,DT_MB90_0120_NO_NAME,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(2)

Call SelectRowGuiGridbyRowNo("Output from Goods Movements","",1,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Process   \(Shift\+F2\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Print preview   \(Shift\+F4\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
'
'''----------------------Tcode MIGO----------------------------
'Enter the Tcode
Call SetTcode(DT_MB90_0100_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MB90_0100_OKCD_OCC1)
'Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-ACTION",DT_MB90_0010_GODYNPROACTION)
'Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_MB90_2010_GODYNPROMAT_DOC,False)
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,DT_MB90_1000_ARTICLE_DOC_YEAR,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()

Call SelectTab("TS_GOITEM","Output",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Display outputs",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTableCellContent(1,"Status","SAPDV70ATC_NAST3",DT_TOOLTIP_ROW1)
Call VerifyTableCellContent(1,"Output Type","SAPDV70ATC_NAST3",DT_ROW1_OUTPUT_TYPE)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


