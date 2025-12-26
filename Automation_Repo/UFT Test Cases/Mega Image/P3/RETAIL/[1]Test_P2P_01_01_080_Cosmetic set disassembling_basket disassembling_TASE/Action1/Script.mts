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

'reload DS to update dates and calculations
''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_P2P_01_01_080_Cosmetic set disassembling_basket disassembling_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 4th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_P2P_01_01_080_basket disassembling_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_P2P_01_01_080_Cosmetic set disassembling  basket disassembling_TASE.xls"
'''''----------------------Login----------------------------

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
''''----------------------Tcode CS03----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Article","RC29N-MATNR","",DT_CS03_0100_ARTICLE,False)
Call SetTextbox("BOM Usage","RC29N-STLAN","",DT_CS03_0100_BOM_USAGE,False)
Call SetTextbox("Valid From","RC29N-DATUV","",ConvertDate(DT_CS03_0100_VALID_FROM),False)
Call SetTextbox("Valid to","RC29N-DATUB","",ConvertDate(DT_CS03_0100_VALID_TO),False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

Call GetTableCellData("SAPLCSDITCMAT","Component",1,"Item","0010",DT_CS03_0152_CHECK_TEXT_OF_TABLECELL_COMPONENT_0_OUTPUT,False)
Call GetTableCellData("SAPLCSDITCMAT","Component",2,"Item","0010",DT_CS03_0152_CHECK_TEXT_OF_TABLECELL_COMPONENT_1_OUTPUT,False)
Call GetTableCellData("SAPLCSDITCMAT","Component",3,"Item","0010",DT_CS03_0152_CHECK_TEXT_OF_TABLECELL_COMPONENT_2_OUTPUT,False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call VerifyTableCellContent(1,"Component","SAPLCSDITCMAT",DT_CS03_0152_CHECK_TEXT_OF_TABLECELL_COMPONENT_0_OCC1)
Call VerifyTableCellContent(2,"Component","SAPLCSDITCMAT",DT_CS03_0152_CHECK_TEXT_OF_TABLECELL_COMPONENT_1_OCC1)
Call VerifyTableCellContent(3,"Component","SAPLCSDITCMAT",DT_CS03_0152_CHECK_TEXT_OF_TABLECELL_COMPONENT_2_OCC1)
''
''''----------------------Tcode MB52----------------------------
'Enter the Tcode
Call SetTcode(DT_CS03_2150_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_CS03_2150_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Article","MATNR-LOW","",DT_CS03_1000_ARTICLE,False)
Call SetTextbox("Site","WERKS-LOW","",DT_CS03_1000_SITE,False)
Call SetTextbox("Storage Location","LGORT-LOW","",DT_CS03_0120_CHECK_TEXT_OF_NO_NAME_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyifGuiLabelExists(DT_CS03_0120_CHECK_TEXT_OF_NO_NAME)
Call VerifyifGuiLabelExists(DT_CS03_0120_CHECK_TEXT_OF_RW04)
Call VerifyifGuiLabelExists(DT_CS03_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
Call VerifyifGuiLabelExists(DT_CS03_0120_CHECK_TEXT_OF_HEALTHY_STOCK)
'''
'''''----------------------Tcode ZMDIM_ART_DISASSEMBL----------------------------
'Enter the Tcode
Call SetTcode(DT_CS03_0120_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_CS03_0120_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Article","P_MATNR","",DT_CS03_1000_ARTICLE_OCC1,False)
Call SetTextbox("Site","S_WERKS-LOW","",DT_CS03_1000_SITE_OCC1,False)
Call SetTextbox("Document Date","P_BLDAT","",ConvertDate(DT_CS03_1000_DOCUMENT_DATE),False)
Call SetTextbox("Posting Date","P_BUDAT","",ConvertDate(DT_CS03_1000_POSTING_DATE),False)
Call SelectCheckbox("P_TEST",0,DT_CS03_1000_TEST_RUN,False)
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectRowGuiGridbyRowNo("","",1,False)
Call ClickButton("Show / Hide Components for Selected Article   \(F2\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Yes",True)
'Capture the screenshot
Call TakeScreenShot()

Call GetLabelContentByRefLabel("Message","0","-32","DT_CS03_0120_CHECK_TEXT_OF_GOODS_MOVEMENT_NUMBER_4943934352_WAS_CREATED_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Capture the screenshot
Call TakeScreenShot()

'Call VerifyifGuiLabelExists(LCase(DT_CS03_0120_CHECK_TEXT_OF_GOODS_MOVEMENT_NUMBER_4943934352_WAS_CREATED))

''
''''----------------------Tcode MIGO----------------------------
'Enter the Tcode
Call SetTcode(DT_CS03_0120_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_CS03_0120_OKCD_OCC1)
'Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-ACTION",DT_CS03_0010_GODYNPROACTION)
'Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_CS03_2010_GODYNPROMAT_DOC,False)
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,DT_CS03_2010_YEAR,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()
Call PressEnter()

Call VerifyTableCellContent(1,"Movement type","SAPLMIGOTV_GOITEM",DT_CS03_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_0)
Call VerifyTableCellContent(2,"Movement type","SAPLMIGOTV_GOITEM",DT_CS03_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_1)
Call VerifyTableCellContent(3,"Movement type","SAPLMIGOTV_GOITEM",DT_CS03_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_2)
Call VerifyTableCellContent(4,"Movement type","SAPLMIGOTV_GOITEM",DT_CS03_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_3)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


