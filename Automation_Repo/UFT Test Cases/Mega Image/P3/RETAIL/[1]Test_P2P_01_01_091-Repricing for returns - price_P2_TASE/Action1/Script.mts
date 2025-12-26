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
'.................Test Script Name : Test_P2P_01_01_091-Repricing for returns - price_P2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 14th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_P2P_01_01_091_P2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_P2P_01_01_091-Repricing for returns - price_P2_TASE.xls"

''''----------------------Login----------------------------

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()   

'''Increment the parameter/reload
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call WriteRunTimeDataToExcelGlobalSheet ("DT_XYZ",1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode ME12----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxNoLabel("EINA-LIFNR", "", DT_ME12_0100_VENDOR, False)

''Call SetTextbox("Vendor","EINA-LIFNR","",DT_ME12_0100_VENDOR,False)
Call SetTextbox("Article","EINA-MATNR","",DT_ME12_0100_ARTICLE,False)
Call SetTextbox("Purchasing Org\.","EINE-EKORG","",DT_ME12_0100_PURCHASING_ORG,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 

Call ClickButtonIfExist("Conditions   \(F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Select time frame   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Valid From","RV13A-DATAB","",DT_ME12_0201_VALID_FROM,False)
Call SetTextbox("Valid to","RV13A-DATBI","",DT_ME12_0201_VALID_TO,False)

Call SetTableDataNoRef("SAPMV13ATCTRL_D0201","Amount",1,DT_ME12_0201_TABLECELL_AMOUNT_0,False)

Call PressEnter() 
Call PressEnter() 
'Capture the screenshot
Call TakeScreenShot()

'Click on Save Buton
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Enter   \(F5\)",True)

'save it to data sheet
Call GetStatusBar("item1","DT_DOCUMENT_NO_OUTPUT")
'reload data sheet
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'verify statusbar with datasheet feed
Call VerifyStatusBar(DT_ME12_0100_CHECK_TEXT_OF_STATUSBAR_OCC1)
'
'''----------------------Tcode ME13----------------------------
'Enter the Tcode
Call SetTcode(DT_ME12_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ME12_0100_OKCD)
'Capture the screenshot
Call TakeScreenShot()

''Call SetTextbox("Vendor","EINA-LIFNR","",DT_ME12_0100_VENDOR_OCC1,False)
Call SetTextboxnoLabel("EINA-LIFNR","",DT_ME12_0100_VENDOR_OCC1,False)
Call SetTextbox("Article","EINA-MATNR","",DT_ME12_0100_ARTICLE_OCC1,False)
Call SetTextbox("Purchasing Org\.","EINE-EKORG","",DT_ME12_0100_PURCHASING_ORG_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter() 

Call ClickButtonIfExist("Conditions   \(F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectCellGuiTable("SAPLV14ATCTRL_D0102","Valid From","Valid From",DT_ME12_0201_CHECK_TEXT_OF_VALID_FROM,True)

Call ClickButtonIfExist("Select time frame   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Valid From","RV13A-DATAB","",DT_ME12_0201_CHECK_TEXT_OF_VALID_FROM,False)
Call VerifyTextBoxContent("Valid to","RV13A-DATBI","",DT_ME12_0201_CHECK_TEXT_OF_VALID_TO,False)

Call VerifyGridCellContent("",1,"Supplier","",DT_ME12_0010_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIFNR)
'Call VerifyGridCellContent("",1,"Vendor","",DT_ME12_0010_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIFNR)
Call VerifyGridCellContent("",1,"Article","",DT_ME12_0010_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MATNR)
Call VerifyGridCellContent("",1,"Purch. Organization","",DT_ME12_0010_CHECK_GETCELLVALUE_OF_GRIDCELL_0_EKORG)

Call VerifyTableCellContent(1,"CnTy","SAPMV13ATCTRL_D0201",DT_ME12_0201_CHECK_TEXT_OF_TABLECELL_CNTY_0)
Call VerifyTableCellContent(1,"Name","SAPMV13ATCTRL_D0201",Lcase(DT_ME12_0201_CHECK_TEXT_OF_TABLECELL_NAME_0))
Call VerifyTableCellContent(1,"Amount","SAPMV13ATCTRL_D0201",DT_ME12_0201_CHECK_TEXT_OF_TABLECELL_AMOUNT_0)
'
'''----------------------Tcode MEI1----------------------------
'Enter the Tcode
Call SetTcode(DT_ME12_0201_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ME12_0201_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Purchasing Documents","S_EBELN-LOW","",DT_ME12_1000_PURCHASING_DOCUMENTS_OCC1,False)
Call SetTextbox("Document Type","S_BSART-LOW","",DT_ME12_1000_DOCUMENT_TYPE,False)
Call SetTextBoxnoLabel("S_LIFNR-LOW","",DT_ME12_1000_VENDOR,False)
'Call SetTextbox("Vendor","S_LIFNR-LOW","",DT_ME12_1000_VENDOR,False)
Call SetTextbox("Purchasing Organization","S_EKORG-LOW","",DT_ME12_1000_PURCHASING_ORGANIZATION,False)
Call SetTextbox("Material","S_MATNR-LOW","","",False)

SAPGuiSession("transaction:=MEI1").SAPGuiWindow("transaction:=MEI1").SAPGuiButton("tooltip:=Multiple selection","index:=3").Click
Wait(2)
Call SelectTab("TAB_STRIP","Exclude Single Values",True)
Call SetTableDataNoRef("SAPLALDBSINGLE_E","Single value",1,DT_ME12_3030_TABLECELL_SINGLE_VALUE_0,True)'exclude
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",True)
Wait(1)

Call ClickButton("Execute   \(F8\)",False)

'Capture the screenshot
Call TakeScreenShot()
Call VerifyTextBoxContent("Information Message","MESSTXT1","",Lcase(DT_ME12_0010_CHECK_TEXT_OF_MESSTXT1),True)
Call ClickButton("Continue   \(Enter\)",True)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


