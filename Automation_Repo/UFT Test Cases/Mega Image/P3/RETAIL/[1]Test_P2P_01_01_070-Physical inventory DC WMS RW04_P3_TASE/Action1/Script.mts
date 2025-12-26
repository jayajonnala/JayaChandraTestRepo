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
'.................Test Script Name : Test_P2P_01_01_070-Physical inventory DC WMS RW04_P3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 21th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_P2P_01_01_070-WMS RW04_P3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_P2P_01_01_070-Physical inventory DC WMS RW04_P3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
''''----------------------Login----------------------------

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'''Increment the parameter/reload
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call WriteRunTimeDataToExcelGlobalSheet ("DT_XYZ",1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''
''''----------------------Tcode MI24----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Article","IM_MATNR-LOW","",DT_MI24_1000_ARTICLE,False)
Call SetTextbox("Site","IM_WERKS-LOW","",DT_MI24_1000_SITE,False)
SAPGuiSession("transaction:=MI24").SAPGuiWindow("transaction:=MI24").SAPGuiButton("tooltip:=Multiple selection","index:=4").Click
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_MI24_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_MI24_3010_TABLECELL_SINGLE_VALUE_1,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",True)
Wait(1)

'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)

'Capture the screenshot
Call TakeScreenShot()

Call SelectCheckBoxByGuiLabel(DT_MI24_3010_TABLECELL_SINGLE_VALUE_0,"17","1",DT_MI24_0120_NO_NAME)
Call ClickLabel(DT_MI24_3010_TABLECELL_SINGLE_VALUE_0,"",False)

'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Physical Inventory History   \(F2\)",False)

'Capture the screenshot
Call TakeScreenShot()
Call VerifyTextBoxContent("Article","ISEG-MATNR","",DT_MI24_0710_CHECK_TEXT_OF_ARTICLE,False)
Call VerifyTextBoxContent("Site","IKPF-WERKS","",DT_MI24_0710_CHECK_TEXT_OF_SITE,False)
Call VerifyTextBoxContent("Storage Loc\.","IKPF-LGORT","",DT_MI24_0710_CHECK_TEXT_OF_STOR_LOCATION,False)
Call VerifyTextBoxContent("Phys\.inv\.status","T064T-STEXT","",Lcase(DT_MI24_0710_CHECK_TEXT_OF_PHYSINVSTATUS),False)
Call VerifyTextBoxContent("Difference qty","VM07I-DIFMG","",DT_MI24_0710_CHECK_TEXT_OF_DIFFERENCE_QTY,False)

Call ClickButton("Display Article Document   \(F5\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Overview   \(F5\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Accounting Documents   \(F7\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call GetTextboxValue("BKPF-BELNR","","DT_MI24_0750_GET_TEXT_OF_DOCUMENT_NUMBER_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Verify details
Call VerifyTextBoxContent("Company Code","BKPF-BUKRS","",DT_MI24_0750_CHECK_TEXT_OF_COMPANY_CODE,False)

Call VerifyGridCellContent("",1,"Posting Key","",DT_MI24_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",1,"Account","",DT_MI24_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",1,"Profit Center","",DT_MI24_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)

Call VerifyGridCellContent("",2,"Posting Key","",DT_MI24_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("",2,"Account","",DT_MI24_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("",2,"Profit Center","",DT_MI24_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)

Call SelectCheckBoxByGuiLabel(DT_MI24_3010_TABLECELL_SINGLE_VALUE_0,"17","1",DT_MI24_0120_NO_NAME_OCC1)
Call SelectCheckBoxByGuiLabel(DT_MI24_3010_TABLECELL_SINGLE_VALUE_1,"17","1",DT_MI24_0120_NO_NAME_OCC2)
Call ClickLabel(DT_MI24_3010_TABLECELL_SINGLE_VALUE_1,"",False)

'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Physical Inventory History   \(F2\)",False)

'Capture the screenshot
Call TakeScreenShot()
Call VerifyTextBoxContent("Site","IKPF-WERKS","",DT_MI24_0710_CHECK_TEXT_OF_SITE_OCC1,False)
Call VerifyTextBoxContent("Storage Loc\.","IKPF-LGORT","",DT_MI24_0710_CHECK_TEXT_OF_STOR_LOCATION_OCC1,False)
Call VerifyTextBoxContent("Phys\.inv\.status","T064T-STEXT","",Lcase(DT_MI24_0710_CHECK_TEXT_OF_PHYSINVSTATUS_OCC1),False)
Call VerifyTextBoxContent("Difference qty","VM07I-DIFMG","",DT_MI24_0710_CHECK_TEXT_OF_DIFFERENCE_QTY_OCC1,False)

Call ClickButton("Display Article Document   \(F5\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Overview   \(F5\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Accounting Documents   \(F7\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call GetTextboxValue("BKPF-BELNR","","DT_MI24_0750_GET_TEXT_OF_DOCUMENT_NUMBER_OCC1_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Verify details
Call VerifyTextBoxContent("Company Code","BKPF-BUKRS","",DT_MI24_0750_CHECK_TEXT_OF_COMPANY_CODE_OCC1,False)

Call VerifyGridCellContent("",1,"Posting Key","",DT_MI24_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
Call VerifyGridCellContent("",1,"Account","",DT_MI24_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR_OCC1)
Call VerifyGridCellContent("",1,"Profit Center","",DT_MI24_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR_OCC1)

Call VerifyGridCellContent("",2,"Posting Key","",DT_MI24_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL_OCC1)
Call VerifyGridCellContent("",2,"Account","",DT_MI24_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR_OCC1)
Call VerifyGridCellContent("",2,"Profit Center","",DT_MI24_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR_OCC1)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

