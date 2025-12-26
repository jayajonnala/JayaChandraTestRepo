
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA039 Capitalization of AUC last year acqusition
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


gstrTestCaseName = "Test_AA039 Capitalization of AUC last year acqusition"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA010- Asset Creation - create asset directly in SAP.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


''--------TransactionCode-S_ALR_87011966----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Company Code","BUKRS-LOW","",DT_S_ALR_87011966_1000_COMPANY_CODE,False)
Call TakeScreenShot
Call ClickButton("All Selections   \(Shift\+F7\)",False)
Call TakeScreenShot
Call SelectRadioButton("XEINZEL","List assets",False)
Call TakeScreenShot
Call SetTextbox("Asset class","SO_ANLKL-LOW","",DT_S_ALR_87011966_1000_ASSET_CLASS,False)
Call SetTextbox("Depreciation area","BEREICH1","",DT_S_ALR_87011966_1000_DEPRECIATION_AREA,False)
Call TakeScreenShot
Call PressEnter()     

Call SetTextbox("Sort Variant","SRTVR","",DT_S_ALR_87011966_1000_SORT_VARIANT,False)
Call SetTextbox("Display variant","P_VARI","",DT_S_ALR_87011966_1000_DISPLAY_VARIANT,False)
Call SetTextbox("Capitalization date","SO_AKTIV-LOW","",ConvertDate(DT_S_ALR_87011966_1000_CAPITALIZATION_DATE),False)
Call TakeScreenShot
Call SetTextboxNoLabel("SO_AKTIV-HIGH",0,ConvertDate(DT_S_ALR_87011966_1000_TO),False)
Call TakeScreenShot
Call ClickButton("%_SO_KANSW_%_APP_%-VALU_PUSH",False)
Call TakeScreenShot
Call SelectTab("TAB_STRIP", "Exclude Single Values", True)
Call TakeScreenShot
Call ClickCellTableByRowNo("SAPLALDBSINGLE_E","Selection Options",1,True)
Call TakeScreenShot
Call ClickButton("Copy   \(Enter\)",True) ' RGB 'RGD
'Call ClickButton("Continue   \(Enter\)",True) 'R1E
Call TakeScreenShot
Call ClickButton("Copy   \(F8\)",True)
Call TakeScreenShot

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call GetGridContentByTitle("",0,"Asset",3,"DT_S_ALR_87011966_GETCELLVALUE_GRIDCELL_1_ANLN0_OUTPUT")
Call GetGridContentByTitle("",0,"Sub-number",3,"DT_S_ALR_87011966_GETCELLVALUE_GRIDCELL_1_ANLN2_OUTPUT")
'
'
''''--------TransactionCode-AS01----------''''

Call SetTcode(DT_S_ALR_87011966_0500_OKCD)     
Call PressEnter()     
Call CheckTCodeScreen(DT_S_ALR_87011966_0500_OKCD)
Call TakeScreenShot

Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_S_ALR_87011966_0105_ASSET_CLASS,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_S_ALR_87011966_0105_COMPANY_CODE,False)
Call TakeScreenShot
Call SetTextbox("Asset","RA02S-RANL1","",DT_S_ALR_87011966_GETCELLVALUE_GRIDCELL_1_ANLN0_OUTPUT,False)
Call SetTextbox("Sub-number","RA02S-RANL2","",DT_S_ALR_87011966_GETCELLVALUE_GRIDCELL_1_ANLN2_OUTPUT,False)
Call TakeScreenShot

Call PressEnter() 
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot

Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call PressEnter()
Call SelectTab("TABSTRIP100", "Allocations", False)
Call TakeScreenShot
Call PressEnter() 
Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_S_ALR_87011966_1160_EVALUATION_GROUP_2,False)
Call SetTextbox("Evaluation group 1","ANLA-ORD41","",FormatBlank(DT_S_ALR_87011966_1160_EVALUATION_GROUP_1),False)

Call TakeScreenShot
Call PressEnter()
'
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
''Call PressEnter() 
Call TakeScreenShot
wait 5
''Call PressEnter()
Call GetStatusBar("item1", "DT_S_ALR_87011966_0105_STATUSBAR_OUTPUT")
Call VerifyStatusBar("The asset "& DT_S_ALR_87011966_0105_STATUSBAR_OUTPUT & " 0 is created")
'
''''--------TransactionCode-/AIAB----------''''
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_S_ALR_87011966_0105_OKCD)     
Call PressEnter()     
Call CheckTCodeScreen(DT_S_ALR_87011966_0105_OKCD)
Call TakeScreenShot

Call SetTextbox("Company Code","AICOM-BUKRS","",DT_S_ALR_87011966_0110_COMPANY_CODE,False)
Call SetTextbox("Asset","AICOM-ANLN1","",DT_S_ALR_87011966_0110_ASSET,False)
Call SetTextbox("Sub-number","AICOM-ANLN2","",DT_S_ALR_87011966_0110_SUBNUMBER,False)
Call SetTextbox("Addit. area","ANLB-AFABE","",Formatblank(DT_S_ALR_87011966_0110_ADDIT_AREA),False)

Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call GetGridContentByTitle("",0,"Amount posted",1,"DT_AMOUNT_AIAB_OUTPUT")

Call SelectRowGuiGridbyRowNo("",0,1,False)
Call TakeScreenShot
Call ClickButton("Enter distribution rules   \(Shift\+F6\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Change",True)
Call TakeScreenShot
Call SetTableData("SAPLKOBSTC_RULES","Settlement Receiver","1","","",DT_S_ALR_87011966_0130_TABLECELL_SETTLEMENT_RECEIVER_0,False)
Call SetTableData("SAPLKOBSTC_RULES","%","1","","",DT_S_ALR_87011966_0130_TABLECELL__0,False)

Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Back   \(F3\)",False)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "AMPEL", 0, DT_S_ALR_87011966_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AMPEL)
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot

''--------TransactionCode-/AIBU----------''''
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_S_ALR_87011966_0500_OKCD_OCC1)     
Call PressEnter()     
Call CheckTCodeScreen(DT_S_ALR_87011966_0500_OKCD_OCC1)
Call TakeScreenShot

Call SelectCheckbox("LKO74-TESTLAUF",0,DT_S_ALR_87011966_0100_TEST_RUN,False)
Call TakeScreenShot
Call SetTextbox("Company code","ANLA-BUKRS","",DT_S_ALR_87011966_0100_COMPANY_CODE,False)

Call SetTextbox("Asset Val. Date","ANEP-BZDAT","",ConvertDate(DT_S_ALR_87011966_0100_ASSET_VAL_DATE),False)
Call SetTextbox("Posting Date","ANEK-BUDAT","",ConvertDate(DT_S_ALR_87011966_0100_POSTING_DATE),False)
Call SetTextbox("Period","ANBZ-MONAT","",DT_S_ALR_87011966_0100_PERIOD,False)
Call SetTextbox("Document Date","ANEK-BLDAT","",ConvertDate(DT_S_ALR_87011966_0100_DOCUMENT_DATE),False)
Call SetTextbox("Document type","\*KOMK1-BLART","",DT_S_ALR_87011966_0100_DOCUMENT_TYPE,False)
Call SetTextbox("Asset","ANLA-ANLN1","",DT_S_ALR_87011966_0100_ASSET,False)
Call SetTextbox("Subnumber","ANLA-ANLN2","",DT_S_ALR_87011966_0100_SUBNUMBER,False)

Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
Call VerifyGridCellContent("", 1, "BWASL", 0, DT_S_ALR_87011966_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL)
Call GetStatusBar("item1", "DT_S_ALR_87011966_0500_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Asset transaction posted with document no. " & DT_S_ALR_87011966_0500_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)



''''''''--------TransactionCode-AS03----------''''
''
Call SetTcode(DT_S_ALR_87011966_0500_OKCD_OCC2)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Asset","ANLA-ANLN1","",DT_S_ALR_87011966_0100_ASSET_OCC1,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_S_ALR_87011966_0100_SUBNUMBER_OCC1,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_S_ALR_87011966_0100_COMPANY_CODE_OCC1,False)
Call TakeScreenShot
Call ClickButton("Asset values   \(Ctrl\+F1\)",False)
Call TakeScreenShot
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_VALDATE_1))
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AMOUNT_1)
Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_S_ALR_87011966_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL)
'
Call ClickButton("Exit   \(Shift\+F3\)",False)
Call TakeScreenShot

Call SetTextbox("Asset","ANLA-ANLN1","",DT_S_ALR_87011966_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ANLN0,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_S_ALR_87011966_0100_SUBNUMBER_OCC1,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_S_ALR_87011966_0100_COMPANY_CODE_OCC1,False)
Call TakeScreenShot
Call ClickButton("Asset values   \(Ctrl\+F1\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_VALDATE_2))
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AMOUNT_2)
Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_TTYPE_2)

Call LogOff()

Call FinalStatus ()






