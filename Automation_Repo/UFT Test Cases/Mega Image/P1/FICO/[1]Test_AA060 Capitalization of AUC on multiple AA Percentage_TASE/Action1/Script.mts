
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA060 Capitalization of AUC on multiple AA Percentage
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


gstrTestCaseName = "Test_AA060 Capitalization of AUC on multiple AA Percentage"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA060 Capitalization of AUC on multiple AA Percentage.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''--------TransactionCode-ZMDPU_INFOREC_COPY----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS01_0105_ASSET_CLASS,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE,False)
Call SetTextbox("Number of similar assets","RA02S-NASSETS","",DT_AS01_0105_NUMBER_OF_SIMILAR_ASSETS,False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE_OCC1,False)
Call TakeScreenShot
Call PressEnter() 
Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION,False)
Call TakeScreenShot
Call PressEnter() 
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call SetTextbox("Business Area","ANLZ-GSBER","",DT_AS01_1145_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_COST_CENTER,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Allocations", False)
Call TakeScreenShot
Call SetTextbox("Evaluation group 1","ANLA-ORD41","",DT_AS01_1160_EVALUATION_GROUP_1,False)
Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_AS01_1160_EVALUATION_GROUP_2,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Origin", False)
Call TakeScreenShot
Call SetTextbox("Vendor","ANLA-LIFNR","",DT_AS01_1181_VENDOR,False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("WBS element","ANLA-POSNR","",DT_AS01_1182_WBS_ELEMENT,False)
Call TakeScreenShot
Call PressEnter()
Call SelectTab("TABSTRIP100", "Leasing", False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Deprec. Areas", False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call VerifyCheckBoxValue("ANLZ-XSTIL", DT_AS01_1145_CHECK_SELECTED_OF_ASSET_SHUTDOWN)
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call ClickButtonIfExist("Yes", True)

''''''Call GetTextStatusBar("DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT")
''''''Call WriteRunTimeDataToExcelGlobalSheet ("DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT",DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR_OCC1)
''''''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
''''''Call VerifyStatusBar(DT_ASSET_CHECK)

Call GetStatusBar("item1", "DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT")
Call VerifyStatusBar("The asset "&DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT&" 0 is created")
'
'''''''--------TransactionCode-/nF-90----------''''
'
Call SetTcode(DT_AS01_0105_OKCD)     
Call TakeScreenShot
Call PressEnter()     

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_AS01_0100_COMPANY_CODE,False)
Call SetTextbox("Type","BKPF-BLART","",DT_AS01_0100_TYPE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_AS01_0100_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0100_ACCOUNT,False)
Call SetTextbox("Document Date","BKPF-BLDAT","",ConvertDate(DT_AS01_0100_DOCUMENT_DATE),False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_AS01_0100_CURRENCYRATE,False)
Call TakeScreenShot
Call PressEnter()  
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0302_AMOUNT,False)
Call SetTextbox("Tax Amount","BSEG-WMWST","",DT_AS01_0302_TAX_AMOUNT,False)
Call SetTextbox("Tax Code","BSEG-MWSKZ","",DT_AS01_0302_TAX_CODE,False)
Call SetTextbox("Bus. Area","BSEG-GSBER","",DT_AS01_0302_BUS_AREA,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_AS01_0302_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT,False)
Call TakeScreenShot
Call FocusTextBox("Disc. base", "BSEG-SKFBT", False)
Call PressEnter()  
Call TakeScreenShot
Call PressEnter()  
Call SetTextbox("Bus. Area","BSEG-GSBER","",DT_AS01_0302_BUS_AREA_OCC1,False)
Call PressEnter() 
Call TakeScreenShot
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call SetTextbox("TType","RF05A-NEWBW","",DT_AS01_0302_TTYPE,False)
Call TakeScreenShot
Call PressEnter() 

Call SetTextbox("Amount","BSEG-WRBTR","",DT_AS01_0305_AMOUNT,False)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_AS01_0305_ASSIGNMENT,False)
Call TakeScreenShot
Call FocusTextBox("Text","BSEG-SGTXT",False)
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
Call VerifyTextBoxContent("C","RF05A-AZSAL","",trim(DT_AS01_0700_CHECK_TEXT_OF_C),False)
Call PressEnter()
Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call PressEnter()
Call GetStatusBar("item1", "DT_AS01_0100_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT")
Call VerifyStatusBar("Document "&DT_AS01_0100_CHECK_TEXT_OF_STATUSBAR_OCC1_OUTPUT&" was posted in company code RO02")
Call TakeScreenShot

''''''--------TransactionCode-/nAS03----------''''

Call SetTcode(DT_AS01_0100_OKCD)  
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Asset values   \(Ctrl\+F1\)", False)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR)
Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL)
Call VerifyGridCellContent("Transactions", 1, "BWATXT", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWATXT)

''''''--------TransactionCode- /NAIAB----------''''

Call SetTcode(DT_AS01_0100_OKCD_OCC1)  
Call TakeScreenShot
Call PressEnter()     

''Call SetTextbox("Addlt. area","ANLB-AFABE","","",False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)", False)
Call SelectRowGuiGrid("", 0, "Assignment", DT_AS01_0305_ASSIGNMENT, False)
Call TakeScreenShot
Call ClickButton("btn\[18\]", False)
Call SetTableData("SAPLKOBSTC_RULES", "Settlement Receiver", 1, "", "", DT_AS01_0130_TABLECELL_SETTLEMENT_RECEIVER_0, False)
Call SetTableData("SAPLKOBSTC_RULES", "Settlement Receiver", 2, "", "", DT_AS01_0130_TABLECELL_SETTLEMENT_RECEIVER_1, False)
Call SetTableData("SAPLKOBSTC_RULES", "Settlement Receiver", 3, "", "", DT_AS01_0130_TABLECELL_SETTLEMENT_RECEIVER_2, False)
Call SetTableData("SAPLKOBSTC_RULES", "%", 1, "", "", DT_AS01_0130_TABLECELL__0, False)
Call SetTableData("SAPLKOBSTC_RULES", "%", 2, "", "", DT_AS01_0130_TABLECELL__1, False)
Call SetTableData("SAPLKOBSTC_RULES", "%", 3, "", "", DT_AS01_0130_TABLECELL__2, False)
Call TakeScreenShot
Call PressEnter()
Call ClickButton("Back   \(F3\)", False)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)", False)
Call VerifyStatusBar(DT_AS01_0500_CHECK_TEXT_OF_STATUSBAR)

''''''--------TransactionCode- /NAIBU----------''''

Call SetTcode(DT_AS01_0500_OKCD)  
Call TakeScreenShot
Call PressEnter()     

Call SetTextbox("Text","\*KOMK3-SGTXT",0,DT_AS01_0100_TEXT,False)
Call SetTextbox("Document type","\*KOMK1-BLART","",DT_AS01_0100_DOCUMENT_TYPE,False)
Call SetTextbox("Reference","\*KOMK1-XBLNR","",DT_REFERENCE,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)", False)
Call VerifyGridCellContent("", 8, "ANBTR", 0, DT_AS01_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_7_ANBTR)
Call VerifyGridCellContent("", 6, "ANBTR", 0, DT_AS01_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_ANBTR)
Call VerifyGridCellContent("", 1, "ANBTR", 0, DT_AS01_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ANBTR)
Call VerifyGridCellContent("", 2, "ANBTR", 0, DT_AS01_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ANBTR)
Call VerifyGridCellContent("", 3, "ANBTR", 0, DT_AS01_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_ANBTR)
Call VerifyGridCellContent("", 4, "ANBTR", 0, DT_AS01_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ANBTR)
Call VerifyGridCellContent("", 5, "ANBTR", 0, DT_AS01_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_ANBTR)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
Call SelectCheckbox("LKO74-TESTLAUF", 0, DT_AS01_0100_TEST_RUN, False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)", False)
Call GetStatusBar("item1", "DT_AS01_0500_CHECK_TEXT_OF_STATUSBAR_OCC2_OUTPUT")
Call VerifyStatusBar("Asset transaction posted with document no. "&DT_AS01_0500_CHECK_TEXT_OF_STATUSBAR_OCC2_OUTPUT)

''''''--------TransactionCode- /NAS03----------''''

Call SetTcode(DT_AS01_0500_OKCD_OCC1)  
Call TakeScreenShot
Call PressEnter() 
Call PressEnter() 
Call ClickButton("Asset values   \(Ctrl\+F1\)", False)
Call VerifyGridCellContent("Transactions", 4, "VERAENDERUNG", 0, "")
'Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, trim(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR))
'Call VerifyGridCellContent("Transactions", 2, "BWASL", 0, trim(DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWASL))
Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS01_0100_ASSET,False)
Call TakeScreenShot
Call PressEnter() 
Call VerifyTextBoxContent("Capitalized on","ANLA-AKTIV","",ConvertDate(DT_AS01_1142_CHECK_TEXT_OF_CAPITALIZED_ON_OCC1),False)
Call VerifyTextBoxContent("First acquisition on","ANLA-ZUGDT","",ConvertDate(DT_AS01_1142_CHECK_TEXT_OF_FIRST_ACQUISITION_ON_OCC1),False)
Call TakeScreenShot
Call ClickButton("Asset values   \(Ctrl\+F1\)", False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
Call ClickButton("Back   \(F3\)", False)
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS01_0100_ASSET_OCC1,False)
Call TakeScreenShot
Call PressEnter() 
Call VerifyTextBoxContent("Capitalized on","ANLA-AKTIV","",ConvertDate(DT_AS01_1142_CHECK_TEXT_OF_CAPITALIZED_ON_OCC2),False)
Call VerifyTextBoxContent("First acquisition on","ANLA-ZUGDT","",ConvertDate(DT_AS01_1142_CHECK_TEXT_OF_FIRST_ACQUISITION_ON_OCC2),False)
Call TakeScreenShot
Call ClickButton("Asset values   \(Ctrl\+F1\)", False)
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_2)
Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_AS01_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL_OCC2)
Call TakeScreenShot


Call LogOff()

Call FinalStatus ()






'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [8,3437477]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//




