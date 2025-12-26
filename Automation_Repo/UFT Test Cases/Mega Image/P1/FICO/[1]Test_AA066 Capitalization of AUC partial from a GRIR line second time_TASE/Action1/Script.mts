
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA066 Capitalization of AUC partial from a GRIR line second time
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

gstrTestCaseName = "Test_AA066"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA066 Capitalization of AUC partial from a GRIR line second time.xls"
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

Call SetTextbox("Company Code","AICOM-BUKRS","",DT_AIAB_0110_COMPANY_CODE,False)
Call SetTextbox("Sub-number","AICOM-ANLN2","",DT_AIAB_0110_SUBNUMBER,False)
Call SetTextbox("Layout","P_DISVAR","",DT_AIAB_0110_LAYOUT,False)
Call SetTextbox("Addit\. area","ANLB-AFABE","","",False)
Call SetTextbox("Asset","AICOM-ANLN1","",DT_AIAB_0110_ASSET,False)
Call TakeScreenShot
Call ClickButton("btn\[6\]", False)  ''Further Selection Criteria
Call SetTextbox("Document Number","%%DYN006-LOW","",DT_AIAB_1105_DOCUMENT_NUMBER,True)
Call ClickButton("Save   \(Ctrl\+S\)", True)
Call ClickButton("Execute   \(F8\)", False)

Call SelectRowGuiGridbyRowNo("", 0, 1, False)

Call ClickButton("Enter distribution rules   \(Shift\+F6\)", False)			
Call SetTableData("SAPLKOBSTC_RULES", "Settlement Receiver", 1, "", "", DT_AIAB_0130_TABLECELL_SETTLEMENT_RECEIVER_0, False)
Call SetTableData("SAPLKOBSTC_RULES", "Settlement Receiver", 2, "", "", DT_AIAB_0130_TABLECELL_SETTLEMENT_RECEIVER_1, False)
Call SetTableData("SAPLKOBSTC_RULES", "Settlement Receiver", 3, "", "", DT_AIAB_0130_TABLECELL_SETTLEMENT_RECEIVER_2, False)
Call SetTableData("SAPLKOBSTC_RULES", "Amount", 1, "", "", DT_AIAB_0130_TABLECELL_AMOUNT_0, False)
Call SetTableData("SAPLKOBSTC_RULES", "Amount", 2, "", "", DT_AIAB_0130_TABLECELL_AMOUNT_1, False)
Call SetTableData("SAPLKOBSTC_RULES", "Amount", 3, "", "", DT_AIAB_0130_TABLECELL_AMOUNT_2, False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
Call TakeScreenShot
Call ClickButton("Save   \(Ctrl\+S\)", False)
Call VerifyStatusBar(DT_AIAB_0500_CHECK_TEXT_OF_STATUSBAR)

''''''--------TransactionCode- /NAIBU----------''''

Call SetTcode(DT_AIAB_0500_OKCD)  
Call TakeScreenShot
Call PressEnter()     

Call SetTextbox("Asset Val\. Date","ANEP-BZDAT","",ConvertDate(DT_AIAB_0100_ASSET_VAL_DATE),False)
Call SetTextbox("Posting Date","ANEK-BUDAT","",ConvertDate(DT_AIAB_0100_POSTING_DATE),False)
Call SetTextbox("Period","ANBZ-MONAT","",DT_AIAB_0100_PERIOD,False)
Call SetTextbox("Document Date","ANEK-BLDAT","",ConvertDate(DT_AIAB_0100_DOCUMENT_DATE),False)
Call SetTextboxNoLabel("ANLA-ANLN2","",DT_AIAB_0100_SUBNUMBER,False)
'Call SetTextbox("Text","*KOMK3-SGTXT","",DT_AS01_0100_TEXT,False)
Call SetTextbox("Document type","\*KOMK1-BLART","",DT_AIAB_0100_DOCUMENT_TYPE,False)
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AIAB_0100_ASSET,False)
Call SetTextbox("Company code","ANLA-BUKRS","",DT_AIAB_0100_COMPANY_CODE,False)
Call SelectCheckbox("LKO74-TESTLAUF", 0, DT_AIAB_0100_TEST_RUN, False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)", False)
Call GetStatusBar("item1", "DT_AIAB_0500_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Asset transaction posted with document no. "&DT_AIAB_0500_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)
Call VerifyGridCellContent("", 1, "BWASL", 0, DT_AIAB_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL)
Call VerifyGridCellContent("", 3, "BWASL", 0, DT_AIAB_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BWASL)
Call VerifyGridCellContent("", 5, "BWASL", 0, DT_AIAB_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BWASL)
Call VerifyGridCellContent("", 2, "ANBTR", 0, DT_AIAB_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_ANBTR)
Call VerifyGridCellContent("", 4, "ANBTR", 0, DT_AIAB_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_ANBTR)
Call VerifyGridCellContent("", 6, "ANBTR", 0, DT_AIAB_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_ANBTR)
Call TakeScreenShot

''''''--------TransactionCode- /NAS03----------''''

Call SetTcode(DT_AIAB_0500_OKCD_OCC1)  
Call PressEnter() 
Call TakeScreenShot

Call SetTextbox("Asset","ANLA-ANLN1","",DT_AIAB_0100_ASSET_OCC2,False)
Call SetTextboxNoLabel("ANLA-ANLN2","",DT_AIAB_0100_SUBNUMBER_OCC2,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AIAB_0100_COMPANY_CODE_OCC1,False)
Call TakeScreenShot
Call ClickButton("Asset values   \(Ctrl\+F1\)", False)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE)

Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR)
Call VerifyGridCellContent("Transactions", 3, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BUBTR)
Call VerifyGridCellContent("Transactions", 4, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BUBTR)
Call VerifyGridCellContent("Transactions", 5, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BUBTR)
Call VerifyGridCellContent("Transactions", 6, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BUBTR)
Call VerifyGridCellContent("Transactions", 7, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_6_BUBTR)

Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL)
Call VerifyGridCellContent("Transactions", 2, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWASL)
Call VerifyGridCellContent("Transactions", 3, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BWASL)
Call VerifyGridCellContent("Transactions", 4, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BWASL)
Call VerifyGridCellContent("Transactions", 5, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BWASL)
Call VerifyGridCellContent("Transactions", 6, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BWASL)
Call VerifyGridCellContent("Transactions", 7, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_6_BWASL)
Call ActivateNodeGuiTree(0,"#1;#2")
Call TakeScreenShot
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE)

Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR)
Call VerifyGridCellContent("Transactions", 3, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BUBTR)
Call VerifyGridCellContent("Transactions", 4, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BUBTR)
Call VerifyGridCellContent("Transactions", 5, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BUBTR)
Call VerifyGridCellContent("Transactions", 6, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BUBTR)
Call VerifyGridCellContent("Transactions", 7, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_6_BUBTR)

Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL)
Call VerifyGridCellContent("Transactions", 2, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWASL)
Call VerifyGridCellContent("Transactions", 3, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BWASL)
Call VerifyGridCellContent("Transactions", 4, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BWASL)
Call VerifyGridCellContent("Transactions", 5, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BWASL)
Call VerifyGridCellContent("Transactions", 6, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BWASL)
Call VerifyGridCellContent("Transactions", 7, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_6_BWASL)

Call ActivateNodeGuiTree(0,"#1;#1")
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AIAB_0201_ASSET,False)
Call TakeScreenShot
Call PressEnter()

Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC2)
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT))
Call VerifyGridCellContent("Transactions", 2, "BZDAT", 0, ConvertDate(DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BZDAT))
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC2)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC2)
Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL_OCC2)
Call VerifyGridCellContent("Transactions", 2, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWASL_OCC2)
Call TakeScreenShot
Call ActivateNodeGuiTree(0,"#1;#2")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC2)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC2)
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT))
Call VerifyGridCellContent("Transactions", 2, "BZDAT", 0, ConvertDate(DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BZDAT))
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC2)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC2)
Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL_OCC2)
Call VerifyGridCellContent("Transactions", 2, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWASL_OCC2)
Call TakeScreenShot
Call ActivateNodeGuiTree(0,"#1;#1")
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AIAB_0201_ASSET_OCC1,False)
Call TakeScreenShot
Call PressEnter()
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC4)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC4)
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT_OCC2))
Call VerifyGridCellContent("Transactions", 2, "BZDAT", 0, ConvertDate(DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BZDAT_OCC2))
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC4)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC4)
Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL_OCC4)
Call VerifyGridCellContent("Transactions", 2, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWASL_OCC4)
Call TakeScreenShot
Call ActivateNodeGuiTree(0,"#1;#2")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC4)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC4)
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT_OCC2))
Call VerifyGridCellContent("Transactions", 2, "BZDAT", 0, ConvertDate(DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BZDAT_OCC2))
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC4)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC4)
Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL_OCC4)
Call VerifyGridCellContent("Transactions", 2, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWASL_OCC4)
Call TakeScreenShot

Call ActivateNodeGuiTree(0,"#1;#1")
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AIAB_0201_ASSET_OCC2,False)
Call TakeScreenShot
Call PressEnter()
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC6)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC6)
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT_OCC4))
Call VerifyGridCellContent("Transactions", 2, "BZDAT", 0, ConvertDate(DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BZDAT_OCC4))
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC6)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC6)
Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL_OCC6)
Call VerifyGridCellContent("Transactions", 2, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWASL_OCC6)
Call TakeScreenShot
Call ActivateNodeGuiTree(0,"#1;#2")
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "VERAENDERUNG", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_VERAENDERUNG_OCC6)
Call VerifyGridCellContent("Planned values IFRS APC, depreciation", 4, "JENDE", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_3_JENDE_OCC6)
Call VerifyGridCellContent("Transactions", 1, "BZDAT", 0, ConvertDate(DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BZDAT_OCC4))
Call VerifyGridCellContent("Transactions", 2, "BZDAT", 0, ConvertDate(DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BZDAT_OCC4))
Call VerifyGridCellContent("Transactions", 1, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUBTR_OCC6)
Call VerifyGridCellContent("Transactions", 2, "BUBTR", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR_OCC6)
Call VerifyGridCellContent("Transactions", 1, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BWASL_OCC6)
Call VerifyGridCellContent("Transactions", 2, "BWASL", 0, DT_AIAB_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BWASL_OCC6)

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




