
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA011- Asset Creation - create asset directly in SAP with refera
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


gstrTestCaseName = "Test_AA011- Asset Creation - create asset directly in SAP with refera"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA011- Asset Creation - create asset directly in SAP with refera.xls"
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
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+5)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS01_0105_ASSET_CLASS,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE,False)
Call SetTextbox("Number of similar assets","RA02S-NASSETS","",DT_AS01_0105_NUMBER_OF_SIMILAR_ASSETS,False)
Call SetTextbox("Asset","RA02S-RANL1","",DT_AS01_0105_ASSET,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)", True)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_COST_CENTER,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Allocations", False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Evaluation group 1","ANLA-ORD41","",DT_AS01_1160_EVALUATION_GROUP_1,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Origin", False)
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Maintain",True)
Call TakeScreenShot
Call SetTableData("SAPLAISTTC_MULTIPLE_ASSETS", "Inventory no.", 1, "", "", DT_AS01_2300_TABLECELL_INVENTORY_NO_0, False)
Call SetTableData("SAPLAISTTC_MULTIPLE_ASSETS", "Inventory no.", 2, "", "", DT_AS01_2300_TABLECELL_INVENTORY_NO_1, False)
Call SetTableData("SAPLAISTTC_MULTIPLE_ASSETS", "Inventory no.", 3, "", "", DT_AS01_2300_TABLECELL_INVENTORY_NO_2, False)
Call SetTableData("SAPLAISTTC_MULTIPLE_ASSETS", "Inventory no.", 4, "", "", DT_AS01_2300_TABLECELL_INVENTORY_NO_3, False)
Call SetTableData("SAPLAISTTC_MULTIPLE_ASSETS", "Inventory no.", 5, "", "", DT_AS01_2300_TABLECELL_INVENTORY_NO_4, False)
Call SetTableData("SAPLAISTTC_MULTIPLE_ASSETS", "WBS element", 1, "", "", DT_AS01_2300_TABLECELL_WBS_ELEMENT_0, False)
Call SetTableData("SAPLAISTTC_MULTIPLE_ASSETS", "WBS element", 2, "", "", DT_AS01_2300_TABLECELL_WBS_ELEMENT_1, False)
Call SetTableData("SAPLAISTTC_MULTIPLE_ASSETS", "WBS element", 3, "", "", DT_AS01_2300_TABLECELL_WBS_ELEMENT_2, False)
Call SetTableData("SAPLAISTTC_MULTIPLE_ASSETS", "WBS element", 4, "", "", DT_AS01_2300_TABLECELL_WBS_ELEMENT_3, False)
Call SetTableData("SAPLAISTTC_MULTIPLE_ASSETS", "WBS element", 5, "", "", DT_AS01_2300_TABLECELL_WBS_ELEMENT_4, False)
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item1", "DT_NUMBER_ASSET_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_NUMBER_ASSET_OUTPUT",DT_NUMBER_ASSET)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call GetStatusBar("item2", "DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT",DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call GetStatusBar("item3", "DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC2_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC2_OUTPUT",DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Call VerifyStatusBar("Assets "&DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT&" to "&DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC2_OUTPUT&" have been created")
Call VerifyStatusBar(Lcase(DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR))
Call ClickButton("Exit   \(Shift\+F3\)", False)
''Call WriteRunTimeDataToExcelGlobalSheet ("DT_COUNT",Cint(DT_COUNT)-Cint(DT_COUNT))

''''''--------TransactionCode-AS03----------''''

Call SetTcode(DT_AS01_0100_OKCD)  
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS01_0100_ASSET,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS01_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0100_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter()
Call VerifyTextBoxContent("Description","ANLA-TXT50","",lcase(DT_AS01_1140_CHECK_TEXT_OF_DESCRIPTION),False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call VerifyTextBoxContent("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_CHECK_TEXT_OF_COST_CENTER,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Allocations", False)
Call VerifyTextBoxContent("Evaluation group 1","ANLA-ORD41","",DT_AS01_1160_CHECK_TEXT_OF_EVALUATION_GROUP_1,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Origin", False)
Call VerifyTextBoxContent("WBS element","ANLA-POSNR","",DT_AS01_1182_CHECK_TEXT_OF_WBS_ELEMENT,False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
'Call WriteRunTimeDataToExcelGlobalSheet ("DT_COUNT",Cint(DT_COUNT)+1)

''''''--------TransactionCode-AS03----------''''

Call SetTcode(DT_AS01_0100_OKCD)  
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS01_0100_ASSET,False)
Call SetTextbox("Asset","ANLA-ANLN1","",DT_ASSET_1,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS01_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0100_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter()
Call VerifyTextBoxContent("Description","ANLA-TXT50","",lcase(DT_AS01_1140_CHECK_TEXT_OF_DESCRIPTION),False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call VerifyTextBoxContent("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_CHECK_TEXT_OF_COST_CENTER,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Allocations", False)
Call VerifyTextBoxContent("Evaluation group 1","ANLA-ORD41","",DT_AS01_1160_CHECK_TEXT_OF_EVALUATION_GROUP_1,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Origin", False)
Call VerifyTextBoxContent("WBS element","ANLA-POSNR","",DT_AS01_1182_CHECK_TEXT_OF_WBS_ELEMENT,False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
''Call WriteRunTimeDataToExcelGlobalSheet ("DT_COUNT",Cint(DT_COUNT)+1)

''''''--------TransactionCode-AS03----------''''

Call SetTcode(DT_AS01_0100_OKCD)  
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS01_0100_ASSET,False)
Call SetTextbox("Asset","ANLA-ANLN1","",DT_ASSET_2,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS01_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0100_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter()
Call VerifyTextBoxContent("Description","ANLA-TXT50","",lcase(DT_AS01_1140_CHECK_TEXT_OF_DESCRIPTION),False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call VerifyTextBoxContent("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_CHECK_TEXT_OF_COST_CENTER,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Allocations", False)
Call VerifyTextBoxContent("Evaluation group 1","ANLA-ORD41","",DT_AS01_1160_CHECK_TEXT_OF_EVALUATION_GROUP_1,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Origin", False)
Call VerifyTextBoxContent("WBS element","ANLA-POSNR","",DT_AS01_1182_CHECK_TEXT_OF_WBS_ELEMENT,False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
''Call WriteRunTimeDataToExcelGlobalSheet ("DT_COUNT",Cint(DT_COUNT)+1)

''''''--------TransactionCode-AS03----------''''

Call SetTcode(DT_AS01_0100_OKCD)  
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS01_0100_ASSET,False)
Call SetTextbox("Asset","ANLA-ANLN1","",DT_ASSET_3,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS01_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0100_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter()
Call VerifyTextBoxContent("Description","ANLA-TXT50","",lcase(DT_AS01_1140_CHECK_TEXT_OF_DESCRIPTION),False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call VerifyTextBoxContent("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_CHECK_TEXT_OF_COST_CENTER,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Allocations", False)
Call VerifyTextBoxContent("Evaluation group 1","ANLA-ORD41","",DT_AS01_1160_CHECK_TEXT_OF_EVALUATION_GROUP_1,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Origin", False)
Call VerifyTextBoxContent("WBS element","ANLA-POSNR","",DT_AS01_1182_CHECK_TEXT_OF_WBS_ELEMENT,False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
''Call WriteRunTimeDataToExcelGlobalSheet ("DT_COUNT",Cint(DT_COUNT)+1)

''''''--------TransactionCode-AS03----------''''

Call SetTcode(DT_AS01_0100_OKCD)  
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS01_0100_ASSET,False)
Call SetTextbox("Asset","ANLA-ANLN1","",DT_ASSET_4,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS01_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0100_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter()
Call VerifyTextBoxContent("Description","ANLA-TXT50","",lcase(DT_AS01_1140_CHECK_TEXT_OF_DESCRIPTION),False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call VerifyTextBoxContent("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_CHECK_TEXT_OF_COST_CENTER,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Allocations", False)
Call VerifyTextBoxContent("Evaluation group 1","ANLA-ORD41","",DT_AS01_1160_CHECK_TEXT_OF_EVALUATION_GROUP_1,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Origin", False)
Call VerifyTextBoxContent("WBS element","ANLA-POSNR","",DT_AS01_1182_CHECK_TEXT_OF_WBS_ELEMENT,False)
Call TakeScreenShot
Call ClickButton("Back   \(F3\)", False)
''Call WriteRunTimeDataToExcelGlobalSheet ("DT_COUNT",Cint(DT_COUNT)+1)
Call ClickButton("Exit   \(Shift\+F3\)", False)



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




