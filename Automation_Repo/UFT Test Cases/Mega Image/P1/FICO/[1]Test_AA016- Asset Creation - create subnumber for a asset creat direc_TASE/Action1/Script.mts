
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA016- Asset Creation - create subnumber for a asset creat direc
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


gstrTestCaseName = "Test_AA016- Asset Creation - create subnumber for a asset creat direc"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA016- Asset Creation - create subnumber for a asset creat direc.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''''--------TransactionCode-AS11----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS11_0110_ASSET,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS11_0110_COMPANY_CODE,False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Description","ANLA-TXT50","",DT_AS11_1140_DESCRIPTION,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot

Call SelectTab("TABSTRIP100", "Allocations", False)
Call SetTextbox("Evaluation group 1","ANLA-ORD41","",DT_EVAL1,False)
Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_EVAL2,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Origin", False)
Call SetTextbox("Vendor","ANLA-LIFNR","",DT_AS11_1181_VENDOR,False)
Call TakeScreenShot

Call SelectTab("TABSTRIP100", "Deprec. Areas", False)
Call TakeScreenShot

Call SetTableData("SAPLAISTTC_ANLB","UseLife","1","","",DT_AS11_1190_TABLECELL_USELIFE_0,False) 
Call SetTableData("SAPLAISTTC_ANLB","UseLife","2","","",DT_AS11_1190_TABLECELL_USELIFE_1,False) 
Call SetTableData("SAPLAISTTC_ANLB","Prd","1","","",DT_AS11_1190_TABLECELL_PRD_0,False) 
Call SetTableData("SAPLAISTTC_ANLB","Prd","2","","",DT_AS11_1190_TABLECELL_PRD_1,False) 


''Call VerifyTableCellContent(1, "UseLife", "SAPLAISTTC_ANLB", DT_AS11_1190_TABLECELL_USELIFE_0)
''Call VerifyTableCellContent(2, "UseLife", "SAPLAISTTC_ANLB", DT_AS11_1190_TABLECELL_USELIFE_1)
''Call VerifyTableCellContent(1, "Prd", "SAPLAISTTC_ANLB", DT_AS11_1190_TABLECELL_PRD_0)
''Call VerifyTableCellContent(2, "Prd", "SAPLAISTTC_ANLB", DT_AS11_1190_TABLECELL_PRD_1)
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call PressEnter()

Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Create",True)
Call PressEnter()
Call TakeScreenShot
Call GetStatusBar("item1", "DT_AS11_ASSET_OUTPUT")
Call GetStatusBar("item2", "DT_AS11_0110_ASSET_SUBNUMBER_OUTPUT")
Call VerifyStatusBar("The asset "&DT_AS11_ASSET_OUTPUT&" "&DT_AS11_0110_ASSET_SUBNUMBER_OUTPUT&" is created")


''''''--------TransactionCode-/nAS03----------''''

Call SetTcode(DT_AS11_0110_OKCD)  
Call PressEnter() 
Call TakeScreenShot
Call PressEnter()  

Call VerifyTextBoxContent("Description","ANLA-TXT50","",lcase(DT_AS11_1140_CHECK_TEXT_OF_DESCRIPTION),False)
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call VerifyTextBoxContent("Cost Center","ANLZ-KOSTL","",DT_AS11_1145_CHECK_TEXT_OF_COST_CENTER,False)
Call VerifyCheckBoxValue("ANLZ-XSTIL", DT_AS11_1145_CHECK_SELECTED_OF_ASSET_SHUTDOWN)
Call SelectTab("TABSTRIP100", "Allocations", False)
Call TakeScreenShot
Call VerifyTextBoxContent("Evaluation group 1","ANLA-ORD41","",DT_AS11_1160_CHECK_TEXT_OF_EVALUATION_GROUP_1,False)
Call VerifyTextBoxContent("Evaluation group 2","ANLA-ORD42","",DT_AS11_1160_CHECK_TEXT_OF_EVALUATION_GROUP_2,False)
Call SelectTab("TABSTRIP100", "Origin", False)
Call TakeScreenShot
Call VerifyTextBoxContent("Vendor","ANLA-LIFNR","",DT_AS11_1181_CHECK_TEXT_OF_VENDOR,False)
Call VerifyTextBoxContent("WBS element","ANLA-POSNR","",DT_AS11_1182_CHECK_TEXT_OF_WBS_ELEMENT,False)
Call SelectTab("TABSTRIP100", "Deprec. Areas", False)
Call TakeScreenShot
Call VerifyTableCellContent(1, "UseLife", "SAPLAISTTC_ANLB", DT_AS11_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_0)
Call VerifyTableCellContent(1, "Prd", "SAPLAISTTC_ANLB", DT_AS11_1190_CHECK_TEXT_OF_TABLECELL_PRD_0)
Call VerifyTableCellContent(2, "UseLife", "SAPLAISTTC_ANLB", DT_AS11_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_1)
Call VerifyTableCellContent(2, "Prd", "SAPLAISTTC_ANLB", DT_AS11_1190_CHECK_TEXT_OF_TABLECELL_PRD_1)
Call VerifyTableCellContent(5, "UseLife", "SAPLAISTTC_ANLB", DT_AS11_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_4)
Call VerifyTableCellContent(5, "Prd", "SAPLAISTTC_ANLB", DT_AS11_1190_CHECK_TEXT_OF_TABLECELL_PRD_4)


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




