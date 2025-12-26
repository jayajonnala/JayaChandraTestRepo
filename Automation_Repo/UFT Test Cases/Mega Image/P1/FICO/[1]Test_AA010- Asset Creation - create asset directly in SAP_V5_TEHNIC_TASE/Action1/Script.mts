
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AA010- Asset Creation - create asset directly in SAP_V5_TEHNIC
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

gstrTestCaseName = "Test_AA010- Asset Creation - create asset directly in SAP_V5_TEHNIC"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\FICO\TASE_DT_AA010- Asset Creation - create asset directly in SAP_V5_TEHNIC.xls"
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
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call SetTextbox("Business Area","ANLZ-GSBER","",DT_AS01_1145_BUSINESS_AREA,False)
Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_COST_CENTER,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Allocations", False)
Call TakeScreenShot
Call PressEnter()
Call SetTextbox("Evaluation group 1","ANLA-ORD41","",DT_AS01_1160_EVALUATION_GROUP_1,False)
Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_AS01_1160_EVALUATION_GROUP_2,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Origin", False)
Call TakeScreenShot
Call SetTextbox("Vendor","ANLA-LIFNR","",DT_AS01_1181_VENDOR,False)
Call SetTextbox("Type name","ANLA-TYPBZ","",DT_AS01_1181_TYPE_NAME,False)
Call TakeScreenShot
Call SetTextbox("WBS element","ANLA-POSNR","",DT_AS01_1182_WBS_ELEMENT,False)
Call TakeScreenShot
Call PressEnter()
Call SelectTab("TABSTRIP100", "Deprec. Areas", False)
Call SelectTab("TABSTRIP100", "Leasing", False)
Call SelectTab("TABSTRIP100", "Origin", False)
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetStatusBar("MessageType", "DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetStatusBar("item1", "DT_DOC_NR_OUTPUT")


'''''''--------TransactionCode-AS03----------''''

Call SetTcode(DT_AS01_1000_OKCD)  
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call PressEnter()

Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call SelectTab("TABSTRIP100", "Allocations", False)
Call GetTableCellData("SAPLAISTTC_EQUI","Equipment number",1,"","","DT_AS01_1170_CHECK_TEXT_OF_TABLECELL_EQUIPMENT_NUMBER_0_OUTPUT",False)
Call VerifyTableCellContent(1, "Equipment number", "SAPLAISTTC_EQUI", DT_AS01_1170_CHECK_TEXT_OF_TABLECELL_EQUIPMENT_NUMBER_0_OUTPUT)
Call VerifyTextBoxContent("Class","ANLA-ANLKL","",DT_AS01_1000_CHECK_TEXT_OF_CLASS,False)
Call SelectTab("TABSTRIP100", "Origin", False)
Call VerifyTextBoxContent("Vendor","ANLA-LIFNR","",DT_AS01_1181_CHECK_TEXT_OF_VENDOR,False)
Call VerifyTextBoxContent("WBS element","ANLA-POSNR","",DT_AS01_1182_CHECK_TEXT_OF_WBS_ELEMENT,False)
Call SelectTab("TABSTRIP100", "Deprec. Areas", False)
Call VerifyTextBoxContent("Class","ANLA-ANLKL","",DT_AS01_1000_CHECK_TEXT_OF_CLASS_OCC1,False)
Call VerifyTableCellContent(1, "UseLife", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_0)
Call VerifyTableCellContent(2, "UseLife", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_1)
Call VerifyTableCellContent(3, "UseLife", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_2)
Call VerifyTableCellContent(4, "UseLife", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_3)
Call VerifyTableCellContent(5, "UseLife", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_4)



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




