'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01.04.01.01.01 Maintain Asset Masterdata_Asset Creation Manual
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
	GetRowNo= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrTestCaseName = "Test_01.04.01.01.01 Maintain Asset Masterdata_Asset Creation Manual"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''--------TransactionCode-AS01 ----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS01_0105_ASSET_CLASS,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE,False)
Call SetTextbox("Number of similar assets","RA02S-NASSETS","",DT_AS01_0105_NUMBER_OF_SIMILAR_ASSETS,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION,False)
Call SetTextbox("Quantity","ANLA-MEINS","",DT_AS01_1140_QUANTITY,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Origin", False)
Call TakeScreenShot
Call SetTextbox("WBS element","ANLA-POSNR","",DT_AS01_1182_WBS_ELEMENT,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "General", False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Allocations", False)
Call TakeScreenShot
Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_AS01_1160_EVALUATION_GROUP_2,False)
Call TakeScreenShot
Call SelectTab("TABSTRIP100", "Deprec. Areas", False)
Call TakeScreenShot

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType("S")

Call GetStatusBar("item1","DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_AS01_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

'''''---------AS01 ----------''''
Call SetTcode(DT_AS01_1000_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_AS01_1000_OKCD)

Call PressEnter() 
Call TakeScreenShot

Call VerifyTextBoxContent("Description","ANLA-TXT50", "", Lcase(DT_AS01_1140_CHECK_TEXT_OF_DESCRIPTION), False)
Call VerifyTextBoxContent("Asset main no\. text","ANLH-ANLHTXT", "", Lcase(DT_AS01_1140_CHECK_TEXT_OF_ASSET_MAIN_NO_TEXT), False)
Call VerifyTextBoxContent("Quantity","ANLA-MEINS", "", DT_AS01_1140_CHECK_TEXT_OF_QUANTITY, False)
''Call VerifyTextBoxContent("Acct determination","ANLA-KTOGR", "", DT_AS01_1140_CHECK_TEXT_OF_ACCT_DETERMINATION, False)
Call VerifyTextBoxContent("Acct determination","ANLA-KTOGR", "", DT_CHECK_TEXT_AD, False)
Call TakeScreenShot

Call SelectTab("TABSTRIP100", "Time-dependent", False)
Call TakeScreenShot
Call VerifyTextBoxContent("Business Area","ANLZ-GSBER", "", DT_AS01_1145_CHECK_TEXT_OF_BUSINESS_AREA, False)
Call VerifyTextBoxContent("Cost Center","ANLZ-KOSTL", "", DT_AS01_1145_CHECK_TEXT_OF_COST_CENTER, False)
Call SelectTab("TABSTRIP100", "Allocations", False)
Call TakeScreenShot
Call VerifyTextBoxContent("Evaluation group 2","ANLA-ORD42", "", DT_AS01_1160_CHECK_TEXT_OF_EVALUATION_GROUP_2, False)
Call SelectTab("TABSTRIP100", "Origin", False)
Call TakeScreenShot
Call VerifyTextBoxContent("WBS element","ANLA-POSNR", "", DT_AS01_1182_CHECK_TEXT_OF_WBS_ELEMENT, False)
Call SelectTab("TABSTRIP100", "Deprec. Areas", False)
Call TakeScreenShot

Call VerifyTableCellContent(1, "DKey", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_DKEY_0)
Call VerifyTableCellContent(2, "DKey", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_DKEY_1)
Call VerifyTableCellContent(3, "DKey", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_DKEY_2)
Call VerifyTableCellContent(4, "DKey", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_DKEY_3)

Call VerifyTableCellContent(1, "UseLife", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_0)
Call VerifyTableCellContent(2, "UseLife", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_1)
Call VerifyTableCellContent(3, "UseLife", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_2)
Call VerifyTableCellContent(4, "UseLife", "SAPLAISTTC_ANLB", DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_3)

Call VerifyTableCellContent(1, "Prd", "SAPLAISTTC_ANLB", DT_PERIOD)
Call VerifyTableCellContent(2, "Prd", "SAPLAISTTC_ANLB", DT_PERIOD)
Call VerifyTableCellContent(3, "Prd", "SAPLAISTTC_ANLB", DT_PERIOD)
Call VerifyTableCellContent(4, "Prd", "SAPLAISTTC_ANLB", DT_PERIOD)

Call LogOff'
Call FinalStatus()



