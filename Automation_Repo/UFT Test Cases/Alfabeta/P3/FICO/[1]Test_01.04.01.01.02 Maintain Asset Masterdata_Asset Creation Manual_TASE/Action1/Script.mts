

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01.04.01.01.02 Maintain Asset Masterdata_Asset Creation Manual
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_01.04.01.01.02 Maintain Asset Masterdata_Asset Creation Manual
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 2nd Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_01.04.01.01.02 Maintain Asset Masterdata_Asset Creation Manual"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR10_002_LocVend_w_Subrange_DSD_GR_Deliv_Note_w_Trading_Goods.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)


'gstrresultFolderPath = ReadTxtFileResult("S:\Rsutls_Optimized\GlobalRunTimeResultFolderPath\ResultFolderPath.txt")
'
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Login to SAP System
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode AS01----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))

'Enter details
Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS01_0105_ASSET_CLASS,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE,False)
Call TakeScreenShot()
Call PressEnter() 

Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION,False)
Call SetTextbox("Quantity","ANLA-MEINS","",DT_AS01_1140_QUANTITY,False)
Call TakeScreenShot()

'Navigate to Origin Tab
Call SelectTab("TABSTRIP100","Origin",False)

Call SetTextbox("Type name","ANLA-TYPBZ","",DT_AS01_1181_TYPE_NAME,False)
Call SetTextbox("WBS element","ANLA-POSNR","",DT_AS01_1182_WBS_ELEMENT,False)
Call TakeScreenShot()

'Navigate to Time-dependent Tab
Call SelectTab("TABSTRIP100","Time-dependent",False)

Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_COST_CENTER,False)
Call TakeScreenShot()
Call PressEnter() 

'Navigate to Allocation Tab
Call SelectTab("TABSTRIP100","Allocations",False)

Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_AS01_1160_EVALUATION_GROUP_2,False)
''Call SetTextbox("Evaluation group 4","ANLA-ORD44","",DT_AS01_1160_EVALUATION_GROUP_4,False)
Call TakeScreenShot()

'Navigate to Deprec. Areas Tab
Call SelectTab("TABSTRIP100","Deprec. Areas",False)
Call TakeScreenShot()

'Post the Document
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True) 
Wait(2)
Call GetStatusBar("item1","DT_ASSET_NO_OUTPUT")
Call VerifyStatusBar("The asset "&DT_ASSET_NO_OUTPUT&" 0 is created" )

''----------------------Tcode AS01----------------------------

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'Create Purchase Order
Call SetTcode(DT_AS01_0105_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_AS01_0105_OKCD)
Call PressEnter()

Call  VerifyTextBoxContent("Description","ANLA-TXT50",0,Lcase(DT_AS01_1140_CHECK_TEXT_OF_DESCRIPTION),False)
'''Call  VerifyTextBoxContent("Acct determination","ANLA-KTOGR",0,DT_AS01_1140_CHECK_TEXT_OF_ACCT_DETERMINATION,False)
Call  VerifyTextBoxContent("Acct determination","ANLA-KTOGR",0,DT_AS01_CHECK_TEXT_OF_AD,False)
Call  VerifyTextBoxContent("Asset main no\. text","ANLH-ANLHTXT",0,Lcase(DT_AS01_1140_CHECK_TEXT_OF_ASSET_MAIN_NO_TEXT),False)
Call  VerifyTextBoxContent("Quantity","ANLA-MEINS",0,DT_AS01_1140_CHECK_TEXT_OF_QUANTITY,False)

'Navigate to Time-dependent Tab
Call SelectTab("TABSTRIP100","Time-dependent",False)

Call  VerifyTextBoxContent("Business Area","ANLZ-GSBER",0,DT_AS01_1145_CHECK_TEXT_OF_BUSINESS_AREA,False)
Call  VerifyTextBoxContent("Cost Center","ANLZ-KOSTL",0,DT_AS01_1145_CHECK_TEXT_OF_COST_CENTER,False)
Call  VerifyTextBoxContent("Profit Center","ANLZ-PRCTR",0,DT_AS01_1145_CHECK_TEXT_OF_PROFIT_CENTER,False)


'Navigate to Allocation Tab
Call SelectTab("TABSTRIP100","Allocations",False)

Call VerifyTextBoxContent("Evaluation group 2","ANLA-ORD42",0,DT_AS01_1160_CHECK_TEXT_OF_EVALUATION_GROUP_2,False)
''Call VerifyTextBoxContent("Evaluation group 4","ANLA-ORD44",0,DT_AS01_1160_CHECK_TEXT_OF_EVALUATION_GROUP_4,False)
''Call VerifyTableCellContent(1,"Category","SAPLAISTTC_EQUI",DT_AS01_1170_CHECK_TEXT_OF_TABLECELL_CATEGORY_0)
''Call VerifyTableCellContent(1,"Description of Technical Object","SAPLAISTTC_EQUI",Lcase(DT_AS01_1170_CHECK_TEXT_OF_OF_TECHNICAL_OBJECT_0))
''
'Navigate to Origin Tab
Call SelectTab("TABSTRIP100","Origin",False)


Call VerifyTextBoxContent("Type name","ANLA-TYPBZ",0,DT_AS01_1181_CHECK_TEXT_OF_TYPE_NAME,False)
Call VerifyTextBoxContent("WBS element","ANLA-POSNR",0,DT_AS01_1182_CHECK_TEXT_OF_WBS_ELEMENT,False)


'Navigate to Deprec. Areas Tab
Call SelectTab("TABSTRIP100","Deprec. Areas",False)
Call TakeScreenShot()


Call VerifyTableCellContent(1,"DKey","SAPLAISTTC_ANLB",DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_DKEY_0)
Call VerifyTableCellContent(2,"DKey","SAPLAISTTC_ANLB",DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_DKEY_1)
Call VerifyTableCellContent(3,"DKey","SAPLAISTTC_ANLB",DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_DKEY_2)
Call VerifyTableCellContent(4,"DKey","SAPLAISTTC_ANLB",DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_DKEY_3)

Call VerifyTableCellContent(1,"UseLife","SAPLAISTTC_ANLB",DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_0)
Call VerifyTableCellContent(2,"UseLife","SAPLAISTTC_ANLB",DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_1)
Call VerifyTableCellContent(3,"UseLife","SAPLAISTTC_ANLB",DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_2)
Call VerifyTableCellContent(4,"UseLife","SAPLAISTTC_ANLB",DT_AS01_1190_CHECK_TEXT_OF_TABLECELL_USELIFE_3)

Call VerifyTableCellContent(1,"Prd","SAPLAISTTC_ANLB",DT_PERIOD)
Call VerifyTableCellContent(2,"Prd","SAPLAISTTC_ANLB",DT_PERIOD)
Call VerifyTableCellContent(3,"Prd","SAPLAISTTC_ANLB",DT_PERIOD)
Call VerifyTableCellContent(4,"Prd","SAPLAISTTC_ANLB",DT_PERIOD)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

