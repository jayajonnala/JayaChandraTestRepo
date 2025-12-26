

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01.04.01.01.03 Maintain Asset Masterdata_Change asset_P1
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
'.................Test Script Name : Test_01.04.01.01.03 Maintain Asset Masterdata_Change asset_P1
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 2nd Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_01.04.01.01.03 Maintain Asset Masterdata_Change asset_P1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_02GR10_002_LocVend_w_Subrange_DSD_GR_Deliv_Note_w_Trading_Goods.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'----------------------Tcode AS02----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


'Enter details
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS02_0100_ASSET,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS02_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS02_0100_COMPANY_CODE,False)
Call TakeScreenShot()
Call PressEnter() 


Call SetTextbox("Description","ANLA-TXT50","",DT_AS02_1140_DESCRIPTION,False)
Call SetTextbox("Asset main no\. text","ANLH-ANLHTXT",""," ",False)
Call SetTextbox("Serial number","ANLA-SERNR","",DT_AS02_1140_SERIAL_NUMBER,False)
Call SetTextbox("Inventory number","ANLA-INVNR","",DT_AS02_1140_INVENTORY_NUMBER,False)
Call SetTextbox("Quantity","ANLA-MENGE","",DT_AS02_1140_QUANTITY,False)
Call TakeScreenShot()
Call PressEnter() 

'Post the Document
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True) 
Wait(2)
Call GetStatusBar("item1","DT_ASSET_NO_OUTPUT")
Call VerifyStatusBar("The asset "&DT_ASSET_NO_OUTPUT&" 0 is changed" )


''----------------------Tcode AS03----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_AS02_0100_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_AS02_0100_OKCD)
Call PressEnter()  

Call  VerifyTextBoxContent("Description","ANLA-TXT50",0,Lcase(DT_AS02_1140_CHECK_TEXT_OF_DESCRIPTION_OCC1),False)
Call  VerifyTextBoxContent("Asset main no\. text","ANLH-ANLHTXT",0,Lcase(DT_AS02_1140_CHECK_TEXT_OF_ASSET_MAIN_NO_TEXT),False)
Call  VerifyTextBoxContent("Serial number","ANLA-SERNR",0,DT_AS02_1140_CHECK_TEXT_OF_SERIAL_NUMBER,False)
Call  VerifyTextBoxContent("Inventory number","ANLA-INVNR",0,DT_AS02_1140_CHECK_TEXT_OF_INVENTORY_NUMBER,False)
Call  VerifyTextBoxContent("Quantity","ANLA-MENGE",0,DT_AS02_1140_CHECK_TEXT_OF_QUANTITY,False)


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

