

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_01.04.01.01.03 Maintain Asset Masterdata_Change asset_P2
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
'.................Test Script Name : Test_01.04.01.01.03 Maintain Asset Masterdata_Change asset_P2
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 2nd Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_01.04.01.01.03 Maintain Asset Masterdata_Change asset_P2"
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

'Enter details
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS02_0100_ASSET,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS02_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS02_0100_COMPANY_CODE,False)
Call TakeScreenShot()
Call PressEnter() 


'Navigate to Time-dependent Tab
Call SelectTab("TABSTRIP100","Time-dependent",False)


Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS02_1145_COST_CENTER,False)
Call SetTextbox("Profit Center","ANLZ-PRCTR","",DT_AS02_1145_PROFIT_CENTER,False)
Call SetTextbox("License plate number","ANLZ-KFZKZ","",DT_AS02_1145_LICENSE_PLATE_NUMBER,False)
Call TakeScreenShot()
Call PressEnter() 
Call PressEnter()

'Click on More Interval
Call ClickButton("More Intervals",False)
Call PressEnter()

'Click on Add Interval
Call ClickButton("Insert Row",False)
Call PressEnter()

Call SetTextbox("From-date of new interval","ANLZ-ADATU","",ConvertDate(DT_AS02_3010_FROMDATE_OF_NEW_INTERVAL),True)
Call TakeScreenShot()
Call ClickButton("Yes   \(Enter\)",True)

Call SetTableData("SAPLAISTTIME","Cost Center","1","","",DT_AS02_3000_TABLECELL_COST_CENTER_0,False)
Call SetTableData("SAPLAISTTIME","Profit Center","1","","",DT_AS02_3000_TABLECELL_PROFIT_CENTER_0,False)
Call TakeScreenShot()
Call PressEnter()


'Post the Document
Call ClickButton("Save   \(Ctrl\+S\)",False)
Call PressEnter()
Call PressEnter()
Call PressEnter()
Call PressEnter()
Wait(2)
Call GetStatusBar("item1","DT_ASSET_NO_OUTPUT")
Call VerifyStatusBar("The asset "&DT_ASSET_NO_OUTPUT&" 0 is changed" )


''----------------------Tcode AS03----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_AS02_0100_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_AS02_0100_OKCD)
Call PressEnter() 

'Navigate to Time-dependent Tab
Call SelectTab("TABSTRIP100","Time-dependent",False)

'Click on More Interval
Call ClickButton("More Intervals",False)
Call PressEnter()


Call VerifyTableCellContent(1,"Fr","SAPLAISTTIME",ConvertDate(DT_AS02_3000_CHECK_TEXT_OF_TABLECELL_FR_0))
Call VerifyTableCellContent(1,"Cost Center","SAPLAISTTIME",DT_AS02_3000_CHECK_TEXT_OF_TABLECELL_COST_CENTER_0)
Call VerifyTableCellContent(1,"Profit Center","SAPLAISTTIME",DT_AS02_3000_CHECK_TEXT_OF_TABLECELL_PROFIT_CENTER_0)
Call VerifyTableCellContent(1,"License plate","SAPLAISTTIME",DT_AS02_3000_CHECK_TEXT_OF_TABLECELL_LICENSE_PLATE_0)
Call VerifyTableCellContent(2,"Fr","SAPLAISTTIME",DT_AS02_3000_CHECK_TEXT_OF_TABLECELL_FR_1)
Call VerifyTableCellContent(2,"Cost Center","SAPLAISTTIME",DT_AS02_3000_CHECK_TEXT_OF_TABLECELL_COST_CENTER_1)
Call VerifyTableCellContent(2,"Profit Center","SAPLAISTTIME",DT_AS02_3000_CHECK_TEXT_OF_TABLECELL_PROFIT_CENTER_1)
Call VerifyTableCellContent(2,"License plate","SAPLAISTTIME",DT_AS02_3000_CHECK_TEXT_OF_TABLECELL_LICENSE_PLATE_1)


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

