

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.11.01.10.02 Retire Assets_Asset Retirem. frm Sale with Customer [Unique]_P3
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
'.................Test Script Name : Test_09.11.01.10.02 Retire Assets_Asset Retirem. frm Sale with Customer [Unique]_P3
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 2nd Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_09.11.01.10.02 Retire Assets_Asset Retirem. frm Sale with Customer [Unique]_P3"
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

'----------------------Tcode FD10N----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Enter details
Call SetTextbox("Customer","SO_KUNNR-LOW","",DT_FD10N_1000_CUSTOMER,False)
Call SetTextbox("Fiscal year","GP_GJAHR","",DT_FD10N_1000_FISCAL_YEAR,False)
Call SetTextbox("Company Code","SO_BUKRS-LOW","",DT_FD10N_1000_COMPANY_CODE,False)
Call TakeScreenShot()


'Click on Execute
Call ClickButton("Execute   \(F8\)",False)


Call SelectRowGuiGrid("",0,"Period",DT_FD10N_0100_GRIDCELL_7_PERIOD,False)
Call DoubleClickGuiGridCell("",0,DT_FD10N_0100_GRIDCELL_7_PERIOD,"Period",False)
Call TakeScreenShot()

Call SelectColumnGuiGrid("",0,"Document Number",False)
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)


Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FD10N_1105_DOCUMENT_NUMBER,True)
Call ClickButton("Execute   \(Enter\)",True)

'Verify the content
Call VerifyGridCellContent("",1,"Document type",0,DT_FD10N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("",1,"Document Date",0,DT_FD10N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT)
Call VerifyGridCellContent("",1,"Amount in local currency",0,DT_FD10N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("",1,"Local Currency",0,DT_FD10N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)

Call ClickButton("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call SelectRowGuiGrid("Column Set",0,"Column Name","Profit Center",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
Call SelectRowGuiGrid("Column Set",0,"Column Name","G/L Account",True)
Call ClickButton("Show Selected Fields \(F7\)",True)
'Call SelectRowGuiGrid("Column Set",0,"Column Name","Text",True)
'Call ClickButton("Show Selected Fields \(F7\)",True)

Call ClickButton("Transfer   \(Enter\)",True)
Call TakeScreenShot()

'Verify the content
Call VerifyGridCellContent("",1,"Profit Center",0,DT_FD10N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("",1,"G/L Account",0,DT_FD10N_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

