

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.11.01.10.02 Retire Assets_Asset Retirem. frm Sale with Customer [Unique]_P8
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
'.................Test Script Name : Test_09.11.01.10.02 Retire Assets_Asset Retirem. frm Sale with Customer [Unique]_P8
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 2nd Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_09.11.01.10.02 Retire Assets_Asset Retirem. frm Sale with Customer [Unique]_P8"
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

'----------------------Tcode FAGLB03----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Enter details
Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_FAGLB03_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_FAGLB03_1000_FISCAL_YEAR,False)
Call TakeScreenShot()
Call PressEnter()

'Click on Execute 
Call ClickButton("Execute   \(F8\)",False)

Call SelectRowGuiGrid("",0,"Period",DT_FAGLB03_0030_GRIDCELL_7_PERIOD,False)
Call DoubleClickGuiGridCell("",0,DT_FAGLB03_0030_GRIDCELL_7_PERIOD,"Period",False)
Call TakeScreenShot()

Call SelectColumnGuiGrid("",0,"Document Number",False)
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)


Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FAGLB03_1105_DOCUMENT_NUMBER,True)
Call ClickButton("Execute   \(Enter\)",True)

Call ClickButton("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call SelectRowGuiGrid("Column Set",0,"Column Name","Posting Date",True)
Call ClickButton("Show Selected Fields \(F7\)",True)

Call ClickButton("Transfer   \(Enter\)",True)
Call TakeScreenShot()


'Verify the content
Call VerifyGridCellContent("",1,"Document type",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("",1,"Document Date",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT)
Call VerifyGridCellContent("",1,"Posting Key",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",1,"Amount in local currency",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("",1,"Local Currency",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("",1,"Text",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("",1,"Profit Center",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("",1,"Posting Date",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT)

'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)

'Click on Back
Call ClickButtonIfExist("Back   \(F3\)",False)
Wait(2)

'Click on Choose Ledger
Call ClickButtonIfExist("Choose Ledger   \(Ctrl\+F4\)",False)
Wait(2)

Call SetTextbox("Ledger","SVALD-VALUE","",DT_FAGLB03_0300_LEDGER,True)
Call ClickButton("Cont\.   \(Enter\)",True)

'Click on Execute 
Call ClickButton("Execute   \(F8\)",False)

Call SelectRowGuiGrid("",0,"Period",DT_FAGLB03_0030_GRIDCELL_7_PERIOD,False)
Call DoubleClickGuiGridCell("",0,DT_FAGLB03_0030_GRIDCELL_7_PERIOD,"Period",False)
Call TakeScreenShot()

Call SelectColumnGuiGrid("",0,"Document Number",False)
Call ClickButton("Set filter   \(Ctrl\+Shift\+F2\)",False)


Call SetTextbox("Document Number","%%DYN001-LOW","",DT_FAGLB03_1105_DOCUMENT_NUMBER_OCC1,True)
Call ClickButton("Execute   \(Enter\)",True)


Call ClickButton("Change layout\.\.\.   \(Ctrl\+F8\)",False)

Call SelectRowGuiGrid("Column Set",0,"Column Name","Posting Date",True)
Call ClickButton("Show Selected Fields \(F7\)",True)

Call ClickButton("Transfer   \(Enter\)",True)
Call TakeScreenShot()

'Verify the content
Call VerifyGridCellContent("",1,"Document type",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART_OCC1)
Call VerifyGridCellContent("",1,"Posting Key",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL_OCC1)
Call VerifyGridCellContent("",1,"Document Date",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT_OCC1)
Call VerifyGridCellContent("",1,"Amount in local currency",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB_OCC1)
Call VerifyGridCellContent("",1,"Local Currency",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER_OCC1)
Call VerifyGridCellContent("",1,"Profit Center",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR_OCC1)
Call VerifyGridCellContent("",1,"Posting Date",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BUDAT_OCC1)
Call VerifyGridCellContent("",1,"Text",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT_OCC1)


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

