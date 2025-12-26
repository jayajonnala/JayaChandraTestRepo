

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.11.01.10.02 Retire Assets_Asset Retirem. frm Sale with Customer [Unique]_P2
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
'.................Test Script Name : Test_09.11.01.10.02 Retire Assets_Asset Retirem. frm Sale with Customer [Unique]_P2
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 2nd Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_09.11.01.10.02 Retire Assets_Asset Retirem. frm Sale with Customer [Unique]_P2"
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

'----------------------Tcode AS03----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'Enter details
Call SetTextbox("Asset","ANLA-ANLN1","",DT_AS03_0100_ASSET,False)
Call SetTextbox("Sub-number","ANLA-ANLN2","",DT_AS03_0100_SUBNUMBER,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS03_0100_COMPANY_CODE,False)
Call TakeScreenShot()
Call PressEnter()

Call VerifyTextBoxContent("Quantity","ANLA-MENGE",0,"0,000",False)
Call VerifyTextBoxContent("Deactivation on","ANLA-DEAKT",0,ConvertDate(DT_AS03_1142_CHECK_TEXT_OF_DEACTIVATION_ON),False)


'Post the Asset values
Call ClickButton("Asset values   \(Ctrl\+F1\)",False)

Call ActivateNodeGuiTree(0,"Depreciation Areas;0L Leading Ledger;01 IFRS APC, depreciation")


Call VerifyGridCellContent("Planned values IFRS APC, depreciation",3,"Change",0,"")
Call VerifyGridCellContent("Transactions",1,"Amount posted",0,DT_AS03_0301_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BUBTR)

Call SelectRowGuiGrid("Transactions",0,"Transaction type name","Retirement of current-year acquis. with revenue",False)
Call DoubleClickGuiGridCell("Transactions",0,2,"Transaction type name",False)
Call TakeScreenShot()


'Verify the Cell Content
Call VerifyGridCellContent("",1,"Posting Key",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",1,"Account",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",1,"Alternative Account No.",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LOKKT)
Call VerifyGridCellContent("",1,"Amount",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("",1,"Currency",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW)
Call VerifyGridCellContent("",1,"Profit Center",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)

Call VerifyGridCellContent("",2,"Posting Key",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("",2,"Account",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("",2,"Alternative Account No.",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_LOKKT)
Call VerifyGridCellContent("",2,"Tax Code",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)
Call VerifyGridCellContent("",2,"Currency",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PSWSL)
Call VerifyGridCellContent("",2,"Amount",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("",2,"Profit Center",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)

Call VerifyGridCellContent("",3,"Posting Key",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("",3,"Account",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("",3,"Alternative Account No.",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_LOKKT)
Call VerifyGridCellContent("",3,"Tax Code",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MWSKZ)
Call VerifyGridCellContent("",3,"Currency",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_RF05A_UBAZW)
Call VerifyGridCellContent("",3,"Amount",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)
Call VerifyGridCellContent("",3,"Profit Center",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PRCTR)


Call VerifyGridCellContent("",4,"Account",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)
Call VerifyGridCellContent("",4,"Alternative Account No.",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_LOKKT)
Call VerifyGridCellContent("",4,"Tax Code",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_MWSKZ)
Call VerifyGridCellContent("",4,"Description",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KOBEZ)
Call VerifyGridCellContent("",4,"Currency",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_RF05A_UBAZW)
Call VerifyGridCellContent("",4,"Amount",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_AZBET)
Call VerifyGridCellContent("",4,"Profit Center",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_PRCTR)

Call VerifyGridCellContent("",5,"Posting Key",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BSCHL)
Call VerifyGridCellContent("",5,"Account",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_KTONR)
Call VerifyGridCellContent("",5,"Alternative Account No.",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_LOKKT)
Call VerifyGridCellContent("",5,"Tax Code",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_MWSKZ)
Call VerifyGridCellContent("",5,"Currency",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_RF05A_UBAZW)
Call VerifyGridCellContent("",5,"Amount",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_AZBET)
Call VerifyGridCellContent("",5,"Profit Center",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_PRCTR)

Call VerifyGridCellContent("",6,"Posting Key",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BSCHL)
Call VerifyGridCellContent("",6,"Account",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_KTONR)
Call VerifyGridCellContent("",6,"Alternative Account No.",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_LOKKT)
Call VerifyGridCellContent("",6,"Description",0,Lcase(DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_KOBEZ))
Call VerifyGridCellContent("",6,"Tax Code",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_MWSKZ)
Call VerifyGridCellContent("",6,"Currency",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_RF05A_UBAZW)
Call VerifyGridCellContent("",6,"Amount",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_AZBET)
Call VerifyGridCellContent("",6,"Profit Center",0,DT_AS03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_PRCTR)

'Click on Exit
Call ClickButtonIfExist("Exit   \(Shift\+F3\)",False)
Wait(2)

Call ActivateNodeGuiTree(0,"Depreciation Areas;D1 Non-Leading (Local);06 Local GAAP APC, depreciation")
Call VerifyGridCellContent("Transactions",1,"Amount posted",0,DT_CHECK_20)
Call VerifyGridCellContent("Transactions",2,"Amount posted",0,DT_CHECK_22)


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

