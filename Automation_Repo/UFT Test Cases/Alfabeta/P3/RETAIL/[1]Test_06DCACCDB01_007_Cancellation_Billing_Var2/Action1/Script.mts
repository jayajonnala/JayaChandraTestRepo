	

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_POST_DeleteVAT_from_Customer
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
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_06DCACCDB01_Cancellation_Var2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_06DCACCDB01_007_Cancellation_Billing_Var2_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'//-----------------------------------VF11 -----------------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call SetTableData("SAPMV60ATCTRL_ERF_FAKT","Document",1,"","",DT_VF11_0102_TABLECELL_DOCUMENT_0,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call VerifyTableCellContent(1,"Billing Type","SAPMV60ATCTRL_ERZ_FAKT",Lcase(DT_VF11_0103_CHECK_TEXT_OF_TABLECELL_BILLING_TYPE_0))
Call VerifyTableCellContent(2,"Billing Type","SAPMV60ATCTRL_ERZ_FAKT",Lcase(DT_VF11_0103_CHECK_TEXT_OF_TABLECELL_BILLING_TYPE_1))


Call VerifyTableCellContent(1,"Net Value","SAPMV60ATCTRL_ERZ_FAKT",DT_VF11_0103_CHECK_TEXT_OF_TABLECELL_NET_VALUE_0)
Call VerifyTableCellContent(2,"Net Value","SAPMV60ATCTRL_ERZ_FAKT",DT_VF11_0103_CHECK_TEXT_OF_TABLECELL_NET_VALUE_1)
'Call VerifyTableCellContent(1,"Net Value","SAPMV60ATCTRL_ERZ_FAKT",Trim(DT_VF11_0103_CHECK_TEXT_OF_TABLECELL_NET_VALUE_0))
'Call VerifyTableCellContent(2,"Net Value","SAPMV60ATCTRL_ERZ_FAKT",Trim(DT_VF11_0103_CHECK_TEXT_OF_TABLECELL_NET_VALUE_1))
'''''''Call VerifyTableCellContent(1,"Net Value","SAPMV60ATCTRL_ERZ_FAKT",Lcase(DT_VF11_0103_CHECK_TEXT_OF_TABLECELL_NET_VALUE_0))
'''''''Call VerifyTableCellContent(2,"Net Value","SAPMV60ATCTRL_ERZ_FAKT",Lcase(DT_VF11_0103_CHECK_TEXT_OF_TABLECELL_NET_VALUE_1))

Call TakeScreenShot

Call ClickButtonIfExist("Save   \(Ctrl\+S\)",false)
Call TakeScreenShot
Call GetStatusBar("item1","DT_DOC_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_OUTPUT",DT_DOC)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(Lcase(DT_VF11_0102_CHECK_TEXT_OF_STATUSBAR))

'//-----------------------------------VL09 -----------------------------------
Call SetTcode(DT_VF11_0102_OKCD) 
Call PressEnter() 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)
Call TakeScreenShot

Call SetTextbox("Inbound / Outbound Delivery","I_VBELN-LOW","",DT_VF11_1000_INBOUND__OUTBOUND_DELIVERY,False)
Call SetTextbox("Shipping Point","I_VSTEL-LOW","",DT_VF11_1000_SHIPPING_POINT,False)
Call TakeScreenShot

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("",1,"Delivery",0,DT_VF11_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VBELN)

Call SelectRowGuiGridByRefTwoColumns("","","Delivery",DT_VF11_1000_INBOUND__OUTBOUND_DELIVERY,"SPt",DT_VF11_1000_SHIPPING_POINT)
Call TakeScreenShot

Call ClickButton("Reverse Goods Movement   \(F5\)",False)
Call TakeScreenShot

Call VerifyTextBoxContent("Question","G_LINE_1",0,Lcase(DT_VF11_0300_CHECK_TEXT_OF_G_LINE_1),True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",False)
Call TakeScreenShot

Call VerifyifGuiLabelExistsByRelativeid("Goods issue for delivery "&DT_VF11_1000_INBOUND__OUTBOUND_DELIVERY&" canceled","wnd\[1\]/usr/lbl\[5,3\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_VF11_0120_CHECK_TEXT_OF_NO_PROBLEMS_HAVE_BEEN_LOGGED,"wnd\[1\]/usr/lbl\[5,4\]")
Call ClickButtonIfExist("Continue   \(Enter\)",False) 
Call TakeScreenShot

Call LogOff()
Call FinalStatus ()



'//------------------------------------------(       ......        UTILITY STATEMENTS    ......        )---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Call CreateRunTimeExcelFile(strFileName)       ................Can use this function if user want to Create Run Time Excel Sheet which captures the run time data 
'Call GetRunTimeDataFromExcel(strRunTimeExcelFileName,IterationIndex)          ................Can use this function if user want to Get Run Time captured data from run time excel sheet 
'Call WriteRunTimeScenarioData(strRunTimeExcelFileName,strVariableName,strVariableValue)          ................Can use this function if user want to Write Run Time captured data to run time excel sheet 

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


'// ---- Script Generated in [0] Minutes , [12,8640008]  Seconds ---- //
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
' ................NOTE: 
'.................1		This file is auto converted code from pre-recorded QTP script suitable for TASE Framework only .Please verify each function for applicability
'.................2		Default Index value 0 is used. If Multiple objects with similar names exists in application,replace 0 with 1/2 etc in case of failure.See the comment line
'.................3		User supplied Data is auto-parametized with relevant variable Names.See the comment line for details
'.................4		Input test data excel file is auto generated along with this script in the same location as this file.Input excel file contains all variable names and use defined data as appearing in this script initially
' ................5		If required additional logic  like  IF - Else , While Loop etc ,can be inserted in between lines  
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


