

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_040-Cancelation Inv from EU Vendors_P3_Chk_PO_Clear   
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



gstrTestCaseName = "Test_P2P_01_01_040_P3_Chk_PO_Clear"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_040-Cancelation Inv from EU Vendors_P3_Chk_PO_Clear.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


''''--------TransactionCode-MB51----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call ClickButton("Other Purchase Order   \(Shift\+F5\)",false)
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_ME23N_0003_PUR_ORDER,True)
Call ClickButton("Other Document   \(Enter\)",True)
Call TakeScreenShot
Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
Call SelectTab("ITEM_DETAIL","Purchase Order History",False)
Call TakeScreenShot
'''''Call ClickButtonToolBar("&MB_FILTER",0)
'''''''Call SelectMenuIdToolBar("&FILTER",1)
'''''
'''''Call SelectCellGuiGrid("Column Set",0,4,"Column Name",True)
'''''Call ClickButton("Add Filter Criterion \(F7\)",True)
'''''Call ClickButton("Define Filter Values",True)
'''''Call SetTextbox("Article Document","%%DYN001-LOW",0,DT_ME23N_1105_ARTICLE_DOCUMENT,True)
'''''Call ClickButton("Execute   \(Enter\)",True)
'''''Call VerifyGridCellContent("", 1, "BELNR", 0, DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
'''''Call TakeScreenShot
'''''Call SelectMenuIdToolBar("&DELETE_FILTER",0)
'''''
'''''Call ClickButtonToolBar("&MB_FILTER",0)
'''''Call SelectCellGuiGrid("Column Set",0,4,"Column Name",True)
'''''Call ClickButton("Add Filter Criterion \(F7\)",True)
'''''Call ClickButton("Define Filter Values",True)
'''''Call SetTextbox("Article Document","%%DYN001-LOW",0,DT_ME23N_1105_ARTICLE_DOCUMENT_OCC1,True)
'''''Call ClickButton("Execute   \(Enter\)",True)
'''''Call VerifyGridCellContent("", 1, "BELNR", 0, DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC1)
'''''Call TakeScreenShot
'''''Call SelectMenuIdToolBar("&DELETE_FILTER",0)
'''''
'''''Call ClickButtonToolBar("&MB_FILTER",0)
'''''Call SelectCellGuiGrid("Column Set",0,4,"Column Name",True)
'''''Call ClickButton("Add Filter Criterion \(F7\)",True)
'''''Call ClickButton("Define Filter Values",True)
'''''Call SetTextbox("Article Document","%%DYN001-LOW",0,DT_ME23N_1105_ARTICLE_DOCUMENT_OCC2,True)
'''''Call ClickButton("Execute   \(Enter\)",True)
'''''Call VerifyGridCellContent("", 1, "BELNR", 0, DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC2)
'''''Call TakeScreenShot
'''''Call SelectMenuIdToolBar("&DELETE_FILTER",0)
'''''
'''''Call ClickButtonToolBar("&MB_FILTER",0)
'''''Call SelectCellGuiGrid("Column Set",0,4,"Column Name",True)
'''''Call ClickButton("Add Filter Criterion \(F7\)",True)
'''''Call ClickButton("Define Filter Values",True)
'''''Call SetTextbox("Article Document","%%DYN001-LOW",0,DT_ME23N_1105_ARTICLE_DOCUMENT_OCC3,True)
'''''Call ClickButton("Execute   \(Enter\)",True)
'''''Call VerifyGridCellContent("", 1, "BELNR", 0, DT_ME23N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OCC3)
'''''Call TakeScreenShot
Call GetTextboxValue("MEPO_TOPLINE-SUPERFIELD",0,"DT_VENDOR_NUMBER_OUTPUT",False)

Call SetTcode(DT_ME23N_0014_OKCD)     
Call PressEnter() 
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTextbox("Account","RF05A-AGKON","",DT_ME23N_0131_ACCOUNT,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_ME23N_0131_COMPANY_CODE,False)
Call SelectRadioButton("RF05A-XPOS1","Reference",False)
Call TakeScreenShot
Call PressEnter() 
Call TakeScreenShot
Call SetTextbox("From","RF05A-SEL01","",DT_ME23N_0731_FROM,False)
Call PressEnter() 
Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot
Call ClickButton("ICON_SELECT_ALL",False)
Call ClickButton("Activate Items",False)
Call GetTextBoxValue("RF05A-DIFFB",0,"DT_VAL_OUT",False)
''Call VerifyTextBoxContent("Not assigned","RF05A-DIFFB",0,DT_ME23N_6102_CHECK_TEXT_OF_NOT_ASSIGNED,False)
Call VerifyTextBoxContent("Not assigned","RF05A-DIFFB",0,trim(DT_VAL_OUT),False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_ME23N_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_ME23N_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code RO02")


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


