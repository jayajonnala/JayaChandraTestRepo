
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0273-Return process in DC_P1
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :[1]Test_P2P_01_01_0268-Self consumption for between Store and DC_P3_Cancell_All
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//





gstrTestCaseName = "Test_P2P_01_01_0268-DC_P3_Cancell_All"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\MI\DT_P2P_01_01_0268-Self consumption for between Store and DC_P3.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


Call SetTcode(DT_SAPTRANSACTIONCODE)     ' - Line (2)
Call PressEnter()     ' - Line (3)

Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call TakeScreenShot()

Call SetTableData("SAPMV60ATCTRL_ERF_FAKT", "Document", 1, "", "", DT_VF11_0102_TABLECELL_DOCUMENT_0, False)

Call ClickButton("Save   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_VF11_0102_CHECK_TEXT_OF_STATUSBAR_OUTPUT")

Call VerifyStatusBar("Document "&DT_VF11_0102_CHECK_TEXT_OF_STATUSBAR_OUTPUT&" has been saved")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Call SetTcode(DT_VF11_0102_OKCD)

Call PressEnter()

Call TakeScreenShot()

Call SetTextBox("Inbound / Outbound Delivery","I_VBELN-LOW",0,DT_VF11_1000_INBOUND__OUTBOUND_DELIVERY,False)

Call SetTextBox("Shipping Point","I_VSTEL-LOW",0,DT_VF11_1000_SHIPPING_POINT,False)

Call PressEnter()

Call CLickButton("Execute   \(F8\)",False)

Call TakeScreenShot()

Call SelectRowGuiGrid("", 0, "SPt", "RW05", False)


Call ClickButton("Reverse Goods Movement   \(F5\)",False)

Call ClickButton("Continue   \(Enter\)",False)


Call TakeScreenShot()
wait 5
Call ClickButton("Continue   \(Enter\)",False)

''''''''Call VerifyifGuiLabelExists(DT_VF11_0120_CHECK_TEXT_OF_GOODS_ISSUE_FOR_DELIVERY_85002522_CANCELED)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Call SetTcode(DT_VF11_0100_OKCD)

Call PressEnter()

Call TakeScreenShot()

Call SetTextBox("Outbound Delivery","LIKP-VBELN",0,DT_VF11_4004_OUTBOUND_DELIVERY,False)

Call PressEnter()

Call CLickButton("Delete   \(Shift\+F2\)",False)

Call ClickButton("SPOP-OPTION1",True)

Call ClickButton("SPOP-VAROPTION1",True)

Call VerifyStatusBar(DT_VF11_4004_CHECK_TEXT_OF_STATUSBAR)

Call TakeScreenShot()

Call LogOff()

Call FinalStatus()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




























'Call SetTcode(DT_VF11_4004_OKCD)
'
'Call PressEnter()
'
'Call TakeScreenShot()
'
'Call SetTextBox("Order","VBAK-VBELN",0,DT_VF11_0102_ORDER,False)
'
'Call PressEnter()
'

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




