
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_AT_P2P_01_01_0261 - Self cons for marketing dept 65-11_P2     
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




gstrTestCaseName = "Test_AT_P2P_01_01_0261 -  dept 65-11_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_P2P_01_01_0261-Self consumption for marketing department in DC for consumables SL0004_P2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 


'''''''--------TransactionCode-VA03----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call PressEnter()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Order","VBAK-VBELN","",DT_SO_NUMBER,False)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call ClickButtonIfExist("Display document flow   \(F5\)", False)
Call ActivateNodeGuiTree(0, "#1;#1")
Call GetGridContent("Self Consumption.*", 0, "DOCNUM", 1, "<NA>", "<NA>", "DT_VA03_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1_OUTPUT")
'''Call GetNodeTextGuiTree(0, "Self Consumption 00"&DT_SO_NUMBER&" / 10", 1, "DT_VA03_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1_OUTPUT")

Call TakeScreenShot

'''''''--------TransactionCode-/nVL02N----------''''
Call SetTcode(DT_VL01N_4001_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call GetInputFromExcel(gstrInputExcelFilePathAndName, "Global",DataRowSet )
Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VA03_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1_OUTPUT,False)
Call PressEnter()     
Call TakeScreenShot

Call ClickButtonIfExist("Post Goods Issue   \(Shift\+F8\)",False)
Call TakeScreenShot
Call GetStatusBar("item2", "DT_VA03_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1_OUTPUT")
Call GetStatusBar("item3", "DT_VA03_0100_CHECK_TEXT_OF_TREE_DOCNUM_OUTPUT")
Call VerifyStatusBar("Self Consumption "&DT_VA03_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1_OUTPUT&" saved, article document "&DT_VA03_0100_CHECK_TEXT_OF_TREE_DOCNUM_OUTPUT&" created")


'''''''--------TransactionCode-/nVL03N----------''''
Call SetTcode(DT_VL01N_4004_OKCD)     
Call PressEnter()     
Call TakeScreenShot


''ENTER OUTBOUND DELIVERY
Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VA03_0100_CHECK_TEXT_OF_TREE_SAPTABLETREECONTROL1_OUTPUT,False)
Call PressEnter()  
Call TakeScreenShot

''Click on Document Flow (F7)
Call ClickButtonIfExist("Document Flow   \(F7\)",False)
Call TakeScreenShot

'' Select  node3 form tree
Call ActivateNodeGuiTree(0, "#1;#1;#1")

Call ClickButtonIfExist("Details   \(F2\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("Self-consumption.*", 1, "Status", 0, DT_VL01N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_STATUS)
Call VerifyGridCellContent("Self-consumption.*", 2, "Status", 0, DT_VL01N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_1_STATUS)


Call GetGridContent("Self-consumption.*", 0, "Doc.no.", 1, "<NA>", "<NA>", "DT_VL01N_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DOCNUM_OUTPUT")

'' Click on Display document
Call ClickButtonIfExist("Display document   \(F8\)",False)
Call TakeScreenShot
'' Get art.Slip value
Call GetTextboxValue("RM07M-MTSNR", 0, "DT_VL01N_0420_CHECK_TEXT_OF_ART_SLIP_OUTPUT", False)
''Verify Art.Slip
Call VerifyTextBoxContent("Art\. Slip", "RM07M-MTSNR", 0, DT_VL01N_0420_CHECK_TEXT_OF_ART_SLIP_OUTPUT, False)

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


