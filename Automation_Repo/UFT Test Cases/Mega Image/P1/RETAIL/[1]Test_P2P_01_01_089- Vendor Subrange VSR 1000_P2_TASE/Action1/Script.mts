
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

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0234-Scrapping of products - non-deductible cancellation MVT 985_Dry_P1
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_P2P_01_01_089- VSR 1000_P2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\MI\DT_P2P_01_01_0234-Scrapping of products - non-deductible cancellation MVT 985_Dry_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter() 
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION)

Call SetComboByKey("GODYNPRO-REFDOC", DT_MIGO_0010_GODYNPROREFDOC) 

Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER", 0, DT_MIGO_2000_GODYNPROPO_NUMBER, False)

Call SetTextBox("Delivery Note","GOHEAD-LFSNR",0,DT_MIGO_0110_DELIVERY_NOTE,False)

Call PressEnter()

Call SelectCheckbox("GODYNPRO-DETAIL_TAKE", 0, "ON", False)

Call ClickButton("Check Entries   \(F7\)",False)

Call TakeScreenShot()

Call ClickButton("Post Document   \(Shift\+F11\)",False)

Call TakeScreenShot()

Call GetStatusBar("item1","DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")

'Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR_OCC1)

Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_0010_GODYNPROACTION_OCC1)

'Call SetTextboxNoLabel("GODYNPRO-PO_NUMBER", 0, DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT, False)

Call TakeScreenShot()

Call PressEnter()



Call  GetTextboxValue("GOHEAD-LFSNR", 0, "DT_MIGO_0110_CHECK_TEXT_OF_DELIVERY_NOTE_OUTPUT", False)

Call SetTcode(DT_MIGO_0001_OKCD)

Call PressEnter()

Call TakeScreenShot()

Call SetTextBox("Sort order","PM_NSORT",0,DT_MIGO_1000_SORT_ORDER,False)

Call SetTextBox("Processing Mode","PM_VERMO",0,DT_MIGO_1000_PROCESSING_MODE,False)

Call SetTextBox("Article Doc\. Year","PM_MJAHR",0,DT_MIGO_1000_ARTICLE_DOC_YEAR,False)

'Call SetTextBox("Article Document","RG_MBLNR-LOW",0,DT_MIGO_1000_ARTICLE_DOCUMENT,False)

Call TakeScreenShot()

Call PressEnter()

Call ClickButton("Execute   \(F8\)",False)

Call SelectCheckboxNoLabel(0, "ON", False)

Call ClickButton("Print preview   \(Shift\+F4\)",False)

Call TakeScreenShot()

Call LogOff()

Call FinalStatus()






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




