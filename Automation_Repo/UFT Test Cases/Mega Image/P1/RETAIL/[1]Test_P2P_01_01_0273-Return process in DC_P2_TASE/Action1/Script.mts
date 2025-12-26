
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0273-Return process in DC_P2
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
'.................Test Script Name :Test_P2P_01_01_0273-Return process in DC_P2
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_P2P_01_01_0273-Return process in DC_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\smasu\Documents\TASE\Data Input\MI\TASE_DT_P2P_01_01_0273-Return process in DC_P2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)

Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)     

Call PressEnter()     

Call TakeScreenShot()
'
Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION)

Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_0010_GODYNPROREFDOC)

Call TakeScreenShot()

Call SetTextBoxNolabel("GODEFAULT_TV-BWART",0,DT_MIGO_0010_GODEFAULT_TVBWART,False)

Call SetTextBoxNolabel("GODYNPRO-PO_NUMBER",0,DT_MIGO_2000_GODYNPROPO_NUMBER,False)

 Call TakeScreenShot()
 
 Call SetTextBox("Delivery Note","GOHEAD-LFSNR",0,DT_MIGO_0110_DELIVERY_NOTE,False)
 
Call TakeScreenShot()

Call PressEnter()
Call SelectCheckBox("GODYNPRO-DETAIL_TAKE",0,"ON",False)

Call SetTableData("SAPLMIGOTV_GOITEM", "OK", 2, "", "", "ON", False)

Call ClickButton("Check Entries   \(F7\)",False)

Call VerifyStatusBar(DT_MIGO_0001_CHECK_TEXT_OF_STATUSBAR)

Call TakeScreenShot()

Call SelectTab("TS_GOITEM","Where",False)

Call TakeScreenShot()

Call VerifyTextBoxContent("Movement type", "GOITEM-BWART", 0, DT_MIGO_0325_CHECK_TEXT_OF_MOVEMENT_TYPE, False)

Call ClickButton("btn\[23\]",False)

Call GetStatusBar("item1","DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")

Call VerifyStatusBar("Article Document "&DT_MIGO_0001_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" posted")

Call TakeScreenShot()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Call SetComboByKey("GODYNPRO-ACTION",DT_MIGO_0010_GODYNPROACTION_OCC1)

Call TakeScreenShot()




Call TakeScreenShot()

Call PressEnter()

Call GetTextboxValue("GOHEAD-LFSNR", 0," DT_MIGO_0110_CHECK_TEXT_OF_DELIVERY_NOTE_OUTPUT", False)

Call TakeScreenShot()

Call SelectTab("TS_GOHEAD","Doc. info",False)

Call TakeScreenShot()

Call ClickButton("OK_FI_DOC",False)

Call TakeScreenShot()
''Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)
''Call ClickButtonIfExist("Display Document   \(F2\)",True)
If SAPGuisession(sessionObject).sapguiwindow(windowObject).sapguiedit("guicomponenttype:=31","name:=BKPF-BELNR","attachedtext:=Document Number","Index:=0").Exist(1) Then
	wait 1
Else
Call SelectRowGuiGrid("Documents in Accounting", 0, "Object type text", "Accounting document", True)
Call ClickButtonIfExist("Display Document   \(F2\)",True)
End If
Call GetTextboxValue("BKPF-BELNR",0,"DT_MIGO_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT",False)

Call VerifyGridCellContent("", 1, "Posting Key", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)

Call VerifyGridCellContent("", 2, "Posting Key", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

Call VerifyGridCellContent("", 1, "Account", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)

Call VerifyGridCellContent("", 2, "Account", 0, DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)

Call LogOff()

Call FinalStatus()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''











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




