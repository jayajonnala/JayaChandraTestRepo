
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_P2P_01_01_0114-Number Range for manual Return DN done in SAP_P2_CreateRDN_First
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
'.................Test Script Name :[1]Test_PRE_MI10_Add_Stock_to_SLOC
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)


gstrTestCaseName = "Test_P2P_01_01_0114-Number  in SAP_P2_CreateRDN_First"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_P2P_01_01_0114-Number Range for manual Return DN done in SAP_P2_CreateRDN_X2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)

Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter()
Call TakeScreenShot()

Call SetTextbox("Article","MATNR-LOW",1,DT_MB52_1000_ARTICLE,False)
Call SetTextbox("Site","WERKS-LOW",2,DT_MB52_1000_SITE,False)
Call SetTextbox("Storage Location","LGORT-LOW",3,DT_MB52_1000_STORAGE_LOCATION,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call GetLabelContentByRefLabel("Unrestr\.",0,-32,"DT_MB52_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT",False)
Call TakeScreenShot
Call SetTcode(DT_MB52_0120_OKCD)
Call PressEnter()
Call TakeScreenShot()
Call SetTextbox("Movement Type","RM07M-BWARTWE","",DT_MB52_0200_MOVEMENT_TYPE,False)
Call SetTextbox("Site","RM07M-WERKS","",DT_MB52_0200_SITE,False)
Call SetTextbox("Storage Location","RM07M-LGORT","",DT_MB52_0200_STORAGE_LOCATION,False)
Call PressEnter()


If SAPGuisession(sessionObject).sapguiwindow(windowObject).sapguiedit("guicomponenttype:=32","name:=MSEGK-LIFNR","Index:=0").Exist(5) Then
wait 1
Else
Call PressEnter()
End If
Call SetTextboxNoLabel("MSEGK-LIFNR","","10003277",False)

Call SetTextbox("Article","MSEG-MATNR","",DT_MB52_0241_ARTICLE,False)
Call SetTextbox("Quantity","MSEG-ERFMG","",DT_MB52_0241_QUANTITY,False)
Call PressEnter()
Call TakeScreenShot()
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_MB52_0200_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_MB52_0200_CHECK_TEXT_OF_STATUSBAR_OUTPUT&" posted")
Call SetTcode(DT_MB52_0200_OKCD)
Call PressEnter()
Call TakeScreenShot()
Call SetTextbox("Article Doc\.","RM07M-MBLNR","",DT_MB52_0200_CHECK_TEXT_OF_STATUSBAR_OUTPUT,False)
Call TakeScreenShot()
Call PressEnter()
Call GetTextboxValue("RM07M-LFSNR",0,"DT_MB52_0221_CHECK_TEXT_OF_DELIVNOTE_OUTPUT",False)
Call VerifyTextBoxContent("Deliv\.Note","RM07M-LFSNR",0,DT_MB52_0221_CHECK_TEXT_OF_DELIVNOTE_OUTPUT,False)


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




