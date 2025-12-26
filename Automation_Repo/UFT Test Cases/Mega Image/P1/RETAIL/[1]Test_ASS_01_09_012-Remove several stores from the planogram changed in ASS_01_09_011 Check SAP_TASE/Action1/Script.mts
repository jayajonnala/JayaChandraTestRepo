
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_ASS_01_09_012-Remove several stores from the planogram changed in ASS_01_09_011 Check SAP
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

gstrTestCaseName = "Test_ASS_01_09_012-  ASS_01_09_011 Check SAP"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\MI\RETAIL\TASE_DT_ASS_01_09_012-Remove several stores from the planogram changed in ASS_01_09_011  Check SAP.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''--------------login----------------'''''

''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


'''--------TransactionCode-ZMDPU_INFOREC_COPY----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Assortment","V_WRS1-ASORT","",DT_WSOA2_0001_ASSORTMENT,False)
Call PressEnter()
Call TakeScreenShot
Call SelectTab("ASORT_TAB", "Assortment User", False)
Call TakeScreenShot

Call SelectRowGuiTableByRow("WRFM_WSO6DYN040",6,False)
Call ClickButton("Delete",false)
Call SelectRowGuiTableByRow("WRFM_WSO6DYN040",7,False)
Call ClickButton("Delete",false)
Call SelectRowGuiTableByRow("WRFM_WSO6DYN040",8,False)
Call ClickButton("Delete",false)

Call ClickButton("Save   \(Ctrl\+S\)",false)

Call VerifyStatusBar(DT_WSOA2_0001_CHECK_TEXT_OF_STATUSBAR)
Call SetTcode(DT_WSOA2_0001_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Assortment","S_FILIA-LOW","",DT_WSOA2_1000_ASSORTMENT,False)
Call SetTextbox("Article","S_ARTNR-LOW","",DT_WSOA2_1000_ARTICLE,False)
Call SetTextbox("Valid From","P_DATAB","",ConvertDate(DT_WSOA2_1000_VALID_FROM),False)
Call SetTextbox("Customer No\. - Site","S_LOCNR-LOW","",DT_WSOA2_1000_CUSTOMER_NO__SITE,False)
Call SetTextbox("to","S_LOCNR-HIGH","",DT_WSOA2_1000_TO,False)
Call PressEnter()   

Call ClickButton("Execute   \(F8\)",false)
Call VerifyStatusBar(DT_WSOA2_1000_CHECK_TEXT_OF_STATUSBAR)
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




