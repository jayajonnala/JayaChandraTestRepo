

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.05.01.01.02 Preparation before Mass Payment Run - Update ven_old
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

gstrTestCaseName = "TC01_Test_09.05.01.01.02 - Update ven_old"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

''--------------------------------------------  ZFIAP_OPEN_GR_IMP_DN------------------------------------------------

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT_REFERENCE",Cint(DT_INCREMENT_REFERENCE)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SelectCheckbox("P_TEST",0,DT_ZFIAP_OPEN_GR_IMP_DN_1000_TEST_MODE,False)
Call SetTextbox("Company Code","P_BUKRS","",DT_ZFIAP_OPEN_GR_IMP_DN_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","P_GJAHR","",DT_ZFIAP_OPEN_GR_IMP_DN_1000_FISCAL_YEAR,False)
Call SetTextbox("G/L Account","S_HKONT-LOW","",DT_ZFIAP_OPEN_GR_IMP_DN_1000_GL_ACCOUNT,False)
Call SetTextbox("Posting Date","S_BUDAT-LOW","",ConvertDate(DT_ZFIAP_OPEN_GR_IMP_DN_1000_POSTING_DATE),False)
Call SetTextbox("to","S_BUDAT-HIGH","",ConvertDate(DT_ZFIAP_OPEN_GR_IMP_DN_1000_TO),False)
Call SetTextbox("Document type","S_BLART-LOW","",DT_ZFIAP_OPEN_GR_IMP_DN_1000_DOCUMENT_TYPE,False)
Call SetTextbox("Posting Key","S_BSCHL-LOW","",DT_ZFIAP_OPEN_GR_IMP_DN_1000_POSTING_KEY,False)

Call SetTextbox("Posting Date","P_BUDAT","",ConvertDate(DT_ZFIAP_OPEN_GR_IMP_DN_1000_POSTING_DATE_OCC1),False)
Call SetTextbox("Document Date","P_BLDAT","",ConvertDate(DT_ZFIAP_OPEN_GR_IMP_DN_1000_DOCUMENT_DATE),False)
Call SetTextbox("Document type","P_BLART","",DT_ZFIAP_OPEN_GR_IMP_DN_1000_DOCUMENT_TYPE_OCC1,False)
Call SetTextbox("Document Header Text","P_BKTXT","",DT_ZFIAP_OPEN_GR_IMP_DN_1000_DOCUMENT_HEADER_TEXT,False)
Call SetTextbox("Reference","P_XBLNR","",DT_ZFIAP_OPEN_GR_IMP_DN_1000_REFERENCE,False)
Call SetTextbox("Payment Block","P_ZLSPR","",DT_ZFIAP_OPEN_GR_IMP_DN_1000_PAYMENT_BLOCK,False)
Call SetTextbox("Text","P_SGTXT","",DT_ZFIAP_OPEN_GR_IMP_DN_1000_TEXT,False)

Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()
Call SelectAllRowGuiGrid("",0,False)
Call ClickButton("Save   \(Ctrl\+S\)",False)
'Call SendKey("{F5}")

Call GetGridContent("",0,"LIFNR",1,"<NA>","<NA>","DT_ZFIAP_OPEN_GR_IMP_DN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_LIFNR_OUTPUT")
Call GetGridContent("",0,"TOTAL_AMNT",1,"<NA>","<NA>","DT_ZFIAP_OPEN_GR_IMP_DN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_TOTAL_AMNT_OUTPUT")
Call GetGridContent("",0,"MESSAGE",1,"<NA>","<NA>","DT_ZFIAP_OPEN_GR_IMP_DN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MESSAGE_OUTPUT")

Call VerifyGridCellContent("",1,"MESSAGE",0,DT_ZFIAP_OPEN_GR_IMP_DN_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MESSAGE_OUTPUT)


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


