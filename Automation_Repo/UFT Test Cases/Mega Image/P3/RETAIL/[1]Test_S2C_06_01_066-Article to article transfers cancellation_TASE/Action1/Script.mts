

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_S2C_06_01_066-Article to article transfers cancellation 
      
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



gstrTestCaseName = "Test_S2C_06_01_066-Article to article transfers cancellation"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\TASE_DT_S2C_06_01_066-Article to article transfers cancellation.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet = 2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''

'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

''''--------TransactionCode-MIGO----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call ClickButtonIfExist("btn\[21\]", False)
Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_TR_0010_GODYNPROREFDOC) 
Call SetComboByKey("GODYNPRO-REFDOC",DT_MIGO_TR_0010_GODYNPROREFDOC_OCC1)
Call SelectTab("TS_GOITEM","Where", False)
Call SetTextbox("Movement type","GOITEM-BWART","",DT_MIGO_TR_0010_GODEFAULT_TVBWART,False)
Call TakeScreenShot
Call PressEnter()
Call SelectTab("TS_GOITEM","Transfer Posting", False)
Call TakeScreenShot
Call SetTextbox("Article","GODYNPRO-MAKTX","",DT_MIGO_TR_0390_ARTICLE,False)
'Call SetTextbox("Qty in UnE","GODYNPRO-ERFME","",DT_MIGO_TR_0390_QTY_IN_UNE,False)
Call SetTextbox("Qty in UnE","GODYNPRO-ERFMG","",DT_MIGO_TR_0390_QTY_IN_UNE,False)
'''Call SetTextbox("Article Trfr Pstg","GOITEM-UMMAKTX","",DT_MIGO_TR_0390_GOITEMUMMAKTX,False)
Call SetTextboxNoLabel("GOITEM-UMMAKTX","",DT_MIGO_TR_0390_GOITEMUMMAKTX,False)
Call SetTextbox("Site","GODYNPRO-NAME1","",DT_MIGO_TR_0390_SITE,False)
'''Call SetTextbox("Site Trfr Pstg","GOITEM-UMNAME1","",DT_MIGO_TR_0390_GOITEMUMNAME1,False)
Call SetTextboxNoLabel("GOITEM-UMNAME1","",DT_MIGO_TR_0390_GOITEMUMNAME1,False)
Call SetTextbox("Stor\. Loc\.","GODYNPRO-LGOBE","",DT_MIGO_TR_0390_STOR_LOC,False)
'''Call SetTextbox("SLoc Transfer Pstg","GOITEM-UMLGOBE","",DT_MIGO_TR_0390_STOR_LOC,False)
Call SetTextboxNoLabel("GOITEM-UMLGOBE","",DT_MIGO_TR_0390_STOR_LOC,False)

Call TakeScreenShot
Call PressEnter() 
Call ClickButton("Post Document   \(Shift\+F11\)",false)
Call GetStatusBar("item1","DT_GET_GR_TR_ARTICLE_DOCUMENT_NO_OUTPUT")
Call VerifyStatusBar("Article document "&DT_GET_GR_TR_ARTICLE_DOCUMENT_NO_OUTPUT&" posted")


''''--------TransactionCode-MBST----------''''

Call SetTcode(DT_MIGO_TR_0001_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call SetTextbox("Article Doc\.","RM07M-MBLNR","",DT_GET_GR_TR_ARTICLE_DOCUMENT_NO_OUTPUT,False)
Call SetTextbox("Art\. Doc\. Year","RM07M-MJAHR","",DT_MIGO_TR_0460_ART_DOC_YEAR,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call GetStatusBar("item1","DT_GET_TR_CANCEL_ARTICLE_DOCUMENT_NO_OUTPUT")
Call VerifyStatusBar("Document "&DT_GET_TR_CANCEL_ARTICLE_DOCUMENT_NO_OUTPUT&" posted")


''''--------TransactionCode-MBST----------''''

Call SetTcode(DT_MIGO_TR_0460_OKCD)     
Call PressEnter()     
Call SetComboByKey("GODYNPRO-ACTION", DT_MIGO_TR_0010_GODYNPROACTION)
Call TakeScreenShot
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC", 0, DT_GET_TR_CANCEL_ARTICLE_DOCUMENT_NO_OUTPUT, False)
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR", 0, DT_MIGO_TR_2010_GODYNPRODOC_YEAR, False)
Call TakeScreenShot
Call ClickButton("Execute   \(Enter\)",False)
Call TakeScreenShot
Call SelectTab("TS_GOHEAD","Doc. info", False)
Call ClickButtonIfExist("FI Documents",False)
Call TakeScreenShot


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


