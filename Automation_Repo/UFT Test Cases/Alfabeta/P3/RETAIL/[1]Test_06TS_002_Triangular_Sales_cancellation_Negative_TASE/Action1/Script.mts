

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_06TS_002_Triangular_Sales_cancellation_Negative
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
	GetRowNo= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
gstrTestCaseName = "Test_06TS_002_Cancel_Negative"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\jjonn\Desktop\TASEWork\Data\TASE_DT_09.04.01.01.13 Manage Manual Foreign Vendor Crediting.xls"
'strResultFolderPath = "C:\Users\jjonn\Desktop\TASEWork\Results"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'GetRowNo=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''


'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''''--------------------------------ZMDSO_TRIANG_SO--------------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)  
Call PressEnter()   
call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call ClickButtonIfExist("Execute   \(F8\)",True) 

Call SelectRowGuiGridbyRowNo("Variant Catalog for Program ZMDSO_TRIANGULAR_SO_FROM_GR", "", 1, Truew)
Call ClickButtonIfExist("Choose   \(F2\)",True) 

Call SelectCheckbox("P_TEST",0,DT_ZMDSO_TRIANG_SO_1000_TEST_MODE,False)
Call SelectCheckbox("P_DISPSO",0,DT_ZMDSO_TRIANG_SO_1000_DISPLAY_EXISTING_SO,False)

call SetComboByKey("P_ACTION",DT_ZMDSO_TRIANG_SO_1000_ACTION)
Call SetTextbox("Article Document","P_MBLNR","",DT_ZMDSO_TRIANG_SO_1000_ARTICLE_DOCUMENT,False)
Call SetTextbox("Article Doc\. Year","P_MJAHR","",Year(DT_ZMDSO_TRIANG_SO_1000_ARTICLE_DOC_YEAR),False)
Call SetTextbox("Posting Date","S_BUDAT-LOW","",ConvertDate(DT_ZMDSO_TRIANG_SO_1000_POSTING_DATE),False)
Call SetTextbox("to","S_BUDAT-HIGH","",ConvertDate(DT_ZMDSO_TRIANG_SO_1000_TO),False)
Call SetTextbox("Site","S_WERKS-LOW","",DT_ZMDSO_TRIANG_SO_1000_SITE,False)
Call SetTextbox("to","S_WERKS-HIGH","",DT_ZMDSO_TRIANG_SO_1000_TO_OCC1,False)
Call SetTextbox("Movement type","S_BWART-LOW","",DT_ZMDSO_TRIANG_SO_1000_MOVEMENT_TYPE,False)
Call SetTextbox("to","S_BWART-HIGH","",DT_ZMDSO_TRIANG_SO_1000_TO_OCC2,False)
Call SetTextbox("Trans\./event type","S_VGABE-LOW","",DT_ZMDSO_TRIANG_SO_1000_TRANSEVENT_TYPE,False)

Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)

Call FindRowNumber("","Article Document",DT_ZMDSO_TRIANG_SO_1000_ARTICLE_DOCUMENT,"DT_ROW_NO_PO_OUTPUT")

Call VerifyGridCellContentbyName("shell",DT_ROW_NO_PO_OUTPUT,"STATUS","",DT_ZMDSO_TRIANG_SO_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_STATUS)
Call VerifyGridCellContentbyName("shell",DT_ROW_NO_PO_OUTPUT,"Article Document","",DT_ZMDSO_TRIANG_SO_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MBLNR)

Call LogOff'
Call FinalStatus()
