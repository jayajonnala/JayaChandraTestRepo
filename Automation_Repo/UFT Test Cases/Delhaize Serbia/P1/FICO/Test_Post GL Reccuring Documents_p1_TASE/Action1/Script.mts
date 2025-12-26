
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Post GL Reccuring Documents_p1
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
'.................Test Script Name : Test_Post GL Reccuring Documents_p1
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 19th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Post GL Reccuring Documents_p1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Post GL Reccuring Documents_p1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode FBD1----------------------------

'Create Purchase Order
Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))

'Enter the details
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FBD1_0106_COMPANY_CODE,False)   
Call SetTextbox("First run on","BKDF-DBBDT","",ConvertDate(DT_FBD1_0106_FIRST_RUN_ON),False)
Call SetTextbox("Last run on","BKDF-DBEDT","",ConvertDate(DT_FBD1_0106_LAST_RUN_ON),False)   
Call SetTextbox("Interval in months","BKDF-DBMON","",DT_FBD1_0106_INTERVAL_IN_MONTHS,False)   
Call SetTextbox("Document type","BKPF-BLART","",DT_FBD1_0106_DOCUMENT_TYPE,False) 
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_FBD1_0106_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_FBD1_0106_REFERENCE,False)
'Call SetTextbox("Translation dte","BKPF-WWERT","",ConvertDate(DT_FBD1_0106_TRANSLATION_DTE),False)
Call SetTextboxNoLabel("BKPF-WWERT","",ConvertDate(DT_FBD1_0106_TRANSLATION_DTE),False)
Call SetTextbox("Document Header Text","BKPF-BKTXT","",DT_FBD1_0106_DOCUMENT_HEADER_TEXT,False)
Call TakeScreenShot()


'Click on Fast Entry
Call ClickButton("Call Up Fast Entry   \(Shift\+F8\)",False) 
Wait(1)
Call TakeScreenShot()


Call SetTextbox("PK","BSEG-BSCHL","",DT_FBD1_0310_PK,False)
Call FocusTextBoxByIndex("Account","RF05A-KONTO",0,False)
Call SendKey("{F4}")
Call SelectTab("G_SELONETABSTRIP",DT_FBD1_SELECT_TAB,True)
Call SetTextbox("Language Key","G_SELFLD_TAB-LOW","","",True)
Call SetTextbox("G/L account","G_SELFLD_TAB-LOW","",DT_FBD1_0220_GL_ACCOUNT,True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call ClickButton("Find   \(Ctrl\+F\)",True)
Call SetTextbox("Find","RSYSF-STRING","",DT_FBD1_0800_FIND,True)
Call TakeScreenShot()
Call ClickButton("Find   \(Enter\)",True)
Call TakeScreenShot()
Call SetFocusGuiLabel(DT_FBD1_0800_FIND,11,40,True)
Wait 5
Call SendKey("{F2}")
Call ClickButton("Copy   \(Enter\)",True)
'''Call SetTextbox("Account","RF05A-KONTO","",DT_FBD1_0220_GL_ACCOUNT,False)
Call SetTextbox("Amount","BSEG-WRBTR","",DT_FBD1_0310_AMOUNT,False)
Call SetTextbox("B\.Area","BSEG-GSBER","",DT_FBD1_0310_BAREA,False) 
Call SetTextbox("Cost Center","BSEG-KOSTL","",DT_FBD1_0310_COST_CENTER,False) 

Call SetTextbox("PK","BSEG-BSCHL",1,DT_FBD1_0310_PK_OCC1,False)
Call FocusTextBoxByIndex("Account","RF05A-KONTO",1,False)
Wait 5
Call SendKey("{F4}")
Call SetTextbox("Language Key","G_SELFLD_TAB-LOW","","",True)
Call SetTextbox("G/L account","G_SELFLD_TAB-LOW","",DT_FBD1_0220_GL_ACCOUNT_OCC1,True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call ClickButton("Find   \(Ctrl\+F\)",True)
Call SetTextbox("Find","RSYSF-STRING","",DT_FBD1_0800_FIND_OCC1,True)
Call TakeScreenShot()
Call ClickButton("Find   \(Enter\)",True)
Call TakeScreenShot()
Call SetFocusGuiLabel(DT_FBD1_0800_FIND_OCC1,11,40,True)
Wait 5
Call SendKey("{F2}")
Call ClickButton("Copy   \(Enter\)",True)
'''Call SetTextbox("Account","RF05A-KONTO",1,DT_FBD1_0220_GL_ACCOUNT_OCC1,False)
Call SetTextbox("Amount","BSEG-WRBTR",1,DT_FBD1_0310_AMOUNT_OCC1,False)
Call SetTextbox("B\.Area","BSEG-GSBER",1,DT_FBD1_0310_BAREA_OCC1,False)
Call SetTextbox("Cost Center","BSEG-KOSTL",1,DT_FBD1_0310_COST_CENTER_OCC1,False)
Call TakeScreenShot()
Call PressEnter() 
Wait(2)
Call TakeScreenShot()


'Post the Document No
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)  
Call GetStatusBar("item1","DT_DOCUMENT_NO_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_FBD1_0106_CHECK_TEXT_OF_STATUSBAR)


''----------------------Tcode F.14----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_SAPTRANSACTIONCODE2) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE2)
Call TakeScreenShot()


'Enter the details
Call SetTextbox("Company code","BR_BUKRS-LOW","",DT_F14_1000_COMPANY_CODE,False)
Call SetTextbox("Document Number","BR_BELNR-LOW","",DT_F14_1000_DOCUMENT_NUMBER,False)
Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","",DT_F14_1000_FISCAL_YEAR,False)
Call SetTextbox("Document type","BR_BLART-LOW","",DT_F14_1000_DOCUMENT_TYPE,False)
Call SetTextbox("Settlement period","AUSFDATE-LOW","",ConvertDate(DT_F14_1000_SETTLEMENT_PERIOD),False)
Call SetTextbox("to","AUSFDATE-HIGH","",ConvertDate(DT_F14_1000_TO),False)
Call SetTextbox("Batch input session name","MAPNAME","",DT_F14_1000_BATCH_INPUT_SESSION_NAME,False)
Call SelectCheckbox("BDC_KEEP",0,"ON",False)
Call TakeScreenShot()

'Click on Execute Button
Call ClickButton("Execute   \(F8\)",False) 
Wait(1)
Call VerifyStatusBar(DT_F14_1000_CHECK_TEXT_OF_STATUSBAR)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

