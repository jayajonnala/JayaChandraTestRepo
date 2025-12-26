'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Fixed Assets Sales_p6_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 22th March
'.................Modified By :
'.................Modified Date/Details :
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

gstrTestCaseName = "Test_Fixed Assets Sales_p6_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Fixed Assets Sales_p7_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
''''Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode ZFIAR_RS_RFKORD50PDF----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call SetTextbox("Variant","V-LOW","",DT_REF_VALUE,True)
Call SetTextbox("Created By","ENAME-LOW","",DT_CREATED_BY,True)

'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)

'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RBELNR-LOW","",DT_ZFIAR_RS_RFKORD50PDF_1000_DOCUMENT_NUMBER,False)
Call SetTextbox("Fiscal Year","RGJAHR-LOW","",DT_ZFIAR_RS_RFKORD50PDF_0010_FISCAL_YEAR,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Wait(2)
'Capture the screenshot
Call TakeScreenShot()
''
''''----------------------Tcode AB08----------------------------
''
''Enter the Tcode
Call SetTcode(DT_ZFIAR_RS_RFKORD50PDF_0500_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ZFIAR_RS_RFKORD50PDF_0500_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Company Code","RLAB01-BUKRS","",DT_ZFIAR_RS_RFKORD50PDF_0010_COMPANY_CODE,False)
Call SetTextbox("Asset","RLAB01-ANLN1","",DT_ZFIAR_RS_RFKORD50PDF_0010_ASSET,False)
Call SetTextbox("Sub-number","RLAB01-ANLN2","",DT_ZFIAR_RS_RFKORD50PDF_0010_SUBNUMBER,False)
Call SetTextbox("Fiscal Year","RLAB01-GJAHR","",DT_ZFIAR_RS_RFKORD50PDF_0010_FISCAL_YEAR,False)

Call PressEnter() 

'Capture the screenshot
Call TakeScreenShot()

'select apporiapiate row as per data sheet
If DT_ZFIAR_RS_RFKORD50PDF_0100_SAPLAB01TC_ZAB01_D0100 = "True" Then
	Call SelectRowGuiTableByRow("SAPLAB01TC_ZAB01_D0100",1,False)
End If

If DT_ZFIAR_RS_RFKORD50PDF_0100_SAPLAB01TC_ZAB01_D0100_OCC1 = "True" Then
	Call SelectRowGuiTableByRow("SAPLAB01TC_ZAB01_D0100",2,False)
End If
'set focus on the celll
Call SelectCellGuiTable("SAPLAB01TC_ZAB01_D0100","AstValDt","TTy",DT_Tty,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Reverse Document   \(F6\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Continue   \(Enter\)",True)
'Capture the screenshot
Call TakeScreenShot()

wait(1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Reversal Reason","UF05A-STGRD","",DT_ZFIAR_RS_RFKORD50PDF_0105_REVERSAL_REASON,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If asset is created
Call GetStatusBar("item1","DT_ZFIAR_RS_RFKORD50PDF_0010_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_ZFIAR_RS_RFKORD50PDF_0010_CHECK_TEXT_OF_STATUSBAR)

'
'''----------------------Tcode FB03----------------------------
'
'Enter the Tcode
Call SetTcode(DT_ZFIAR_RS_RFKORD50PDF_0010_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ZFIAR_RS_RFKORD50PDF_0010_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Document Number","RF05L-BELNR","",DT_ZFIAR_RS_RFKORD50PDF_0100_DOCUMENT_NUMBER,False)
Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_ZFIAR_RS_RFKORD50PDF_0010_FISCAL_YEAR,False)
Call SetTextbox("Company Code","RF05L-BUKRS","",DT_ZFIAR_RS_RFKORD50PDF_0010_COMPANY_CODE,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

'Validate text boxes
Call VerifyTextBoxContent("Document Number","BKPF-BELNR","",DT_ZFIAR_RS_RFKORD50PDF_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER,False)
Call VerifyTextBoxContent("Company Code","BKPF-BUKRS","",DT_ZFIAR_RS_RFKORD50PDF_0750_CHECK_TEXT_OF_COMPANY_CODE,False)

'validate grid components
call VerifyGridCellContent("",1,"Posting Key","",DT_ZFIAR_RS_RFKORD50PDF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
call VerifyGridCellContent("",2,"Posting Key","",DT_ZFIAR_RS_RFKORD50PDF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
call VerifyGridCellContent("",3,"Posting Key","",DT_ZFIAR_RS_RFKORD50PDF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
call VerifyGridCellContent("",4,"Posting Key","",DT_ZFIAR_RS_RFKORD50PDF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BSCHL)
call VerifyGridCellContent("",5,"Posting Key","",DT_ZFIAR_RS_RFKORD50PDF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_BSCHL)
call VerifyGridCellContent("",6,"Posting Key","",DT_ZFIAR_RS_RFKORD50PDF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_BSCHL)

call VerifyGridCellContent("",1,"Account","",DT_ZFIAR_RS_RFKORD50PDF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
call VerifyGridCellContent("",2,"Account","",DT_ZFIAR_RS_RFKORD50PDF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
call VerifyGridCellContent("",3,"Account","",DT_ZFIAR_RS_RFKORD50PDF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
call VerifyGridCellContent("",4,"Account","",DT_ZFIAR_RS_RFKORD50PDF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_3_KTONR)
call VerifyGridCellContent("",5,"Account","",DT_ZFIAR_RS_RFKORD50PDF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_4_KTONR)
call VerifyGridCellContent("",6,"Account","",DT_ZFIAR_RS_RFKORD50PDF_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_5_KTONR)

'Capture the screenshot
Call TakeScreenShot()

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

