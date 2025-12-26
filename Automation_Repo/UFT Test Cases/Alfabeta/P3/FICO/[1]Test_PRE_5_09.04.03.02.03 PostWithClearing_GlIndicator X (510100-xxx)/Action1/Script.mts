	

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_POST_DeleteVAT_from_Customer
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

gstrTestCaseName = "Test_PRE_5_09.04.03.02.03 PostWithClearing_GlIndicator X (510100-xxx)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_PRE_5_09.04.03.02.03 PostWithClearing_GlIndicator X (510100-xxx)_Output.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------MIGO -----------------------------------

Call SetTextbox("Document Date","BKPF-BLDAT","",Replace(DT_F40_0122_DOCUMENT_DATE,"/","."),False)     ' - Line (19)
Call SetTextbox("Period","BKPF-MONAT","",DT_F40_0122_PERIOD,False)     ' - Line (19)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F40_0122_COMPANY_CODE,False)     ' - Line (19)

Call SetTextbox("Clearing text","RF05A-AUGTX","",DT_F40_0122_CLEARING_TEXT,False)     ' - Line (19)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F40_0122_REFERENCE,False)     ' - Line (19)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F40_0122_DOCHEADER_TEXT,False)     ' - Line (19)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F40_0122_ACCOUNT,False)     ' - Line (19)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_F40_0122_SGL_IND,False)     ' - Line (19)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F40_0122_CURRENCYRATE,False)     ' - Line (19)
Call PressEnter() 
'Call PressEnter()
'call SendKey("{ENTER}0")
call VerifyTextBoxContent("Vendor","RF05A-KTNRA","",DT_F40_0320_CHECK_TEXT_OF_VENDOR,False)
TakeScreenShot()

'''Bill exchange add vendor item '''

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F40_0320_AMOUNT,False)     ' - Line (19)
wait(2)
Call SetTextbox("Assignment","BSEG-ZUONR","",DT_F40_0320_ASSIGNMENT,False)     ' - Line (19)
wait(2)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F40_0320_TEXT,False)     ' - Line (19)
wait(2)
Call SetTextbox("Due On","BSEG-ZFBDT","",replace(DT_F40_0320_DUE_ON,"/","."),False)     ' - Line (19)
wait(2)
Call SetTextbox("Check No\.","BSED-BOENO","",DT_F40_0320_CHECK_NO,False)     ' - Line (19)
wait(2)
Call PressEnter()
call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
TakeScreenShot()

''Bill of Exchange Payment Display Overview '''
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_PRE_5_09.04.03.02.03 PostWithClearing_GlIndicator X (510100-xxx).xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
'
'Call StartExecution(gstrInputExcelFilePathAndName,"Global",2)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

call VerifyTextBoxContent("Document Date","BKPF-BLDAT","",REPLACE(DT_F40_0700_CHECK_TEXT_OF_DOCUMENT_DATE,"/","."),False)
call VerifyTextBoxContent("Posting Date","BKPF-BUDAT","",Replace(DT_F40_0700_CHECK_TEXT_OF_POSTING_DATE,"/","."),False)
call VerifyTextBoxContent("D","RF05A-AZSOL","",DT_F40_0700_CHECK_TEXT_OF_D,False)
call VerifyTextBoxContent("C","RF05A-AZHAB","",(DT_F40_0700_CHECK_TEXT_OF_C &",00"),False)
call VerifyTextBoxContent("C","RF05A-AZSAL","",Replace(DT_F40_0700_CHECK_TEXT_OF_C_OCC1,".",","),False)
call VerifyTextBoxContent("C","RF05A-AZSAL","",Replace(DT_F40_0700_CHECK_TEXT_OF_C_OCC1,".",","),False)

' VerifyifGuiLabelExists(Content)
call VerifyifGuiLabelExists(DT_F40_0700_CHECK_TEXT_OF_TEXT1)
call ClickButton("Process Open Items   \(Shift\+F4\)",False)
call ClickButton("btn\[16\]",False)
TakeScreenShot()

'''Bill of Exchange Payment Select open items '''
'''need to check
'Call SetTextbox("Account Type","RF05A-AGKOA","",DT_AccountType,False)     ' - Line (19)
'Call PressEnter()
'''check'''
TakeScreenShot()
call ClickButton("ICON_SELECT_ALL",False)
call ClickButton("Deactivate Items",False)
call ClickButton("ICON_SEARCH",False)
call SelectRadioButton("RF05A-XPOS1","Posting Date",True)
call ClickButton("btn\[0\]",False)
Call SetTextbox("From","RF05A-VONDT","",replace(DT_F40_0732_FROM,"/","."),true)     ' - Line (19)
Call SetTextbox("To","RF05A-BISDT","",replace(DT_F40_0732_TO,"/","."),True)     ' - Line (19)
call ClickButton("btn\[0\]",False)
TakeScreenShot()
call ClickButton("ICON_SELECT_ALL",False)
call ClickButton("Activate Items",False)
call GetTableCellData("SAPDF05XTC_6102","Document Number",1,"","",DT_F40_6102_CHECK_TEXT_OF_TABLECELL_DOCUMENT_NUMBER_0,False)
TakeScreenShot()
call VerifyTableCellContent(1,"Document Date","SAPDF05XTC_6102",replace(DT_F40_6102_CHECK_TEXT_OF_TABLECELL_DOCUMENT_DATE_0,"/","."))
call VerifyTableCellContent(1,"EUR Gross","SAPDF05XTC_6102",DT_F40_6102_CHECK_TEXT_OF_TABLECELL______EUR_GROSS_0)
call ClickButton("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_PO_NUMBER")  


'''close application'''
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


