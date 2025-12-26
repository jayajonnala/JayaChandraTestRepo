	

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

gstrTestCaseName = "Test_09.04.03.02.03 Post With Clearing_Gl indicator X (510100-x)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_09.04.03.02.03PostWithClearingusingglindicatorX510100-x_Output.xls"


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
Call SetTextbox("Document Date","BKPF-BLDAT","",Replace(DT_F51_0122_DOCUMENT_DATE,"/","."),False)     ' - Line (19)
Call SetTextbox("Period","BKPF-MONAT","",DT_F51_0122_PERIOD,False)     ' - Line (19)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F51_0122_COMPANY_CODE,False)     ' - Line (19)
Call SetTextbox("Clearing text","RF05A-AUGTX","",DT_F51_0122_CLEARING_TEXT,False)     ' - Line (19)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F51_0122_REFERENCE,False)     ' - Line (19)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F51_0122_DOCHEADER_TEXT,False)     ' - Line (19)
Call SetTextbox("Type","BKPF-BLART","",DT_F51_0122_TYPE,False)     ' - Line (19)
Call SetTextbox("SGL Ind","RF05A-NEWUM","",DT_F51_SGL_INDIC,False)     ' - Line (19)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F51_0122_CURRENCYRATE,False)     ' - Line (19)
'Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F40_0122_CURRENCYRATE,False)     ' - Line (19)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F51_0122_PSTKY,False)     ' - Line (19)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F51_0122_ACCOUNT,False)     ' - Line (19)
Call PressEnter() 
Call PressEnter()
'call SendKey("{ENTER}0")

'''Nithya
'TakeScreenShot()



'''nithya




'''Bill exchange add vendor item '''

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F51_0300_AMOUNT,False)     ' - Line (19)
wait(2)
Call PressEnter()
Call PressEnter()
Call ClickButtonIfExist("Continue   \(Enter\)",False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F51_0300_TEXT,False)     ' - Line (19)
wait(2)
Call SetTextbox("Value date","BSEG-VALUT","",replace(DT_F51_0300_VALUE_DATE,"/","."),False)     ' - Line (19)
wait(2)
call ClickButton("btn\[7\]",False)

''''Post with clearing correct G/L account item'''

'Call SetTextbox("Value date","BSEG-VALUT","",DT_F51_0300_TEXT,False)     ' - Line (19)
Call SetTextbox("House Bank","BSEG-HBKID","",DT_F51_0330_HOUSE_BANK,False)     ' - Line (19)
wait(2)
Call SetTextbox("/","BSEG-HKTID","",DT_F51_0330_BSEGHKTID,False)     ' - Line (19)
wait(2)
'call ClickButton("btn\[16\]",False)   ''process open items
call ClickButton("Process Open Items   \(Shift\+F4\)",False) 
'''Post with clearing  select open items'''

Call SetTextbox("Special G/L ind","RF05A-AGUMS","",DT_F51_0710_SPECIAL_GL_IND,False)     ' - Line (19)
wait(2)
Call SetTextbox("Account Type","RF05A-AGKOA","",DT_F51_0710_ACCOUNT_TYPE,False)     ' - Line (19)
wait(2)
Call SetTextbox("Company Code","RF05A-AGBUK","",DT_F51_0710_COMPANY_CODE,False)     ' - Line (19)
wait(2)
Call SetTextbox("Account","RF05A-AGKON","",DT_F51_0710_ACCOUNT,False)     ' - Line (19)
wait(2)
'call SelectRadioButton("RF05A-XPOS1","Document Number",false)  ''uncomment
' SelectRadioButton(radiobuttonName, radiobuttonAttachedtext, blnIsItPopup)
'Call PressEnter()
'Call PressEnter()
'Call SetTextbox("From","RF05A-SEL01","",DT_F51_0731_FROM,False)     ' - Line (19)
'Call SetTextbox("To","RF05A-SEL02","",DT_F51_0731_TO,True)     ' - Line (19)
'Call PressEnter()   '
'TakeScreenShot()
'call ClickButton("Process Open Items   \(Shift\+F4\)",False)
'call VerifyTableCellContent(1,"Document Number","SAPDF05XTC_6102",DT_F51_6102_CHECK_TEXT_OF_TABLECELL_DOCUMENT_NUMBER_0)  ''uncomment
'
'call ClickButton("Save   \(Ctrl\+S\)",false)  'uncomment

' VerifyStatusBar(Content)
'call VerifyStatusBar("posted")   'uncomment
'Call GetStatusBar("item1","DT_PO_NUMBER")  'uncomment

'''second tcode'''
Call SetTcode(DT_F51_0100_OKCD)     ' - Line (12)
Call PressEnter() 



call SelectRadioButton("X_AISEL","All items",false)
' SelectRadioButton(radiobuttonName, radiobuttonAttachedtext, blnIsItPopup)

Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_F51_1000_VENDOR_ACCOUNT,false)  


' SelectCheckbox(checkboxName, checkBoxIndex, OnOffStatus, blnIsItPopup)
call SelectCheckbox("X_SHBV",2,"ON",false)
call ClickButton("Dynamic selections   \(Shift\+F4\)",false)
    

Call SetTextbox("Document Number","%%DYN012-LOW","",DT_F51_1106_DOCUMENT_NUMBER,false)  
Call SetTextbox("Posting Date","%%DYN013-LOW","",DT_F51_1106_POSTING_DATE,false)  
call ClickButton("Execute   \(F8\)",false)
' VerifyGridCellContent(gridTitle, gridRowNumber, gridColumnName, gridIndex, expectedValue)
''''com
'call VerifyGridCellContent("",1,"Local Currency",10,DT_F51_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
'
'
''''third tcode'''
'Call SetTcode(DT_SAPTRANSACTIONCODE1)     ' - Line (12)
'Call PressEnter() 
'
'
'Call SetTextbox("Document Number","RF05L-BELNR","",DT_FB03_0100_DOCUMENT_NUMBER,false)  
'Call SetTextbox("Company Code","RF05L-BUKRS","",DT_FB03_0100_COMPANY_CODE,false)  
'Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FB03_0100_FISCAL_YEAR,false)  
'
'Call PressEnter() 
'
'' VerifyTextBoxContent(textboxAttachedText, textboxName, textboxIndex, expectedValue, blnIsItPopup)
''call VerifyTextBoxContent("Currency","BKPF-WAERS",9,DT_FB03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_RF05A_UBAZW,false)
'
'
'comm
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


