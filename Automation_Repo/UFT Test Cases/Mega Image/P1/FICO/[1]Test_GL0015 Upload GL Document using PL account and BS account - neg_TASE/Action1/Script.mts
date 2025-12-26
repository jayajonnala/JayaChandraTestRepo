
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_GL0015 Upload GL Document using PL account and BS account - neg
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


gstrTestCaseName = "Test_GL0015 neg"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()


''----------------------Tcode FAGLB03----------------------------

'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)


'Enter the details
Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER,False)    
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_FAGLB03_1000_COMPANY_CODE,False)    
Call SetTextbox("Fiscal Year","RYEAR","",DT_FAGLB03_1000_FISCAL_YEAR,False)    
Call TakeScreenShot()
Wait(2)
Call PressEnter()

'Click execute
Call ClickButton("Execute   \(F8\)",False) 

Call GetGridContent("",0,"Debit",MOnth(Date)+2,"Period",MOnth(Date),"DT_DEBIT_AMNT")
Call GetGridContent("",0,"Balance",MOnth(Date)+2,"Period",MOnth(Date),"DT_BALANCE_AMNT")


Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call ClickButtonIfExist("Back   \(F3\)",False)

'----------------------Tcode ZFIGL_UPLOAD_POST----------------------------

'Enter the Tcode
Call SetTcode(DT_FAGLB03_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

'Enter the Details
Call SetTextbox("File path name","P_FILE","",DT_FAGLB03_1000_FILE_PATH_NAME,False)   
Call SetTextbox("Session","P_SESS","",DT_FAGLB03_UPLOAD_POST_1000_SESSION,False) 
'Capture the screenshot
Call TakeScreenShot()

'Click execute
Call ClickButton("Execute   \(F8\)",False) 

'
'Set ODialog=Dialog("regexpwndtitle:=Microsoft Excel","text:=Microsoft Excel")
'ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").Highlight
''Capture the screenshot
'Call TakeScreenShot()
'ODialog.WinButton("regexpwndtitle:=&Yes","attached text:=There is a large amount of information on the Clipboard.*","text:=&Yes").Click
'Wait(2)

'Verify the details
Call VerifyGridCellContent("",1,"Account",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_NEWKO)
Call VerifyGridCellContent("",2,"Account",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_NEWKO)
Call VerifyGridCellContent("",3,"Account",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_NEWKO)
Call VerifyGridCellContent("",4,"Account",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_NEWKO)
Call VerifyGridCellContent("",5,"Account",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_NEWKO)
Call VerifyGridCellContent("",6,"Account",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_NEWKO)
Call VerifyGridCellContent("",7,"Account",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_6_NEWKO)
Call VerifyGridCellContent("",8,"Account",0,DT_FAGLB03_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_7_NEWKO)

'Click execute
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)


'Click execute
Call ClickButton("Yes",True) 
Wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)
'----------------------Tcode SM35----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_FAGLB03_0100_OKCD_OCC1) 
Call PressEnter()     ' 
Call CheckTCodeScreen(EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot()
Wait(2)

Call SetTextbox("Sess\.:","D0100-MAPN","",DT_FAGLB03_1005_SESS,False)    
Call SetTextbox("From:","D0100-VON","",DT_FAGLB03_1005_FROM,False)    
Call SetTextbox("Created by:","D0100-CREATOR","",DT_FAGLB03_1005_CREATED_BY,False)    
Call TakeScreenShot()
Wait(1)
Call PressEnter()
Wait(3)

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1,False)

'Click execute
Call ClickButton("Process session   \(F8\)",False) 
Wait(2)
Call SelectRadioButtonByIndexIfPopupExists("D0300-ERROR",0)
Call TakeScreenShot()
Wait(2)
'Click execute
Call ClickButton("Process   \(Enter\)",True) 
Wait(5)

Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)

Call SetTcode("/n") 
Call PressEnter()     ' 
Call CheckTCodeScreen("S000")

'Click on Session overview"
Call ClickButtonIfExist("Go back to batch input session overview   \(Enter\)",True) 
Wait(2)

Call SetTextbox("Sess\.:","D0100-MAPN","",DT_FAGLB03_1005_SESS,False)    
Call SetTextbox("From:","D0100-VON","",DT_FAGLB03_1005_FROM,False)    
Call SetTextbox("Created by:","D0100-CREATOR","",DT_FAGLB03_1005_CREATED_BY,False)    
Call TakeScreenShot()
Wait(1)
Call PressEnter()
Wait(3)


Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI",1,False)
Call SendKey("{F2}")
Set objWsh = CreateObject("WScript.Shell") 
objWsh.SendKeys "{TAB}"
objWsh.SendKeys "{F2}"
Set objWsh=nothing

'ScreenShot
Call TakeScreenShot()

'Navigate to Log Created Tab
Call SelectTab("TAB_DYNPRO"," Log created on "&Replace((DT_DATE),"/","."),False)
Wait(1)
Call TakeScreenShot()

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)



'----------------------Tcode FAGLB03----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Enter the Tcode
Call SetTcode(DT_FAGLB03_0100_OKCD_OCC2) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC9)


'Enter the details
Call SetTextbox("Account Number","RACCT-LOW","",DT_FAGLB03_1000_ACCOUNT_NUMBER_OCC1,False)    
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_FAGLB03_1000_COMPANY_CODE_OCC1,False)    
Call SetTextbox("Fiscal Year","RYEAR","",DT_FAGLB03_1000_FISCAL_YEAR_OCC1,False)    
Call TakeScreenShot()
Wait(2)


'Click execute
Call ClickButton("Execute   \(F8\)",False) 

''Verify the content
'Call VerifyGridCellContent("",6,"Debit",0," ")
'Call VerifyGridCellContent("",6,"Balance",0," ")


Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

'----------------------Tcode FAGLL03----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_FAGLB03_0100_OKCD_OCC3) 
Call PressEnter() 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC11)


Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FAGLB03_1000_GL_ACCOUNT,False)
Call SelectRadioButton("X_AISEL","All Items",False)

Call SetTextbox("Posting Date","SO_BUDAT-LOW","",DT_FAGLB03_1000_POSTING_DATE,False)
Call SetTextbox("to","SO_BUDAT-HIGH","",DT_FAGLB03_1000_TO,False)

'Click on Execute Button
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)


Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FAGLB03_1000_GL_ACCOUNT_OCC1,False)
Call SelectRadioButton("X_AISEL","All Items",False)

Call SetTextbox("Posting Date","SO_BUDAT-LOW","",DT_FAGLB03_1000_POSTING_DATE,False)
Call SetTextbox("to","SO_BUDAT-HIGH","",DT_FAGLB03_1000_TO,False)

'Click on Execute Button
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

'***********************************************************************************************************************************************
'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()
'***********************************************************************************************************************************************

