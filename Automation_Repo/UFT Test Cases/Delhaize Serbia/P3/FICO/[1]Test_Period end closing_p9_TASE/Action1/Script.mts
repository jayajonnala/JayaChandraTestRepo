
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Period end closing_p9
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
'.................Test Script Name :Test_Period end closing_p9
'.................Author : TCS 	   :Raushan
'................ Creation Date    :16th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Period end closing_p9"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.7.1.2.1. Create Low Level Header ZRPC.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode KSU3----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

'Enter the details
Call SetTextbox("Cycle","RKAL1-KSCYC","",DT_KSU3_0103_CYCLE,FALSE)
Call SetTextbox("Start Date","T811C-SDATE","",ConvertDate(DT_KSU3_0103_START_DATE),FALSE)
Call TakeScreenShot()
Call PressEnter()     ' 

Call TakeScreenShot()

''Click on First segment
Call ClickButton("First segment   \(F6\)",False)
Call TakeScreenShot()

'Navigate to Senders/Receivers Tab
Call SelectTab("SEG_TABSTRIP","Senders/Receivers",False)
Wait(1)
Call TakeScreenShot()

Call  FocusTextBoxWithIndex("Cost Center","KGALK-SETID",0,False)
Call DoubleClick()
Call TakeScreenShot()

Call SetFocusGuiLabel("4","","",True)
Call SendKey("{F2}")
Call TakeScreenShot()

Call SetFocusGuiLabel("S01","","",True)
Call ClickBUtton("Set focus   \(F5\)",True)
Call TakeScreenShot()

Call SetFocusGuiLabel("S01O","","",True)
Call ClickBUtton("Set focus   \(F5\)",True)
Call TakeScreenShot()

Call SetFocusGuiLabel("S01OS","","",True)
Call ClickBUtton("Set focus   \(F5\)",True)
Call TakeScreenShot()

Call SetFocusGuiLabel("S01OS1","","",True)
Call ClickBUtton("Set focus   \(F5\)",True)
Call TakeScreenShot()

Call SetFocusGuiLabel("Integrated","","",True)

Call SendKey("^f")
Call TakeScreenShot()

Call SetTextbox("Find","RSYSF-STRING","",DT_KSU3_0800_FIND,True)
Call SelectCheckbox("SCAN_STRING-START",0,"OFF",True)
Call TakeScreenShot()

'Click on Find
Call ClickButton("Find   \(Enter\)",True)

Call VerifyifGuiLabelExists(DT_KSU3_0800_FIND)

'Click on Cancel
Call ClickButton("Cancel   \(F12\)",True)

'Click on Cancel
Call ClickButton("Cancel   \(F12\)",True)

'Click on Cancel
Call ClickButton("Cancel   \(F12\)",True)


Call FocusTextBoxWithIndex("Cost Element","KGALK-SETID",0,False)
Call DoubleClick()
Call TakeScreenShot()

Call SetFocusGuiLabel("S2440","","",True)
Call ClickBUtton("Set focus   \(F5\)",True)
Call TakeScreenShot()

Call SetFocusGuiLabel("Corporate sponsorship","","",True)

Call SendKey("^f")
Call TakeScreenShot()


Call SetTextbox("Find","RSYSF-STRING","",DT_KSU3_0800_FIND_OCC1,True)
Call SelectCheckbox("SCAN_STRING-START",0,"OFF",True)
Call TakeScreenShot()

'Click on Find
Call ClickButton("Find   \(Enter\)",True)

Call VerifyifGuiLabelExists(DT_KSU3_0120_CHECK_TEXT_OF_NO_NAME)

'Click on Cancel
Call ClickButton("Cancel   \(F12\)",True)

'Click on Cancel
Call ClickButton("Cancel   \(F12\)",True)

'Click on Cancel
Call ClickButton("Cancel   \(F12\)",True)


Call FocusTextBoxWithIndex("Cost Center","KGALK-SETID",1,False)
Call DoubleClick()
Call TakeScreenShot()

Call SetFocusGuiLabel("RS Format CC Marketing-Alokacija procena","","",True)

Call SendKey("^f")
Call TakeScreenShot()


Call SetTextbox("Find","RSYSF-STRING","",DT_KSU3_0800_FIND_OCC2,True)
Call SelectCheckbox("SCAN_STRING-START",0,"OFF",True)
Call TakeScreenShot()

'Click on Find
Call ClickButton("Find   \(Enter\)",True)

Call VerifyifGuiLabelExists(DT_KSU3_0120_CHECK_TEXT_OF_SS000400)

'Click on Cancel
Call ClickButton("Cancel   \(F12\)",True)

'Click on Cancel
Call ClickButton("Cancel   \(F12\)",True)

'Click on Cancel
Call ClickButton("Cancel   \(F12\)",True)



'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

