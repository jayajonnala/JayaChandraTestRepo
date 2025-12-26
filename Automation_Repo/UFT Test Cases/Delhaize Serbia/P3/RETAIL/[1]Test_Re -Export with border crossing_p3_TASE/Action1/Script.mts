
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_2.1.1.5.1. Reclassify an Article (merchandise category)
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
'.................Test Script Name :Test_Re -Export with border crossing_p3
'.................Author : TCS 	   :Raushan
'................ Creation Date    :20th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Re -Export with border crossing_p3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Re -Export with border crossing_p3.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode MIGO----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()


Call SetCombo("GODYNPRO-ACTION","Display")
''Call SetCombo("GODYNPRO-REFDOC","Article Document")
Call SetComboByKey("GODYNPRO-REFDOC","R02")
Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_MIGO_2010_GODYNPROMAT_DOC,False)
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,DT_MIGO_2010_GODYNPRODOC_YEAR,False)
Call PressEnter() 
Wait(2)
Call TakeScreenShot()


'Verify the Movement Type
Call VerifyTableCellContent(1,"Movement Type","SAPLMIGOTV_GOITEM",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_0)

'Navigate to the Document Details
Call SelectTab("TS_GOHEAD","Doc. info",False)
Wait(1)
Call TakeScreenShot()
Call ClickButton("FI Documents",False)
Wait(5)
Call TakeScreenShot()

'Get Document No
Call GetGridContentByTitle("Documents in Accounting",0,"Document Number",1,"DT_ARTICLE_DOC_NUM")

Call DoubleClickGuiGridCell("Documents in Accounting",0,1,"Document Number",True)


Call VerifyGridCellContent("",1,"Profit Center",0,DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)

''----------------------Tcode VF01----------------------------

Call SetTcode(DT_MIGO_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MIGO_0100_OKCD)

'
Call SetTableData("SAPMV60ATCTRL_ERF_FAKT","Document","1","","",DT_MIGO_0102_TABLECELL_DOCUMENT_0,False)
Call TakeScreenShot()
Call PressEnter()
Wait(2)
 
'Click on Save
Call ClickButton("Save   \(Ctrl\+S\)",False) 
Wait(2)

'Verify message Type
Call VerifyStatusBarMessageType("S")

'Validate If Document No is generated
Call GetStatusBar("item1","DT_DOC_NUMBER_OUTPUT")
VerifyStatusBar("Document "&DT_DOC_NUMBER_OUTPUT&" has been saved")



'----------------------Tcode VF03----------------------------

Call SetTcode(DT_MIGO_0100_OKCD_OCC1) 
Call PressEnter() 
Call CheckTCodeScreen(DT_MIGO_0100_OKCD_OCC1)

'Display the Invoice Details
Call SetTextbox("Billing document","VBRK-VBELN","",DT_DOC_NUMBER_OUTPUT,False)
Call TakeScreenShot()



'Navigate to Billing document;Issue Output To
Call SelectMenuBar("Billing document;Issue Output To")
Call TakeScreenShot()


'Verify the Message type
Call VerifyTableCellContent(1,"Message type","SAPLVMSGTABCONTROL",DT_MIGO_0200_CHECK_TEXT_OF_TABLECELL_MESSAGE_TYPE_0)


'Select Row
Call SelectRowGuiTable("SAPLVMSGTABCONTROL","Message type","ZGRS",True)
Call TakeScreenShot()

'Click on Print
Call ClickButton("Print preview   \(Ctrl\+Shift\+F1\)",True)
Wait(1)
Call TakeScreenShot()

'Click on Back
Call ClickButton("Back   \(F3\)",False)
Wait(2)


'Click on Cancel
Call ClickButtonIfExist("Cancel   \(F12\)",True)
Wait(2)

'Click on Accounting overview
Call ClickButton("Accounting overview   \(F6\)",False)
Wait(1)
Call TakeScreenShot()

Call GetGridContentByTitle("Documents in Accounting",0,"Document Number",1,"DT_ACCOUNTING_DOC_NO")
Call TakeScreenShot()
Call  DoubleClickGuiGridCell("Documents in Accounting",0,1,"Document Number",True)


'Verify the profit center
Call VerifyGridCellContent("",2,"Profit Center",0,DT_MIGO_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR_OCC1)

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************


