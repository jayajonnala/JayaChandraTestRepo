
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_159_Update Account Completion Table (All Opcos)_TASE
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)

'----------------------------------------------------------------------------------------------------------------------------
gstrTestCaseName = "Test_06-02-01-02-Generate OBD"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''Login''''
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'-------------------------VL10G----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


Call FocusTextBox("Deliv\. Creation Date","ST_LEDAT-LOW",False)
Call SelectTab("TABSTRIP_ORDER_CRITERIA","Sales Orders",False)
'Call SetTextboxNoLabel("ST_VBELN-LOW","",DT_SALES_1,False)
Call SetTextboxNoLabel("ST_VBELN-LOW","",DT_VL10G_1030_PURCHASING_DOCUMENT,False)
'SetTextboxNoLabel is used for chnage in textbox name from Sales Document to SD Document.
Call TakeScreenShot()
Call SelectTab("TABSTRIP_ORDER_CRITERIA","Purchase Orders",False)
Call SetTextbox("Purchasing Document","ST_EBELN-LOW","",DT_VL10G_1030_PURCHASING_DOCUMENT,False)
Call SelectTab("TABSTRIP_ORDER_CRITERIA","Sales Orders",False)

Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)", False)
Call TakeScreenShot()

' SelectRowGuiGrid(gridTitle, gridIndex, columnName, refValue, blnIsItPopup)
'Call SelectRowGuiGrid("","","Originating Document",DT_SALES_1,False)
Call SelectRowGuiGrid("","","Originating Document",DT_VL10G_1030_PURCHASING_DOCUMENT,False)

Call TakeScreenShot()
Call ClickButton("Create Delivery in Background   \(Shift\+F7\)", False)
Call VerifyStatusBarMessageType("S")

Call TakeScreenShot()
call VerifyGridCellContent("",2,"Traffic Light","",DT_VL10G_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AMPEL)
call VerifyGridCellContent("",1,"Traffic Light","",DT_VL10G_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AMPEL)
call GetGridContentByRefColumn("","0","Traffic Light","@08\QGreen Light@","SD Document","DT_VL10G_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VBELN_OUTPUT")
'Call GetGridContent("","0","Sales Document","1","","","DT_VL10G_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VBELN_OUTPUT")
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()




