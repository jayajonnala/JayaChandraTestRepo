'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test_F.14-SM35-Create posting recurring entry  
'.................Author : TCS 
'................ Creation Date :
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_F.14-SM35-Create posting recurring"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\aprus\Desktop\DLL_P3\Data\TASE_DT_02-04-01-05-03-Create new assortment-BASA.xls"

If qtpParamExist("gstrInputExcelFilePathAndName") Then
	gstrInputExcelFilePathAndName= Parameter("gstrInputExcelFilePathAndName")	
End If

If qtpParamExist("gstrresultFolderPath") Then
	gstrresultFolderPath= Parameter("gstrresultFolderPath")	
End If

If qtpParamExist("datatable_row") Then
	datatable_row= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",datatable_row,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call CloseSessionsSAP()
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

'''''--------TransactionCode-F.14----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Settlement period","AUSFDATE-LOW","",ConvertDate(DT_F14_1000_SETTLEMENT_PERIOD),False)
Call SetTextbox("to","AUSFDATE-HIGH","",ConvertDate(DT_F14_1000_TO),False)
Call SetTextbox("User name","BDC_USER","",DT_F14_1000_USER_NAME,False)
Call SetTextbox("Company Code","BR_BUKRS-LOW","",DT_F14_1000_COMPANY_CODE,False)
Call SetTextbox("Document Number","BR_BELNR-LOW","",DT_F14_1000_DOCUMENT_NUMBER,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)  
Call VerifyStatusBar(DT_F14_1000_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot

''--------TransactionCode-SM35----------''''

Call SetTcode(DT_F14_1000_OKCD)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

Call SelectRowGuiTableByRow("SAPMSBDC_CCTC_APQI", 1, False)
Call ClickButton("Process session   \(F8\)",False)
Call TakeScreenShot
Call SelectRadioButton("D0300-ERROR", "Display errors only", True)
Call TakeScreenShot
Call ClickButton("Process   \(Enter\)",True)
Call TakeScreenShot
Call ClickButton("Go back to batch input session overview   \(Enter\)",True)
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

'''''''--------TransactionCode-/faglb03----------''''

Call SetTcode(DT_F14_1000_OKCD_OCC1)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC5)

Call SetTextbox("Account Number","RACCT-LOW","",DT_F14_1000_ACCOUNT_NUMBER,False)
Call SetTextbox("Company Code","RBUKRS-LOW","",DT_F14_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Fiscal Year","RYEAR","",DT_F14_1000_FISCAL_YEAR,False)
Call TakeScreenShot
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot
'''Call DoubleClickGuiGridCell("",1,5,"Period",False)
Call DoubleClickGuiGridCell("",1,DT_F14_0030_GRIDCELL_8_BALANCE,"Period",False)
Call TakeScreenShot
Call SelectMenuBar("Settings;Switch List")
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
'''''''Added 90 to 96 acc to TAO by KGARA on 19/07/2022'''''''
Call ClickButtonToolBar("&FIND",0)
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_F14_0841_SEARCH_TERM,True)
Call PressEnter()
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Add Filter Criterion \(F7\)",True)
Call TakeScreenShot
Call ClickButton("Define Filter Values",True)
''''Call PressEnter()
'''''''Added 99 to 103 acc to TAO by ARUD on 12/08/2022'''''''
'Call ClickButton("Find",True)
'Call SetTextbox("Find","GD_SEARCHSTR",0,DT_F14_0841_SEARCH_TERM,True)
'Call ClickButton("Continue   \(Enter\)",True)
'Call ClickButton("Show sel\. fields \(CTRL\+F3\)",True)
'Call ClickButton("Copy   \(Enter\)",True)

Call SetTextbox("Text","%%DYN001-LOW","",DT_F14_1105_TEXT,True)
Call PressEnter()
'Call ClickButton("Execute   \(Enter\)",True)  
Call TakeScreenShot
' VerifyGridCellContentbyName(gridName, gridRowNumber, gridColumnName, gridIndex, expectedValue)
Call VerifyGridCellContentbyName("shell",1,"Profit Center","",DT_F14_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContentbyName("shell",1,"Text","",DT_F14_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContentbyName("shell",1,"Amount in local currency","",DT_F14_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
' VerifyifGuiLabelExistsByRelativeid(Content, Relativeid)
'Call VerifyifGuiLabelExistsByRelativeid(DT_F14_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT,"wnd\[0\]/usr/lbl\[128,8\]")
'Call VerifyifGuiLabelExistsByRelativeid(DT_F14_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR,"wnd\[0\]/usr/lbl\[179,8\]")
'Call VerifyifGuiLabelExistsByRelativeid(DT_F14_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB,"wnd\[0\]/usr/lbl\[101,8\]")
'''Call VerifyifGuiLabelExists(DT_F14_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
'''Call VerifyifGuiLabelExists(DT_F14_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
''''Call VerifyifGuiLabelExists(DT_F14_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call TakeScreenShot

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

