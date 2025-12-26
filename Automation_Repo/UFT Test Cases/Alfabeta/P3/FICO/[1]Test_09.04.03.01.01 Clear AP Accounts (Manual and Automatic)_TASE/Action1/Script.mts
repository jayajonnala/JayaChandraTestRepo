

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.03.01.01 Clear AP Accounts (Manual and Automatic)
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

gstrTestCaseName = "Test_09.04.03.01.01 Clear AP Accounts (Manual and Automatic)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\jjonn\Desktop\TASEWork\Data\TASE_DT_09.04.03.01.01 Clear AP Accounts (Manual and Automatic).xls"
'strResultFolderPath = "C:\Users\jjonn\Desktop\TASEWork\Results"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'GetRowNo=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''


'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()
'
'''''''--------TransactionCode-F44----------''''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Currency","BKPF-WAERS","",DT_F44_0131_CURRENCY,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F44_0131_COMPANY_CODE,False)
Call SetTextbox("Account","RF05A-AGKON","",DT_F44_0131_ACCOUNT,False)
Call SelectRadioButton("RF05A-XPOS1", "Document Number", False)
Call TakeScreenShot
Call PressEnter()

Call SetTextbox("From","RF05A-SEL01",0,DT_F44_0731_FROM_OCC2,False)
Call SetTextbox("From","RF05A-SEL01",1,DT_F44_0731_FROM_OCC3,False)
Call TakeScreenShot
Call ClickButtonIfExist("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot

Call ClickButtonIfExist("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetStatusBar("item1","DT_F44_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_F44_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_F44_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_F44_0131_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTcode(DT_F44_0131_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectRadioButton("X_AISEL", "All items", False)
Call SetTextbox("Vendor account","KD_LIFNR-LOW",0,DT_F44_1000_VENDOR_ACCOUNT,False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Wait(3)
Call ClickButton("%_%%DYN012_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE", "Single value", "1", "", "", DT_F44_3010_TABLECELL_SINGLE_VALUE_0, False)
Call SetTableData("SAPLALDBSINGLE", "Single value", "2", "", "", DT_F44_3010_TABLECELL_SINGLE_VALUE_1, False)
Call SetTableData("SAPLALDBSINGLE", "Single value", "3", "", "", DT_F44_3010_TABLECELL_SINGLE_VALUE_2, False)

Call ClickButton("Copy   \(F8\)",True)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "ICO_AUGP", 0, DT_F44_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICO_AUGP)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_F44_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_F44_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call VerifyGridCellContent("", 3, "BELNR", 0, DT_F44_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_F44_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 2, "BLART", 0, DT_F44_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLART)
Call VerifyGridCellContent("", 3, "BLART", 0, DT_F44_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BLART)
Call VerifyGridCellContent("", 4, "BLART", 0, DT_F44_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_BLART)
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_F44_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 2, "DMSHB", 0, DT_F44_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
Call VerifyGridCellContent("", 3, "DMSHB", 0, DT_F44_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_DMSHB)
Call VerifyGridCellContent("", 4, "DMSHB", 0, DT_F44_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_3_DMSHB)
Call VerifyGridCellContent("", 1, "HWAER", 0, DT_F44_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 1, "AUGBL", 0, DT_F44_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AUGBL)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_F44_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("", 2, "SGTXT", 0, DT_F44_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)

Call LogOff
Call FinalStatus()
