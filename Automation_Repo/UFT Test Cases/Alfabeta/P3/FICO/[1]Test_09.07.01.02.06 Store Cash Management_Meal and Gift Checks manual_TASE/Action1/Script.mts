

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.02.06 Store Cash Management_Meal and Gift Checks manual
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

gstrTestCaseName = "Test_09.07.01.02.06 Store Cash Management_Meal and Gift Checks manual"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
Call StartExecution1(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'----------------------Login----------------------------
''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

''''''--------TransactionCode-F-21----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Document Date","BKPF-BLDAT", "", ConvertDate(DT_F21_0100_DOCUMENT_DATE), False)
Call SetTextbox("Type","BKPF-BLART","",DT_F21_0100_TYPE,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F21_0100_COMPANY_CODE,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F21_0100_PSTKY,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F21_0100_CURRENCYRATE,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F21_0100_REFERENCE,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F21_0100_DOCHEADER_TEXT,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F21_0100_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F21_0300_AMOUNT,False)
Call SetTextbox("Profit Ctrs","COBL-PRCTR","",DT_F21_1016_PROFIT_CTRS,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F21_0300_TEXT,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F21_0300_PSTKY,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F21_0300_ACCOUNT,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F21_0301_AMOUNT,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F21_0301_TEXT,False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call VerifyStatusBarMessageType("W")
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyStatusBarMessageType("S")
Call GetStatusBar("item1","DT_F21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Document "&DT_F21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet("DT_F21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_F21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Document Date","BKPF-BLDAT", "", ConvertDate(DT_F21_0100_DOCUMENT_DATE_OCC1), False)
Call SetTextbox("Type","BKPF-BLART","",DT_F21_0100_TYPE_OCC1,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_F21_0100_COMPANY_CODE_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F21_0100_PSTKY_OCC1,False)
Call SetTextbox("Currency/Rate","BKPF-WAERS","",DT_F21_0100_CURRENCYRATE_OCC1,False)
Call SetTextbox("Reference","BKPF-XBLNR","",DT_F21_0100_REFERENCE_OCC1,False)
Call SetTextbox("Doc\.Header Text","BKPF-BKTXT","",DT_F21_0100_DOCHEADER_TEXT_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F21_0100_ACCOUNT_OCC1,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F21_0300_AMOUNT_OCC1,False)
Call SetTextbox("Profit Ctrs","COBL-PRCTR","",DT_F21_1016_PROFIT_CTRS_OCC1,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F21_0300_TEXT_OCC1,False)
Call SetTextbox("PstKy","RF05A-NEWBS","",DT_F21_0300_PSTKY_OCC1,False)
Call SetTextbox("Account","RF05A-NEWKO","",DT_F21_0300_ACCOUNT_OCC1,False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR","",DT_F21_0301_AMOUNT_OCC1,False)
Call SetTextbox("Text","BSEG-SGTXT","",DT_F21_0301_TEXT_OCC1,False)

Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call VerifyStatusBarMessageType("W")

Call ClickButton("Display Document Overview   \(Shift\+F2\)",False)
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot
Call VerifyStatusBarMessageType("S")

Call GetStatusBar("item1","DT_F21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT")
Call VerifyStatusBar("Document "&DT_F21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT&" was posted in company code GR02")
Call WriteRunTimeDataToExcelGlobalSheet("DT_F21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1_OUTPUT",DT_F21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call ClickButton("Exit   \(Shift\+F3\)",False)
Call ClickButtonIfExist("Yes",True)
Call TakeScreenShot

'''''''--------TransactionCode-FAGLL03----------''''

Call SetTcode(DT_F21_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectRadioButton("X_AISEL", "All Items", False)
Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_F21_1000_GL_ACCOUNT,False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_F21_1000_COMPANY_CODE,False)

Call ClickButton("Custom Selections   \(Ctrl\+F1\)",False)
Call ActivateNodeGuiTree(0, "#3;#1")

Call ClickButton("%_%%DYN001_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F21_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_F21_3010_TABLECELL_SINGLE_VALUE_1,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)

Call ClickBUtton("Back   \(F3\)",False)
Call ClickBUtton("Yes",True)
Wait 2
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

'''Call VerifyifGuiLabelExistsByRelativeid(ConvertDate(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT), "wnd\[0\]/usr/lbl\[63,8\]")
'''Call VerifyifGuiLabelExistsByRelativeid(ConvertDate(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLDAT), "wnd\[0\]/usr/lbl\[63,9\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB,"wnd\[0\]/usr/lbl\[74,8\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB,"wnd\[0\]/usr/lbl\[74,9\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT,"wnd\[0\]/usr/lbl\[128,8\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT,"wnd\[0\]/usr/lbl\[128,9\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR,"wnd\[0\]/usr/lbl\[179,8\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR,"wnd\[0\]/usr/lbl\[179,9\]")
'
'Call VerifyGridCellContent("",1,"Document Date",0,DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT)
'Call VerifyGridCellContent("",2,"Document Date",0,DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLDAT)
'Call VerifyGridCellContent("",1,"Amount in local currency",0,DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
'Call VerifyGridCellContent("",2,"Amount in local currency",0,DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
'Call VerifyGridCellContent("",1,"Profit Center",0,DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
'Call VerifyGridCellContent("",2,"Profit Center",0,DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
'Call VerifyGridCellContent("",1,"Text",0,DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
'Call VerifyGridCellContent("",2,"Text",0,DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)
'Call VerifyGridCellContent("",4,"Amount in local currency",0,DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_DMSHB)
'Call TakeScreenShot
'
Call VerifyGridCellContent("",5,"Document Date",0,DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT)
Call VerifyGridCellContent("",6,"Document Date",0,DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLDAT)
Call VerifyGridCellContent("",5,"Amount in local currency",0,DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("",6,"Amount in local currency",0,DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
Call VerifyGridCellContent("",5,"Profit Center",0,DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("",6,"Profit Center",0,DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call VerifyGridCellContent("",5,"Text",0,DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("",6,"Text",0,DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)
Call VerifyGridCellContent("",4,"Amount in local currency",0,DT_F21_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_4_DMSHB)
Call TakeScreenShot

Call LogOff()
Call FinalStatus()


