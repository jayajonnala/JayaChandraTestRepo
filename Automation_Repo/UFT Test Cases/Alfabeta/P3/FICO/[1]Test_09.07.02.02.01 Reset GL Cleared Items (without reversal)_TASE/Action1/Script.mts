		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.02.02.01 Reset GL Cleared Items (without reversal)
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

gstrTestCaseName = "Test_09.07.02.02.01 Reset GL Cleared Items (without reversal)"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''''''--------TransactionCode-FBRA----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_ASSORTMENT_UNIQUE",Cint(DT_ASSORTMENT_UNIQUE)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Clearing Document","RF05R-AUGBL","",DT_FBRA_0100_CLEARING_DOCUMENT,False)
Call SetTextbox("Company Code","RF05R-BUKRS","",DT_FBRA_0100_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","RF05R-GJAHR","",Year(DT_FBRA_0100_FISCAL_YEAR),False)
Call TakeScreenShot
Call PressEnter

Call SelectMenuBar("Clearing;Reset cleared items")
Call ClickButtonIfExist("Only Reset", True)

Call VerifyStatusBar(Lcase(DT_FBRA_0100_CHECK_TEXT_OF_STATUSBAR))
Call TakeScreenShot

'''''----------------TransactionCode-FAGLL03-----------------------------------------------

Call SetTcode(DT_FBRA_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FBRA_1000_GL_ACCOUNT,False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FBRA_1000_COMPANY_CODE,False)
Call SelectRadioButton("X_AISEL","All Items", False)
Call TakeScreenShot

Call ClickButton("Custom Selections   \(Ctrl\+F1\)",False)
Call ActivateNodeGuiTree(0, "#4;#1")

Call ClickButton("%_%%DYN001_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FBRA_0100_DOCUMENT_NUMBER,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)

Call ClickBUtton("Back   \(F3\)",False)
Call ClickBUtton("Yes",True)
Wait 2
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot  

'Call VerifyifGuiLabelExistsByRelativeid(DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR,"wnd\[0\]/usr/lbl\[24,8\]")
'Call VerifyifGuiLabelExistsByRelativeid(DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART,"wnd\[0\]/usr/lbl\[9,8\]")
'Call VerifyifGuiLabelExistsByRelativeid(DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB,"wnd\[0\]/usr/lbl\[74,8\]")
'Call VerifyifGuiLabelExistsByRelativeid(ConvertDate(DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT),"wnd\[0\]/usr/lbl\[52,8\]")

Call VerifyGridCellContent("",1,"Document Number",0,DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("",1,"Document Type",0,DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("",1,"Amount in local currency",0,DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("",1,"Document Date",0,ConvertDate(DT_FBRA_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))

Call LogoFF
Call finalstatus
