

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02LIV13_009_LIV_Reverse_in_Closed_WAC_date
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

gstrTestCaseName = "Test_02LIV13_009_LIV_Reverse_in_Closed_WAC_date"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)

SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
'''--------------------------------MIR4-----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE)  
Call PressEnter()   
call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Invoice Document No\.","RBKP-BELNR","",DT_MIR4_6150_INVOICE_DOCUMENT_NO,False) 
Call SetTextbox("Fiscal Year","RBKP-GJAHR","",Year(Date),False)
Call TakeScreenShot()
Call ClickButton("Display Document   \(F2\)",False)
Call TakeScreenShot()
Call GetspecialTextboxValue("INVFO-BUTXT",0,"DT_COMPANY_CODE_OUTPUT",False)
Call GetTextboxValue("INVFO-BUDAT",0,"DT_POSTING_DATE_OUTPUT",False)



''''''''--------------------------------SM34-----------------------------
''''
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MIR4_6000_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_MIR4_6000_OKCD)

Call SetTextbox("View cluster","VCLDIR-VCLNAME","",DT_MIR4_200_VIEW_CLUSTER,False) 
Call TakeScreenShot()
Call  ClickButton("Display",False)
Call TakeScreenShot()
Wait 30
' SelectRowGuiTable(tableName, columnName, cellVal, blnIsItPopup)
Call SelectRowGuiTable("SAPLZACPF_TAB_MNTTCTRL_ZACPF_V_LOG_LOGV","Logical Group",DT_LOGICAL_GROUP_NAME,False)

''Call SelectRowGuiTableByRow("SAPLZACPF_TAB_MNTTCTRL_ZACPF_V_LOG_LOGV",2,False)
Call  ActivateItemGuiTree(0,"#1;#2","Parameters")
Call FindRowNumber("SAPLZACPF_TAB_MNTTCTRL_ZACPF_V_LG_DV_PV","Default Value Parameter Name",DT_PARAMETER_DESCR,"DT_ROW_PARMETER_OUTPUT")
Call GetTableCellData("SAPLZACPF_TAB_MNTTCTRL_ZACPF_V_LG_DV_PV","Parameter Value",CINT(DT_ROW_PARMETER_OUTPUT),"","","DT_WAC_DATE_OUTPUT",False)

'''''''''--------------------------------SM34-----------------------------
'''''
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_MIR4_130_OKCD)  
Call PressEnter()   
call CheckTCodeScreen(DT_MIR4_130_OKCD)

Call  SetTextbox("Reversal Reason","UF05A-STGRD","",DT_MIR4_300_REVERSAL_REASON,False)
Call  SetTextbox("Posting Date","G_BUDAT","",ConvertDate(DT_MIR4_300_POSTING_DATE_OCC1),False)
Call TakeScreenShot()
Call SelectMenuBar("Invoice Document;Reverse")
wait 2
Call VerifyStatusBarMessageType("E")
Call VerifyStatusBar(DT_MIR4_300_CHECK_TEXT_OF_STATUSBAR)


Call LogOff()
Call FinalStatus ()



