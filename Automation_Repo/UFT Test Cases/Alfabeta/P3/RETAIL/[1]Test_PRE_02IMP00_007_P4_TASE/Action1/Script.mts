	

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_PRE_02IMP00_007_P4_TASE
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

gstrTestCaseName = "Test_PRE_02IMP00_007_P4_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\AB\RETAIL\DT_PRE_02IMP00_006_P4_Output.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

'//-----------------------------------MIGO -----------------------------------

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_MIRO_1000_COMPANY_CODE,False)    
Call ClickButton("Continue   \(Enter\)",fALSE)
Call TakeScreenShot()

Call SetCombobyKey("RM08M-VORGANG",1)
Call ClickButton("btn\[12\]",fALSE)
Call TakeScreenShot()


Call SetTextbox("Invoice date","INVFO-BLDAT","",Replace(DT_MIRO_0010_INVOICE_DATE,"/","."),False) 
Call SetTextbox("Reference","INVFO-XBLNR","",DT_MIRO_6211_RM08MEBELN,False)
Call TakeScreenShot()
Call SetComboByKey("RM08M-REFERENZBELEGTYP",DT_MIRO_6020_RM08MREFERENZBELEGTYP)
Call TakeScreenShot()
Call SetTextboxNoLabel("RM08M-EBELN","",DT_MIRO_6211_RM08MEBELN,False)
call SetComboByKey("RM08M-XWARE_BNK",DT_MIRO_6211_RM08MXWARE_BNK)
Call TakeScreenShot()
call SelectCheckbox("INVFO-XMWST",0,DT_MIRO_0010_CALCULATE_TAX,False)
Call TakeScreenShot()
call GetTextboxValue("RM08M-DIFFERENZ",1,"DT_Amount_output",False)
Call TakeScreenShot()
Call SetTextbox("Amount","INVFO-WRBTR","",replace(DT_Amount_output,"-",""),False)   
Call PressEnter()
Call TakeScreenShot()
'Call VerifyTextBoxContentIconName("RM08M-AMPEL","",DT_MIRO_6000_CHECK_ICONNAME_OF_TRANSACTION,False)

Call SelectTab("HEADER","Details",False)
Call TakeScreenShot()
'''Call SetComboByKey("Doc\. Type","RC")
Call SetComboByKey("Doc\. Type",DT_MIRO_DOCUMENT_TYPE_KEY)

Call SelectTab("HEADER","Basic Data",False)
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType("S")

Call GetStatusBar("item1","DT_MIRO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcel("DT_MIRO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT", DT_MIRO_6000_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

call VerifyStatusBar(DT_MIRO_6000_CHECK_TEXT_OF_STATUSBAR)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()
