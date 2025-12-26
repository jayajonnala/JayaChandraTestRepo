'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Intracompany Store returns to DC - SW31_p4_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 21th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Intracompany Store returns to DC - SW31_p4_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DS\Retail\DT_Intracompany Store returns to DC - SW31_p4_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  



'Call WriteRunTimeDataToExcelGlobalSheet ("DT_ME21N_1211_CHECK_TEXT_OF_TABLECELL_REQMT_NO_0",(Cint(DT_ME21N_1211_CHECK_TEXT_OF_TABLECELL_REQMT_NO_0)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'----------------------Tcode MSR_TRC_I----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call SetTextbox("Delivery","SO_DLVNI-LOW","",DT_MSR_TRC_I_0301_DELIVERY,False)
Call PressEnter()  
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
wait(3)
Call TakeScreenShot()

Call VerifyStatusBarMessageType("S")

CAll ActivateItemGuiTree("","Send back to previous Plant;Stock Transport Order for Send Back to Previous Plant","Stock Transport Order for Send Back to Previous Plant")
Call TakeScreenShot()

Call GetTextboxValue("MEPO_TOPLINE-EBELN","","DT_MSR_TRC_I_1105_CHECK_TEXT_OF_MEPO_TOPLINEEBELN_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


'----------------------Tcode MIGO----------------------------

Call SetTcode(DT_MSR_TRC_I_0014_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MSR_TRC_I_0014_OKCD)
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-ACTION",DT_MSR_TRC_I_0010_GODYNPROACTION_OCC1)
Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_MSR_TRC_I_2010_GODYNPROMAT_DOC,False)
'Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,DT_MIGO_2010_FISCAL_YEAR,False)
Call TakeScreenShot()
Call PressEnter()
Call PressEnter()

Call TakeScreenShot()

Call SelectTab("TS_GOHEAD",DT_MSR_TRC_I_0100_DOC_INFO,False)
Call TakeScreenShot()

Call ClickButton("FI Documents",False)
Wait(1)
Call TakeScreenShot()

'
''Call GetGridContent("Documents in Accounting","","Doc. Number",1,"Object type text","Accounting document","DT_MSR_TRC_I_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT")
''''''Call GetTextboxValue("BKPF-BELNR","","DT_MSR_TRC_I_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER_OUTPUT",False)
''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''
''Call SelectRowGuiGrid("Documents in Accounting","","Doc. Number", DT_MSR_TRC_I_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER ,True)
''Call TakeScreenShot()
''Call ClickButton("Display Document   \(F2\)",True)
''Call TakeScreenShot()
Call VerifyTextBoxContent("Reference","BKPF-XBLNR","",DT_MSR_TRC_I_0750_CHECK_TEXT_OF_REFERENCE,False)


Call LogOff()
Call FinalStatus ()
