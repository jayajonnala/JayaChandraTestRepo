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

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'reload DS to update dates and calculations
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Intracompany Store returns to DC - SW31_p2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 20th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Intracompany Store returns to DC - SW31_p2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="S:\TASETestData\DS\Retail\DT_Test_Intracompany Store returns to DC - SW31_p2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'Increment the parameter
'Call WriteRunTimeDataToExcelGlobalSheet ("DT_ME21N_1211_CHECK_TEXT_OF_TABLECELL_REQMT_NO_0",(Cint(DT_ME21N_1211_CHECK_TEXT_OF_TABLECELL_REQMT_NO_0)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode VL32N----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Inbound delivery","LIKP-VBELN","",DT_VL32N_4104_INBOUND_DELIVERY,False) 
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post Goods Receipt   \(Shift\+F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call GetStatusBar("item2","DT_VL32N_4104_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPPUT")
'reload data sheet
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'verify statusbar with datasheet feed
Call VerifyStatusBarMessageType("S")
Call VerifyStatusBar(DT_VL32N_4104_CHECK_TEXT_OF_STATUSBAR)

Call PressEnter() 
Call PressEnter() 
'
'''----------------------Tcode MIGO----------------------------
'Enter the Tcode
Call SetTcode(DT_VL32N_1000_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_VL32N_1000_OKCD)
'Capture the screenshot
Call TakeScreenShot()

Call SetComboByKey("GODYNPRO-ACTION",DT_VL32N_0010_GODYNPROACTION_OCC1)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_VL32N_2010_GODYNPROMAT_DOC,False)
'Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,DT_MIGO_6150_FISCAL_YEAR,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()

Call VerifyTableCellContent(1,"Movement type","SAPLMIGOTV_GOITEM",DT_VL32N_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_0)
Call VerifyTableCellContent(1,"Direction","SAPLMIGOTV_GOITEM",DT_VL32N_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_0)

Call SelectTab("TS_GOHEAD","General",False)
Call VerifyTextBoxContent("Delivery Note","GOHEAD-LFSNR","",DT_VL32N_0110_CHECK_TEXT_OF_DELIVERY_NOTE,False)


Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_VL32N_2010_GODYNPROMAT_DOC_OCC1,False)
'Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,DT_MIGO_6150_FISCAL_YEAR,False)
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()

Call SelectTab("TS_GOHEAD","General",False)
Call VerifyTextBoxContent("Delivery Note","GOHEAD-LFSNR","",DT_VL32N_0110_CHECK_TEXT_OF_DELIVERY_NOTE_OCC1,False)

Call SelectTab("TS_GOHEAD","Doc. info",False)
Call VerifyTableCellContent(1,"Movement type","SAPLMIGOTV_GOITEM",DT_VL32N_0200_CHECK_TEXT_OF_TABLECELL_MOVEMENT_TYPE_0_OCC1)
Call VerifyTableCellContent(1,"Direction","SAPLMIGOTV_GOITEM",DT_VL32N_0200_CHECK_TEXT_OF_TABLECELL_DIRECTION_0_OCC1)

Call SelectTab("TS_GOHEAD",DT_VL32N_0100_DOC_INFO,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS_GOHEAD",DT_VL32N_0100_GENERAL,False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TS_GOITEM",DT_VL32N_0305_MESSAGES,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Display outputs",False)
Wait(1)
'Capture the screenshot
Call TakeScreenShot()
Call VerifyTableCellContent(1,"Output Type","SAPDV70ATC_NAST3",DT_VL32N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)

Call ClickButton("Back   \(F3\)",False)
'
'''----------------------Tcode VL03N----------------------------
'Enter the Tcode
Call SetTcode(DT_VL32N_0001_OKCD) 
Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()

'Enter data
Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VL32N_4004_OUTBOUND_DELIVERY,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Header Details   \(F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TAXI_TABSTRIP_HEAD",DT_VL32N_2000_ADMINISTRATION,False)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Ext\. Delivery","LIKP-LIFEX","",DT_VL32N_2110_CHECK_TEXT_OF_EXT_DELIVERY,False)
wait(2)

'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

