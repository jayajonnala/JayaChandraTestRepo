'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Impairment_p3_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 19th March
'.................Modified By :
'.................Modified Date/Details :
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
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Impairment_p3_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Fixed Assets Sales_p3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
''''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call TakeScreenShot()
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SelectRadioButton("XEINZEL","List assets",False)
Call SelectCheckbox("P_GRID",0,"ON",False)

Call SetTextbox("Company code","BUKRS-LOW","",DT_AR01_1000_COMPANY_CODE,False)
Call SetTextbox("Asset number","ANLAGE-LOW","",DT_AR01_1000_ASSET_NUMBER,False)
Call SetTextbox("Report date","BERDATUM","",Replace(DT_AR01_1000_REPORT_DATE,"/","."),False)
Call SetTextbox("Depreciation area","BEREICH1","",DT_AR01_1000_DEPRECIATION_AREA,False)
Call SetTextbox("Sort Variant","SRTVR","",DT_AR01_1000_SORT_VARIANT,False)
Call TakeScreenShot()
Call PressEnter()    
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

Call ClickButton("Create worklist   \(Ctrl\+Shift\+F8\)",False)
Call SetTextbox("WL name","G_WI_TEXT","",DT_AR01_0700_WL_NAME,True)
Call SelectRowGuiTableByRow("SAPLSTC1GENERIC_TABLE_16",5,True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)

Call SelectRadioButton("FIAA_IMPAIR_PARA-IMPRST","NetBkVal.",True)
Call SelectCheckbox("FIAA_IMPAIR_PARA-XAUFWN",0,"ON",True)

Call SetTextbox("Document Date","FIAA_IMPAIR_PARA-BLDAT","",Replace(DT_AR01_0500_DOCUMENT_DATE,"/","."),True)
'Call SetTextbox("Posting Date","FIAA_IMPAIR_PARA-BUDAT","",Replace(DT_AR01_0500_POSTING_DATE,"/","."),True)
'Call SetTextbox("Asset Val. Date","FIAA_IMPAIR_PARA-BZDAT","",Replace(DT_AR01_0500_ASSET_VAL_DATE,"/","."),True)
Call SetTextbox("Posting Date","FIAA_IMPAIR_PARA-BUDAT","",Replace(DT_AR01_0500_DOCUMENT_DATE,"/","."),True)
Call SetTextbox("Asset Val. Date","FIAA_IMPAIR_PARA-BZDAT","",Replace(DT_AR01_0500_DOCUMENT_DATE,"/","."),True)

Call SetTextbox("Prior-Year Acquis. Trans. Type","FIAA_IMPAIR_PARA-BWASL_ALT","",DT_AR01_0500_PRIORYEAR_ACQUIS_TRANS_TYPE,True)
Call SetTextbox("Current-Year Acquis. Trans. Type","FIAA_IMPAIR_PARA-BWASL_NEU","",DT_AR01_0500_CURRENTYEAR_ACQUIS_TRANS_TYPE,True)
Call SetTextbox("Retirement Simulation Trans. Type","FIAA_IMPAIR_PARA-BWASL_ABG","",DT_AR01_0500_RETIREMENT_SIMULATION_TRANS_TYPE,True)
Call SetTextbox("Impairment Amt","FIAA_IMPAIR_PARA-IMPBTR","",DT_AR01_0500_IMPAIRMENT_AMT,True)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call GetStatusBar("item1","DT_AR01_0500_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''VerifyStatusBar(DT_AR01_0500_CHECK_TEXT_OF_STATUSBAR)
'
''''''''''''''''''AR31'''''''''''''''''''
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_SAPTRANSACTIONCODE_OCC1) 
Call PressEnter()     ' 
Call TakeScreenShot()
Call SetTextbox("Worklist","PA_AI_ID","",DT_AR31_1000_WORKLIST,False)
Call SetTextbox("Report date","BERDATUM","",Replace(DT_AR01_1000_REPORT_DATE,"/","."),False)

Call SetTextbox("Depreciation area","BEREICH1","",DT_AR31_1000_DEPRECIATION_AREA,False)
Call SetTextbox("Sort Variant","SRTVR","",DT_AR31_1000_SORT_VARIANT,False)
Call TakeScreenShot()
Call PressEnter()   
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

Call VerifyGridCellContent("",1,"Planned Imp.",0,DT_AR31_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BTR4)
Call VerifyGridCellContent("",2,"Planned Imp.",0,DT_AR31_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BTR4)
Call VerifyGridCellContent("",3,"Planned Imp.",0,DT_AR31_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BTR4)
Call VerifyGridCellContent("",4,"Planned Imp.",0,DT_AR31_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BTR4)
Call VerifyGridCellContent("",5,"Planned Imp.",0,DT_AR31_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BTR4)
Call VerifyGridCellContent("",6,"Planned Imp.",0,DT_AR31_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BTR4)
Call ClickButton("Release worklist   \(Shift\+F10\)",False)
Call GetStatusBar("text","DT_AR31_0500_CHECK_TEXT_OF_STATUSBAR_OUTPUT")
'''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call VerifyStatusBar(DT_AR31_0500_CHECK_TEXT_OF_STATUSBAR_OUTPUT)


Call LogOff()
Call FinalStatus ()




