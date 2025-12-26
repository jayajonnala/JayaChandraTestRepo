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
'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_P2P_01_01_072-Physical inventory Non VBS -SAP only RW04_P1_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 9th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_P2P_01_01_072-Physical inventory_P1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_P2P_01_01_072-Physical inventory  Non VBS -SAP only  RW04_P1_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''Login'''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'''Increment the parameter/reload
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''
''''----------------------Tcode MI31----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Article","R_MATNR-LOW","",DT_MI31_1000_ARTICLE,False)
'index need to be checked
SAPGuiSession("transaction:=MI31").SAPGuiWindow("transaction:=MI31").SAPGuiButton("tooltip:=Multiple selection","index:=3").Click
Wait(2)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_MI31_3010_TABLECELL_SINGLE_VALUE_1,True)'select
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",True)
Wait(1)

Call SetTextbox("Site","R_WERKS-LOW","",DT_MI31_1000_SITE,False)
Call SetTextbox("Storage Location","R_LGORT-LOW","",DT_MI31_1000_STORAGE_LOCATION,False)

SAPGuiSession("transaction:=MI31").SAPGuiWindow("transaction:=MI31").SAPGuiButton("tooltip:=Multiple selection","index:=4").Click
Wait(2)
Call SelectTab("TAB_STRIP",DT_MI31_3000_EXCLUDE_SINGLE_VALUES,True)
Call SetTableDataNoRef("SAPLALDBSINGLE_E","Single value",1,DT_MI31_3030_TABLECELL_SINGLE_VALUE_0,True)'exclude
Call SetTableDataNoRef("SAPLALDBSINGLE_E","Single value",2,DT_MI31_3030_TABLECELL_SINGLE_VALUE_1,True)'exclude
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Copy   \(F8\)",True)
Wait(1)

Call SetTextbox("Name of Session","MAPPE","",DT_SET_SESSION_NAME,False)

Call SelectCheckbox("XKEEP","1",DT_MI31_1000_HOLD_PROCESSED_SESSIONS,False)
Call SelectCheckbox("XBUFI","1",DT_MI31_1000_FREEZE_BOOK_INVBAL,False)

Call SetTextbox("Physical Inventory Number","INVNU","",DT_MI31_1000_PHYSICAL_INVENTORY_NUMBER,False)
Call SetTextbox("Phys\. Inventory Ref\.","XBLNI","",DT_MI31_1000_PHYS_INVENTORY_REF,False)

Call ClickButton("Acc\. to Stck",False)
Wait(2)
Call SelectCheckbox("XIMAT","1",DT_MI31_1000_INCL_ARTLS_SUBJ_TO_PHYS_INV,False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function VerifyifGuiLabelWithIndexExists(Content,Index)
	
 If Not (Environment.Value("blnFatalError") or Content= DS_SKIP) Then
	If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : VerifyifGuiLabelWithIndexExists"
	
	strStepName = "Verify if Gui Label exists with index"

    If Content <>"" Then
     set   objLabel = SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiLabel("content:="&Content,"guicomponenttype:=30","index:="&Index)
               If objLabel.Exist Then
					Call ReporterFunction(strLibraryFileName,"VerifyifGuiLabelWithIndexExists","2",Content,"Gui Label with value "&Content &" index : "&Index &" exists in the screen")	
					strStatus = "PASS"
					strMsg = "Gui Label with value "&Content&"index : "&Index &" exists in the screen"	
					blnCaptureFlag = True
			              	If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
		            	  		ImagePath=CaptureScreenshot(strStepName,objLabel,False,False,False)
                              End If
					
					Else
					Call ReporterFunction(strLibraryFileName,"VerifyifGuiLabelWithIndexExists","1"," Status bar Content","Gui Label with value "&Content &"index : "&Index &" doesn't exist  in the screen")	
					strStatus = "FAIL"
					blndefectFlag =True
					strMsg = "Gui Label with value "&Content&" doesn't exist  in the screen"
					blnObjectError=True
				End If
       	Else
    	Call ReporterFunction(strLibraryFileName,"VerifyifGuiLabelWithIndexExists","1","Gui Label","Function Parameter Not Passed Properly. Check the --VerifyifGuiLabelWithIndexExists-- Function Call")
			strStatus = "FAIL" 
			strMsg = "Function Parameter Not Passed Properly. Check the --VerifyifGuiLabelWithIndexExists-- Function Call-"
	End if


If strStatus = "FAIL"  Then
		VerifyifGuiLabelWithIndexExists = strMsg
		blnMainFailFlag = True
		ImagePath=CaptureScreenshot(strStepName,objLabel,False,False,False)
    
	Else
		VerifyifGuiLabelWithIndexExists = True
	End If
	If blnDefault_eSwiftReporting Then  
		Call UpdateResultHtml(strStepName,Content,strMsg,strStatus,"")
	End If

End If
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Call VerifyifGuiLabelWithIndexExists(DT_MI31_0120_CHECK_TEXT_OF_RW22,"0")
Call VerifyifGuiLabelWithIndexExists(DT_MI31_0120_CHECK_TEXT_OF_NO_NAME,"0")
Call VerifyifGuiLabelExists(DT_MI31_0120_CHECK_TEXT_OF_NO_NAME_OCC1)

Call VerifyifGuiLabelWithIndexExists(DT_MI31_0120_CHECK_TEXT_OF_RW22_OCC1,"1")
Call VerifyifGuiLabelWithIndexExists(DT_MI31_0120_CHECK_TEXT_OF_NO_NAME_OCC2,"1")
Call VerifyifGuiLabelExists(DT_MI31_0120_CHECK_TEXT_OF_NO_NAME_OCC3)
Wait(2)

Call ClickButton("Process Session   \(Shift\+F2\)",False)
''Capture the screenshot
Call TakeScreenShot()
'______________________________________________________________SM35___________________________________________________________
'select row
Call SelectRowGuiTable("SAPMSBDC_CCTC_APQI","Session Name",DT_SET_SESSION_NAME,False)
Call ClickButton("Process session   \(F8\)",False)

'Capture the screenshot
Call TakeScreenShot()
Call SelectRadioButtonIfPopupExists("D0300-ERROR",DT_MI31_0300_DISPLAY_ERRORS_ONLY)
Call ClickButton("Process   \(Enter\)",True)
'Capture the screenshot
call TakeScreenShot()

Call VerifyTextBoxContent("Information Message","MESSTXT1","",LCase(DT_MI31_0010_CHECK_TEXT_OF_MESSTXT1),True)
Call ClickButton("Go back to batch input session overview   \(Enter\)",True)
'Capture the screenshot
call TakeScreenShot()

Call SelectTab("D1000_TABSTRIP",DT_MI31_1000_PROCESSED,False)
Call SelectRowGuiTable("SAPMSBDC_CCTC_APQI","Session Name",DT_SET_SESSION_NAME,False)
'Capture the screenshot
call TakeScreenShot()

Call ClickButton("Log   \(F7\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call SelectRowGuiTableByRow("RSBDC_PROTOCOLTC_PROTOCOL",1,False)

Call ClickButton("Analyze session and logs   \(Shift\+F6\)",False)
'Capture the screenshot
call TakeScreenShot()

'Navigate to Log Created Tab
Call SelectTab("TAB_DYNPRO"," Log created on "&Replace((DT_MI31_0100_LOG_CREATED_ON),"/","."),False)

'fetch oupput mesage
Call GetTableCellData("RSBDC_ANALYSETC_PROTOCOL","Message",2,"Index","1","DT_MI31_0400_GET_TEXT_OF_TABLECELL_MESSAGE_1_OUTPUT",False)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
''''----------------------Tcode MI03----------------------------
'Enter the Tcode
Call SetTcode(DT_MI31_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_MI31_0100_OKCD)
'Capture the screenshot
Call TakeScreenShot()
'
Call SetTextbox("Phys\. Inventory Doc\.","RM07I-IBLNR","",DT_MI31_0701_PHYS_INVENTORY_DOC,False)
Call SetTextbox("Fiscal Year","RM07I-GJAHR","",DT_MI31_0701_FISCAL_YEAR,False)

Call PressEnter()   
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Stor\. Loc\.","IKPF-LGORT","",DT_CHECK_STORAGE_LOCATION,False)
Call VerifyTextBoxContent("Site","IKPF-WERKS","",DT_CHECK_SITE,False)
Call VerifyTextBoxContent("Article","ISEG-MATNR","0",DT_CHECK_ARTICLE_1,False)
Call VerifyTextBoxContent("Article","ISEG-MATNR","1",DT_CHECK_ARTICLE_2,False)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()

