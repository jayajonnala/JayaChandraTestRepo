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
'.................Test Script Name : Test_P2P_01_01_071-Physical inventory DC WMS RW22_P2_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 8th June
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_P2P_01_01_071- RW22_P2_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_P2P_01_01_071-Physical inventory DC WMS RW22_P2_TASE.xls"
''gstrInputExcelFilePathAndName="S:\TASETestData\P3\MI\RETAIL\DT_P2P_01_01_071-Physical inventory DC WMS RW22_P2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''Login'''

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''
''''----------------------Tcode MI24----------------------------
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Article","IM_MATNR-LOW","",DT_MI24_1000_ARTICLE,False)
Call SetTextbox("Site","IM_WERKS-LOW","",DT_MI24_1000_SITE,False)
Call SetTextbox("Layout","P_VARI","",DT_MI24_1000_LAYOUT,False)
Call SelectCheckbox("IM_SELKB","1",DT_MI24_1000_PHYS_INVENTORY_DOCUMENTS,False)
Call SelectCheckbox("IM_SELKP","1",DT_MI24_1000_PHYS_INVENTORY_ITEMS,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)

Call SelectCheckbox("IX-SELP1","1",DT_MI24_0101_NOT_YET_COUNTED,True)
Call SelectCheckbox("IX-SELP3","1",DT_MI24_0101_POSTED,True)
Call SelectCheckbox("IX-SELP2","1",DT_MI24_0101_ONLY_CNTD,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)
Call ClickButton("Execute   \(F8\)",False)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Capture the screenshot
Call TakeScreenShot()

Call SelectCheckBoxByGuiLabel(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC12,"17","1",DT_MI24_0120_NO_NAME_OCC2)

Call VerifyifGuiLabelExists(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME)
Call VerifyifGuiLabelExists(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC1)

Call VerifyifGuiLabelExists(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC2)
Call VerifyifGuiLabelExists(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC3)

Call VerifyifGuiLabelExists(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC4)
Call VerifyifGuiLabelExists(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC5)

Call VerifyifGuiLabelExists(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC8)
Call VerifyifGuiLabelExists(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC9)

SAPGuiSession("transaction:=MI24").SAPGuiWindow("transaction:=MI24").HorizontalScrollbarPosition 157
Call VerifyifGuiLabelWithIndexExists(Lcase(DT_MI24_0120_CHECK_TEXT_OF_NOT_YET_COUNTED_OCC4),"3")
Call VerifyifGuiLabelWithIndexExists(Lcase(DT_MI24_0120_CHECK_TEXT_OF_NOT_YET_COUNTED_OCC2),"2")
Call VerifyifGuiLabelWithIndexExists(Lcase(DT_MI24_0120_CHECK_TEXT_OF_NOT_YET_COUNTED_OCC1),"1")
Call VerifyifGuiLabelWithIndexExists(Lcase(DT_MI24_0120_CHECK_TEXT_OF_NOT_YET_COUNTED),"0")

Call SelectCheckBoxByGuiLabel(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC20,"17","1",DT_MI24_0120_NO_NAME_OCC3)
Call SelectCheckBoxByGuiLabel(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC12,"17","1",DT_MI24_0120_NO_NAME_OCC2)

'Capture the screenshot
Call TakeScreenShot()
Wait(2)

Call ClickButton("Enter Count   \(Ctrl\+Shift\+F11\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyTextBoxContent("Site","IKPF-WERKS","",DT_MI24_0731_CHECK_TEXT_OF_SITE,False)
Call VerifyTextBoxContent("Stor\. Loc\.","IKPF-LGORT","",DT_MI24_0731_CHECK_TEXT_OF_STOR_LOC,False)
Call SetTextboxNoLabel("ISEG-ERFMG","",DT_MI24_0731_ISEGERFMG,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Post   \(Ctrl\+S\)",False)

'Capture the screenshot
Call TakeScreenShot()
Wait(2)

Call VerifyTextBoxContent("Site","IKPF-WERKS","",DT_MI24_0731_CHECK_TEXT_OF_SITE_OCC1,False)
Call VerifyTextBoxContent("Stor\. Loc\.","IKPF-LGORT","",DT_MI24_0731_CHECK_TEXT_OF_STOR_LOC_OCC1,False)
Call SetTextboxNoLabel("ISEG-ERFMG","",DT_MI24_0731_ISEGERFMG_OCC1,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Post   \(Ctrl\+S\)",False)

'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Refresh   \(Ctrl\+Shift\+F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call SelectCheckBoxByGuiLabel(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC14,"17","1",DT_MI24_0120_NO_NAME_OCC5)
Call SelectCheckBoxByGuiLabel(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC16,"17","1",DT_MI24_0120_NO_NAME_OCC6)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Enter Count   \(Ctrl\+Shift\+F11\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call SelectCheckbox("ISEG-XNULL","1",DT_MI24_0731_CHECK_SELECTED_OF_ISEGXNULL,False)
Call VerifyTextBoxContent("Site","IKPF-WERKS","",DT_MI24_0731_CHECK_TEXT_OF_SITE_OCC2,False)
Call VerifyTextBoxContent("Stor\. Loc\.","IKPF-LGORT","",DT_MI24_0731_CHECK_TEXT_OF_STOR_LOC_OCC2,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Post   \(Ctrl\+S\)",False)

'Capture the screenshot
Call TakeScreenShot()
Wait(2)

Call SelectCheckbox("ISEG-XNULL","1",DT_MI24_0731_CHECK_SELECTED_OF_ISEGXNULL_OCC1,False)
Call VerifyTextBoxContent("Site","IKPF-WERKS","",DT_MI24_0731_CHECK_TEXT_OF_SITE_OCC3,False)
Call VerifyTextBoxContent("Stor\. Loc\.","IKPF-LGORT","",DT_MI24_0731_CHECK_TEXT_OF_STOR_LOC_OCC3,False)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Post   \(Ctrl\+S\)",False)

'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Refresh   \(Ctrl\+Shift\+F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyifGuiLabelExists(Lcase(DT_MI24_0120_CHECK_TEXT_OF_LIST_DOES_NOT_CONTAIN_ANY_DATA))
'
Call ClickButton("Back   \(F3\)",False)
'Capture the screenshot
Call TakeScreenShot()
'
Call SetTextbox("Article","IM_MATNR-LOW","",DT_MI24_1000_ARTICLE_OCC1,False)
Call SetTextbox("Site","IM_WERKS-LOW","",DT_MI24_1000_SITE_OCC1,False)
Call SetTextbox("Layout","P_VARI","",DT_MI24_1000_LAYOUT,False)
Call SelectCheckbox("IM_SELKB","1",DT_MI24_1000_PHYS_INVENTORY_ITEMS_OCC1,False)
Call SelectCheckbox("IM_SELKP","1",DT_MI24_1000_PHYS_INVENTORY_DOCUMENTS_OCC1,False)
Call SetTextbox("Physical Inventory Document","IM_IBLNR-LOW","","",False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)

Call SelectCheckbox("IX-SELP1","1","OFF",True)
Call SelectCheckbox("IX-SELP3","1",DT_MI24_0101_POSTED_OCC1,True)
Call SelectCheckbox("IX-SELP2","1",DT_MI24_0101_ONLY_CNTD_OCC1,True)
'Capture the screenshot
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",True)
Call ClickButton("Execute   \(F8\)",False)

'Capture the screenshot
Call TakeScreenShot()

Call VerifyifGuiLabelExists(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC12)
Call VerifyifGuiLabelExists(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC13)

Call VerifyifGuiLabelExists(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC14)
Call VerifyifGuiLabelExists(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC15)

Call VerifyifGuiLabelExists(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC16)
Call VerifyifGuiLabelExists(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC17)

Call VerifyifGuiLabelExists(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC20)
Call VerifyifGuiLabelExists(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC21)

SAPGuiSession("transaction:=MI24").SAPGuiWindow("transaction:=MI24").HorizontalScrollbarPosition 157
Call VerifyifGuiLabelWithIndexExists(Lcase(DT_MI24_0120_CHECK_TEXT_OF_COUNTED_OCC4),"3")
Call VerifyifGuiLabelWithIndexExists(Lcase(DT_MI24_0120_CHECK_TEXT_OF_COUNTED_OCC2),"2")
Call VerifyifGuiLabelWithIndexExists(Lcase(DT_MI24_0120_CHECK_TEXT_OF_COUNTED_OCC1),"1")
Call VerifyifGuiLabelWithIndexExists(Lcase(DT_MI24_0120_CHECK_TEXT_OF_COUNTED),"0")

Call SelectCheckBoxByGuiLabel(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC12,"17","1",DT_MI24_0120_NO_NAME_OCC9)
Call SelectCheckBoxByGuiLabel(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC14,"17","1",DT_MI24_0120_NO_NAME_OCC10)
Call SelectCheckBoxByGuiLabel(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC16,"17","1",DT_MI24_0120_NO_NAME_OCC11)
Call SelectCheckBoxByGuiLabel(DT_MI24_0120_CHECK_TEXT_OF_NO_NAME_OCC20,"17","1",DT_MI24_0120_NO_NAME_OCC12)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post Difference   \(Ctrl\+Shift\+F9\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Post   \(Ctrl\+S\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call GetTextStatusBar("DT_MI24_0701_GET_TEXT_OF_STATUSBAR_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickButton("Post   \(Ctrl\+S\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call GetTextStatusBar("DT_MI24_0701_GET_TEXT_OF_STATUSBAR_OCC2_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickButton("Post   \(Ctrl\+S\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call GetTextStatusBar("DT_MI24_0701_GET_TEXT_OF_STATUSBAR_OCC4_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call ClickButton("Post   \(Ctrl\+S\)",False)
'Capture the screenshot
Call TakeScreenShot()
Call GetTextStatusBar("DT_MI24_0701_GET_TEXT_OF_STATUSBAR_OCC8_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

wait(2)
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Refresh   \(Ctrl\+Shift\+F8\)",False)
'Capture the screenshot
Call TakeScreenShot()

Call VerifyifGuiLabelExists(Lcase(DT_MI24_0120_CHECK_TEXT_OF_LIST_DOES_NOT_CONTAIN_ANY_DATA_OCC1))


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()
