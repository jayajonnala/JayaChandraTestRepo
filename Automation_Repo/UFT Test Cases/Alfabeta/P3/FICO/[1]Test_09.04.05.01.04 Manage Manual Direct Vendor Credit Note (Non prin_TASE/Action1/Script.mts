

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.05.01.04 Manage Manual Direct Vendor Credit Note (Non prin

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

gstrTestCaseName = "Test_09.04.05.01.04 Manage Manual Direct Vendor Credit Note (Non prin"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\jjonn\Desktop\TASEWork\Data\TASE_DT_09.04.01.01.03 Manage Manual Direct Domestic Interco.xls"
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

'''--------TransactionCode-FB60 ----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB60_1000_COMPANY_CODE,True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot

Call SetComboByKey("Transactn",DT_FB60_1100_TRANSACTN)
Call TakeScreenShot

Call SetTextboxNoLabel("INVFO-ACCNT","",DT_FB60_0010_VENDOR,False)
Call SetTextbox("Amount","INVFO-WRBTR","",DT_FB60_0010_AMOUNT,False)
Call SetTextbox("Invoice date", "INVFO-BLDAT", "", ConvertDate(DT_FB60_0010_INVOICE_DATE), False)
Call SetTextbox("Posting Date", "INVFO-BUDAT", "", ConvertDate(DT_FB60_0010_POSTING_DATE), False)
Call SetTextbox("Reference", "INVFO-XBLNR", "", DT_FB60_0010_REFERENCE, False)
Call SetTextbox("Amount", "INVFO-WAERS", "", DT_CURR, False)

Call TakeScreenShot
Call SelectCheckbox("INVFO-XMWST","0","ON",False)
Call SetComboByKey("INVFO-MWSKZ",DT_FB60_0010_INVFOMWSKZ)
Call TakeScreenShot
Call SetTextbox("Text", "INVFO-SGTXT", "", DT_FB60_0010_TEXT, False)

Call SetTableData("SAPLFSKBTABLE", "G/L acct", "1", "", "", DT_FB60_0100_TABLECELL_GL_ACCT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Amount in doc.curr.", "1", "", "", DT_FB60_0100_TABLECELL_AMOUNT_IN_DOCCURR_0, False)
Call SetTableData("SAPLFSKBTABLE", "Text", "1", "", "", DT_FB60_0100_TABLECELL_TEXT_0, False)
Call SetTableData("SAPLFSKBTABLE", "Business area", "1", "", "", DT_FB60_0100_TABLECELL_BUSINESS_AREA_0, False)
Call SetTableData("SAPLFSKBTABLE", "Cost center", "1", "", "", DT_FB60_0100_TABLECELL_COST_CENTER_0, False)
Call TakeScreenShot
Call SetComboByKey("INVFO-BLART",DT_FB60_0010_DOCUMENT_TYPE)
Call TakeScreenShot
Call PressEnter()     
Call TakeScreenShot

Call SelectTab("TS","Details",False)
Call TakeScreenShot
Call SetTextbox("G/L", "INVFO-HKONT", "", DT_FB60_0050_GL, False)
Call SendKey("{F4}")
Wait(10)
Call ClickButton("Restrict Values   \(Shift\+F5\)",True)
Call TakeScreenShot
Call ClickButtonByIndex("Multiple selection",0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FB60_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_FB60_3010_TABLECELL_SINGLE_VALUE_1,True)
Call TakeScreenShot
Call ClickButton("Copy   \(F8\)",True)
Call TakeScreenShot
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot
Call VerifyifGuiLabelExists(DT_FB60_0120_CHECK_TEXT_OF_NO_NAME)
Call VerifyifGuiLabelExists(DT_FB60_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
Call ClickButton("Copy   \(Enter\)",True)
Call SetTextbox("G/L", "INVFO-HKONT", "", DT_FB60_0050_GL, False)
Call PressEnter()     
Call TakeScreenShot

Call SelectTab("TS","Payment",False)
Call TakeScreenShot

Call SelectTab("TS","Payment",False)
Call TakeScreenShot
Call SetTextbox("Pmnt Terms", "INVFO-ZTERM", "", DT_FB60_0020_PAYT_TERMS, False)
Call PressEnter()     
Call TakeScreenShot

Call ClickButton("Simulate Document Posting   \(F9\)",False)
Call PressEnter()
Call TakeScreenShot
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_DOC_1_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_1_OUTPUT&" was posted in company code GR02" )
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_1_OUTPUT",DT_DOC_1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call ClickButton("Cancel   \(F12\)",False)
'
''''--------TransactionCode-FBL1N ----------''''

Call SetTcode(DT_FB60_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectRadioButton("X_AISEL","All items",False)
Call SetTextbox("Company Code","KD_BUKRS-LOW","",DT_FB60_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_FB60_1000_VENDOR_ACCOUNT,False)
Call TakeScreenShot

Call ClickButton("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButton("%_%%DYN012_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_FB60_3010_TABLECELL_SINGLE_VALUE_0_OCC1,True)
Call ClickButton("Copy   \(F8\)",True)
Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot

Call ClickButton("Change layout\.\.\.   \(Ctrl\+F8\)",False)
Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB60_0841_SEARCH_TERM,True)
Call SetComboByKey("Search Direction",DT_FB60_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)

Call ClickButtonToolBar("&FIND",0)
Call TakeScreenShot
Call SetTextbox("Search Term:","GS_SEARCH-VALUE","",DT_FB60_0841_SEARCH_TERM_OCC1,True)
Call SetComboByKey("Search Direction",DT_FB60_0841_SEARCH_DIRCT)
Call TakeScreenShot
Call PressEnter
Call ClickButton("Cancel   \(F12\)",True)
Call ClickButton("Show Selected Fields \(F7\)",True)

Call ClickButton("Transfer   \(Enter\)",True)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BELNR", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
Call VerifyGridCellContent("", 1, "DMSHB", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 1, "HWAER", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 1, "SGTXT", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
Call VerifyGridCellContent("", 1, "HKONT", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("", 1, "ZTERM", 0, DT_FB60_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ZTERM)

Call LogOff
Call FinalStatus()

' **********************************************************************************
'Name of the Function    :  ClickButtonByIndex
'Description                      : Clicks on buttons like "save","execute","copy","post goods issue","post goods receipt","complete data".
'                                                   tooltipOrButtonName :- either value of tooltip property  or name property  of button
'                                                 blnIsItPopup:- A  parameter which  checks if the SAP window is main window or a popup 
'                                                =Yes - if it is popup ;   No/""=If it is main SAP  window
'Input Parameters         :  tooltipOrButtonName,blnIsItPopup
'Output Parameters        :Null
'***********************************************************************************
'Function ClickButton(tooltipOrButtonName,blnIsItPopup)
'***********************************************************************************
Function ClickButtonByIndex(tooltipOrButtonName,BtnIndex,blnIsItPopup)

If Not Environment.Value("blnFatalError") Then


    If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : ClickButtonByIndex"
    strFuncName = "ClickButton"
    'Val = "<b>Button Component :</b>" & tooltipOrButtonName
    strStepName = "Click " & "'" & Replace(tooltipOrButtonName,"\","") & "'" & " Button"
    
    Dim objButton,objWindow
    Set objWindow= SetSAPwindowObj(blnIsItPopup)
                If  tooltipOrButtonName <>"" Then
                            
                             If SAPGuiSession(sessionObject).sapguiwindow(objWindow).SAPGuiButton("guicomponenttype:=40","tooltip:="& tooltipOrButtonName,"index:="&BtnIndex).Exist(2) Then
                                     Set objButton=SAPGuiSession(sessionObject).sapguiwindow(objWindow).SAPGuiButton("guicomponenttype:=40","tooltip:="& tooltipOrButtonName,"index:="&BtnIndex)
                                     blnObjectFound = True
                             ElseIf SAPGuiSession(sessionObject).sapguiwindow(objWindow).SAPGuiButton("guicomponenttype:=40","name:="&tooltipOrButtonName,"index:="&BtnIndex).Exist(2) Then
                                     Set objButton=SAPGuiSession(sessionObject).sapguiwindow(objWindow).SAPGuiButton("guicomponenttype:=40","name:="&tooltipOrButtonName,"index:="&BtnIndex)
                                     blnObjectFound = True
                             End If                
                                 If  blnObjectFound = True Then
    

                                            If  objButton.Exist  Then     
                                                        blnObjectFound = True
                                                        
                                                                    If blnCaptureFlag = True  or  blnCreateImageEachStep = True or blnCreateTrainingDoc Then
                                                                                ImagePath=CaptureScreenshot(strStepName,objButton,True,True,True)
                                                                    End If    
                                                                    objButton.Click
                                                                    strStatus = "DONE"
                                                                    strMsg = "Button with :-"& vbNewLine & "tooltipOrButtonName:"& Replace(tooltipOrButtonName,"\","") &" has been clicked  successfully"
                                                                    Call ReporterFunction(strLibraryFileName,strFuncName,"2",strStepName,strMsg)
                                                                    
                                            Else 
                                                        strStatus = "FAIL"
                                                        strMsg = "The button: "& Replace(tooltipOrButtonName,"\","") & "is not found"
                                                        blnObjectError=True
                                                        Call ReporterFunction(strLibraryFileName,strFuncName,"1",strStepName,strMsg)

                                            End If            
                                    
                        
               Else 
                            blnObjectError=True
                            strStatus = "FAIL"
                            strMsg = "The button:: "& tooltipOrButtonName & " is not Found--Check the Screen Properly"
                            Call ReporterFunction(strLibraryFileName,strFuncName,"1",strStepName,strMsg)
                
               End If
                            

     Else
                strStatus = "FAIL"
                strMsg = "Function Parameter Not Passed Properly. Check the --ClickButtonByIndex-- Function Call--"
                Call ReporterFunction(strLibraryFileName,strFuncName,"1",strStepName,strMsg)
    End If

                            If  blnObjectError  Then
                                Environment.Value("blnFatalError")=True
                            End If

                            If strStatus = "FAIL"  Then
                                ClickButtonByIndex = strMsg
                                blnMainFailFlag = True
                                ImagePath=CaptureScreenshot(strStepName,objButton,True,True,True)
                            Else
                                ClickButtonByIndex = True
                            End If

                            If blnDefault_eSwiftReporting Then  
                                    Call UpdateResultHtml (strStepName,Val,strMsg,strStatus,"")
                            End If

                        Set objButton = Nothing
                        Set objWindow= Nothing
End If

End Function
'************************************************************************************
'End Function - ClickButtonByIndex
'************************************************************************************

