		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.12.02.03.04 Manage Account Receivable Monthly Closing
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

gstrTestCaseName = "Test_09.12.02.03.04 Manage Account Receivable Monthly Closing"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'

''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

'''''''--------TransactionCode-F-30----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Document Date","BKPF-BLDAT", "", ConvertDate(DT_F30_0122_DOCUMENT_DATE), False)
Call SetTextbox("Type","BKPF-BLART", "", DT_F30_0122_TYPE, False)
Call SetTextbox("Company Code","BKPF-BUKRS","", DT_F30_0122_COMPANY_CODE, False)
Call SetTextbox("Currency/Rate","BKPF-WAERS", "", DT_F30_0122_CURRENCYRATE, False)
Call SetTextbox("Clearing text","RF05A-AUGTX", "", DT_F30_0122_CLEARING_TEXT, False)
Call SetTextbox("PstKy","RF05A-NEWBS", "", DT_F30_0122_PSTKY, False)
Call SetTextbox("Account","RF05A-NEWKO", "", DT_F30_0122_ACCOUNT, False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot

Call SetTextbox("Amount","BSEG-WRBTR", "", DT_F30_0300_AMOUNT, False)
Call SetTextbox("Profit Ctrs","COBL-PRCTR", "", DT_F30_1016_PROFIT_CTRS, False)
Call SetTextbox("Text","BSEG-SGTXT", "", DT_F30_0300_TEXT, False)
Call TakeScreenShot

Call ClickButton("Display Additional Data for Document Item   \(F7\)",False)
Call TakeScreenShot

Call SetTextbox("House Bank","BSEG-HBKID", "", DT_F30_0330_HOUSE_BANK, False)
Call SetTextbox("/","BSEG-HKTID", "", DT_F30_0330_BSEGHKTID, False)
Call TakeScreenShot
Call ClickButton("Choose open items   \(F6\)",False)

Call SetTextbox("Account","RF05A-AGKON", "", DT_F30_0710_ACCOUNT, False)
Call SetTextbox("Account Type","RF05A-AGKOA", "", DT_F30_0710_ACCOUNT_TYPE, False)
Call SetTextbox("Special G/L ind","RF05A-AGUMS", "", DT_F30_0710_SPECIAL_GL_IND, False)
Call SelectRadioButton("RF05A-XPOS1", "Document Number", False)
Call TakeScreenShot
Call PressEnter()
Call TakeScreenShot
Call SetTextbox("From","RF05A-SEL01", "0", DT_F30_0731_FROM, False)
Call SetTextbox("From","RF05A-SEL01", "1", DT_F30_0731_FROM_OCC1, False)
Call ClickButton("Process Open Items   \(Shift\+F4\)",False)
Call TakeScreenShot
Call VerifyTextBoxContent("Not assigned","RF05A-DIFFB", "0",DT_F30_6102_CHECK_TEXT_OF_NOT_ASSIGNED, False)
Call TakeScreenShot

Call ClickButton("Post   \(Ctrl\+S\)",False)
Call TakeScreenShot

Call GetStatusBar("item1","DT_F30_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR")
Call VerifyStatusBar("Document "&DT_F30_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR&" was posted in company code GR02")
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet("DT_F30_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR",DT_F30_0122_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

'''''''--------TransactionCode-zfiar_f104----------''''

Call SetTcode(DT_F30_0122_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectCheckbox("FOR_CUR", "0", DT_F30_1000_CARRY_OUT_POSTINGS, False)

Call SetTextbox("Customer account","DD_KUNNR-LOW", "0", DT_F30_1000_CUSTOMER_ACCOUNT, False)
Call SetTextbox("Company code","DD_BUKRS-LOW", "0", DT_F30_1000_COMPANY_CODE, False)
Call SetTextbox("Open items at key date","DD_STIDA", "0", ConvertDate(DT_F30_1000_OPEN_ITEMS_AT_KEY_DATE), False)
Call SetTextbox("Key Date","STICHTAG", "0", ConvertDate(DT_F30_1000_KEY_DATE), False)
Call SetTextbox("Valuation Area","BWBER", "0", DT_F30_1000_VALUATION_AREA, False)
Call SetTextbox("Document Type","B_BLART", "0", DT_F30_1000_DOCUMENT_TYPE, False)
Call SetTextbox("Posting date","B_BUDAT","0", ConvertDate(DT_F30_1000_POSTING_DATE), False)
Call TakeScreenShot
Call ClickButton("%_BELNR_%_APP_%-VALU_PUSH",False)
Call TakeScreenShot
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F30_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_F30_3010_TABLECELL_SINGLE_VALUE_1,True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call ClickBUtton("Back   \(F3\)",False)
Call TakeScreenShot
Call SelectCheckbox("B_PST_FL", "0", DT_F30_1000_CARRY_OUT_POSTINGS, False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyifGuiLabelExistsByRelativeid(DT_F30_0120_CHECK_TEXT_OF_NO_NAME_OCC4,"wnd\[0\]/usr/lbl\[49,7\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_F30_0120_CHECK_TEXT_OF_NO_NAME_OCC5,"wnd\[0\]/usr/lbl\[49,8\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_F30_0120_CHECK_TEXT_OF_NO_NAME_OCC2,"wnd\[0\]/usr/lbl\[131,7\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_F30_0120_CHECK_TEXT_OF_NO_NAME_OCC3,"wnd\[0\]/usr/lbl\[131,8\]")

Call SetVerticalScrollBar("5",False)

Call GetLabelContentByRefLabel("Text", 441, -32, "DT_F30_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_F30_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT",DT_F30_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
Call GetLabelContentByRefLabel("Text", 441, -48, "DT_F30_0120_CHECK_TEXT_OF_NO_NAME_OCC1_OUTPUT", False)
Call WriteRunTimeDataToExcelGlobalSheet("DT_F30_0120_CHECK_TEXT_OF_NO_NAME_OCC1_OUTPUT",DT_F30_0120_CHECK_TEXT_OF_NO_NAME_OCC1_OUTPUT)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''''--------TransactionCode-faglb03----------''''

Call SetTcode(DT_F30_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Account Number","RACCT-LOW", "0", DT_F30_1000_ACCOUNT_NUMBER, False)
Call SetTextbox("Company Code","RBUKRS-LOW", "0", DT_F30_1000_COMPANY_CODE_OCC1, False)
Call SetTextbox("Fiscal Year","RYEAR", "0", Year(DT_F30_1000_FISCAL_YEAR), False)
Call PressEnter()
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

If Cint(Month(Date)) < 12  Then
	DT_PERIOD_ROWNUMBER = Cint(Month(Date))+1
	Else
	DT_PERIOD_ROWNUMBER = 2
End If

Call SelectCellGuiGrid("", "0", DT_PERIOD_ROWNUMBER, "Period", False)
Call DoubleClick()
Call TakeScreenShot

Call SelectColumnGuiGrid("", "0", "Document Number", False)
Call ClickButtonIfExist("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Document Number","%%DYN001-LOW", "0", DT_F30_3010_TABLECELL_SINGLE_VALUE_0_OCC1,True)
Call SetTextbox("to","%%DYN001-HIGH", "0", DT_F30_3010_TABLECELL_SINGLE_VALUE_1_OCC1, True)
Call ClickButton("Execute   \(Enter\)",True)


Call VerifyGridCellContent("", 1, "BELNR", 0, DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 2, "BLART", 0, DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLART)
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 1, "DMSHB", 0,DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 2, "DMSHB", 0,DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)
Call VerifyGridCellContent("", 1, "HWAER", 0,DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HWAER)
Call VerifyGridCellContent("", 2, "HWAER", 0,DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HWAER)
Call VerifyGridCellContent("", 1, "PRCTR", 0,DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("", 2, "PRCTR", 0,DT_F30_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call LogOff'



''**********************************************************************************
'Name of the Function   :  SetVerticalScrollBar
'Description             : Selects all rows in a SAP Gui Grid        
'Input Parameters         : gridTitle : Title property of the GuiGrid 
'                                                 blnIsItPopup-A  parameters whiich  checks if the SAP window is main window or a popup 
'                                                                            =Yes - if it is popup ;   No/"" =If it is main SAP  window                                      
'Output Parameters        :Null
'***********************************************************************************
'Function SetVerticalScrollBar(gridTitle,columnName,refValue)
'***********************************************************************************
Public Function SetVerticalScrollBar(ScrollBarPosition,blnIsItPopup)


 If Not (Environment.Value("blnFatalError")) Then
       If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : SetVerticalScrollBar"
   Dim objTextbox32,objTextbox31,objTextbox,objWindow
   Set objWindow= SetSAPwindowObj(blnIsItPopup)
   strStepName="Set Vertical Scroll Bar "
   If ScrollBarPosition<>"" Then
      Set objWindow = SAPGuisession(sessionObject).sapguiwindow(objWindow)
            
             If objWindow.Exist Then
              objWindow.VerticalScrollbarPosition ScrollBarPosition
                  strStatus = "DONE"
                        If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
                              ImagePath=CaptureScreenshot(strStepName,objWindow,False,False,False)
                          End if
                      Call ReporterFunction(strLibraryFileName,"SetVerticalScrollBar","2","Window","Vertical Scroll bar was set at position " & ScrollBarPosition)
                      strMsg="Vertical Scroll bar was set at position " & ScrollBarPosition
             Else
                   Call ReporterFunction(strLibraryFileName,"SetVerticalScrollBar","1","Window"," Object Missing")
                   strStatus = "FAIL"
                   strMsg = " The Window Object is not Found, Please verify the Screen"
                   blnObjectError=True
             End If
   Else
                   Call ReporterFunction(strLibraryFileName,"SetVerticalScrollBar","1","Text box","Function Parameter Not Passed Properly. Check the --SetVerticalScrollBar-function")
                   strStatus="FAIL"
                   strMsg = "Function Parameter Not Passed Properly. Check the --SetVerticalScrollBar-- Function Call-"
   End If
       If  blnObjectError  Then
            Environment.Value("blnFatalError")=True
       End If


       If strStatus = "FAIL"  Then
            SetVerticalScrollBar = strMsg
            blnMainFailFlag = True
            ImagePath=CaptureScreenshot(strStepName,objTextbox,False,False,False)
       Else
            SetVerticalScrollBar = strMsg
       End If
       
        
'        If blnWriteDataToOutputSheet Then
'                                        strStepName = "Retrieve '"&strCapturedText&"' value from window and store in '"&gstrOutputSheetName&"' sheet under column '"& dataTableColumnName & "'"                                
'                                        call WriteRunTimeDataToExcel (dataTableColumnName,strCapturedText)
'                            Else
'                                        strStepName = "Retrieve '"&strCapturedText&"' value from textbox"
'                            End If 
                            
    
        If blnDefault_eSwiftReporting Then  
           Call UpdateResultHtml(strStepName,"",strMsg,strStatus,strCapturedText)
        End If
      
              
               Set objwindowcaptureSAP =Nothing
'               Set objTextbox31=Nothing
'               Set objTextbox = Nothing
               Set objWindow=Nothing
   End If
   
   
End Function
'************************************************************************************
'End Function - SetVerticalScrollBar
'************************************************************************************

