		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.12.02.03.03 Manage Account Receivable Monthly Closing
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

gstrTestCaseName = "Test_09.12.02.03.03 Manage Account Receivable Monthly Closing"
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

'''''--------TransactionCode-ZFIAR_F104----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SelectCheckbox("B_PST_FL", "0",DT_ZFIAR_F104_1000_CARRY_OUT_POSTINGS, False)
Call SetTextbox("Customer account","DD_KUNNR-LOW","",DT_ZFIAR_F104_1000_CUSTOMER_ACCOUNT,False)
Call SetTextbox("Company code","DD_BUKRS-LOW","",DT_ZFIAR_F104_1000_COMPANY_CODE,False)
Call SetTextbox("Valuation Area","BWBER","",DT_ZFIAR_F104_1000_VALUATION_AREA,False)
Call SetTextbox("SpG/L ind\.doub\.rec\.","P_UMSKZ","",DT_ZFIAR_F104_1000_SPGL_INDDOUBREC,False)
Call SetTextbox("Posting date","B_BUDAT","",ConvertDate(DT_ZFIAR_F104_1000_POSTING_DATE),False)
Call SetTextbox("Document date","B_BLDAT","",ConvertDate(DT_ZFIAR_F104_1000_DOCUMENT_DATE),False)
Call TakeScreenShot
Call ClickButton("%_BELNR_%_APP_%-VALU_PUSH",False)
Call TakeScreenShot
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_ZFIAR_F104_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_ZFIAR_F104_3010_TABLECELL_SINGLE_VALUE_1,True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyifGuiLabelExistsByRelativeid(DT_ZFIAR_F104_0120_CHECK_TEXT_OF_C,"wnd\[0\]/usr/lbl\[6,7\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_ZFIAR_F104_0120_CHECK_TEXT_OF_C_OCC1,"wnd\[0\]/usr/lbl\[6,8\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_ZFIAR_F104_0120_CHECK_TEXT_OF_NO_NAME_OCC2,"wnd\[0\]/usr/lbl\[49,7\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_ZFIAR_F104_0120_CHECK_TEXT_OF_NO_NAME_OCC3,"wnd\[0\]/usr/lbl\[49,8\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_ZFIAR_F104_0120_CHECK_TEXT_OF_NO_NAME_OCC4,"wnd\[0\]/usr/lbl\[131,7\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_ZFIAR_F104_0120_CHECK_TEXT_OF_NO_NAME_OCC5,"wnd\[0\]/usr/lbl\[131,8\]")
Call SetVerticalScrollBar("5",False)
Call VerifyifGuiLabelExistsByRelativeid(DT_ZFIAR_F104_0120_CHECK_TEXT_OF_NO_NAME_OCC6,"wnd\[0\]/usr/lbl\[61,29\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_ZFIAR_F104_0120_CHECK_TEXT_OF_NO_NAME_OCC7,"wnd\[0\]/usr/lbl\[61,30\]")
Call GetLabelContentByRefLabel("Text", 441, -32, "DT_ZFIAR_F104_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT", False)
Call GetLabelContentByRefLabel("Text", 441, -48, "DT_ZFIAR_F104_0120_CHECK_TEXT_OF_NO_NAME_OCC1_OUTPUT", False)


'''''''----------------TransactionCode-FAGLL03-----------------------------------------------

Call SetTcode(DT_ZFIAR_F104_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_ZFIAR_F104_1000_GL_ACCOUNT,False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_ZFIAR_F104_1000_COMPANY_CODE_OCC1,False)
Call SelectRadioButton("X_AISEL","All Items", False)
Call TakeScreenShot
Call ClickButton("Custom Selections   \(Ctrl\+F1\)",False)
Call TakeScreenShot

Call ActivateNodeGuiTree(0, "#4;#1")
Call ClickButton("%_%%DYN001_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_ZFIAR_F104_3010_TABLECELL_SINGLE_VALUE_0_OCC1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_ZFIAR_F104_3010_TABLECELL_SINGLE_VALUE_1_OCC1,True)
Call TakeScreenShot
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "BELNR", 0, DT_ZFIAR_F104_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_ZFIAR_F104_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call VerifyGridCellContent("", 1, "BLART", 0, DT_ZFIAR_F104_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
Call VerifyGridCellContent("", 2, "BLART", 0, DT_ZFIAR_F104_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLART)
'Call VerifyGridCellContent("", 1, "BLDAT", 0, ConvertDate(DT_ZFIAR_F104_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT))
'Call VerifyGridCellContent("", 2, "BLDAT", 0, ConvertDate(DT_ZFIAR_F104_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLDAT))
Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_ZFIAR_F104_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_ZFIAR_F104_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 1, "DMSHB", 0,DT_ZFIAR_F104_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
Call VerifyGridCellContent("", 2, "DMSHB", 0,DT_ZFIAR_F104_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)

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

