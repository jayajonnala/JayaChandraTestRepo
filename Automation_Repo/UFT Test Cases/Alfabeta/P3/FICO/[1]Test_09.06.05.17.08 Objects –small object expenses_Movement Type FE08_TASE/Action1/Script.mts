		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.06.05.17.08 Objects –small object expenses_Movement Type FE08
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_09.06.05.17.08 Objects –small object expenses_Movement Type FE08_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'GetRowNo =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''''''''------------tcode-FAGLL03--------------''''''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FAGLL03_1000_GL_ACCOUNT,False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_FAGLL03_1000_COMPANY_CODE,False)
Call SelectRadioButton("X_AISEL","All items", False)
DT_FAGLL03_1000_POSTING_DATE= "01.01."& CSTR(Year(Date)-5)
Call SetTextbox("Posting Date","SO_BUDAT-LOW","",ConvertDate(DT_FAGLL03_1000_POSTING_DATE),False)
Call SetTextbox("to","SO_BUDAT-HIGH","",ConvertDate(DT_FAGLL03_1000_TO),False)

Call ClickButton("Custom Selections   \(Ctrl\+F1\)",false)
Call ActivateNodeGuiTree("0","#4;#8")
Call SetTextBox("Profit Center","%%DYN001-LOW","",DT_FAGLL03_0100_PROFIT_CENTER,False)
Call ClickButton("Save   \(Ctrl\+S\)",False)

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call VerifyStatusBarMessageType("S")
Call TakeScreenShot

Call SetHorizontalScrollBar(75,False)
Call TakeScreenshot

Call ClickLabel("Text", "1", False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Call SetTextbox("Text","%%DYN001-LOW","",DT_FAGLL03_1105_TEXT,True)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot

Call CLickButton("Display Document   \(Ctrl\+Shift\+F7\)",False)
Call TakeScreenshot
Call ClickButton("Call Up Document Overview   \(F9\)",False)
Call TakeScreenshot

Call GetGridContentByTitle("", 0, "Amount", 1, "DT_FAGLL03_0500_GRIDCELL_0_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FAGLL03_0500_GRIDCELL_0_OUTPUT",DT_FAGLL03_0500_GRIDCELL_0_)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetGridContentByTitle("", 0, "Amount", 2, "DT_FAGLL03_0500_GRIDCELL_0_OCC1_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FAGLL03_0500_GRIDCELL_0_OCC1_OUTPUT",DT_FAGLL03_0500_GRIDCELL_0_OCC1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetGridContentByTitle("", 0, "Amount", 3, "DT_FAGLL03_0500_GRIDCELL_0_OCC2_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FAGLL03_0500_GRIDCELL_0_OCC2_OUTPUT",DT_FAGLL03_0500_GRIDCELL_0_OCC2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)


Call VerifyGridCellContent("", 1, "BSCHL", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("", 1, "KTONR", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("", 1, "MWSKZ", 0, "")
Call VerifyGridCellContent("", 1, "Amount", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)

Call VerifyGridCellContent("", 2, "BSCHL", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("", 2, "KTONR", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
'Call VerifyGridCellContent("", 2, "MWSKZ", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)
Call VerifyGridCellContent("", 2, "Amount", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)

Call VerifyGridCellContent("", 3, "BSCHL", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("", 3, "KTONR", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
'Call VerifyGridCellContent("", 3, "MWSKZ", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_MWSKZ)
Call VerifyGridCellContent("", 3, "Amount", 0, DT_FAGLL03_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)



Call ClickButton("Display Document Header   \(F5\)",False)
Call VerifyTextBoxContent("Document type","BKPF-BLART", 0, DT_FAGLL03_1710_CHECK_TEXT_OF_DOCUMENT_TYPE, True)
Call CLickButton("Cancel   \(F12\)",True)
Call LogOff'
Call FinalStatus()







































































''**********************************************************************************
'Name of the Function   :  SetHorizontalScrollBar
'Description             : Selects all rows in a SAP Gui Grid        
'Input Parameters         : gridTitle : Title property of the GuiGrid 
'                                                 blnIsItPopup-A  parameters whiich  checks if the SAP window is main window or a popup 
'                                                                            =Yes - if it is popup ;   No/"" =If it is main SAP  window                                      
'Output Parameters        :Null
'***********************************************************************************
'Function SetHorizontalScrollBar(gridTitle,columnName,refValue)
'***********************************************************************************
Public Function SetHorizontalScrollBar(ScrollBarPosition,blnIsItPopup)

 


 If Not (Environment.Value("blnFatalError")) Then
       If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : SetHorizontalScrollBar"
   Dim objTextbox32,objTextbox31,objTextbox,objWindow
   Set objWindow= SetSAPwindowObj(blnIsItPopup)
   strStepName="Set Horizontal Scroll Bar "
   If ScrollBarPosition<>"" Then
      Set objWindow = SAPGuisession(sessionObject).sapguiwindow(objWindow)
            
             If objWindow.Exist Then
              objWindow.HorizontalScrollbarPosition ScrollBarPosition
                  strStatus = "DONE"
                        If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
                              ImagePath=CaptureScreenshot(strStepName,objWindow,False,False,False)
                          End if
                      Call ReporterFunction(strLibraryFileName,"SetHorizontalScrollBar","2","Window","Horizontal Scroll bar was set at position " & ScrollBarPosition)
                      strMsg="Horizontal Scroll bar was set at position " & ScrollBarPosition
             Else
                   Call ReporterFunction(strLibraryFileName,"SetHorizontalScrollBar","1","Window"," Object Missing")
                   strStatus = "FAIL"
                   strMsg = " The Window Object is not Found, Please verify the Screen"
                   blnObjectError=True
             End If
   Else
                   Call ReporterFunction(strLibraryFileName,"SetHorizontalScrollBar","1","Text box","Function Parameter Not Passed Properly. Check the --SetHorizontalScrollBar-function")
                   strStatus="FAIL"
                   strMsg = "Function Parameter Not Passed Properly. Check the --SetHorizontalScrollBar-- Function Call-"
   End If
       If  blnObjectError  Then
            Environment.Value("blnFatalError")=True
       End If

 


       If strStatus = "FAIL"  Then
            SetHorizontalScrollBar = strMsg
            blnMainFailFlag = True
            ImagePath=CaptureScreenshot(strStepName,objTextbox,False,False,False)
       Else
            SetHorizontalScrollBar = strMsg
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
'End Function - SetHorizontalScrollBar
'************************************************************************************


