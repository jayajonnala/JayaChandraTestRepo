

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.03.01.03 Clear AP Accounts (Manual and Automatic)_P3
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

gstrTestCaseName = "Test_09.04.03.01.03 Clear AP Accounts (Manual and Automatic)_P3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Users\jjonn\Desktop\TASEWork\Data\TASE_DT_09.04.01.01.01 Manage Manual Post  Direct Domestic Vendor Invoic.xls"
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

'''''''--------TransactionCode-FB65----------''''
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)


Call SetTextbox("Company Code","BUKRX-LOW","",DT_F13_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","GJAHX-LOW","",DT_F13_1000_FISCAL_YEAR,False)

Call ClickButton("%_DOCNR_%_APP_%-VALU_PUSH",False)
Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F13_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_F13_3010_TABLECELL_SINGLE_VALUE_1,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)

Call SelectCheckbox("X_LIFNR", 0, "ON", False)
Call SelectCheckbox("X_TESTL", 0, "ON", False)

Call SetTextbox("Posting Date","POSTDATE-LOW","",ConvertDate(DT_F13_1000_POSTING_DATE),False)
Call TakeScreenShot

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyifGuiLabelExistsByRelativeid(DT_F13_0120_CHECK_TEXT_OF_NO_NAME,"wnd\[0\]/usr/lbl\[47,12\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_F13_0120_CHECK_TEXT_OF_NO_NAME_OCC1,"wnd\[0\]/usr/lbl\[1,10\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_F13_0120_CHECK_TEXT_OF_NO_NAME_OCC2,"wnd\[0\]/usr/lbl\[1,11\]")

Call ClickButton("Back   \(F3\)",False)
Call SelectCheckbox("X_TESTL", 0, "OFF", False)

Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot
Call PressEnter()
Call GetStatusBar("item1","DT_DOC_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_OUTPUT&" was posted in company code GR02" )
Call TakeScreenShot
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_OUTPUT",DT_DOC)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)

''--------------------TCODE-FBL1N-----------------------

Call SetTcode(DT_F13_0100_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SelectRadioButton("X_CLSEL","Cleared items",False)
Call SetTextbox("Company Code","KD_BUKRS-LOW","",DT_F13_1000_COMPANY_CODE_OCC1,False)
Call SetTextbox("Vendor account","KD_LIFNR-LOW","",DT_F13_1000_VENDOR_ACCOUNT,False)
Call TakeScreenShot

Call ClickButtonIfExist("Dynamic selections   \(Shift\+F4\)",False)
Call ClickButtonIfExist("%_%%DYN012_%_APP_%-VALU_PUSH",False)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_F13_3010_TABLECELL_SINGLE_VALUE_0_OCC1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_F13_3010_TABLECELL_SINGLE_VALUE_1_OCC1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","3","","",DT_F13_3010_TABLECELL_SINGLE_VALUE_2,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call VerifyGridCellContent("", 1, "ICO_AUGP", 0, DT_F13_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_ICO_AUGP)
Call VerifyGridCellContent("", 1, "BELNR", 0, DT_F13_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
Call VerifyGridCellContent("", 2, "BELNR", 0, DT_F13_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)
Call VerifyGridCellContent("", 3, "BELNR", 0, DT_F13_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BELNR)
Call VerifyGridCellContent("", 5, "DMSHB", 0, DT_F13_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_5_DMSHB)

Call LogOff'
Call FinalStatus()

Function VerifyifGuiLabelExistsByRelativeid(Content,Relativeid)

 

 If Not (Environment.Value("blnFatalError") or Content= DS_SKIP) Then
    If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : VerifyifGuiLabelExists"
   
    strStepName = "Verify if Gui Label exists "

 

    If Content <>"" Then
     set   objLabel = SAPGuiSession(sessionObject).SapGuiWindow(windowobject).SAPGuiLabel("content:="&Content,"guicomponenttype:=30","relativeid:="&Relativeid)
               If objLabel.Exist Then
                    Call ReporterFunction(strLibraryFileName,"VerifyifGuiLabelExistsByRelativeid","2",Content,"Gui Label with value "&Content &" exists in the screen")    
                    strStatus = "PASS"
                    strMsg = "Gui Label with value "&Content&" exists in the screen"    
                    blnCaptureFlag = True
                              If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
                                  ImagePath=CaptureScreenshot(strStepName,objLabel,False,False,False)
                              End If
                   
                    Else
                    Call ReporterFunction(strLibraryFileName,"VerifyifGuiLabelExistsByRelativeid","1"," Label Content","Gui Label with value "&Content &" doesn't exist  in the screen")    
                    strStatus = "FAIL"
                    blndefectFlag =True
                    strMsg = "Gui Label with value "&Content&" doesn't exist  in the screen"
                    blnObjectError=True
                End If
           Else
        Call ReporterFunction(strLibraryFileName,"VerifyifGuiLabelExistsByRelativeid","1","Gui Label","Function Parameter Not Passed Properly. Check the --VerifyifGuiLabelExistsByRelativeid-- Function Call")
            strStatus = "FAIL"
            strMsg = "Function Parameter Not Passed Properly. Check the --VerifyifGuiLabelExistsByRelativeid-- Function Call-"
    End if

 


If strStatus = "FAIL"  Then
        VerifyifGuiLabelExistsByRelativeid = strMsg
        blnMainFailFlag = True
        ImagePath=CaptureScreenshot(strStepName,objLabel,False,False,False)
   
    Else
        VerifyifGuiLabelExistsByRelativeid = True
    End If
    If blnDefault_eSwiftReporting Then  
        Call UpdateResultHtml(strStepName,Content,strMsg,strStatus,"")
    End If

 

End If
End Function

