

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.04.05.01 Manage AP Invoice Number Allocation issues
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


gstrTestCaseName = "Test_09.04.04.05.01 Manage AP Invoice Number Allocation issues"
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

'''--------TransactionCode-S_ALR_87012341 ----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Company Code","BR_BUKRS-LOW","",DT_S_ALR_87012341_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","BR_GJAHR-LOW","",(Year(DT_S_ALR_87012341_1000_FISCAL_YEAR)),False)
Call SetTextbox("Posting date","BR_BUDAT-LOW","",ConvertDate(DT_S_ALR_87012341_1000_POSTING_DATE),False)

Call SetTextbox("to","BR_BUDAT-HIGH","",ConvertDate(DT_S_ALR_87012341_1000_TO),False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)

Call GetLabelContentByRefLabel("DocumentNo", 0, -48,"DT_S_ALR_87012341_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT" , False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_S_ALR_87012341_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT",DT_S_ALR_87012341_0120_CHECK_TEXT_OF_NO_NAME)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetLabelContentByRefLabel("DocumentNo", -77, -48,"DT_S_ALR_87012341_0120_CHECK_TEXT_OF_NO_NAME_OCC2_OUTPUT" , False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_S_ALR_87012341_0120_CHECK_TEXT_OF_NO_NAME_OCC2_OUTPUT",DT_S_ALR_87012341_0120_CHECK_TEXT_OF_NO_NAME_OCC2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetLabelContentByRefLabel("DocumentNo", -259, -48,"DT_S_ALR_87012341_0120_CHECK_TEXT_OF_NO_NAME_OCC6_OUTPUT" , False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_S_ALR_87012341_0120_CHECK_TEXT_OF_NO_NAME_OCC6_OUTPUT",DT_S_ALR_87012341_0120_CHECK_TEXT_OF_NO_NAME_OCC6)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call GetLabelContentByRefLabel("DocumentNo", -301, -48,"DT_S_ALR_87012341_0120_CHECK_TEXT_OF_NO_NAME_OCC8_OUTPUT" , False)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_S_ALR_87012341_0120_CHECK_TEXT_OF_NO_NAME_OCC8_OUTPUT",DT_S_ALR_87012341_0120_CHECK_TEXT_OF_NO_NAME_OCC8)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call VerifyifGuiLabelExistsByRelativeid(ConvertTripledigit(DT_S_ALR_87012341_0120_CHECK_TEXT_OF_NO_NAME_OCC2),"wnd\[0\]/usr/lbl\[12,7\]")
Call VerifyifGuiLabelExistsByRelativeid(Cstr(DT_S_ALR_87012341_0120_CHECK_TEXT_OF_NO_NAME_OCC6),"wnd\[0\]/usr/lbl\[38,7\]")
Call VerifyifGuiLabelExistsByRelativeid(Cstr(DT_S_ALR_87012341_0120_CHECK_TEXT_OF_NO_NAME_OCC8),"wnd\[0\]/usr/lbl\[44,7\]")
Call VerifyifGuiLabelExistsByRelativeid(Cstr(DT_S_ALR_87012341_0120_CHECK_TEXT_OF_EUR),"wnd\[0\]/usr/lbl\[63,7\]")

Call LogOff'
Call FinalStatus()























































Function ConvertTripledigit(StrSingle)
If Len (StrSingle)<3 Then
	ConvertTripledigit= "00" + StrSingle
Else
    ConvertTripledigit= StrSingle
End If

End Function



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

