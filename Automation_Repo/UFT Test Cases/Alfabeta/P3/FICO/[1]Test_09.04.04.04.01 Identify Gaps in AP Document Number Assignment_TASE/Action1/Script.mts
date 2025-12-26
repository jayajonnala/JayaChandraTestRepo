

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.04.04.04.01 Identify Gaps in AP Document Number Assignment
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

gstrTestCaseName = "Test_09.04.04.04.01 Identify Gaps in AP Document Number Assignment"
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

'''--------TransactionCode-S_ALR_87012342 ----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call SetTextbox("Company Code","SO_BUKRS-LOW","",DT_S_ALR_87012342_1000_COMPANY_CODE,False)
Call SetTextbox("Fiscal Year","PA_GJAHR","",((DT_S_ALR_87012342_1000_FISCAL_YEAR)),False)
Call TakeScreenShot

Call SelectCheckbox("PA_BUFF","0","ON",False)
Call TakeScreenShot
Call SelectCheckbox("PA_GONLY","0","ON",False)
Call TakeScreenShot
Call ClickButtonIfExist("Execute   \(F8\)",False)
Call TakeScreenShot

Call GetLabelContentByRefLabel("Gap From", -441, -48,"DT_S_ALR_87012342_0120_CHECK_TEXT_OF_NO_NAME_OUTPUT" , False)
Call GetLabelContentByRefLabel("Gap From", 56, -48,"DT_S_ALR_87012342_0120_CHECK_TEXT_OF_GR02_OUTPUT" , False)

Call VerifyifGuiLabelExistsByRelativeid(Cstr(DT_S_ALR_87012342_0120_CHECK_TEXT_OF_NO_NAME),"wnd[0]/usr/lbl[77,8]")
Call VerifyifGuiLabelExistsByRelativeid(Cstr(DT_S_ALR_87012342_0120_CHECK_TEXT_OF_GR02),"wnd\[0\]/usr/lbl\[6,8\]")


'''' Dynamic values are coming up in every release and with every changing year and hence verification is failing.

'Call GetLabelContentByRefLabel("Gap From", 0, -48,"DT_S_ALR_87012342_0120_CHECK_TEXT_OF_NO_NAME_OCC1_OUTPUT" , False)
'Call WriteRunTimeDataToExcelGlobalSheet ("DT_S_ALR_87012342_0120_CHECK_TEXT_OF_NO_NAME_OCC1_OUTPUT",DT_S_ALR_87012342_0120_CHECK_TEXT_OF_NO_NAME_OCC1)
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)
'
'Call GetLabelContentByRefLabel("Gap From", 21, -48,"DT_S_ALR_87012342_0120_CHECK_TEXT_OF_NO_NAME_OCC2_OUPTUT" , False)
'Call WriteRunTimeDataToExcelGlobalSheet ("DT_S_ALR_87012342_0120_CHECK_TEXT_OF_NO_NAME_OCC2_OUPTUT",DT_S_ALR_87012342_0120_CHECK_TEXT_OF_NO_NAME_OCC2)
'Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

'Call VerifyifGuiLabelExistsByRelativeid(Cstr(DT_S_ALR_87012342_0120_CHECK_TEXT_OF_NO_NAME_OCC3),"wnd\[0\]/usr/lbl\[14,8\]")

Call LogOff
Call FinalStatus()

































































''*********************************************************************************
'Functions Definition
'**********************************************************************************
'Name of the Function   : VerifyifGuiLabelExistsByRelativeid
'Author             :Shriya
'Description             : Verifies if Gui Label exiss   
'Input Parameters         :  Content to be verfied
'Output Parameters        : Null
'Creation Date : 
'***********************************************************************************
'Function VerifyifGuiLabelExistsByRelativeid(Content,Relativeid)
'***********************************************************************************

 

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

 


''************************************************************************************
'End Function - VerifyifGuiLabelExistsByRelativeid
'************************************************************************************
