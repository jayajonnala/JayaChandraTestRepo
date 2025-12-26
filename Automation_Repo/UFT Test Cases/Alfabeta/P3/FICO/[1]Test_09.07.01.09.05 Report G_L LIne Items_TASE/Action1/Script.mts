		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.09.05 Report G_L LIne Items
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

gstrTestCaseName = "Test_09.07.01.09.05 Report G_L LIne Items"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
''''''''--------TransactionCode-S_ALR_87012280----------''''

Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_S_ALR_87012280_1000_GL_ACCOUNT,False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_S_ALR_87012280_1000_COMPANY_CODE,False)
Call ClickButton("Execute   \(F8\)",False)
Call ClickLabel("DocumentNo", "0", False)
Call ClickButton("Set Filter   \(Ctrl\+F5\)",False)
Call ClickButton("%_%%DYN001_%_APP_%-VALU_PUSH",True)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_S_ALR_87012280_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_S_ALR_87012280_3010_TABLECELL_SINGLE_VALUE_1,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)

Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot

Call ClickButton("Change Layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButton("Find",True)
Call SetTextbox("Find","GD_SEARCHSTR","",DT_S_ALR_87012280_0850_FIND,True)

Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenshot
Call ClickButton("Show Sel\. Fields \(Ctrl\+F3\)",True)
Call ClickButton("Copy   \(Enter\)",True)
'
Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_AB, "wnd\[0\]/usr/lbl\[16,8\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_AB_OCC1, "wnd\[0\]/usr/lbl\[16,7\]")

Call VerifyifGuiLabelExistsByRelativeid(ConvertDate(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC2), "wnd\[0\]/usr/lbl\[32,7\]")
Call VerifyifGuiLabelExistsByRelativeid(ConvertDate(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC3), "wnd\[0\]/usr/lbl\[32,8\]")

Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC4, "wnd\[0\]/usr/lbl\[43,8\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC5, "wnd\[0\]/usr/lbl\[43,7\]")

Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_EUR, "wnd\[0\]/usr/lbl\[90,7\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_EUR_OCC1, "wnd\[0\]/usr/lbl\[90,8\]")

Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC7, "wnd[0]/usr/lbl[115,7]")
'Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC7, "wnd\[0\]/usr/lbl\[115,8\]")

Call SetHorizontalScrollBar(10, False)
wait(5)

Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_GZ000018, "wnd[0]/usr/lbl[177,7]")
Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_GZ000018_OCC1, "wnd[0]/usr/lbl[177,8]")

''''''''--------TransactionCode-/nS_ALR_87012281----------''''
'
Call SetTcode(DT_S_ALR_87012280_0120_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Chart of accounts","SD_KTOPL-LOW","",DT_S_ALR_87012280_1000_CHART_OF_ACCOUNTS,False)
Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_S_ALR_87012280_1000_GL_ACCOUNT_OCC1,False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_S_ALR_87012280_1000_COMPANY_CODE_OCC1,False)

Call ClickButton("Execute   \(F8\)",False)
Call ClickLabel("DocumentNo", "0", False)
Call ClickButton("Set Filter   \(Ctrl\+F5\)",False)
Call ClickButton("%_%%DYN001_%_APP_%-VALU_PUSH",True)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_S_ALR_87012280_3010_TABLECELL_SINGLE_VALUE_0_OCC1,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_S_ALR_87012280_3010_TABLECELL_SINGLE_VALUE_1_OCC1,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)

Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot

Call ClickButton("Change Layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButton("Find",True)
Call SetTextbox("Find","GD_SEARCHSTR","",DT_S_ALR_87012280_0850_FIND_OCC1,True)

Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenshot
Call ClickButton("Show Sel\. Fields \(Ctrl\+F3\)",True)
Call ClickButton("Copy   \(Enter\)",True)


Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_AB_OCC2, "wnd\[0\]/usr/lbl\[16,8\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_AB_OCC3, "wnd\[0\]/usr/lbl\[16,9\]")

Call VerifyifGuiLabelExistsByRelativeid(ConvertDate(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC10), "wnd\[0\]/usr/lbl\[32,8\]")
Call VerifyifGuiLabelExistsByRelativeid(ConvertDate(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC11), "wnd\[0\]/usr/lbl\[32,9\]")

Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC12, "wnd\[0\]/usr/lbl\[43,9\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC13, "wnd\[0\]/usr/lbl\[43,8\]")

Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC8, "wnd\[0\]/usr/lbl\[21,8\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC9, "wnd\[0\]/usr/lbl\[21,9\]")

Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_EUR_OCC2, "wnd\[0\]/usr/lbl\[90,8\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_EUR_OCC3, "wnd\[0\]/usr/lbl\[90,9\]")

''Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC6, "wnd\[0\]/usr/lbl\[115,8\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC15, "wnd\[0\]/usr/lbl\[115,8\]")

Call SetHorizontalScrollBar(10, False)

Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_GZ000018_OCC2, "wnd[0]/usr/lbl[177,8]")
Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_GZ000018_OCC3, "wnd[0]/usr/lbl[177,9]")

'Call VerifyifGuiLabelExists(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_GZ000018_OCC2)
'Call VerifyifGuiLabelExists(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_GZ000018_OCC3)

' CheckifGuiLabelExists(Content)


''''''''--------TransactionCode-//nS_ALR_87012282----------''''

Call SetTcode(DT_S_ALR_87012280_1000_OKCD)     
Call PressEnter()     
Call TakeScreenShot

Call SetTextbox("Chart of accounts","SD_KTOPL-LOW","",DT_S_ALR_87012280_1000_CHART_OF_ACCOUNTS_OCC1,False)
Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_S_ALR_87012280_1000_GL_ACCOUNT_OCC2,False)
Call SetTextbox("Company code","SD_BUKRS-LOW","",DT_S_ALR_87012280_1000_COMPANY_CODE_OCC2,False)

Call ClickButton("Execute   \(F8\)",False)
Call ClickLabel("DocumentNo", "0", False)
Call ClickButton("Set Filter   \(Ctrl\+F5\)",False)
Call ClickButton("%_%%DYN001_%_APP_%-VALU_PUSH",True)

Call SetTableData("SAPLALDBSINGLE","Single value","1","","",DT_S_ALR_87012280_3010_TABLECELL_SINGLE_VALUE_0_OCC2,True)
Call SetTableData("SAPLALDBSINGLE","Single value","2","","",DT_S_ALR_87012280_3010_TABLECELL_SINGLE_VALUE_1_OCC2,True)
Call ClickButtonIfExist("Copy   \(F8\)",True)

Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenshot

Call ClickButton("Change Layout\.\.\.   \(Ctrl\+F8\)",False)

Call ClickButton("Find",True)
Call SetTextbox("Find","GD_SEARCHSTR","",DT_S_ALR_87012280_0850_FIND_OCC2,True)

Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenshot
Call ClickButton("Show Sel\. Fields \(Ctrl\+F3\)",True)
Call ClickButton("Copy   \(Enter\)",True)

Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_AB_OCC4, "wnd\[0\]/usr/lbl\[16,7\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_AB_OCC5, "wnd\[0\]/usr/lbl\[16,8\]")

Call VerifyifGuiLabelExistsByRelativeid(ConvertDate(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC18), "wnd\[0\]/usr/lbl\[32,7\]")
Call VerifyifGuiLabelExistsByRelativeid(ConvertDate(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC19), "wnd\[0\]/usr/lbl\[32,8\]")

Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC21, "wnd\[0\]/usr/lbl\[43,7\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC20, "wnd\[0\]/usr/lbl\[43,8\]")

Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC16, "wnd\[0\]/usr/lbl\[21,7\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC17, "wnd\[0\]/usr/lbl\[21,8\]")

Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_EUR_OCC4, "wnd\[0\]/usr/lbl\[90,7\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_EUR_OCC5, "wnd\[0\]/usr/lbl\[90,8\]")

'Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC6, "wnd\[0\]/usr/lbl\[115,8\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_NO_NAME_OCC23, "wnd\[0\]/usr/lbl\[115,7\]")

Call SetHorizontalScrollBar(10, False)

Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_GZ00, "wnd\[0\]/usr/lbl\[177,7\]")
Call VerifyifGuiLabelExistsByRelativeid(DT_S_ALR_87012280_0120_CHECK_TEXT_OF_GZ00_OCC1, "wnd\[0\]/usr/lbl\[177,8\]")

Call logOff()
Call FinalStatus()




