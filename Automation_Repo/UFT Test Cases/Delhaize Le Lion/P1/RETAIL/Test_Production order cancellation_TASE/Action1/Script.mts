
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Production order cancellation _TASE
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
	DataRowSet= Parameter("datatable_row")	
End If

If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If


gstrTestCaseName = "TC_08_Test_Production order cancellation _TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_MD_ABI067_001 Create ZCXT Retail Customer Local or Foreign_TASE.xls"

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''----------------------Tcode CO13----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call SetTextbox("Order","CORUF-AUFNR", "", DT_ORDER_NUMMBER, False)
Call TakeScreenShot()
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()
Call ClickButton("Yes",False)
Call TakeScreenShot()

'Call VerifyTextBoxContent("Article Number","CAUFVD-MATNR", 0, DT_ARTICLE_NUMBER, False)
Call VerifyTextBoxContent("Yield","AFRUD-LMNGA", 0, DT_YIELD, False)
Call VerifyTextBoxContent("Work Center","AFVGD-ARBPL", 0, DT_WORK_CENTER, False)
Call VerifyTextBoxContent("Site","AFVGD-WERKS", 0, DT_SITE, False)

Call  ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()
Wait 30
'Call ValueUpdateInWord("Sanity Test")
Call TakeScreenShot()

Call  ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

Call VerifystatusBar("Confirmation of order "&DT_ORDER_NUMMBER&" is cancelled")

Call LogOff()
Call FinalStatus()

















'Function ValueUpdateInWord(StrText)
'
'If Not Environment.Value("blnFatalError") Then
'
'
'	If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : ValueUpdateInWord"
'	strFuncName = "ValueUpdateInWord"
'	strStepName = "Update text "&StrText&"in Word Document"
'    
'    Dim ObjWord,ObjDoc
'    
'		Set ObjWord = GetObject(,"Word.Application")
'		
'				If  StrText <>"" Then
'								ObjWord.Visible = True
'								Set ObjDoc = ObjWord.ActiveDocument
'								ObjDoc.Content.Delete
'								ObjDoc.Content.InsertAfter StrText
'								strStatus = "DONE"
'								strMsg = "Word Document is updated with text" &StrText
'								Call ReporterFunction(strLibraryFileName,strFuncName,"2",strStepName,strMsg)
'								
'								If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
'								ImagePath=CaptureScreenshot(strStepName,ObjDoc,True,False,False)
'								End if
'																	
'				Else 
'								strStatus = "FAIL"
'								strMsg = "Please enter the text to be updated in word document"
'								blnObjectError=True
'								Call ReporterFunction(strLibraryFileName,strFuncName,"1",strStepName,strMsg)
'
'				End If			
'									
'						
'			 
'
'							If  blnObjectError  Then
'								Environment.Value("blnFatalError")=True
'							End If
'
'							If strStatus = "FAIL"  Then
'								ValueUpdateInWord = strMsg
'								blnMainFailFlag = True
'								ImagePath=CaptureScreenshot(strStepName,objButton,True,True,True)
'							Else
'								ValueUpdateInWord = True
'							End If
'
'							If blnDefault_eSwiftReporting Then  
'									Call UpdateResultHtml (strStepName,StrText,strMsg,strStatus,"")
'							End If
'
'						Set ObjWord = Nothing
'						Set ObjDoc= Nothing
'End If
'
'End Function
