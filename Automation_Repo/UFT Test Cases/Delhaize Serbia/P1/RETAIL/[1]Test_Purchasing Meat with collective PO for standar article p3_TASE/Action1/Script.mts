
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Purchasing Meat with collective PO for standar article p3
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

gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Purchasing Meat with collective PO for standar article p3
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 19th Nov
'.................Modified By :
'.................Modified Date/Details :
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Purchasing Meat with collective PO for standar article p3"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Regular purchasing in SW31 dry goods_p2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode ME23N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call TakeScreenShot()

Call ClickButtonIfExist("Other Purchase Order   \(Shift\+F5\)",False)
Call TakeScreenShot()
wait(2)
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_PUR_ORDER,True)
'''''''''''''''''Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN",0,DT_ME23N_0003_PUR_ORDER,True)
Call TakeScreenShot()

Call ClickButtonIfExist("Other Document   \(Enter\)",True)
wait(2)
Call TakeScreenShot()

Call ClickButtonIfExist("Expand Header Ctrl\+F2",False)
wait(2)
Call TakeScreenShot()

'Click on Messages button and get the output type.
Call ClickButtonIfExist("Messages   \(Shift\+F9\)",False)
wait(2)
Call TakeScreenShot()

Call VerifyTableCellContent(3,"Output Type","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_2)
Call VerifyTableCellContent(4,"Output Type","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_3)

''----------------------Tcode ME9F----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_ME23N_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ME23N_0100_OKCD)
Call TakeScreenShot()

Call SetTextbox("Document Number","S_EBELN-LOW","",DT_ME23N_1000_DOCUMENT_NUMBER,False)   
Call SetTextbox("Processing Status","P_VSTAT",""," ",False)  
Call SetTextbox("Time Created","P_ERUHR",""," ",False)  
Call TakeScreenShot()

'Click on Execute Button
Call ClickButton("Execute   \(F8\)",False) 
Wait(2)
Call TakeScreenShot()

Call SelectCheckBoxByGuiLabel("ZMAI","38","1","ON") 
Call TakeScreenShot()
Wait(1)

'Click on Trial Printout  Button
Call ClickButton("Trial Printout   \(Shift\+F4\)",False) 
Wait(2)
Call TakeScreenShot()


'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

