
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Intracompany store (DS) replenishment from DC - SW41 (Fresh m_p2
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
'.................Test Script Name :Test_Intracompany store (DS) replenishment from DC - SW41 (Fresh m_p2
'.................Author : TCS 	   :Raushan
'................ Creation Date    :4th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Intracompany store (DS) replenishment from DC - SW41 (Fresh m_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Intracompany store (DS) replenishment from DC - SW41 (Fresh m_p2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode ME23N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

Call ClickButtonIfExist("Other Purchase Order   \(Shift\+F5\)",False)
wait(2)
Call TakeScreenShot()
Call SetTextbox("Pur\. Order", "MEPO_SELECT-EBELN", "", DT_ME23N_0003_PUR_ORDER, True)
Call TakeScreenShot()
Call PressEnter()
Wait(2)
Call TakeScreenShot()

Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
Call TakeScreenShot()
Call SelectTab("ITEM_DETAIL","Purchase Order History",False)
Wait(1)
Call TakeScreenShot()

'Call SelectColumnGuiGrid("",0,"Short Text",False)
'Call ClickButtonToolBar("&MB_FILTER",0)
'Call SetTextboxPopupIfExist("%%DYN001-LOW","Short Text","Lfs")
'Wait(1)
'Call ClickButton("Execute   \(Enter\)",True)
'Wait(2)

Call GetGridContent("",0,"Article Document",1,"Short Text","Lfs","DT_PO_ARTICLE_DOCUMENT")
Call GetGridContent("",0,"Posting Date",1,"Short Text","Lfs","DT_PO_POSTING_DATE")

Call ClickButton("Messages   \(Shift\+F9\)",False)
Wait(2)
Call TakeScreenShot()

Call VerifyTableCellContent(1,"Output Type","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_0)
Call VerifyTableCellContent(2,"Output Type","SAPDV70ATC_NAST3",DT_ME23N_0100_CHECK_TEXT_OF_TABLECELL_OUTPUT_TYPE_1)

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

