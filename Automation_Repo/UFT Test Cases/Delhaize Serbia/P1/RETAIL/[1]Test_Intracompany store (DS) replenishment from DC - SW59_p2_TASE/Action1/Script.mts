
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Intracompany store (DS) replenishment from DC - SW31_p2
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
'.................Test Script Name :Test_Intracompany store (DS) replenishment from DC - SW33_p2
'.................Author : TCS 	   :Raushan
'................ Creation Date    :4th Nov
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_Intracompany store (DS) replenishment from DC - SW59_p2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Intracompany store (DS) replenishment from DC - SW31_p2.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'Login to SAP System
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  

''----------------------Tcode ME23N----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

'Click on Other Purchase order and enter purchase order
Call ClickButtonIfExist("Other Purchase Order   \(Shift\+F5\)",False)
wait(2)
Call SetTextboxPopupIfExist("MEPO_SELECT-EBELN","Pur\. Order",DT_ME23N_0003_PUR_ORDER)
Call PressEnter()
Wait(2)
Call TakeScreenShot()

Call ClickButtonIfExist("Expand Item Details Ctrl\+F4",False)
wait(2)
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()    
'Navigate to Purchase Order History Tab
Call SelectTab("ITEM_DETAIL","Purchase Order History",False)
Wait(1)
Call TakeScreenShot()

'Get The Article Document
Call GetGridContent("",0,"Article Document",1,"Short Text","Lfs","DT_PO_ARTICLE_DOCUMENT")

'----------------------Tcode VL03N----------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

Call SetTcode(DT_ME23N_0014_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_ME23N_0014_OKCD)
Call TakeScreenShot()

'Enter the Outbound Delivery No
Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_ME23N_4004_OUTBOUND_DELIVERY,FALSE)
Call TakeScreenShot()
Call PressEnter()
Wait(3)
Call TakeScreenShot()

'Click on Document Flow
Call ClickButton("Document Flow   \(F7\)",False)
Wait(2)
Call TakeScreenShot()

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

'navigate to "Extras;Delivery Output;Header"
Call SelectMenuBar("Extras;Delivery Output;Header")
Call TakeScreenShot()

'Get The Output Type
Call GetTableCellData("SAPDV70ATC_NAST3","Output Type","1","Description","Del. DC2DC Supplying","DT_OUTPUT_TYPE_0",False)

'------------------------'Log Off From Applicaton--------------------------------
Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************

