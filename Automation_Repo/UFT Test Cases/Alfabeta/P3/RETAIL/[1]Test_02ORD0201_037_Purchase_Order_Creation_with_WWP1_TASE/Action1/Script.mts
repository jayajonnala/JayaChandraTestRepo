

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_02ORD0201_037_Purchase_Order_Creation_with_WWP1
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

gstrTestCaseName = "Test_02ORD0201_037_Purchase_Order_Creation_with_WWP1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
'
'--------------------------------------------  WWP1------------------------------------------------

Call SelectMenuBar("Goto;Variants;Get...")
Call SetTextbox("Created By","ENAME-LOW","",FormatBlank(DT_VARIANT),True) 
Call SetTextbox("Original Language","MLANGU-LOW","","",True)
Call ClickButton("Execute   \(F8\)",True)     
Call SelectRowGuiGrid("Variant Catalog for Program.*","","Variant name","AB",True)                  
Call ClickButton("Choose   \(F2\)",True)     

'

'Call SetTextbox("Vendor","01SLIFNR-LOW","",DT_WWP1_0001_VENDOR,False)
Call SetTextboxNoLabel("01SLIFNR-LOW","",DT_WWP1_0001_VENDOR,False)
Call SetTextbox("Site","01SWERKS-LOW","",DT_WWP1_0001_SITE,False)  
     
Call SetTextbox("Purchasing Group","01SSELEG-LOW","",DT_WWP1_0001_PURCHASING_GROUP,False)    
Call SetTextbox("Delivery date","01PEINDT","",ConvertDate(DT_WWP1_0001_DELIVERY_DATE),False)  
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False) 

Call PressEnter() 
Call ClickButtonToolBar("SEA","")    
Call SetTextbox("Find:","G_STRING","",DT_WWP1_102_FIND,True)  
Call ClickButton("OK   \(Enter\)",True)   
Call TakeScreenShot()
Call ClickButtonToolBar("EXE","")  
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)
Call SetTextbox("Positioning By Material No\.","WOD10001-WOD1_POSITION_BY_MATNR","",DT_WWP1_100_POSITIONING_BY_ARTICLE_NO,False)     
Call SetTableData("SAPLWOD1CONTROL_ARTICLES_TC","PO Quantity",1,"","",DT_WWP1_100_TABLECELL_PO_QUANTITY_0,False)
Call ClickButton("Store Temporarily   \(Shift\+Ctrl\+0\)",False)  
Call GetTextboxValue("G_KOPF_0100_EBELN_ANZEIGE",0,"DT_PO_NUMBER_OUTPUT",False)
Call VerifyTextBoxNoLabelContent("G_KOPF_0100_EBELN_ANZEIGE",0,DT_PO_NUMBER_OUTPUT,False)

Call ClickButton("Save   \(Ctrl\+S\)",False)  
Call TakeScreenShot()

Call LogOff()
Call FinalStatus ()





