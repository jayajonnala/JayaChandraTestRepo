'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Depreciate Assets_p1_TASE
'.................Author : TCS        :Bitan
'................ Creation Date    : 19th March
'.................Modified By :
'.................Modified Date/Details :
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

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

gstrTestCaseName = "Test_Depreciate Assets_p1_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Fixed Assets Sales_p3_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'
''''Login'''
'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM) 
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call TakeScreenShot()

Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS01_0105_ASSET_CLASS,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE,False)
Call TakeScreenShot()

Call PressEnter()  
Call TakeScreenShot()

Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION,False)
Call SetTextbox("Inventory number","ANLA-INVNR","",DT_AS01_1140_INVENTORY_NUMBER,False)
Call TakeScreenShot()

Call PressEnter()  
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Origin",False)
Call TakeScreenShot()
Call SelectTab("TABSTRIP100","Time-dependent",False)
Call SetTextbox("WBS element","ANLA-POSNR","",DT_WBS_ELEMENT,False)
Call TakeScreenShot()

Call PressEnter()  
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Time-dependent",False)
Call TakeScreenShot()
Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_COST_CENTER,False)
Call SetTextbox("Business Area","ANLZ-GSBER","",DT_AS01_1145_BUSINESS_AREA,False)
Call TakeScreenShot()

Call PressEnter()  

Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Allocations",False)

Call TakeScreenShot()
Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_AS01_1160_EVALUATION_GROUP_2,False)
Call SetTextbox("Evaluation group 1","ANLA-ORD41","",DT_AS01_1160_EVALUATION_GROUP_1,False)

Call TakeScreenShot()

Call PressEnter()  

Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Deprec. Areas",False)

Call TakeScreenShot()
Call SetTableDataNoRef("SAPLAISTTC_ANLB","DKey",5,DT_AS01_1190_TABLECELL_DKEY_4,False)

Call TakeScreenShot()

Call PressEnter()  

Call TakeScreenShot()
Call PressEnter()  

Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()

Call GetStatusBar("item1","DT_NEW_ASSET_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR)


Call LogOff()
Call FinalStatus ()




