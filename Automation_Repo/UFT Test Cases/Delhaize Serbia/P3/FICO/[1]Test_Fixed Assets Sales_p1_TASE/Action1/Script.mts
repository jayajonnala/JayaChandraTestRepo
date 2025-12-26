'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Fixed Assets Sales_p1_TASE
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

gstrTestCaseName = "Test_Fixed Assets Sales_p1_TASE"
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

'''Increment the parameter
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
'
'''----------------------Tcode AS01----------------------------
'
'Enter the Tcode
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Asset Class","ANLA-ANLKL","",DT_AS01_0105_ASSET_CLASS,False)
Call SetTextbox("Company Code","ANLA-BUKRS","",DT_AS01_0105_COMPANY_CODE,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SetTextbox("Description","ANLA-TXT50","",DT_AS01_1140_DESCRIPTION,False)
Call SetTextbox("Inventory number","ANLA-INVNR","",DT_AS01_1140_INVENTORY_NUMBER,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Origin",False)
'Capture the screenshot
Call TakeScreenShot()
Call SelectTab("TABSTRIP100","Time-dependent",False)
Call SetTextbox("WBS element","ANLA-POSNR","",DT_WBS_ELEMENT,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Time-dependent",False)
'Capture the screenshot
Call TakeScreenShot()
Call SetTextbox("Cost Center","ANLZ-KOSTL","",DT_AS01_1145_COST_CENTER,False)
Call SetTextbox("Business Area","ANLZ-GSBER","",DT_AS01_1145_BUSINESS_AREA,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Allocations",False)
'Capture the screenshot
Call TakeScreenShot()
Call SetTextbox("Evaluation group 2","ANLA-ORD42","",DT_AS01_1160_EVALUATION_GROUP_2,False)
Call SetTextbox("Evaluation group 1","ANLA-ORD41","",DT_AS01_1160_EVALUATION_GROUP_1,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call SelectTab("TABSTRIP100","Deprec. Areas",False)
'Capture the screenshot
Call TakeScreenShot()
Call SetTableDataNoRef("SAPLAISTTC_ANLB","DKey",5,DT_AS01_1190_TABLECELL_DKEY_4,False)
'Capture the screenshot
Call TakeScreenShot()

Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()
Call PressEnter()  
'Capture the screenshot
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call ClickButtonIfExist("Save",True)
wait(1)
Call TakeScreenShot()
'Validate If asset is created
Call GetStatusBar("item1","DT_NEW_ASSET_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
VerifyStatusBar(DT_AS01_0105_CHECK_TEXT_OF_STATUSBAR_OCC1)


'Log Off From Applicaton
Call LogOff()
Call FinalStatus ()


'''''''
'''''''''----------------------Tcode AIAB----------------------------
'''''''
'''''''Enter the Tcode
''''''Call SetTcode(DT_AS01_0105_OKCD) 
''''''Call PressEnter()     ' 
''''''Call CheckTCodeScreen(DT_AS01_0105_OKCD)
'''''''Capture the screenshot
''''''Call TakeScreenShot()
''''''
''''''Call SetTextbox("Asset","AICOM-ANLN1","",DT_AS01_0110_ASSET,False)
'''''''Capture the screenshot
''''''Call TakeScreenShot()
''''''
''''''Call ClickButton("Execute   \(F8\)",False)
'''''''Capture the screenshot
''''''Call TakeScreenShot()
''''''
''''''Call SelectRowGuiGridbyRowNo("","",DT_AS01_0500_GRIDCELL_0__OCC1,False)
'''''''Capture the screenshot
''''''Call TakeScreenShot()
''''''
''''''Call ClickButton("Enter distribution rules   \(Shift\+F6\)",False)
'''''''Capture the screenshot
''''''Call TakeScreenShot()
''''''
'''''''enter table data
''''''Call SetTableDataNoRef("SAPLKOBSTC_RULES","Cat",1,DT_AS01_0130_TABLECELL_CAT_0,FALSE)
''''''Call SetTableDataNoRef("SAPLKOBSTC_RULES","Settlement Receiver",1,DT_AS01_0130_TABLECELL_SETTLEMENT_RECEIVER_0,FALSE)
''''''Call SetTableDataNoRef("SAPLKOBSTC_RULES","%",1,DT_AS01_0130_TABLECELL__0,FALSE)
'''''''Capture the screenshot
''''''Call TakeScreenShot()
''''''
''''''Call PressEnter() 
'''''''Capture the screenshot
''''''Call TakeScreenShot()
''''''
''''''Call ClickButton("Back   \(F3\)",False)
'''''''Capture the screenshot
''''''Call TakeScreenShot()
''''''
''''''Call ClickButton("Save   \(Ctrl\+S\)",False)
''''''Wait(1)
'''''''Capture the screenshot
''''''Call TakeScreenShot()
'''''''check if data saved successfully
''''''Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE_OCC1)
''''''VerifyStatusBar(DT_AS01_0500_CHECK_TEXT_OF_STATUSBAR)
''''''
''''''Call ClickButton("Execute settlement   \(Shift\+F8\)",False)
'''''''Capture the screenshot
''''''Call TakeScreenShot()
''''''
''''''Call SetTextbox("Asset Val. Date","ANEP-BZDAT","",Replace(DT_AS01_0100_ASSET_VAL_DATE,"/","."),False)
''''''Call SelectCheckbox("LKO74-TESTLAUF",0,DT_AS01_0100_TEST_RUN,False) 'test run checkbox off
'''''''Capture the screenshot
''''''Call TakeScreenShot()
''''''
''''''Call ClickButton("Simulate   \(F9\)",False)
'''''''Capture the screenshot
''''''Call TakeScreenShot()
''''''
''''''Call ClickButton("Back   \(F3\)",False)
'''''''Capture the screenshot
''''''Call TakeScreenShot()
''''''
''''''Call ClickButton("Execute   \(F8\)",False)
''''''Wait(1)
'''''''Capture the screenshot
''''''Call TakeScreenShot()
''''''
'''''''Validate If operation is completed
''''''Call GetStatusBar("item1","DT_AS01_0500_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
''''''Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''''''VerifyStatusBar(DT_AS01_0500_CHECK_TEXT_OF_STATUSBAR_OCC1)
''''''
'''''''
'''''''''----------------------Tcode FB03----------------------------
'''''''
'''''''Enter the Tcode
''''''Call SetTcode(DT_AS01_0500_OKCD) 
''''''Call PressEnter()     ' 
''''''Call CheckTCodeScreen(DT_AS01_0500_OKCD)
'''''''Capture the screenshot
''''''Call TakeScreenShot()
''''''
''''''Call SetTextbox("Document Number","RF05L-BELNR","",DT_AS01_0100_DOCUMENT_NUMBER,False)
''''''Call SetTextbox("Fiscal Year","RF05L-GJAHR","",DT_FISCAL_YEAR,False)
''''''''''following field is not required as it is populated automatically in sap screen
'''''''Call SetTextbox("Company Code","RF05L-BUKRS","",DATA,False)
''''''
''''''Call PressEnter()
'''''''Capture the screenshot
''''''Call TakeScreenShot()
''''''
'''''''Validate text boxes
''''''Call VerifyTextBoxContent("Document Number","BKPF-BELNR","",DT_AS01_0750_CHECK_TEXT_OF_DOCUMENT_NUMBER,False)
''''''Call VerifyTextBoxContent("Company Code","BKPF-BUKRS","",DT_AS01_0750_CHECK_TEXT_OF_COMPANY_CODE,False)
'''''''validate grid components
''''''call VerifyGridCellContent("",1,"Posting Key","",DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
''''''call VerifyGridCellContent("",2,"Posting Key","",DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
''''''call VerifyGridCellContent("",1,"Account","",DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
''''''call VerifyGridCellContent("",2,"Account","",DT_AS01_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
''''''Call TakeScreenShot()
''''''
''''''

