

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_03.03.01.01.02 Manage Equipment _Vehicle - Create Sub-Equipment
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

gstrTestCaseName = "Test_03.03.01.01.02 Manage Equipment _Vehicle - Create Sub-Equipment"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_01PRI00_013_ENA_prices_are_not_higher_than_AB_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''
''----------------------Tcode IE01----------------------------

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


'Enter Details
Call SetTextbox("Equipment category","RM63E-EQTYP","",DT_IE01_0100_EQUIPMENT_CATEGORY,False) 
Call TakeScreenShot()
Call PressEnter() 
Call SetTextbox("Description","ITOB-SHTXT","",DT_IE01_1525_DESCRIPTION,False) 
Call TakeScreenShot()

'Click on Status Button
Call ClickButton("%_AUTOTEXT003",False)
Call SetTableData("SAPLBSVATC_E","X",2,"","","ON",False)
Call TakeScreenShot()

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

'Enter Details
Call SetTextbox("Description","ITOB-SHTXT","",DT_IE01_1525_DESCRIPTION,False)
Call SetTextbox("Size/dimension","ITOB-GROES","",DT_IE01_1020_SIZEDIMENSION,False)
Call SetTextbox("Object type","ITOB-EQART","",DT_IE01_1025_OBJECT_TYPE,False)
Call SetTextbox("Model number","ITOB-TYPBZ","",DT_IE01_1022_MODEL_NUMBER,False)
Call SetTextbox("ManufSerialNo\.","ITOB-SERGE","",DT_IE01_1022_MANUFSERIALNO,False)
Call TakeScreenShot()
Call PressEnter()


'Navigate to Location Tab
Call SelectTab("TABSTRIP","Location",False)
Wait(1)
Call TakeScreenShot()

'Enter Details
Call SetTextbox("MaintSite","ITOB-SWERK","",DT_IE01_1050_MAINTSITE,False)
Call SetTextbox("Work center","ITOBATTR-ARBPL","",DT_IE01_1050_WORK_CENTER,False)
Call TakeScreenShot()
Call PressEnter()

'Navigate to Organization Tab
Call SelectTab("TABSTRIP","Organization",False)
Wait(1)
Call TakeScreenShot()

'Enter Details
Call SetTextbox("Planner group","ITOB-INGRP","",DT_IE01_1062_PLANNER_GROUP,False)
Call SetTextbox("Main WorkCtr","ITOBATTR-GEWRK","",DT_IE01_1062_MAIN_WORKCTR,False)
Call TakeScreenShot()
Call PressEnter()

Call VerifyStatusBarMessageType("W")

Call PressEnter()


'Navigate to Structure Tab
Call SelectTab("TABSTRIP","Structure",False)
Wait(1)
Call TakeScreenShot()

'Click on Change InstLoc
Call ClickButton("Change InstLoc",False)

Call SetTextbox("Superord\.Equip\.","IEQINSTALL-HEQNR","",DT_IE01_0100_SUPERORDEQUIP,True)
Call TakeScreenShot()

'Click on Change InstLoc
Call ClickButton("Confirm   \(Shift\+F4\)",True)

'Enter Details
Call SetTextbox("TechIdentNo\.","ITOB-TIDNR","",DT_IE01_1065_TECHIDENTNO,False)
Call TakeScreenShot()
Call PressEnter()


'Navigate to Warranties Tab
Call SelectTab("TABSTRIP","Warranties",False)
Wait(1)
Call TakeScreenShot()


'Enter Details
Call SetTextbox("Warranty Start","WCHECK_V_H-GWLDT_I","1",DT_IE01_3400_WARRANTY_START,False)
Call SetTextbox("Warranty end","WCHECK_V_H-GWLEN_I","1",DT_IE01_3400_WARRANTY_END,False)
Call TakeScreenShot()
Call PressEnter()

'Click on Class overview
Call ClickButton("Class overview   \(Shift\+F8\)",False)
Call SetTableData("SAPLCLFMTC_OBJ_CLASS","Class","1","","",DT_IE01_1600_TABLECELL_CLASS_0,False) 
Call TakeScreenShot()
Call PressEnter()

Call ClickButtonIfExist("Back   \(F3\)",False)
wait(2)

'Navigate to Classification Tab
Call SelectTab("TABSTRIP","Classification",False)
Wait(1)
Call TakeScreenShot()

Call SetTextbox("SERIAL NUMBER","RCTMS-MWERT","1",DT_IE01_4000_RCTMSMWERT,False)
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_IE01_0100_SUBEQUIPMENT_OUTPUT")
Call VerifyStatusBar("Equipment created with the number "&DT_IE01_0100_SUBEQUIPMENT_OUTPUT)

'------------------------'Log Off From Applicaton--------------------------------

Call LogOff()
Call FinalStatus ()

'*********************************************End Of Script*********************************************************************
