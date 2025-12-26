

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04.01.05.01.01 Request for Equipment Repair - Create _ Change
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

gstrTestCaseName = "Test_04.01.05.01.01 Request for Equipment Repair - Create _ Change"
'gstrTestCaseName = "Test_04.01.05.01.01_Request_for_Equipment_Repair_-_Create___Change"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\P1_FICO\DT_04.01.05.01.01 Request for Equipment Repair - Create _ Change_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)
'DataRowSet =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
'Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
''''

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'
'''--------------------------------------------  IW21------------------------------------------------
'
Call SetTextbox("Notification type","RIWO00-QMART","",DT_IW21_0100_NOTIFICATION_TYPE,False)
Call TakeScreenShot()
Call PressEnter() 
Call TakeScreenShot()
Call SetTextbox("Description","RIWO00-HEADKTXT","",DT_IW21_1050_NOTIFICATION,False)
Call SetTextbox("Functional loc\.","RIWO1-TPLNR","",DT_IW21_0100_FUNCTIONAL_LOC,False)
Call TakeScreenShot()
Call SetTextbox("Planner group","VIQMEL-INGRP","",DT_IW21_7326_PLANNER_GROUP,False)
Call SetTextboxNoLabel("VIQMEL-IWERK",0,DT_IW21_7326_VIQMELIWERK,False)
Call SetTextbox("Main WorkCtr","RIWO00-GEWRK","",DT_IW21_7326_MAIN_WORKCTR,False)
Call SetTextboxNoLabel("RIWO00-SWERK",0,DT_IW21_7326_RIWO00SWERK,False)
Call SetTextbox("Reported by","VIQMEL-QMNAM","",DT_IW21_7326_REPORTED_BY,False)
Call SetTextbox("Notif\.date","VIQMEL-QMDAT","",ConvertDate(DT_IW21_7326_NOTIFDATE),False)
Call SetTextbox("Notif\.date","VIQMEL-MZEIT","",DT_IW21_7326_NOTIFDATE_OCC1,False)
Call TakeScreenShot()

Call SetTableData("SAPLIQS0TEXT","#1",1,"","",DT_IW21_7710_TABLECELL__0,False)
Call SetTextbox("Object part","VIQMFE-OTGRP","",DT_IW21_7324_OBJECT_PART,False)
Call SetTextboxNoLabel("VIQMFE-OTEIL",0,DT_IW21_7324_VIQMFEOTEIL,False)
Call SetTextbox("Defect type","VIQMFE-FEGRP","",DT_IW21_7324_DEFECT_TYPE,False)
Call SetTextboxNoLabel("VIQMFE-FECOD",0,DT_IW21_7324_VIQMFEFECOD,False)
Call SetTextbox("Text","VIQMFE-FETXT","",DT_IW21_7324_TEXT,False)
Call SetTextboxNoLabel("VIQMUR-URGRP",0,DT_IW21_7324_CAUSE_CODE,False)


Call SetTextboxNoLabel("VIQMUR-URCOD",0,DT_IW21_7324_VIQMURURCOD,False)
Call SetTextbox("Cause text","VIQMUR-URTXT","",DT_IW21_7324_CAUSE_TEXT,False)
Call TakeScreenShot()

Call SetTextbox("Required Start","VIQMEL-STRMN","",ConvertDate(DT_IW21_7328_MALFUNCT_START),False)
Call SetTextbox("Required Start","VIQMEL-STRUR","",DT_IW21_7328_MALFUNCT_START_OCC1,False)
Call SetTextbox("Required End","VIQMEL-LTRMN","",ConvertDate(DT_IW21_7328_MALFUNCTEND),False)
Call SetTextbox("Required End","VIQMEL-LTRUR","",DT_IW21_7328_MALFUNCTEND_OCC1,False)
Call TakeScreenShot()

Call SetTextboxNoLabel("VIQMFE-OTEIL",0,DT_IW21_7324_BUILDING,False)
Call SetTextboxNoLabel("VIQMFE-FECOD",0,DT_IW21_7324_PERFORMANCE,False)
Call SetTextboxNoLabel("VIQMUR-URCOD",0,DT_IW21_7324_NATURAL_DISTORTION,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
Call PressEnter()
Call SetTextbox("Planner group","VIQMEL-INGRP","",DT_IW21_7326_PLANNER_GROUP_OCC1,False)
Call TakeScreenShot()

Call SetTextboxNoLabel("VIQMFE-OTEIL",0,DT_IW21_7324_PLUMBING,False)
Call SetTextboxNoLabel("VIQMFE-FECOD",0,DT_IW21_7324_DESIGN__STUDY,False)
Call SetTextbox("Defect type","VIQMFE-FEGRP","",DT_IW21_7324_DEFECT_TYPE_OCC1,False)
Call SetTextboxNoLabel("VIQMUR-URCOD",0,DT_IW21_7324_MISUSE,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
 Call VerifyTextBoxContent("Object part","RIWO00-TXTCDOT",0,UCASE(DT_IW21_7324_CHECK_TEXT_OF_OBJECT_PART),False)
 Call VerifyTextBoxContent("Defect type","RIWO00-TXTCDFE",0,UCASE(DT_IW21_7324_CHECK_TEXT_OF_DEFECT_TYPE),False)
 Call VerifyTextBoxNoLabelContent("RIWO00-TXTCDUR",0,DT_IW21_7324_CHECK_TEXT_OF_CAUSE_CODE,False)
 Call SetTextbox("Planner group","VIQMEL-INGRP","",DT_IW21_7326_PLANNER_GROUP_OCC2,False)
 Call TakeScreenShot()

Call SetTextboxNoLabel("VIQMFE-OTEIL",0,DT_IW21_7324_BUILDING_OCC1,False)
Call SetTextboxNoLabel("VIQMFE-FECOD",0,DT_IW21_7324_MECHANICAL,False)
Call SetTextbox("Defect type","VIQMFE-FEGRP","",DT_IW21_7324_DEFECT_TYPE_OCC2,False)
Call TakeScreenShot()
Call PressEnter()

 Call VerifyTextBoxContent("Object part","RIWO00-TXTCDOT",0,UCASE(DT_IW21_7324_CHECK_TEXT_OF_OBJECT_PART_OCC1),False)
 Call VerifyTextBoxContent("Defect type","RIWO00-TXTCDFE",0,UCASE(DT_IW21_7324_CHECK_TEXT_OF_DEFECT_TYPE_OCC1),False)
 Call VerifyTextBoxNoLabelContent("RIWO00-TXTCDUR",0,DT_IW21_7324_CHECK_TEXT_OF_CAUSE_CODE_OCC1,False)
 Call VerifyTextBoxNoLabelContent("RIWO00-INNAM",0,DT_IW21_7326_CHECK_TEXT_OF_RIWO00INNAM,False)

Call SetTextbox("Equipment","RIWO1-EQUNR","",DT_IW21_0100_EQUIPMENT,False)
Call TakeScreenShot()

Call SetTextbox("Description","RIWO00-HEADKTXT","",DT_IW21_7710_DESCRIPTION_OCC1,False)
Call TakeScreenShot()
Call PressEnter()
Call TakeScreenShot()
Call ClickButtonIfExist("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call ClickButtonIfExist("Delete   \(F5\)",True)
Call ClickButtonIfExist("Delete   \(F5\)",True)
Call ClickButtonIfExist("Delete   \(F5\)",True)
Call ClickButtonIfExist("Delete   \(F5\)",True)

Call SetTextbox("Main WorkCtr","RIWO00-GEWRK","",DT_IW21_7326_MAIN_WORKCTR_OCC1,False)
Call SetTextboxNoLabel("RIWO00-SWERK",0,DT_IW21_7326_RIWO00SWERK_OCC1,False)
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call GetStatusBar("item1","DT_IW21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call VerifyStatusBar("Notification "& DT_IW21_0100_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT &" saved" )

'''--------------------------------------------  IW23------------------------------------------------
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
Call SetTcode(DT_IW21_0100_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_IW21_0100_OKCD)

 Call PressEnter()     ' 
 Call VerifyTextBoxContent("Notification","VIQMEL-QMNUM",0,DT_IW21_1050_CHECK_TEXT_OF_NOTIFICATION,False)
 Call VerifyTextBoxContent("Notification","VIQMEL-QMTXT",0,UCASE(DT_IW21_1050_CHECK_TEXT_OF_NOTIFICATION_OCC1),False)
 Call VerifyTextBoxContent("Description","RIWO00-HEADKTXT",0,UCASE(DT_IW21_7710_CHECK_TEXT_OF_DESCRIPTION),False)
 Call VerifyTableCellContent(1,"#1","SAPLIQS0TEXT",LCASE(DT_IW21_7710_CHECK_TEXT_OF_TABLECELL__0))
 Call VerifyTextBoxContent("Planner group","VIQMEL-INGRP",0,UCASE(DT_IW21_7326_CHECK_TEXT_OF_PLANNER_GROUP),False)
 Call VerifyTextBoxNoLabelContent("VIQMEL-IWERK",0,UCASE(DT_IW21_7326_CHECK_TEXT_OF_VIQMELIWERK),False)
 Call VerifyTextBoxContent("Main WorkCtr","RIWO00-GEWRK",0,UCASE(DT_IW21_7326_CHECK_TEXT_OF_MAIN_WORKCTR),False)
 Call VerifyTextBoxNoLabelContent("RIWO00-SWERK",0,UCASE(DT_IW21_7326_CHECK_TEXT_OF_RIWO00SWERK),False)
 Call VerifyTextBoxContent("Reported by","VIQMEL-QMNAM",0,UCASE(DT_IW21_7326_CHECK_TEXT_OF_REPORTED_BY),False)
' Call VerifyTextBoxContent("Notif\.date","VIQMEL-QMDAT",0,ConvertDate(DT_IW21_7326_CHECK_TEXT_OF_NOTIFDATE),False)
' Call VerifyTextBoxContent("Notif\.date","VIQMEL-MZEIT",0,UCASE(DT_IW21_7326_CHECK_TEXT_OF_NOTIFDATE_OCC1),False) ''-- check

Call VerifyTextBoxContent("Object part","VIQMFE-OTGRP",0,UCASE(DT_IW21_7324_CHECK_TEXT_OF_OBJECT_PART_OCC2),False)
 Call VerifyTextBoxContent("Defect type","VIQMFE-FEGRP",0,UCASE(DT_IW21_7324_CHECK_TEXT_OF_DEFECT_TYPE_OCC2),False)
 Call VerifyTextBoxNoLabelContent("VIQMUR-URGRP",0,UCASE(DT_IW21_7324_CHECK_TEXT_OF_CAUSE_CODE_OCC2),False)
 
 Call VerifyTextBoxContent("Object part","RIWO00-TXTCDOT",0,UCASE(DT_IW21_7324_CHECK_TEXT_OF_OBJECT_PART_OCC3),False)
 Call VerifyTextBoxContent("Defect type","RIWO00-TXTCDFE",0,UCASE(DT_IW21_7324_CHECK_TEXT_OF_DEFECT_TYPE_OCC3),False)
 Call VerifyTextBoxNoLabelContent("RIWO00-TXTCDUR",0,DT_IW21_7324_CHECK_TEXT_OF_CAUSE_CODE_OCC3,False)
 
Call SelectTab("TAB_GROUP_10","Malfunction, breakdown",False)
Call TakeScreenShot()
Call SelectTab("TAB_GROUP_10","Location data",False)
Call TakeScreenShot()

 Call VerifyTextBoxContent("MaintSite","ILOA-SWERK",0,UCASE(DT_IW21_7000_CHECK_TEXT_OF_MAINTSITE),False)
 Call VerifyTextBoxContent("Sort field","ILOA-EQFNR",0,UCASE(DT_IW21_7000_CHECK_TEXT_OF_SORT_FIELD),False) 
 Call VerifyTextBoxContent("Company Code","ILOA-BUKRS",0,UCASE(DT_IW21_7000_CHECK_TEXT_OF_COMPANY_CODE),False)
 Call VerifyTextBoxContent("Business Area","ILOA-GSBER",0,UCASE(DT_IW21_7000_CHECK_TEXT_OF_BUSINESS_AREA),False) 
 Call VerifyTextBoxContent("Cost Center","ILOA-KOSTL",0,UCASE(DT_IW21_7000_CHECK_TEXT_OF_COST_CENTER),False)
 '
Call LogOff()
Call FinalStatus ()



