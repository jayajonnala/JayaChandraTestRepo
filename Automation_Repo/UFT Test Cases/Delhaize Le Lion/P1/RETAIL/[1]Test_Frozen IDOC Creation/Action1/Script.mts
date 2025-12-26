
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_159_Update Account Completion Table (All Opcos)_TASE
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

gstrresultFolderPath =  ReadTxtFileResult(RunTimeResultFolder)



gstrTestCaseName = "Test_Frozen IDOC Creation"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''' Login '''
SAPGuiUtil.OpenConnection("R1E - SAP RETAIL Pre-Production EUROPE")
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  



''SAP Login'''
'
Call SetTcode(DT_SAPTRANSACTIONCODE)     ' - Line (12)
Call PressEnter()     ' - Line (13)

''INPUT''
call ClickButtonIfExist("Continue   \(Enter\)",True)
call SelectRadioButton("MSED7START-SEL_IDOCTP","Basic Type",False)
call SetTextbox("Basic Type","MSED7START-IDOCTYP","",DT_WE19_0010_BASICTYP_OCC1,false) 
Call PressEnter() 
Call TakeScreenShot()
call ClickButton("Create\.\.\.   \(F8\)",false)  '
Call TakeScreenShot()
call ClickLabel("4",0,false)
'Call SendKey("{F2}")
CALL SetFocusGuiLabel("E1EDL22","","",FALSE)
call ClickButton("Delete   \(Shift\+F2\)",false) 


CALL SetFocusGuiLabel("E1EDL21","","",FALSE)
call ClickButton("Expand Subtree   \(Ctrl\+Shift\+F11\)",false) 
'Call SendKey("{F2}")
CALL SetFocusGuiLabel("E1EDL23","","",FALSE)
call ClickButton("Delete   \(Shift\+F2\)",false) 

CALL SetFocusGuiLabel("E1EDL51","","",FALSE)
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
Call TakeScreenShot()
CALL SetFocusGuiLabel("E1EDL55","","",FALSE)
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 

call ClickLabel("4",0,false)
'Call SendKey("{F2}")
CALL SetFocusGuiLabel("E1EDT13","","",FALSE)
call ClickButton("Expand Subtree   \(Ctrl\+Shift\+F11\)",false) 
CALL SetFocusGuiLabel("E1EDL40","","",FALSE)
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 

CALL SetFocusGuiLabel("E1TXTH8","","",FALSE)
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
Call TakeScreenShot()
CALL SetFocusGuiLabel("E1EDL24","","",FALSE)
call ClickButton("Expand Subtree   \(Ctrl\+Shift\+F11\)",false) 
'call ClickLabel("4",0,false)
'Call SendKey("{F2}")
CALL SetFocusGuiLabel("E1EDL25","","",FALSE)
call ClickButton("Delete   \(Shift\+F2\)",false) 
CALL SetFocusGuiLabel("E1EDL52","","",FALSE)
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
'Call SendKey("{Shift+F2}")
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
Call TakeScreenShot()
CALL SetFocusGuiLabel("E1EDL26","","",FALSE)
call ClickButton("Expand Subtree   \(Ctrl\+Shift\+F11\)",false) 
'call ClickLabel("4",0,false)
'Call SendKey("{F2}")
CALL SetFocusGuiLabel("E1EDL27","","",FALSE)
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 

Call SetVerticalScrollBar(0,False)
CALL SetFocusGuiLabel("EDIDC","","",FALSE)
Call SendKey("{F2}")
wait 5
Call TakeScreenShot()
call SetTextbox("Part\. Type","EDI_PRT-RCVPRT","",DT_WE19_0100_PART_TYPE,true) 
call SetTextbox("Partner No\.","G_CONTROL_RECORD-RCVPRN","",DT_WE19_0100_PARTNER_NO,true) 
call SetTextbox("Part\. Type","EDI_PRT-SNDPRT","",DT_WE19_0100_PART_TYPE_OCC1,true) 
'''call SetTextbox("Partner Role","G_CONTROL_RECORD-RCVPFC","","",true) 
call SetTextbox("Partner No\.","G_CONTROL_RECORD-SNDPRN","",DT_WE19_0100_PARTNER_NO_OCC1,true) 
call SetTextbox("Port","G_CONTROL_RECORD-RCVPOR","",DT_WE19_0100_PORT,true) 
call SetTextbox("Port","G_CONTROL_RECORD-SNDPOR","",DT_WE19_0100_PORT_OCC1,true) 
''''call SetTextbox("Partner Role","G_CONTROL_RECORD-SNDPFC","","",true) 
call SetTextbox("Message Variant","G_CONTROL_RECORD-MESCOD","",DT_WE19_0100_MESSAGE_VARIANT,true) 
call SetTextbox("Message Function","G_CONTROL_RECORD-MESFCT","",DT_WE19_0100_MESSAGE_FUNCTION,true) 
call SetTextbox("Message Type","EDI_MSGTYP-MSGTYP","",DT_WE19_0100_MESSAGE_TYPE,true) 
Call TakeScreenShot()
'''''call ClickButton("All Fields   \(F5\)",True) 
call ClickButton("Continue   \(Enter\)",True) 
'
CALL SetFocusGuiLabel("E1EDL20","","",FALSE)
Call SendKey("{F2}")
wait 5

Call SetTextboxNoLabel("SVALD-VALUE", 0, DT_WE19_0400_VBELN, False)
''call SetTextbox("VBELN","SVALD-VALUE",0,DT_WE19_0400_VBELN,true) 

'''call SetTextbox("VBELN","SVALD-VALUE","",DT_WE19_0400_VBELN,true) 
'''call SetTextbox("LIFEX","SVALD-VALUE","",DT_WE19_0400_LIFEX,true)

call SetTextboxNoLabel("SVALD-VALUE","14",DT_WE19_0400_LIFEX,true)
''call SetTextbox("LIFEX","SVALD-VALUE","14",DT_WE19_0400_LIFEX,true)

call ClickButton("Continue   \(Enter\)",True) 
Call TakeScreenShot()
CALL SetFocusGuiLabel("E1EDL18","","",FALSE)
Call SendKey("{F2}")
wait 5
call SetTextbox("Qualifier","SVALD-VALUE","",DT_WE19_0300_QUALIFIER,true) 
call ClickButton("Continue   \(Enter\)",True) 
call ClickButton("Copy   \(Shift\+F5\)",false)
call ClickButton("Insert   \(Shift\+F6\)",false)
call ClickButton("at the same level",false)
CALL ClickLabel("E1EDL18",1,FALSE)
Call TakeScreenShot()
call SetTextbox("Qualifier","SVALD-VALUE","",DT_WE19_0300_QUALIFIER_OCC1,true) 
call ClickButton("Continue   \(Enter\)",True)  

CALL SetFocusGuiLabel("E1EDT13","","",FALSE)
Call SendKey("{F2}")
wait 5

call SetTextboxNoLabel("SVALD-VALUE",0,DT_WE19_0400_QUALIFIER,true) 
''call SetTextbox("Qualifier","SVALD-VALUE","",DT_WE19_0400_QUALIFIER,true) 
call SetTextboxNoLabel("SVALD-VALUE",11,DT_WE19_0400_DATE,true) 
call ClickButton("Continue   \(Enter\)",True) 
Call TakeScreenShot()
CALL SetFocusGuiLabel("E1EDL24","","",FALSE)
call ClickButton("Copy   \(Shift\+F5\)",false)
call ClickButton("Insert   \(Shift\+F6\)",false)
call ClickButton("at the same level",false)
Call TakeScreenShot()
CALL ClickLabel("E1EDL24",0,FALSE)
Call TakeScreenShot()
wait 5
call SetTextboxNoLabel("SVALD-VALUE","0",DT_WE19_0400_POSNR,true)
call SetTextboxNoLabel("SVALD-VALUE","1",DT_WE19_0400_MATNR,true) 
call SetTextboxNoLabel("SVALD-VALUE","4",DT_WE19_0400_LFIMG,true) 
call SetTextboxNoLabel("SVALD-VALUE","5",DT_WE19_0400_VRKME,true) 
call SetTextboxNoLabel("SVALD-VALUE","54",DT_WE19_0400_VFDAT,true)

''''call SetTextbox("POSNR","SVALD-VALUE","0",DT_WE19_0400_POSNR,true)
''''call SetTextbox("MATNR","SVALD-VALUE","1",DT_WE19_0400_MATNR,true) 
''''call SetTextbox("LFIMG","SVALD-VALUE","4",DT_WE19_0400_LFIMG,true) 
''''call SetTextbox("VRKME","SVALD-VALUE","5",DT_WE19_0400_VRKME,true) 
''''call SetTextbox("VFDAT","SVALD-VALUE","54",DT_WE19_0400_VFDAT,true)
'''
''''call SetTextboxNoName("POSNR","",DT_WE19_0400_POSNR,true)
''''call SetTextboxNoName("MATNR","",DT_WE19_0400_MATNR,true) 
''''call SetTextboxNoName("LFIMG","",DT_WE19_0400_LFIMG,true) 
''''call SetTextboxNoName("VRKME","",DT_WE19_0400_VRKME,true) 
''''call SetTextboxNoName("VFDAT","",DT_WE19_0400_VFDAT,true)

''call SetTextbox("POSNR","SVALD-VALUE","",DT_WE19_0400_POSNR,true)
''call SetTextbox("MATNR","SVALD-VALUE","",DT_WE19_0400_MATNR,true) 
''call SetTextbox("LFIMG","SVALD-VALUE","",DT_WE19_0400_LFIMG,true) 
'''call SetTextbox("VRKME","SVALD-VALUE","",DT_WE19_0400_VRKME,true) 
'''call SetTextbox("VFDAT","SVALD-VALUE","",DT_WE19_0400_VFDAT,true)
Call TakeScreenShot()
call ClickButton("Continue   \(Enter\)",True) 

CALL ClickLabel("E1EDL26",0,FALSE)
call SetTextboxNoLabel("SVALD-VALUE",10,DT_WE19_0400_KZUMW,true) 
'''call SetTextbox("KZUMW","SVALD-VALUE","",DT_WE19_0400_KZUMW,true) 
Call TakeScreenShot()
call ClickButton("Continue   \(Enter\)",True) 

CALL ClickLabel("E1EDL24",1,FALSE)
wait 5
call SetTextboxNoLabel("SVALD-VALUE","0",DT_WE19_0400_POSNR_OCC1,true)
call SetTextboxNoLabel("SVALD-VALUE","1",DT_WE19_0400_MATNR_OCC1,true) 
call SetTextboxNoLabel("SVALD-VALUE","4",DT_WE19_0400_LFIMG_OCC1,true) 
call SetTextboxNoLabel("SVALD-VALUE","5",DT_WE19_0400_VRKME_OCC1,true) 
call SetTextboxNoLabel("SVALD-VALUE","54",DT_WE19_0400_VFDAT_OCC1,true)

'''call SetTextbox("POSNR","SVALD-VALUE","",DT_WE19_0400_POSNR_OCC1,true) 
'''call SetTextbox("MATNR","SVALD-VALUE","",DT_WE19_0400_MATNR_OCC1,true) 
'''call SetTextbox("LFIMG","SVALD-VALUE","",DT_WE19_0400_LFIMG_OCC1,true) 
'''call SetTextbox("VRKME","SVALD-VALUE","",DT_WE19_0400_VRKME_OCC1,true) 
'''call SetTextbox("VFDAT","SVALD-VALUE","",DT_WE19_0400_VFDAT_OCC1,true)
Call TakeScreenShot()
call ClickButton("Continue   \(Enter\)",True) 
wait 5

CALL ClickLabel("E1EDL26",1,FALSE)
call SetTextbox("KZUMW","SVALD-VALUE","",DT_WE19_0400_KZUMW_OCC1,true) 
call ClickButton("Continue   \(Enter\)",True) 

call ClickButton("Test Standard Inbound Processing   \(F8\)",false) 
Call TakeScreenShot()
call ClickButton("Continue   \(Enter\)",TRUE) 
Call TakeScreenShot()
call GetTextboxValue("MESSTXT1",0,"DT_IDOC_MESSAGE_OUTPUT",true)
Call TakeScreenShot()
Call WriteRunTimeDataToExcelGlobalSheet ("DT_IDOC_MESSAGE_OUTPUT",DT_IDOC_MESSAGE)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
''call VerifyTextBoxContent("Information Message","MESSTXT1",1,DT_WE19_0010_CHECK_TEXT_OF_MESSTXT1,true)
'call VerifyTextBoxContent("Information Message","MESSTXT2",2,DT_WE19_0010_CHECK_TEXT_OF_MESSTXT2,true)
call ClickButton("Continue   \(Enter\)",TRUE) 
Call SetTcode(DT_WE19_0120_OKCD)     
Call PressEnter() 
Call TakeScreenShot()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)
call SetTextbox("IDoc number","DOCNUM-LOW","",DT_IDOC_NO,False)
Call TakeScreenShot()
call ClickButton("Execute   \(F8\)",False) 
Call TakeScreenShot()
Call LogOff()
Call FinalStatus ()





