
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Split IDOCs and send to PI_P2_TASE
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


gstrTestCaseName = "Test_Split IDOCs and send to PI_P2"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
'
'
''--------------------WE19----------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call PressEnter()     
Call TakeScreenShot
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
   
Call TakeScreenShot
Call ClickButtonifexist("Continue   \(Enter\)",true)
call SelectRadioButton("MSED7START-SEL_IDOCTP","Basic Type",False)
call SetTextbox("Basic Type","MSED7START-IDOCTYP","",DT_WE19_0010_BASICTYP_OCC1,false) 
Call TakeScreenShot
Call PressEnter() 
call ClickButton("Create\.\.\.   \(F8\)",false)
Call TakeScreenShot()
Call SendKey("{F2}")
call ClickLabel("0000000000000000              2",0,false)
Call SendKey("{F2}")

call ClickButton("All Fields   \(F5\)",True) 
Call TakeScreenShot
call SetTextbox("Part\. Type","EDI_PRT-RCVPRT","",DT_WE19_0100_PART_TYPE,true) 
call SetTextbox("Partner No\.","G_CONTROL_RECORD-RCVPRN","",DT_WE19_0100_PARTNER_NO,true) 
call SetTextbox("Part\. Type","EDI_PRT-SNDPRT","",DT_WE19_0100_PART_TYPE_OCC1,true) 
call SetTextbox("Partner Role","G_CONTROL_RECORD-RCVPFC","","",true) 
call SetTextbox("Partner No\.","G_CONTROL_RECORD-SNDPRN","",DT_WE19_0100_PARTNER_NO_OCC1,true) 
call SetTextbox("Port","G_CONTROL_RECORD-SNDPOR","",DT_WE19_0100_PORT_OCC1,true) 
Call TakeScreenShot()
call SetTextbox("Port","G_CONTROL_RECORD-RCVPOR","",DT_WE19_0100_PORT,true) 
''call SetTextbox("Port","G_CONTROL_RECORD-SNDPOR","",DT_WE19_0100_PORT_OCC1,true) 
call SetTextbox("Partner Role","G_CONTROL_RECORD-SNDPFC","","",true) 
call SetTextbox("Message Variant","G_CONTROL_RECORD-MESCOD","",DT_WE19_0100_MESSAGE_VARIANT,true) 
call SetTextbox("Message Function","G_CONTROL_RECORD-MESFCT","",DT_WE19_0100_MESSAGE_FUNCTION,true) 
call SetTextbox("Message Type","EDI_MSGTYP-MSGTYP","",DT_WE19_0100_MESSAGE_TYPE,true) 
call ClickButton("All Fields   \(F5\)",True) 
Call TakeScreenShot
call ClickButton("Continue   \(Enter\)",True) 
Call TakeScreenShot()

CALL SetFocusGuiLabel("E1SHP_OBDLV_CONFIRM_DECENTR","","",FALSE)
'Call SendKey("{F2}")
Call ClickLabel_RelativeID("wnd[0]/usr/lbl[36,3]",False)
'Call SetTextboxNoLabel("SVALD-VALUE", , DT_WE19_0300_DELIVERY, True)

Call TakeScreenShot
call SetTextbox("DELIVERY","SVALD-VALUE","",DT_WE19_0300_DELIVERY,true) 
call ClickButton("Continue   \(Enter\)",true) 
'''SELECT A BRANCH
CALL ClickLabel("4",0,FALSE)
Call TakeScreenShot
CALL SetFocusGuiLabel("/SPE/E1BPOBDLVHDRCTRLCON","","",FALSE)
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
Call TakeScreenShot
CALL SetFocusGuiLabel("E1BPDLVITMSERNO","","",FALSE)
Call TakeScreenShot()
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
Call TakeScreenShot
CALL SetFocusGuiLabel("E1BPDLVHDUNSERNO","","",FALSE)
Call TakeScreenShot()
call ClickButton("Delete   \(Shift\+F2\)",false)
Call TakeScreenShot
CALL SetFocusGuiLabel("E1BPEXT","","",FALSE)
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
call ClickButton("Delete   \(Shift\+F2\)",false) 
CALL SetFocusGuiLabel("E1BPOBDLVHDRCON","","",FALSE)
Call SendKey("{F2}")
Call TakeScreenShot()
wait 5

Call TakeScreenShot
Call SetTextboxNoLabel("SVALD-VALUE",8 , DT_WE19_0400_BILLOFLADING, True)
Call SetTextboxNoLabel("SVALD-VALUE",0 , DT_WE19_0400_DELIV_NUMB_OCC1, True)

'call SetTextbox("BILLOFLADING","SVALD-VALUE","",DT_WE19_0400_BILLOFLADING,true) 
'call SetTextbox("DELIV_NUMB","SVALD-VALUE","",DT_WE19_0400_DELIV_NUMB_OCC1,true) 
call ClickButton("Continue   \(Enter\)",true) 
Call TakeScreenShot
Call SetFocusGuiLabel("E1BPOBDLVHDRCTRLCON","","",FALSE)
Call SendKey("{F2}")
wait 5
Call TakeScreenShot
call SetTextboxNoLabel("SVALD-VALUE",0,DT_WE19_0400_DELIV_NUMB_OCC1,true) 
'call SetTextbox("DELIV_NUMB","SVALD-VALUE","",DT_WE19_0400_DELIV_NUMB_OCC1,true) 
call ClickButton("Continue   \(Enter\)",true) 
Call TakeScreenShot
CALL SetFocusGuiLabel("E1BPDLVCONTROL","","",FALSE)
Call SendKey("{F2}")
Call TakeScreenShot()

call SetTextboxNoLabel("SVALD-VALUE",3,DT_WE19_0300_RECV_SYS,true) 
call SetTextboxNoLabel("SVALD-VALUE",4,DT_WE19_0300_DLV_TYPE,true) 

''call SetTextbox("RECV_SYS","SVALD-VALUE","",DT_WE19_0300_RECV_SYS,true) 
''call SetTextbox("DLV_TYPE","SVALD-VALUE","",DT_WE19_0300_DLV_TYPE,true) 
call ClickButton("Continue   \(Enter\)",true) 
Call TakeScreenShot
CALL SetFocusGuiLabel("E1BPDLVDEADLN","","",FALSE)
Call TakeScreenShot()
Call SendKey("{F2}")

call SetTextboxNoLabel("SVALD-VALUE","0",DT_WE19_0300_DELIV_NUMB,true) 
call SetTextboxNoLabel("SVALD-VALUE","1",DT_WE19_0300_TIMETYPE,true) 
call SetTextboxNoLabel("SVALD-VALUE","2",DT_WE19_0300_TIMESTAMP_UTC,true) 
call SetTextboxNoLabel("SVALD-VALUE","3",DT_WE19_0300_TIMEZONE,true) 

'''call SetTextbox("DELIV_NUMB","SVALD-VALUE","",DT_WE19_0300_DELIV_NUMB,true) 
'''call SetTextbox("TIMETYPE","SVALD-VALUE","",DT_WE19_0300_TIMETYPE,true) 
'''call SetTextbox("TIMESTAMP_UTC","SVALD-VALUE","",DT_WE19_0300_TIMESTAMP_UTC,true) 
'''call SetTextbox("TIMEZONE","SVALD-VALUE","",DT_WE19_0300_TIMEZONE,true) 
call ClickButton("Continue   \(Enter\)",true) 

CALL SetFocusGuiLabel("E1BPOBDLVITEMCON","","",FALSE)
Call SendKey("{F2}")
wait 30
Call TakeScreenShot
call SetTextboxNoLabel("SVALD-VALUE","0",DT_WE19_0400_DELIV_NUMB_OCC2,true) 
call SetTextboxNoLabel("SVALD-VALUE","1",DT_WE19_0400_DELIV_ITEM,true) 
call SetTextboxNoLabel("SVALD-VALUE","31",DT_WE19_0400_DLV_QTY,true) 
call SetTextboxNoLabel("SVALD-VALUE","32",DT_WE19_0400_DLV_QTY_IMUNIT,true) 
call SetTextboxNoLabel("SVALD-VALUE","33",DT_WE19_0400_FACT_UNIT_NOM,true) 
call SetTextboxNoLabel("SVALD-VALUE","34",DT_WE19_0400_FACT_UNIT_DENOM,true) 
Call TakeScreenShot
'''call SetTextbox("DELIV_NUMB","SVALD-VALUE","",DT_WE19_0400_DELIV_NUMB_OCC2,true) 
'''call SetTextbox("DELIV_ITEM","SVALD-VALUE","",DT_WE19_0400_DELIV_ITEM,true) 
'''call SetTextbox("DLV_QTY","SVALD-VALUE","",DT_WE19_0400_DLV_QTY,true) 
'''call SetTextbox("DLV_QTY_IMUNIT","SVALD-VALUE","",DT_WE19_0400_DLV_QTY_IMUNIT,true) 
'''call SetTextbox("FACT_UNIT_NOM","SVALD-VALUE","",DT_WE19_0400_FACT_UNIT_NOM,true) 
'''call SetTextbox("FACT_UNIT_DENOM","SVALD-VALUE","",DT_WE19_0400_FACT_UNIT_DENOM,true) 
call ClickButton("Continue   \(Enter\)",true)
Call TakeScreenShot
CALL SetFocusGuiLabel("E1BPOBDLVITEMCTRLCON","","",FALSE)
Call TakeScreenShot()
Call SendKey("{F2}")
wait 30

call SetTextboxNoLabel("SVALD-VALUE","0",DT_WE19_0300_DELIV_NUMB_OCC1,true) 
call SetTextboxNoLabel("SVALD-VALUE","1",DT_WE19_0300_DELIV_ITEM,true) 
call SetTextboxNoLabel("SVALD-VALUE","2",DT_WE19_0300_CHG_DELQTY,true) 


Call TakeScreenShot
'call SetTextboxNoLabel("DELIV_NUMB","SVALD-VALUE","",DT_WE19_0300_DELIV_NUMB_OCC1,true) 
'call SetTextbox("DELIV_ITEM","SVALD-VALUE","",DT_WE19_0300_DELIV_ITEM,true) 
'call SetTextbox("CHG_DELQTY","SVALD-VALUE","",DT_WE19_0300_CHG_DELQTY,true) 
call ClickButton("Continue   \(Enter\)",true)
Call TakeScreenShot()

CALL SetFocusGuiLabel("E1BPDLVHDUNHDR","","",FALSE)
Call SendKey("{F2}")
Call TakeScreenShot
call SetTextboxNoLabel("SVALD-VALUE","12",DT_WE19_0400_HDL_UNIT_EXID,true) 
call SetTextboxNoLabel("SVALD-VALUE","23",DT_WE19_0400_HDL_UNIT_EXID_TY,true) 
call SetTextboxNoLabel("SVALD-VALUE","19",DT_WE19_0400_SHIP_MAT,true) 
call SetTextboxNoLabel("SVALD-VALUE","33",DT_WE19_0400_PLANT,true) 
call SetTextboxNoLabel("SVALD-VALUE","44",DT_WE19_0400_STGE_LOC,true)
'
''call SetTextbox("HDL_UNIT_EXID","SVALD-VALUE","",DT_WE19_0400_HDL_UNIT_EXID,true) 
''call SetTextbox("HDL_UNIT_EXID_TY","SVALD-VALUE","",DT_WE19_0400_HDL_UNIT_EXID_TY,true) 
''call SetTextbox("SHIP_MAT","SVALD-VALUE","",DT_WE19_0400_SHIP_MAT,true) 
''call SetTextbox("PLANT","SVALD-VALUE","",DT_WE19_0400_PLANT,true) 
''call SetTextbox("STGE_LOC","SVALD-VALUE","",DT_WE19_0400_STGE_LOC,true)

call ClickButtonIfExist("Continue   \(Enter\)",true)
Call TakeScreenShot
CALL SetFocusGuiLabel("E1BPDLVHDUNITM","","",FALSE)
Call TakeScreenShot()
Call SendKey("{F2}")
Call TakeScreenShot()

call SetTextboxNoLabel("SVALD-VALUE","12",DT_WE19_0400_HDL_UNIT_EXID_INTO,true) 
call SetTextboxNoLabel("SVALD-VALUE","14",DT_WE19_0400_DELIV_NUMB_OCC3,true) 
call SetTextboxNoLabel("SVALD-VALUE","15",DT_WE19_0400_DELIV_ITEM_OCC1,true) 
call SetTextboxNoLabel("SVALD-VALUE","16",DT_WE19_0400_PACK_QTY,true) 
call SetTextboxNoLabel("SVALD-VALUE","6",DT_WE19_0400_HU_ITEM_TYPE,true)

Call TakeScreenShot
''call SetTextbox("HDL_UNIT_EXID_INTO","SVALD-VALUE","",DT_WE19_0400_HDL_UNIT_EXID_INTO,true) 
''Call wait (2)
''call SetTextbox("DELIV_NUMB","SVALD-VALUE","",DT_WE19_0400_DELIV_NUMB_OCC3,true) 
''Call wait (2)
''call SetTextbox("DELIV_ITEM","SVALD-VALUE","",DT_WE19_0400_DELIV_ITEM_OCC1,true) 
''Call wait (2)
''call SetTextbox("PACK_QTY","SVALD-VALUE","",DT_WE19_0400_PACK_QTY,true) 
''Call wait (2)
''call SetTextbox("HU_ITEM_TYPE","SVALD-VALUE","",DT_WE19_0400_HU_ITEM_TYPE,true) 
call ClickButton("Continue   \(Enter\)",true)
Call TakeScreenShot
CALL SetFocusGuiLabel("E1BPEXTC","","",FALSE)
Call SendKey("{F2}")
Call TakeScreenShot()
call SetTextbox("FIELD1","SVALD-VALUE","",DT_WE19_0300_FIELD1,true) 
call ClickButton("Continue   \(Enter\)",true)
Call TakeScreenShot()
call ClickButton("Insert   \(Shift\+F6\)",false)
call ClickButton("Copy   \(Shift\+F5\)",false)
Call TakeScreenShot
CALL SetFocusGuiLabel("E1BPOBDLVITEMCON","","",FALSE)
call ClickButton("Copy   \(Shift\+F5\)",false)
call ClickButton("Insert   \(Shift\+F6\)",false)
call ClickButton("at the same level",false)

CALL SetFocusGuiLabel("E1BPOBDLVITEMCTRLCON","","",FALSE)
call ClickButton("Copy   \(Shift\+F5\)",false)
call ClickButton("Insert   \(Shift\+F6\)",false)
call ClickButton("at the same level",false)
Call TakeScreenShot
CALL SetFocusGuiLabel("E1BPDLVHDUNHDR","","",FALSE)
Call TakeScreenShot()
call ClickButton("Copy   \(Shift\+F5\)",false)
call ClickButton("Insert   \(Shift\+F6\)",false)
call ClickButton("at the same level",false)

CALL SetFocusGuiLabel("E1BPDLVHDUNITM","","",FALSE)
call ClickButton("Copy   \(Shift\+F5\)",false)
call ClickButton("Insert   \(Shift\+F6\)",false)
call ClickButton("at the same level",false)

CALL SetFocusGuiLabel("E1BPEXTC","","",FALSE)
call ClickButton("Copy   \(Shift\+F5\)",false)
call ClickButton("Insert   \(Shift\+F6\)",false)
call ClickButton("at the same level",false)
'''''''''''''''''''mention x and y cord'''''''''''''''''''''''''
CALL ClickLabel("E1BPOBDLVITEMCON",1,FALSE)
Call TakeScreenShot()
Call SendKey("{F2}")
wait 10
call SetTextboxNoLabel("SVALD-VALUE","1",DT_WE19_0400_DELIV_ITEM_OCC2,true) 
call SetTextboxNoLabel("SVALD-VALUE","31",DT_DLV_QTY_IMUNIT,true) 
call SetTextboxNoLabel("SVALD-VALUE","32",DT_WE19_0400_DLV_QTY_IMUNIT_OCC1,true) 
call SetTextboxNoLabel("SVALD-VALUE","33",DT_WE19_0400_FACT_UNIT_NOM_OCC1,true) 
call SetTextboxNoLabel("SVALD-VALUE","34",DT_FACT_UNIT_DENOM,true)

Call TakeScreenShot
''call SetTextbox("DELIV_ITEM","SVALD-VALUE","",DT_WE19_0400_DELIV_ITEM_OCC2,true) 
''call SetTextbox("DLV_QTY","SVALD-VALUE","",DT_DLV_QTY_IMUNIT,true) 
''call SetTextbox("DLV_QTY_IMUNIT","SVALD-VALUE","",DT_WE19_0400_DLV_QTY_IMUNIT_OCC1,true) 
''call SetTextbox("FACT_UNIT_NOM","SVALD-VALUE","",DT_WE19_0400_FACT_UNIT_NOM_OCC1,true) 
''call SetTextbox("FACT_UNIT_DENOM","SVALD-VALUE","",DT_FACT_UNIT_DENOM,true) 
call ClickButton("Continue   \(Enter\)",true)
Call TakeScreenShot
CALL ClickLabel("E1BPOBDLVITEMCTRLCON",1,FALSE)
Call TakeScreenShot()
Call SendKey("{F2}")
call SetTextbox("DELIV_ITEM","SVALD-VALUE","",DT_WE19_0300_DELIV_ITEM_OCC1,true) 
call ClickButton("Continue   \(Enter\)",true)

CALL ClickLabel("E1BPDLVHDUNHDR",1,FALSE)
Call TakeScreenShot()
Call SendKey("{F2}")
wait 10
call SetTextboxNoLabel("SVALD-VALUE","12",DT_WE19_0400_HDL_UNIT_EXID_OCC1,true) 

call ClickButton("Continue   \(Enter\)",true)
CALL ClickLabel("E1BPDLVHDUNITM",1,FALSE)
Call TakeScreenShot()
Call SendKey("{F2}")
call SetTextboxNoLabel("SVALD-VALUE","12",DT_WE19_0400_HDL_UNIT_EXID_INTO_OCC1,true) 
call SetTextboxNoLabel("SVALD-VALUE","15",DT_WE19_0400_DELIV_ITEM_OCC3,true) 
call ClickButton("Continue   \(Enter\)",true)
CALL ClickLabel("E1BPEXTC",1,FALSE)

Call TakeScreenShot()
Call SendKey("{F2}")
call SetTextboxNoLabel("SVALD-VALUE","0",DT_WE19_0300_FIELD1_OCC1,true) 
call ClickButton("Continue   \(Enter\)",true)
Call TakeScreenShot()
call ClickButton("Test Standard Inbound Processing   \(F8\)",false)
Call TakeScreenShot()
'' 	S_TL_G and getdata are not correct columnnames
'' Added datacolumnname on line 213 and 215 ' On 17Mar '
'' Author - KGARA
'Call VerifyTextBoxNoLabelContent("EDI_STAT03-ICONDSCRP",0,DT_WE19_0510_CHECK_TEXT_OF_EDI_STAT03ICONDSCRP,True)
call ClickButton("Continue   \(Enter\)",true)
Call TakeScreenShot()
Call GetTextBoxValue("MESSTXT1",0,DT_CHECK_TEXT_OF_MESSTXT1_OUTPUT,True)
Call TakeScreenShot()
call ClickButton("Continue   \(Enter\)",true)
Call TakeScreenShot()
'
Call LogOff()
Call FinalStatus ()




