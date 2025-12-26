
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Split IDOCs and send to PI
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


gstrTestCaseName = "Test_Split IDOCs and send to PI"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
SAPGuiUtil.OpenConnection("R1E - SAP RETAIL Pre-Production EUROPE")
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call TakeScreenshot()
Call PressEnter()  

'------------------------------------------------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)     
Call TakeScreenshot()
Call PressEnter()    

'''outbound delivery
Call ClickButtonifexist("Continue   \(Enter\)",true)
call SelectRadioButton("MSED7START-SEL_MSGTYP","Using message type",False)
call SetTextbox("Using message type","MSED7START-MESSAGETYP","",DT_WE19_0010_VIA_MESSAGE_TYPE_OCC1,false) 
Call TakeScreenshot()
Call PressEnter() 
call ClickButton("Create\.\.\.   \(F8\)",false)  '
Call TakeScreenshot()
' ClickLabel(labelContent, labelIndex, blnIsItPopup)
call ClickLabel("0000000000000000              2",0,false)
call SetTextbox("Part\. Type","EDI_PRT-RCVPRT","",DT_WE19_0100_PART_TYPE,true) 
call SetTextbox("Partner No\.","G_CONTROL_RECORD-RCVPRN","",DT_WE19_0100_PARTNER_NO,true) 
call SetTextbox("Part\. Type","EDI_PRT-SNDPRT","",DT_WE19_0100_PART_TYPE_OCC1,true) 
call SetTextbox("Partner Role","G_CONTROL_RECORD-RCVPFC","","",true) 
call SetTextbox("Partner No\.","G_CONTROL_RECORD-SNDPRN","",DT_WE19_0100_PARTNER_NO_OCC1,true) 
call SetTextbox("Port","G_CONTROL_RECORD-RCVPOR","",DT_WE19_0100_PORT,true) 
call SetTextbox("Port","G_CONTROL_RECORD-SNDPOR","",DT_WE19_0100_PORT_OCC1,true) 
call SetTextbox("Partner Role","G_CONTROL_RECORD-SNDPFC","","",true) 
call SetTextbox("Message Variant","G_CONTROL_RECORD-MESCOD","",DT_WE19_0100_MESSAGE_VARIANT,true) 
call SetTextbox("Message Function","G_CONTROL_RECORD-MESFCT","",DT_WE19_0100_MESSAGE_FUNCTION,true) 
Call TakeScreenshot()
call ClickButton("Continue   \(Enter\)",false) 

Call TakeScreenshot()
call ClickLabel("4",0,false)

'call ClickLabel_NolabelContent(17,false)'AT_BE affiliate - Dry
 Call ClickLabel_RelativeID("wnd[0]/usr/lbl[36,3]", False)
 
call SetTextbox("DELIVERY","SVALD-VALUE","",DT_WE19_0300_DELIVERY,true)
call SetTextbox("DELIVERYNEW","SVALD-VALUE","",DT_WE19_0300_DELIVERYNEW,true) 
call SetTextbox("DELIVERYITEMSBELONGT","SVALD-VALUE","",DT_WE19_0300_DELIVERYITEMSBELONGT,true) 
call SetTextbox("BUSINESSPROCESS","SVALD-VALUE","",DT_WE19_0300_BUSINESSPROCESS,true) 
Call TakeScreenshot()
call ClickButton("Continue   \(Enter\)",false) 
'''SELECT A BRANCH
CALL SetFocusGuiLabel("E1BPPAREX","","",FALSE)
Call TakeScreenshot()
call ClickButton("Delete   \(Shift\+F2\)",false) 
Call TakeScreenshot()
''' ClickSAPUILabel(creationTime, windowText, htmltag, innertext, className, text, index, blnIsItPopup)
'''ClickSAPUILabel
''' ClickLabel(labelContent, labelIndex, blnIsItPopup)
''
'''CALL ClickLabel("4",0,FALSE)
''
'''CALL ClickLabel("0000000000000000              2",0,FALSE)
Call TakeScreenshot()

''''''click blank label
''''Updated as 16 due to label index change
'Call ClickLabel_NolabelContent(16,FALSE) 'AT_BE interco - Dry,'AT_BE affiliate - Dry
'CALL ClickLabel_NolabelContent(17,FALSE)
Call ClickLabel_RelativeID("wnd[0]/usr/lbl[40,5]", False)

call SetTextbox("RECV_SYS","SVALD-VALUE","",DT_WE19_0300_RECV_SYS,true) 
call SetTextbox("DLV_TYPE","SVALD-VALUE","",DT_WE19_0300_DLV_TYPE,true) 
call ClickButton("Continue   \(Enter\)",false) 

Call TakeScreenshot()
'CALL ClickLabel_NolabelContent(15,FALSE)'AT_BE affiliate - Dry
Call ClickLabel_RelativeID("wnd[0]/usr/lbl[40,6]", False)
Call TakeScreenshot()
wait 5
call SetTextbox("DELIV_NUMB","SVALD-VALUE","",DT_WE19_0400_DELIV_NUMB,true) 
call SetTextbox("DELIV_ITEM","SVALD-VALUE","",DT_WE19_0400_DELIV_ITEM,true) 
call SetTextbox("CURR_QTY","SVALD-VALUE","",DT_WE19_0400_CURR_QTY,true) 
call SetTextbox("DLV_QTY_STOCK","SVALD-VALUE","",DT_WE19_0400_DLV_QTY_STOCK,true) 
call SetTextbox("DLV_QTY","SVALD-VALUE","",DT_WE19_0400_DLV_QTY,true) 
call SetTextbox("SALES_QTY_NUM","SVALD-VALUE","",DT_WE19_0400_SALES_QTY_NUM,true) 
call SetTextbox("SALES_QTY_DENOM","SVALD-VALUE","",DT_WE19_0400_SALES_QTY_DENOM,true) 
Call TakeScreenshot()
call ClickButton("Continue   \(Enter\)",false) 

' SetFocusGuiLabel(labelContent, xCord, yCord, blnIsItPopup)
'SetFocusGuiLabel

CALL SetFocusGuiLabel("E1BPOBDLVITEMDSP","","",FALSE)
Call TakeScreenshot()
call ClickButton("Copy   \(Shift\+F5\)",false)
call ClickButton("Insert   \(Shift\+F6\)",false)
call ClickButton("at the same level",false)

Call TakeScreenshot()
'CALL ClickLabel_NolabelContent(18,FALSE)'AT_BE interco - Dry
'CALL ClickLabel_NolabelContent(14,FALSE)'AT_BE affiliate - Dry
wait 5
Call ClickLabel_RelativeID("wnd[0]/usr/lbl[40,7]", False)


'CALL ClickLabel("3246942071000010B4                                            96","",FALSE)
call SetTextbox("DELIV_ITEM","SVALD-VALUE","",DT_WE19_0400_DELIV_ITEM_OCC1,true) 
call SetTextbox("DLV_QTY","SVALD-VALUE","",DT_DEL_QTY,true) 
call SetTextbox("DLV_QTY_STOCK","SVALD-VALUE","",DT_WE19_0400_DLV_QTY_STOCK_OCC1,true) 
call SetTextbox("SALES_QTY_NUM","SVALD-VALUE","",DT_WE19_0400_SALES_QTY_NUM_OCC1,true) 
call SetTextbox("SALES_QTY_DENOM","SVALD-VALUE","",DT_SALES_QTY_DENOM,true) 
Call TakeScreenshot()
call ClickButton("Continue   \(Enter\)",TRUE) 

call ClickButton("Test Standard Inbound Processing   \(F8\)",false) 
Call TakeScreenshot()
''' VerifyTextBoxContent(textboxAttachedText, textboxName, textboxIndex, expectedValue, blnIsItPopup)
'''Call VerifyTextBoxContent("","EDI_STAT03-ICON","",DT_WE19_0510_CHECK_ICONNAME_OF_EDI_STAT03ICON,True)
'''' VerifyTextBoxNoLabelContent(textboxName, textboxIndex, expectedValue, blnIsItPopup)
'''Call VerifyTextBoxNoLabelContent("EDI_STAT03-ICON","",DT_WE19_0510_CHECK_ICONNAME_OF_EDI_STAT03ICON,True)
'''' VerifyTextBoxContentIconName(textboxName, textboxIndex, expectedValue, blnIsItPopup)
Call VerifyTextBoxContentIconName("EDI_STAT03-ICON","",DT_WE19_0510_CHECK_ICONNAME_OF_EDI_STAT03ICON,True)


call ClickButton("Continue   \(Enter\)",TRUE) 

' VerifyTextBoxContent(textboxAttachedText, textboxName, textboxIndex, expectedValue, blnIsItPopup)
' GetTextboxValue(textboxName, textboxIndex, dataTableColumnName, blnIsItPopup)
call GetTextboxValue("MESSTXT1",0,"DT_CHECK_TEXT_OF_MESSTXT1_Output",true)
Call TakeScreenshot()
call ClickButton("Continue   \(Enter\)",TRUE) 
Call TakeScreenshot()
'call VerifyTextBoxContent("Information Message","MESSTXT1",1,

Call LogOff()
Call FinalStatus ()
'''control data


