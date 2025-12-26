
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_04.04.02.46 VIM - PO Precontrole Issue - BR24 - Missing Posting Date (PO)
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


gstrTestCaseName = "Test_Retrieve Cross Delivery"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 

'Call SetTcode(DT_SAPTRANSACTIONCODE)     
'Call PressEnter()     
'Call TakeScreenShot
'Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)


'-------------------------------------------- /nZMDDL_DEL_CROSS_DOCK----------------------------------------------


Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()

'Enter the Outbound Delivery No
Call ClickButton("Get Variant\.\.\.   \(Shift\+F5\)",False)
Call TakeScreenShot()
Call SetTextbox("Created By","ENAME-LOW","","",True)
Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",true)

Call TakeScreenShot()
' SelectRowGuiGrid(gridTitle, gridIndex, columnName, refValue, blnIsItPopup)
'Call SelectRowGuiGrid("Variant Catalog for Program ZMDDL_CREATE_CROSS_DOCK_DELIV",0,"Short Description","No correct lines updated",True)
Call SelectRowGuiGrid("Variant Catalog for Program ZMDDL_CREATE_CROSS_DOCK_DELIV",0,"Short Description","GTE292_BW04_LUX-Background job",True)
Call TakeScreenShot()
Call ClickButton("Choose   \(F2\)",true)
Call TakeScreenShot()
Call SetTextbox("Ship-To Party","S_KUNWE-LOW","",DT_SHIP_TO_PARTY,False)
Call SetTextbox("Distribution Channel","S_VTWEG-LOW","",DT_DISTR_CHANNEL,False)
Call SetTextbox("Act\. Gds Mvmnt Date","S_WADATI-LOW","",ConvertDate(DT_DATE),False)
'Call SetTextbox("to","S_WADATI-HIGH","","",False)
Call SetTextbox("Shipping Point/Receiving Pt","S_VSTEL-LOW","",DT_SHIPPING_POINT,False)
Call SetTextbox("Delivery","S_VBELN-LOW","",DT_ZMDDL_DEL_CROSS_DOCK_1000_DELIVERY,False)

' SelectRadioButton(radiobuttonName, radiobuttonAttachedtext, blnIsItPopup)
Call SelectRadioButton("P_FOR","Interactive mode - \(Online\)",False)

Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",false)
Call ClickButtonIfExist("Continue   \(Enter\)",True)

'VerifyGridCellContentbyName(gridName, gridRowNumber, gridColumnName, gridIndex, expectedValue)
call VerifyGridCellContentbyName("shell",1,"Article",0,DT_ZMDDL_DEL_CROSS_DOCK_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SOURCE_MATNR)
call VerifyGridCellContentbyName("shell",1,"Article",0,DT_ZMDDL_DEL_CROSS_DOCK_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SOURCE_MATNR)
call VerifyGridCellContentbyName("shell",1,"Quantity",0,DT_ZMDDL_DEL_CROSS_DOCK_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SOURCE_LFIMG)
call VerifyGridCellContentbyName("shell",1,"Quantity",0,DT_ZMDDL_DEL_CROSS_DOCK_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SOURCE_LFIMG)

Call TakeScreenShot()
Call LogOff()
Call FinalStatus ()

















'Call SetTextbox("Delivery","S_VBELN-LOW","","",False)
'Call ClickButton("Execute   \(F8\)",true)
'
'' SelectRadioButton(radiobuttonName, radiobuttonAttachedtext, blnIsItPopup)
'SelectRadioButton
'
'Call DoubleClickGuiGridCell("Variant Catalog for Program ZMDDL_CREATE_CROSS_DOCK_DELIV",0,1,"Short Description",True)
'Call SetTextbox("Ship-To Part","S_KUNWE-LOW","",DT_SHIP_TO_PARTY,False)
'Call SetTextbox("Distribution Channel","S_VTWEG-LOW","",DT_DISTR_CHANNEL,False)
'Call SetTextbox("Act\. Gds Mvmnt Date","S_WADATI-LOW","",DT_DATE,False)
'Call SetTextbox("Shipping Point/Receiving Pt","S_VSTEL-LOW","",DT_SHIPPING_POINT,False)
'Call SetTextbox("Delivery","S_VBELN-LOW","",DT_ZMDDL_DEL_CROSS_DOCK_1000_DELIVERY,False)
'
'Call ClickButton("Execute   \(F8\)",true)
'
'' VerifyGridCellContentbyName(gridName, gridRowNumber, gridColumnName, gridIndex, expectedValue)
'call VerifyGridCellContentbyName("shell",1,"Article",0,DT_ZMDDL_DEL_CROSS_DOCK_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SOURCE_MATNR)
'call VerifyGridCellContentbyName("shell",1,"Article",0,DT_ZMDDL_DEL_CROSS_DOCK_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SOURCE_MATNR)
'call VerifyGridCellContentbyName("shell",1,"Quantity",0,DT_ZMDDL_DEL_CROSS_DOCK_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SOURCE_LFIMG)
'call VerifyGridCellContentbyName("shell",1,"Quantity",0,DT_ZMDDL_DEL_CROSS_DOCK_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SOURCE_LFIMG)
'
'
'
'Call LogOff()
'Call FinalStatus ()
'
'
'
