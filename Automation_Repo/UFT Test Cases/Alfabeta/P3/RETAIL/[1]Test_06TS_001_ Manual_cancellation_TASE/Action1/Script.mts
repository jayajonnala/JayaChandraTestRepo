		

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :[1]Test_06TS_001_ Manual_cancellation_TASE
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
	GetRowNo= Parameter("datatable_row")	
End If
If qtpParamExist("RunTimeResultFolder") Then
	RunTimeResultFolder= Parameter("RunTimeResultFolder")	
End If
gstrresultFolderPath = ReadTxtFileResult(RunTimeResultFolder)

gstrTestCaseName = "Test_06TS_001ManualCancel_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"

'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

'GetRowNo =2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''''''--------------login----------------'''''
Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()  
''''--------------------------------VF11-----------------------------
'
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE)

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call SetTextbox("Billing Date","RV60A-FKDAT","",ConvertDate(DT_VF11_0102_BILLING_DATE),False)   
Call SetTableData("SAPMV60ATCTRL_ERF_FAKT","Document","1","","",DT_VF11_0102_TABLECELL_DOCUMENT_0,False) 
Call PressEnter()
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()

Call GetStatusBar("item1","DT_DOC_NO_OUTPUT")
Call VerifyStatusBar("Document "&DT_DOC_NO_OUTPUT&" saved (no automatic clearing of billing document "&DT_VF11_0102_TABLECELL_DOCUMENT_0&")")
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_NO_OUTPUT",DT_DOC_NO)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

''''--------------------------------Vl09-----------------------------

Call SetTcode(DT_VF11_0102_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC1)

Call SetTextbox("Shipping Point","I_VSTEL-LOW","",DT_VF11_1000_SHIPPING_POINT,False)
Call SetTextbox("Inbound / Outbound Delivery","I_VBELN-LOW","",Right(DT_INBOUND,8),False)

Call TakeScreenShot()
Call ClickButton("Execute   \(F8\)",False)

Call VerifyGridCellContent("", 1, "Delivery", 0, DT_VF11_1000_INBOUND__OUTBOUND_DELIVERY)
Call SelectRowGuiGrid("",0,"Delivery",DT_VF11_1000_INBOUND__OUTBOUND_DELIVERY,False)
Call ClickButton("Reverse Goods Movement   \(F5\)",False)
Call TakeScreenShot()
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()
Wait 3
Call ClickButton("Continue   \(Enter\)",True)
Call TakeScreenShot()
''''--------------------------------vl02n-----------------------------

Call SetTcode(DT_VF11_0500_OKCD) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC2)

Call SetTextbox("Outbound Delivery","LIKP-VBELN","",DT_VF11_4004_OUTBOUND_DELIVERY,False)
Call PressEnter()
Call SelectMenuBar("Outbound Delivery;Delete")
Call TakeScreenShot()
Call ClickButtonifExist("Yes", True)
Wait 2
Call ClickButtonifExist("Delete", True)
Call TakeScreenShot()
Call VerifyStatusBar("AB_Outbound Delivery "&DT_VF11_4004_OUTBOUND_DELIVERY&" deleted")
''''--------------------------------VA02-----------------------------

Call SetTcode(DT_VF11_4004_OKCD) 
Call PressEnter()     
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC3)

Call SetTextbox("Order","VBAK-VBELN","",DT_VF11_0102_ORDER,False)
Call PressEnter()  
Call ClickButtonifExist("Continue   \(Enter\)",True)
Call TakeScreenShot()
Call SelectMenuBar("Sales document;Delete")
Call ClickButtonifExist("Yes", True)
Wait 2
Call VerifyStatusBar(Lcase(DT_VF11_0102_CHECK_TEXT_OF_STATUSBAR_OCC1))
'
'''''--------------------------------MIGO-----------------------------

Call SetTcode(DT_VF11_0102_OKCD_OCC1) 
Call PressEnter()     
Call CheckTCodeScreen(DT_EXPECTEDTRANSACTIONCODE_OCC4)

call SetComboByKey("GODYNPRO-ACTION",DT_VF11_0010_GODYNPROACTION)

Call SetTextboxNoLabel("GODYNPRO-MAT_DOC",0,DT_VF11_2010_GODYNPROMAT_DOC,False)
Call SetTextboxNoLabel("GODYNPRO-DOC_YEAR",0,(GODYNPRO-PO_ITEM),False)
Call PressEnter()

Call SetTableData("SAPLMIGOTV_GOITEM","OK","2","","","ON",False)

Call ClickButton("Post   \(Ctrl\+S\)",False)

Call GetStatusBar("item1","DT_DOC_2_OUTPUT")
Call VerifyStatusBar("Article document "& DT_DOC_2_OUTPUT &" posted" )
Call WriteRunTimeDataToExcelGlobalSheet ("DT_DOC_2_OUTPUT",DT_DOC_2)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",GetRowNo)

Call LogOff'
Call FinalStatus()
