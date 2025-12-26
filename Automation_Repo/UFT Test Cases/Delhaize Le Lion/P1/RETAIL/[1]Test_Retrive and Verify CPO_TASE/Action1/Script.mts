
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


gstrTestCaseName = "Test_Retrive and Verify CPO"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_Retrieve and verify child OBD_Output.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//


Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'

Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
'
'--------------------------------------------  SE16----------------------------------------------


Call SetTextbox("Table Name","DATABROWSE-TABLENAME","",DT_SE16_0230_TABLE_NAME,False)
Call TakeScreenShot()
Call PressEnter()
Call SetTextbox("to","I6-HIGH","",ConvertDate(DT_SE16_1000_TO),False)
'Call SetTextbox("Created On","I6-LOW","",ConvertDate(DT_SE16_1000_CREATED_ON),False)
Call SetTextbox("ERDAT","I6-LOW","",ConvertDate(DT_SE16_1000_CREATED_ON),False)
'Call SetTextbox("Preceding Doc\.","I1-LOW","",DT_SE16_1000_PRECEDING_DOC,False)
Call SetTextbox("VBELV","I1-LOW","",DT_SE16_1000_PRECEDING_DOC,False)
Call TakeScreenShot()
Call ClickButtonIfExist("Execute   \(F8\)",false)
Call TakeScreenShot()
'Call GetGridContent("",0,"FollOn doc",1,"<NA>","<NA>","DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VBELN_OUTPUT")
  ' GetLabelContentByRefLabel(refLabelContent, xDifferenceValue, yDifferenceValue, dataTableColumnName, blnIsItPopup)
Call GetLabelContentByRefLabel("VBELN",0,-32,"DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VBELN_OUTPUT",False)


Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Back   \(F3\)",False)

'----------------------------------------------ME23N----------------------------------

Call SetTcode(DT_SE16_0100_OKCD)
Call PressEnter()
Call TakeScreenShot()
Call CheckTCodeScreen(DT_SE16_0100_OKCD)

' GetInputFromExcel(InputFilePath, sheetName, IterationIndex)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global", DataRowSet)

Call ClickButton("Other Purchase Order   \(Shift\+F5\)",false)
wait(2)
Call SetTextbox("Pur\. Order","MEPO_SELECT-EBELN","",DT_SE16_0003_PUR_ORDER,True)
Call TakeScreenShot()
Call PressEnter()
Wait(2)
Call TakeScreenShot()


Call VerifyTextBoxContent("Doc\. date","MEPO_TOPLINE-BEDAT","",ConvertDate(DT_SE16_1105_CHECK_TEXT_OF_DOC_DATE),False)
Call VerifyTextBoxContent("Supplier","MEPO_TOPLINE-SUPERFIELD","",DT_SE16_1105_CHECK_TEXT_OF_VENDOR,False)
Call ClickButtonIfExist("Expand Items Ctrl\+F3",False)
Call TakeScreenShot()
call VerifyTableCellContent(1,"Article","SAPLMEGUITC_1211",DT_SE16_1211_CHECK_TEXT_OF_TABLECELL_ARTICLE_0)
call VerifyTableCellContent(2,"Article","SAPLMEGUITC_1211",DT_SE16_1211_CHECK_TEXT_OF_TABLECELL_ARTICLE_1)
call VerifyTableCellContent(1,"PO Quantity","SAPLMEGUITC_1211",DT_SE16_1211_CHECK_TEXT_OF_TABLECELL_PO_QUANTITY_0)
call VerifyTableCellContent(2,"PO Quantity","SAPLMEGUITC_1211",DT_SE16_1211_CHECK_TEXT_OF_TABLECELL_PO_QUANTITY_1)
call VerifyTableCellContent(1,"Deliv. Date","SAPLMEGUITC_1211",ConvertDate(DT_SE16_1211_CHECK_TEXT_OF_TABLECELL_DELIV_DATE_0))
call VerifyTableCellContent(2,"Deliv. Date","SAPLMEGUITC_1211",ConvertDate(DT_SE16_1211_CHECK_TEXT_OF_TABLECELL_DELIV_DATE_1))
call VerifyTableCellContent(1,"Plnt","SAPLMEGUITC_1211",DT_SE16_1211_CHECK_TEXT_OF_TABLECELL_SITE_0)
call VerifyTableCellContent(2,"Plnt","SAPLMEGUITC_1211",DT_SE16_1211_CHECK_TEXT_OF_TABLECELL_SITE_1)
'call VerifyTableCellContent(1,"Reqmt No\.","SAPLMEGUITC_1211",DT_SE16_1211_CHECK_TEXT_OF_TABLECELL_REQMT_NO_0)
'call VerifyTableCellContent(2,"Requiremt No\.","SAPLMEGUITC_1211",DT_SE16_1211_CHECK_TEXT_OF_TABLECELL_REQMT_NO_1)

call ClickButtonIfExist("Expand Header Ctrl\+F2",false)
Call SelectTab("HEADER_DETAIL","Org. Data",False)
Call TakeScreenShot()
call VerifyTextBoxContent("Purch\. Org\.","MEPO1222-EKORG","",DT_SE16_1221_CHECK_TEXT_OF_PURCH_ORG,False)
call VerifyTextBoxContent("Purch\. Group","MEPO1222-EKGRP","",DT_SE16_1221_CHECK_TEXT_OF_PURCH_GROUP,False)
call VerifyTextBoxContent("Company Code","MEPO1222-BUKRS","",DT_SE16_1221_CHECK_TEXT_OF_COMPANY_CODE,False)

call ClickButtonIfExist("Expand Item Details Ctrl\+F4",false)
Call TakeScreenShot()
Call SelectTab("ITEM_DETAIL","Purchase Order History",False)
Call GetGridContent("",0,"Article Document",1,"<NA>","<NA>","DT_SE16_0100_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR_OUTPUT")

Call LogOff()
Call FinalStatus ()


'call GetLabelContentByRefLabel("VBELN",0,-32,"DT_SE16_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_VBELN_OUTPUT",False)










