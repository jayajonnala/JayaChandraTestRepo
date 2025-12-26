

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_09.07.01.02.09 Store Cash Management_Credit cards man process V5_TASE
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

gstrTestCaseName = "Test_09.07.01.02.09 Store Cash Management_Credit cards man process V5_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_POST_DeleteVAT_from_Customer_TASE.xls"


'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",GetRowNo,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario
'
'''Call LaunchSAPConnection("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe",DT_SAPSYSTEM)
SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
'''
'
'''----------------------Tcode FB08----------------------------
'
Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call SetTextbox("Reversal Reason","UF05A-STGRD","",DT_FB08_0105_REVERSAL_REASON,False)
Call SetTextbox("Fiscal Year","RF05A-GJAHS","",DT_FB08_0105_FISCAL_YEAR,False)
Call SetTextbox("Company Code","BKPF-BUKRS","",DT_FB08_0105_COMPANY_CODE,False)
Call SetTextbox("Document Number","RF05A-BELNS","",DT_FB08_0105_DOCUMENT_NUMBER,False)
Call TakeScreenShot()

Call FocusTextBox("Reversal Reason","UF05A-STGRD",False)
Call ClickButton("Display document before reversal   \(F5\)",fALSE)

Call VerifyGridCellContent("",1,"Posting Key",0,DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
Call VerifyGridCellContent("",2,"Posting Key",0,DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)
Call VerifyGridCellContent("",3,"Posting Key",0,DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_BSCHL)
Call VerifyGridCellContent("",1,"Account",0,DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_KTONR)
Call VerifyGridCellContent("",2,"Account",0,DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_KTONR)
Call VerifyGridCellContent("",3,"Account",0,DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_KTONR)
Call VerifyGridCellContent("",1,"Amount",0,DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_AZBET)
Call VerifyGridCellContent("",2,"Amount",0,DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_AZBET)
Call VerifyGridCellContent("",3,"Amount",0,DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_AZBET)
Call VerifyGridCellContent("",1,"Profit Center",0,DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
Call VerifyGridCellContent("",2,"Profit Center",0,DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)
Call VerifyGridCellContent("",3,"Profit Center",0,DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_PRCTR)
Call VerifyGridCellContent("",1,"Account",0,DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_0_HKONT)
Call VerifyGridCellContent("",2,"Account",0,DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_1_HKONT)
Call VerifyGridCellContent("",3,"Account",0,DT_FB08_0750_CHECK_GETCELLVALUE_OF_GRIDCELL_2_HKONT)

Call ClickButton("Back   \(F3\)",False)
Call ClickButton("Post   \(Ctrl\+S\)",False)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call GetStatusBar("item1","DT_FB08_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT")
Call WriteRunTimeDataToExcelGlobalSheet("DT_FB08_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR_OUTPUT",DT_FB08_0105_CHECK_MESSAGEPARAMETER_OF_STATUSBAR)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


'''''''----------------------Tcode FAGLL03----------------------------


Call SetTcode(DT_FB08_0105_OKCD)
Call PressEnter() 


Call SelectRadioButton("X_AISEL","All Items",False)
Call SetTextbox("Posting Date","SO_BUDAT-LOW","",ConvertDate(DT_FB08_1000_POSTING_DATE),False)
Call SetTextbox("to","SO_BUDAT-HIGH","",ConvertDate(DT_FB08_1000_TO),False)
Call SetTextbox("G/L account","SD_SAKNR-LOW","",DT_FB08_1000_GL_ACCOUNT,False)
Call TakeScreenShot()

Call FocusTextBox("to","SD_SAKNR-HIGH",False)
Call ClickButton("Execute   \(F8\)",fALSE)
wait(10)
Call VerifyStatusBarMessageType(DT_EXPECTEDVALUE)
Call GetStatusBar("item1","DT_DOC_NR_Output")
Call TakeScreenShot()
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Call ClickLabel("DocumentNo",1,False)
Call SelectColumnGuiGrid("",1, "Document Number", False)
Call ClickButton("Set Filter   \(Ctrl\+Shift\+F2\)",False)
Wait(2)
Call ClickButton("Multiple selection",True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",1,DT_FB08_3010_TABLECELL_SINGLE_VALUE_0,True)
Call SetTableDataNoRef("SAPLALDBSINGLE","Single value",2,DT_FB08_3010_TABLECELL_SINGLE_VALUE_1,True)
Call ClickButton("Copy   \(F8\)",True)
wait(2)
Call ClickButton("Execute   \(Enter\)",True)
Call TakeScreenShot()

CAll VerifyGridCellContent("", 1, "BELNR", "", DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR)
CAll VerifyGridCellContent("", 2, "BELNR", "", DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR)

CAll VerifyGridCellContent("", 1, "GSBER", "", DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GSBER)
CAll VerifyGridCellContent("", 2, "GSBER", "", DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_GSBER)

CAll VerifyGridCellContent("", 1, "BLART", "", DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART)
CAll VerifyGridCellContent("", 2, "BLART", "", DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLART)

CAll VerifyGridCellContent("", 1, "BSCHL", "", DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL)
CAll VerifyGridCellContent("", 2, "BSCHL", "", DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL)

CAll VerifyGridCellContent("", 1, "DMSHB", "", DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB)
CAll VerifyGridCellContent("", 2, "DMSHB", "", DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB)

CAll VerifyGridCellContent("", 1, "MWSKZ", "", DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ)
CAll VerifyGridCellContent("", 2, "MWSKZ", "", DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ)

CAll VerifyGridCellContent("", 1, "PRCTR", "", DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR)
CAll VerifyGridCellContent("", 2, "PRCTR", "", DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR)

CAll VerifyGridCellContent("", 1, "SGTXT", "", DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT)
CAll VerifyGridCellContent("", 2, "SGTXT", "", DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT)


'''uncomment belo line if the table displayed instead of grid  ---''Comments added by Jaya
'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BELNR,"wnd\[0\]/usr/lbl\[9,8\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BELNR,"wnd\[0\]/usr/lbl\[9,9\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_GSBER,"wnd\[0\]/usr/lbl\[20,8\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_GSBER,"wnd\[0\]/usr/lbl\[20,9\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLART,"wnd\[0\]/usr/lbl\[25,8\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLART,"wnd\[0\]/usr/lbl\[25,9\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BLDAT,"wnd\[0\]/usr/lbl\[29,8\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BLDAT,"wnd\[0\]/usr/lbl\[29,9\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_BSCHL,"wnd\[0\]/usr/lbl\[40,8\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_BSCHL,"wnd\[0\]/usr/lbl\[40,9\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_DMSHB,"wnd\[0\]/usr/lbl\[43,8\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_DMSHB,"wnd\[0\]/usr/lbl\[43,9\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_MWSKZ,"wnd\[0\]/usr/lbl\[70,8\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_MWSKZ,"wnd\[0\]/usr/lbl\[70,9\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_PRCTR,"wnd\[0\]/usr/lbl\[84,8\]")
'''Call VerifyifGuiLabelExistsByRelativeid(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_PRCTR,"wnd\[0\]/usr/lbl\[84,9\]")
''
''Call SetHorizontalScrollBar(100, False)
''Call takeScreenShot()
''Call VerifyifGuiLabelExistsByRelativeid(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_0_SGTXT,"wnd\[0\]/usr/lbl\[6,8\]")
''Call VerifyifGuiLabelExistsByRelativeid(DT_FB08_0500_CHECK_GETCELLVALUE_OF_GRIDCELL_1_SGTXT,"wnd\[0\]/usr/lbl\[6,9\]")

Call LogOff()
Call FinalStatus ()



