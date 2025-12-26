'
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_Order group maintenance
'.................Author : TCS        :Jaya
'................ Creation Date    : 
'.................Modified By :
'.................Modified Date/Details :

'''''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//


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
'''''//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//



gstrTestCaseName = "Test_Order group maintenance"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\TASE Files\DATA\DT_Clear GL Accounts  Manual and Automatic_p2_TASE.xls"
''//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//
'

'DataRowSet=2
Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''----------------------Login-------------------------'''

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter() 
Call TakeScreenShot()
'''----------------------Tcode KOH1----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE) 
Call PressEnter()     ' 
Call CheckTCodeScreen(DT_SAPTRANSACTIONCODE)
Call TakeScreenShot()


Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",Cint(DT_INCREMENT)+1)
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


Call SetTextbox("Controlling Area","SVALD-VALUE","",(DT_KOH1_0300_CONTROLLING_AREA),True)
Call ClickButtonIfExist("Continue   \(Enter\)",True)

Call SetTextbox("Order Group","GRPDYNP-NAME_COALL","",(DT_KOH1_1000_ORDER_GROUP),False)
Call PressEnter()  
Call TakeScreenShot()
Call SetTextbox(DT_KOH1_1000_ORDER_GROUP,"",0,(DT_KOH1_0120_NO_NAME),False)
Call ClickBUtton("Insert Order   \(Shift\+F4\)",False)

Call SetTextbox(DT_KOH1_1000_ORDER_GROUP,"",6,(DT_KOH1_0120_NO_NAME_OCC1),False)
Call SetTextbox(DT_KOH1_1000_ORDER_GROUP,"",7,(DT_KOH1_0120_NO_NAME_OCC2),False)
Call PressEnter()  
Call TakeScreenShot()
Call ClickBUtton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBar(DT_STATUS_1)

'''----------------------Tcode KOH2----------------------------
Call SetTcode(DT_KOH1_0100_OKCD) 
Call PressEnter()
Call TakeScreenShot()

Call SetTextbox("Order Group","GRPDYNP-NAME_COALL","",DT_KOH1_1000_ORDER_GROUP_OCC1,False)
Call PressEnter()  
Call TakeScreenShot()
Call SetTextbox(DT_KOH1_1000_ORDER_GROUP,"",0,(DT_KOH1_0120_NO_NAME_OCC3),False)
Call ClickBUtton("Insert Order   \(Shift\+F4\)",False)
Call TakeScreenShot()
Call SetTextbox(DT_KOH1_1000_ORDER_GROUP,"",6,(DT_KOH1_0120_NO_NAME_OCC4),False)
Call PressEnter()  
Call TakeScreenShot()
Call ClickBUtton("Save   \(Ctrl\+S\)",False)
Call VerifyStatusBar(DT_STATUS_2)

'''----------------------Tcode KOH3----------------------------
Call SetTcode(DT_KOH1_0100_OKCD_OCC1) 
Call PressEnter()
Call TakeScreenShot()

Call SetTextbox("Order Group","GRPDYNP-NAME_COALL","",DT_KOH1_1000_ORDER_GROUP_OCC2,False)
Call PressEnter()  
Call TakeScreenShot()

Call VerifyifGuiLabelExists("FIAT DUCATO FURGON 2.3 JTD 40 XLH3")
Call VerifyifGuiLabelExists(DT_KOH1_0120_CHECK_TEXT_OF_MERCEDES_BENZ_SPRINTER_308_CDI_OCC1)
Call VerifyifGuiLabelExists("RENAULT MASKOTT 150\.65-TT-3,5M\+BV\+HL\+RV")

Call LogOff()
Call FinalStatus()

