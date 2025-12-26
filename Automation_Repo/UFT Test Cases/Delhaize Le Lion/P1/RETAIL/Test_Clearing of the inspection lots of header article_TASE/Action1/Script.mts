
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_Reverting Goods Receipt using the PO Number and Verification of the net result of material_TASE
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

gstrTestCaseName = "TC_11_Test_Clearing of the inspection lots of header article_TASE"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\DEmopractice\Data\P1_DATA\DT_MD_ABI067_001 Create ZCXT Retail Customer Local or Foreign_TASE.xls"



'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

''''''--------------login----------------'''''

Call SAPGuiUtil.OpenConnection(DT_SAPSYSTEM)
Call Login(DT_SAPUSER,DT_SAPPASSWORD)
Call PressEnter()

''----------------------Tcode QA32----------------------------
Call SetTcode(DT_SAPTRANSACTIONCODE)
Call PressEnter() 
Wait(2)'
Call TakeScreenShot()

Call SetTextboxNoLabel("QL_ENSTD-LOW", 0, ConvertDate(DT_LOT_CREATED_ON), False)
Call SetTextboxNoLabel("QL_ENSTD-HIGH", 0, ConvertDate(DT_LOT_CREATED_ON), False)
Call SetTextboxNoLabel("QL_MATNR-LOW", 0, DT_HEADER_ARTICLE, False)
Call TakeScreenShot()

Call SelectRadioButton("P_ALLES","Select all inspection lots",False)
Call TakeScreenShot()

Call ClickButton("Execute   \(F8\)",False)
Call TakeScreenShot()

Call SelectRowGuiGridbyRowNo("","",1,False)
Call TakeScreenShot()

Call ClickButton("Change inspection lot   \(Ctrl\+Shift\+F6\)",False)
Call TakeScreenShot()

Call SelectMenuBar("Inspection Lot;Functions;Cancel Lot")
Call TakeScreenShot()

Call ClickButton("Save   \(Ctrl\+S\)",False)
Call TakeScreenShot()

Call GetTextStatusBar("DT_STATUS")

Call  ClickButton("Back   \(F3\)",False)
Call TakeScreenShot()

Call LogOff()
Call FinalStatus()
