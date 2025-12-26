
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_03.03.01.01.06 Manage Equipment _Vehicle - Equipment Automatic Creation_p1
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



'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name : Test_03.03.01.01.06 Manage Equipment _Vehicle - Equipment Automatic Creation_p1
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 24th Oct
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_03.03.01.01.06 Manage Equipment _Vehicle - Equipment Automatic Creation_p1"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.1.1.6. Create Article via Web Application - Common article_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

''Login to SAP System

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

'Close All Browser
Call CloseAllBrowsers()
Call LaunchSAPWebApplicationEdge(DT_SAPURL)
Wait(5)

Call LoginSRM(0,DT_SAPUSER,DT_SAPPASSWORD)
Wait(5)
Call CaptureWebScreen(0,"Capture Home Screen")

'Click on Create Professional Shopping Cart 
Call ClickWebButton(0,"",".*","Create Professional Shopping Cart","DIV",0,False)
Wait(10)
Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)


Call SetSAPEdit(1, "Name of shopping cart", "text", 0, DT_NAME_OF_SHOPPING_CART)
''''Enter the Description
Wait(1)

Call CaptureWebScreen(1,"Shopping Cart Description")

'Click on Item and Select From Catalog
Call ClickFrameSAPButton(1,"Create Shopping Cart","Add Item","DIV",0)
Call CaptureWebScreen(1,"Click Add Item")

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{DOWN}",1
Wait(2)
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait(5)

'''Click on AB SRM-MDM PreProduction GR Link
Call ClickWebElement(1,"","SPAN","AB SRM-MDM PreProduction GR","lsLink__text",0,True)
Wait(30)
Call CaptureWebScreen(1,"Capture From Catalog Pop Up")

Call SetWebEdit(1,"","APIDMCJH\.BasicView\.KeywordCriteriaInputField","text","",DT_KEYWORD,False)
Call CaptureWebScreen(1,"Capture screen:Enter KeyWord")

Call ClickFrameSAPButton(1,"ProcurementCatalog7Ehp1","Search","DIV",0)
Wait 5
Call CaptureWebScreen(1,"SearchResult")

Call ClickWebElementFrameNoInnertext(1, "ProcurementCatalog7Ehp1", "SPAN", "APIDMABI\.ResultSetView\.dynamicColumnCellEditorCheckBox\.0", 0)
'Call ClickWebElement(1,"","SPAN",".*",".*lsSelector.*",2,False)
Wait(1)
Call CaptureWebScreen(1,"capture advanced search details")

'Click on Add To Cart
Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Add to Cart", "DIV", 0)
Call CaptureWebScreen(1,"Capture screen:Add To cart")
Call ClickFrameSAPButton(1, "ProcurementCatalog7Ehp1", "Check Out", "DIV", 0)
Wait 30

Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",0)
wait 20
Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Account Assignment", 0)
Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",1)
Call CaptureWebScreen(1,"Capture screen:Account Assignment1")

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{PGDN}",1
Wait(2)
Set wsh= nothing
Wait(5)


Call SetWebEditFrameLogicalName(1, "Shopping Cart", " WBS Element",  "text", 0, DT_WBS_ELEMENT)
Call SetWebEditFrameLogicalName(1, "Shopping Cart",  "Cost Center",  "text", 0, DT_COST_CENTER)

Call CaptureWebScreen(1,"Capture screen:Enter Asset Master details")
'Create Asset Master Button
Call ClickFrameSAPButton(1,"Shopping Cart","Create Asset Master","DIV",0)
Wait(3)
Call CaptureWebScreen(1,"Capture screen:Click Button Create Assset Master")


'Click on Check Button
Call ClickFrameSAPButton(1,"Shopping Cart","Check","DIV",0)
Wait(3)
Call CaptureWebScreen(1,"Capture screen:Click on Check Button")


Call GetValueWebEditByLogicalName(1,"lsField__input","INPUT","Asset",0,"DT_ASSET_NO")

'Click on Check Button
Call ClickFrameSAPButton(1,"Shopping Cart","Check","DIV",0)
Wait(3)
Call CaptureWebScreen(1,"Capture screen:DIV")

Call GetValueWebElementFrame(1, "Shopping Cart", "FPM_OIF_COMPONENT\.ID_0001:CNR_VIEW\.VALUE_1", "SPAN", "ls.*", "DT_SC_NO_OUTPUT")
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Click on Check Button
Call ClickFrameSAPButton(1,"Shopping Cart","Save","DIV",0)
Wait(3)
Call CaptureWebScreen(1,"Capture screen:DIV")

'Get the shopping Cart No generated and Store in Excel
'Call GetValueWebElement(1,"FPM_OIF_COMPONENT\.ID_0001:CNR_VIEW\.VALUE_1",".*","SPAN",0,0,"DT_SC_NO_OUTPUT")



'Click on Edit Button
Call ClickFrameSAPButton(1,"Display Shopping Cart","Edit","DIV",0)
Wait(3)
Call CaptureWebScreen(1,"Capture screen:DIV")

'Click on Order
Call ClickFrameSAPButton(1,"Edit Shopping Cart","Order","DIV",0)
Wait 45
Call CaptureWebScreen(1,"Capture screen:Complete order")
'
'''Verify if order has been plaed succesfully
''Prior to code update
''Call VerifyFrameWebElement(1, "", "Display Document:", "SPAN", "Shopping cart """&DT_NAME_OF_SHOPPING_CART&""" with number "&DT_SC_NUMBER&" ordered successfully", "lsTextView.*", "", False)
''Post code update
Call VerifyFrameWebElement(1, "", "Display Document:", "DIV", "Shopping cart """&DT_NAME_OF_SHOPPING_CART&""" with number "&DT_SC_NUMBER&" ordered successfully", "lsMSGText.*", "", False)
Call ClickFrameSAPButton(1,"Display Document:","Close","DIV",0)
Wait 10
Call CaptureWebScreen(0,"Capture screen:Close button")

''------------------------'Log Off From Applicaton--------------------------------

Call LogoffSRM(0)
Call FinalStatus()


'*********************************************End Of Script*********************************************************************

Public Function GetValueWebElement(creationTime,htmlId,innertext,htmltag,textPosition,Index,excelColumnName)
If Not (Environment.Value("blnFatalError") or excelColumnName = DS_SKIP) Then
   If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : GetValueWebElement"
ALMvalue=getEnvironmentVariable("ALMFlag")
If  ALMvalue=1 Then
	PopUpCreationtime=Cint(creationTime) + 1
else
	PopUpCreationtime=Cint(creationTime)
End If
Dim objWebElement
If htmlId<> ""  and innertext<>""Then
strStepName = "Get WebElement Content "
Set objWebElement =Browser("CreationTime:="& PopUpCreationtime).Page("Index:="& PopUpCreationtime).WebElement("html id:="& htmlId,"innertext:="&innertext,"html tag:="&htmltag,"index:="&Index)
On error resume next 
 If  objWebElement.Exist  then 
	 objWebElement.highlight
     webElementMsg=objWebElement.GetROProperty("outertext")
     MyArray=Split(webElementMsg," ","-1","1")
	 textPosition=Cint(textPosition)
	 val=MyArray(textPosition)
     last  = Right(val,1)
     If last = "-" Then
	     finallast=instr(1,val,"-",1)
         finalval=left(val,finallast-1)
         val=finalval
     End If
	 
               If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
                   ImagePath=CaptureScreenshot(strStepName,objWebElement,False,False,False)
               End if
               
               strStatus = "DONE"
               strMsg= "value captured : "& val
               Call ReporterFunction(strLibraryFileName,"GetValueWebElement","2",strStepName,strMsg)
  else
                blnObjectError=True
              
              strStatus = "FAIL"    
              strMsg = "Web Element  object Missing"
              Call ReporterFunction(strLibraryFileName,"GetValueWebElement","1",strStepName,strMsg)
  end if
else
			  
              strStatus = "FAIL"
              strMsg = "Function Parameter Not Passed Properly. Check the --GetValueWebElement -- Function Call"
              Call ReporterFunction(strLibraryFileName,"GetValueWebElement","1",strStepName,strMsg)
     end if
If  blnObjectError  Then
       Environment.Value("blnFatalError")=True
End If

If strStatus = "FAIL"  Then
       GetValueWebElement = strMsg
       blnMainFailFlag = True
       ImagePath=CaptureScreenshot(strStepName,objWebElement,False,False,False)
Else
       GetValueWebElement = True
End If
If blnWriteDataToOutputSheet Then
	                                    strStepName = "Retrieve '"&val&"' value in table '"&gstrOutputSheetName&"' sheet under column "&excelColumnName
										call WriteRunTimeDataToExcel (excelColumnName,val)
								   ELSE
										strStepName = "Retrieve '"&val&"' value in table"
	 End If
If blnDefault_eSwiftReporting Then  
       Call UpdateResultHtml(strStepName,val,strMsg,strStatus,"")
End If

   Set objWebElement=Nothing
   End If
End Function




'******************************************************************************************************************



