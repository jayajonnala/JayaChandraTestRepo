
'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//
'.................Test Script Name :Test_03.03.01.01.07 Manage Equipment _Vehicle - Equipment Automatic Creation
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
'.................Test Script Name : Test_03.03.01.01.07 Manage Equipment _Vehicle - Equipment Automatic Creation
'.................Author : TCS 	   :Raushan
'................ Creation Date    : 24th Oct
'.................Modified By :
'.................Modified Date/Details :

'//---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------//

gstrTestCaseName = "Test_03.03.01.01.07 Manage Equipment _Vehicle - Equipment Automatic Creation"
gstrTCDescription = ""'.....Please give TestCase Description.
gstrTCPrerequisite = ""'.....Please mention the pre-requisites for TestCase If any.
'gstrInputExcelFilePathAndName="C:\Program Files (x86)\TCS\Test Automation Solution for Enterprise Applications\InputDatasheet\DT_2.1.1.6. Create Article via Web Application - Common article_P1.xls"
'//---------------------------------------------------------------------------------------------------------------------------------------------------AUTOMATION TEST SCRIPT------------------------------------------------------------------------------------------------------------------------------------------------------------//

''Login to SAP System

Call StartExecution(gstrInputExcelFilePathAndName,"Global",DataRowSet,gstrresultFolderPath)  '.......................Mandatory Initial Call only in First Component in a Test Scenario

Call WriteRunTimeDataToExcelGlobalSheet ("DT_INCREMENT",(Cint(DT_INCREMENT)+1))
Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'''Close All Browser
Call CloseAllBrowsers()

''launch adn Login SRM Application on Edge
Call LaunchSAPWebApplicationEdge(DT_SAPURL)
Wait(5)

'''''Login to SRM application
Call LoginSRM(0,DT_SAPUSER,DT_SAPPASSWORD)
Call CaptureWebScreen(0,"Capture Home Screen")

wait (5)

'Click on Create Professional Shopping Cart
Call ClickWebElement(0,"","DIV","Shopping & Receiving","sapUshellAnchorItemInner",0,False)
Call CaptureWebScreen(0,"Shopping & Receiving")
Call ClickWebButton(0,"",".*","Create Professional Shopping Cart","DIV",0,False)
Wait 10
Call CaptureWebScreen(1,"Create Professional Shopping Cart")

'Enter the Description
Call SetSAPEdit(1, "Name of shopping cart", "text", 0, DT_NAME_OF_SHOPPING_CART)
Wait(1)
Call CaptureWebScreen(1,"Shopping Cart Description")

'Click on Item and Select From Catalog
Call ClickFrameSAPButton(1,"Create Shopping Cart","Add Item","DIV",0)
Wait(1)
Call CaptureWebScreen(1,"Click Add Item")
Set wsh = createobject("Wscript.Shell")
'wsh.SendKeys "{DOWN}",1
'Wait(2)
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait(5)
Call CaptureWebScreen(1,"Click Add Item")
Call SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Description", "text", 0, DT_DESCRIPTION)
Call SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Product Category", "text", 0, DT_PRODUCT_CATEGORY)
Call SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Quantity / Unit", "text", 0, DT_QUANTITY__UNIT)
Call SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Unit for Order Quantity", "text", 0, DT_QUANTITY__UNIT_OCC1)
Call SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Net Price / Currency", "text", 0, DT_NET_PRICE__CURRENCY)
Call SetSAPEditSAPFrame(1, "Floor Plan Manager application for OIF", "Currency for Price", "text", 0, DT_SAPSRMWDC_DODC_SC_I_DESID_SC_DOFC_I_DESCRIBEUNIT1)
Call CaptureWebScreen(1,"Capture screen:Enter Description")

Call ClickSAPButtonSAPFrame(1, "Floor Plan Manager application for OIF", "lsButton.*", "Add to Item Overview", "DIV", 0)
Call CaptureWebScreen(1,"Capture screen:Add to Item Overview")
Wait 10
Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",0)
Wait 5

'Click on Sources of Supply
Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Sources of Supply / Service Agents", 0)
Wait(3)
Call CaptureWebScreen(1,"Capture screen: Click on Sources of Supply / Service Agents")

'''''Set the Supplier Value and Click on Assign Supplier button

Call SetWebEditFrameLogicalName(1, "Shopping Cart",  "Supplier",  "text", 0, DT_SUPPLIER)
Call CaptureWebScreen(1,"Cost Center Change")
Call ClickFrameSAPButton(1,"Shopping Cart","Assign Supplier","DIV",0)
Wait 2
Call CaptureWebScreen(1,"Capture screen: Click on Assign supplier")

Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Notes and Attachments", 0)
Wait(2)
Call ClickFrameSAPButton(1,"Shopping Cart","Add Attachment","DIV",0)
Wait 2

Call SetWebFile(1,"Floor Plan Manager application for OIF","/SAPSRM/WDC_DODC_NA\.ID_2979E42ACB25FE185C32B3220196317A:V_DODC_ADDATTACHMENT\.FILEUPLOAD_USAGE","INPUT",0,DT_FREE_TEXT)
Call CaptureWebScreen(1,"Capture screen: Add Attachments")
Call ClickSAPFrameSAPButton(1, "Floor Plan Manager application for OIF", "OK", "DIV", 0)

'Click on Account Assignment
Call ClickWebElementFrame(1, "Shopping Cart", "DIV", "Account Assignment", 0)
Call ClickFrameSAPButton(1,"Shopping Cart","Details","DIV",1)
Call CaptureWebScreen(1,"Capture screen:Account Assignment1")

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{PGDN}",1
Wait(2)
Set wsh= nothing
Wait 5

Call ClickSAPList(1, "Account Assignment Category", "INPUT",1)

Set wsh = createobject("Wscript.Shell")
wsh.SendKeys "{HOME}",1
Wait(2)
wsh.SendKeys "{ENTER}",1
Set wsh= nothing
Wait 5


Call SetWebEditFrameLogicalName(1, "Shopping Cart",  "WBS Element",  "text", 0, DT_WBS_ELEMENT__OCC1)
Call SetWebEditFrameLogicalName(1, "Shopping Cart",  "Cost Center",  "text", 0, DT_COST_CENTER__OCC1)

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

'Click on Check Button
Call ClickSAPButton(1,".*","Save","DIV",0)
Wait(3)
Call CaptureWebScreen(1,"Capture screen:DIV")
'Get the shopping Cart No generated and Store in Excel
Call GetValueWebElementFrame(1, "Display Shopping Cart", "WDR_MESSAGE_AREA\.ID_814D32E25620602B3346FEE322896947:MESSAGE_AREA\.MSG_ID_ADV_2-text", "DIV", "ls.*", "DT_CREATED_CART_OUTPUT")
Call VerifyFrameWebElement(1, "", "Display Shopping Cart", "DIV", DT_CREATED_CART_OUTPUT, "ls.*", 0, False)

'Call GetValueWebElementFrame(1, "Display Document:","WDR_MESSAGE_AREA\.ID_814D32E25620602B3346FEE322896947:MESSAGE_AREA\.MSG_ID_ADV_3","DIV","ls.*", "DT_CREATED_CART_OUTPUT")

Call GetInputFromExcel(gstrInputExcelFilePathAndName,"Global",DataRowSet)

'Click on Edit Button
Call ClickSAPButton(1,".*","Edit","DIV",0)
Wait(3)
Call CaptureWebScreen(1,"Capture screen:DIV")

'Click on Order
Call ClickSAPButton(1,".*","Order","DIV",0)
Wait(5)
Call CaptureWebScreen(1,"Capture screen:DIV")

Call GetValueWebElementFrame(1, "Display Document:", "WDR_MESSAGE_AREA\.ID_814D32E25620602B3346FEE322896947:MESSAGE_AREA\.MSG_ID_ADV_2-text", "DIV", "ls.*", "DT_CREATED_CART_OUTPUT")
Call VerifyFrameWebElement(1, "", "Display Document:", "DIV", DT_CREATED_CART_OUTPUT, "ls.*", 0, False)

'Call VerifyFrameWebElement(1, "", "Display Document:", "SPAN", "Shopping cart """&DT_NAME_OF_SHOPPING_CART&""" with number "&DT_SC_NUMBER&" ordered successfully", "lsTextView.*", "", False)

Call ClickFrameSAPButton(1,"Display Document:","Close","DIV",0)
Wait 10

''------------------------'Log Off From Applicaton--------------------------------

Call  LogoffSRM(0)
Call FinalStatus ()


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


Public Function GetValueWebEdit(creationTime,Webclass,htmltag,WebEditName,WebEditIndex,ExcelColumnName)
If Not (Environment.Value("blnFatalError") or ExcelColumnName = DS_SKIP) Then
   If blnShowNotification Then BalloonTooltip.Show "TASE Automation","Executing Now : GetValueWebEdit"
strStepName = "Get value from the WebEdit"
ALMvalue=getEnvironmentVariable("ALMFlag")
If  ALMvalue=1 Then
	PopUpCreationtime=Cint(creationTime) + 1
else
	PopUpCreationtime=Cint(creationTime)
End If
Dim objWebedit
If Webclass<> ""  and WebEditName <> "" and htmltag<>"" Then
'	columnNumber=Cint(columnNumber)
'	rowNumber=Cint(rowNumber)
On error resume next
Set objWebedit =Browser("CreationTime:="& PopUpCreationtime).Page("Index:="& PopUpCreationtime).WebEdit("html tag:="&htmltag,"class:="&Webclass,"name:=.*"&WebEditName&".*","index:="&WebEditIndex)
 If  objWebedit.Exist  then 
	 objWebedit.Highlight
   	 val=objWebedit.GetROProperty("value")
	 
      If blnCaptureFlag  or  blnCreateImageEachStep or blnCreateTrainingDoc Then
                   ImagePath=CaptureScreenshot(strStepName,objWebedit,False,False,False)
      End if
     
      strStatus = "DONE"
      strMsg="value captured : "& val
      Call ReporterFunction(strLibraryFileName,"GetValueWebEdit","2",strStepName,strMsg)
 else
			  blnObjectError=True
              strStatus = "FAIL"    
              strMsg ="Object Missing"
              Call ReporterFunction(strLibraryFileName,"GetValueWebEdit","1",strStepName,strMsg)
 end if
else
			 
              strStatus = "FAIL"
              strMsg = "Function Parameter Not Passed Properly. Check the --GetValueWebEdit -- Function Call"
              Call ReporterFunction(strLibraryFileName,"GetValueWebEdit","1",strStepName,strMsg)
 end if
If  blnObjectError  Then
       Environment.Value("blnFatalError")=True
End If

If strStatus = "FAIL"  Then
       GetValueWebEdit = strMsg
       blnMainFailFlag = True
       ImagePath=CaptureScreenshot(strStepName,objWebedit,False,False,False)
Else
       GetValueWebEdit = True
End If
If blnWriteDataToOutputSheet Then
	   strStepName = "Retrieve '"&val&"' value in table '"&gstrOutputSheetName&"' sheet under column "&excelColumnName
	   Call WriteRunTimeDataToExcel (excelColumnName,val)
   ELSE
	   strStepName = "Retrieve '"&val&"' value in WebEdit"
	 End If
If blnDefault_eSwiftReporting Then  
       Call UpdateResultHtml(strStepName,ExcelColumnName,strMsg,strStatus,"")
End If

   Set objWebedit=Nothing
   End If
End Function

'******************************************************************************************************************



