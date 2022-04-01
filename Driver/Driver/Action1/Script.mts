'Datatable.AddSheet "Module"
'Datatable.ImportSheet "C:\Users\sfjbs\Desktop\KeyboardDrivenFramework\Organizer\Organizer.xlsx" ,1,"Module"
'Module
'Start Transaction-Shows the time required to complete the entire test case execution
Services.StartTransaction "Tran1"

mrowcount=datatable.GetSheet("Action1").GetRowCount
msgbox mrowcount

For i= 1 To mrowcount  Step 1
Datatable.SetCurrentRow(i)

ModExec=Datatable("Moduleexe","Action1")
'msgbox ModExec
If ModExec="Y" Then
	ModID=Datatable("ModuleID","Action1")
	msgbox ModID

'TestCase
trowcount=datatable.GetSheet("Action2").GetRowCount
msgbox trowcount

For j= 1 To trowcount  Step 1
Datatable.SetCurrentRow(j)

If ModID=Datatable("ModuleID","Action2") and Datatable("Testcaseexe","Action2")="Y" Then
TestcID=Datatable("TestcaseId","Action2")
msgbox TestcID
'Scenario
tsrowcount= Datatable.GetSheet("Action3").GetRowCount
msgbox tsrowcount

For k= 1 To tsrowcount Step 1
datatable.SetCurrentRow(k)
If TestcID=Datatable("TestcaseId","Action3") Then
	keyword=Datatable("Keyword","Action3")
	msgbox keyword
	'Based on keyword display the TestCaseID
	select  case (keyword)
	Case "ln"
	Call login("john","hp")
	
	Case"ca"
	Call Closeapp()
	
	Case"oo"
	Call openorder("5")
	
	Case"uo"
	Call updateorder()
	
	Case "lnd"
	
	drowcount=datatable.GetSheet("Action4").GetRowCount
	For l = 1 To drowcount Step 1
		datatable.SetCurrentRow(l)
		Call Login(datatable("username","Action4"),datatable("password","Action4"))
		Call Closeapp()
	Next
	
	Case "ood"
 @@ hightlight id_;_13173792_;_script infofile_;_ZIP::ssf12.xml_;_
 @@ hightlight id_;_2128468040_;_script infofile_;_ZIP::ssf15.xml_;_
	orrowcount=datatable.GetSheet("Action4").GetRowCount
	For m = 1 To orrowcount Step 1
	datatable.SetCurrentRow(m)
	Call openorder(datatable("orderno","Action4"))
		
	Next
	End select 
	
End If
	
Next
	
End If
Next

End If
Next


Services.EndTransaction "Tran1"
 @@ hightlight id_;_1886165680_;_script infofile_;_ZIP::ssf4.xml_;_
 @@ hightlight id_;_1881744976_;_script infofile_;_ZIP::ssf10.xml_;_
 @@ hightlight id_;_1981894232_;_script infofile_;_ZIP::ssf11.xml_;_
 @@ hightlight id_;_1886182528_;_script infofile_;_ZIP::ssf7.xml_;_

