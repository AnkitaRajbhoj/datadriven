Services.StartTransaction "tr1"

'Datatable.AddSheet "Module"
'Datatable.ImportSheet "E:\capgemini\KeywordDrivenFramework\Organizer\organizer.xlsx",1,"Module"

mrowcount=datatable.GetSheet("Action1").GetRowCount
msgbox mrowcount

For  i = 1 To mrowcount Step 1
	
Datatable.SetCurrentRow(i)

Modexe=Datatable("Moduleexe","Action1")

'msgbox Modexe
If modexe="Y" Then
	
	Modid=Datatable("ModuleID","Action1")
	
	msgbox Modid
	
	trowcount=datatable.GetSheet("Action2").GetRowCount
	
	msgbox trowcount
	
	For j=1  To trowcount Step 1
	Datatable.SetCurrentRow(j)
	
	If Modid=Datatable("ModuleID","Action2") and Datatable("Testcaseexe","Action2")="Y" Then
	testcaseid=Datatable("TestcaseId","Action2")
	msgbox testcaseid
		tsrowcount=Datatable.GetSheet("Action3").GetRowCount
		msgbox tsrowcount
		
		For k = 1 To tsrowcount Step 1
			
		datatable.SetCurrentRow(k)
		
		If testcaseid=Datatable("TestcaseId","Action3") Then
			
		keyword=Datatable("Keyword","Action3")
		msgbox keyword
		
		Select case (keyword)
		
		Case "ln"
		Call Login("john","hp")
		
		Case "ca"
		Call Closeapp()
		
		Case "oo"
		Call Openorder()
		
		Case "uo"
		Call Updateorder()
		
		Case "lnd"
		
		drowcount=datatable.GetSheet("Action4").GetRowCount
		
		For l = 1 To drowcount Step 1
			
			datatable.SetCurrentRow(l)
			
			Call Login(datatable("username","Action4"),datatable("password","Action4"))
			
			Call CloseApp()
		Next
		
		Case "ood"
		orrowcount=datatable.GetSheet("Action4").GetRowCount
		For m = 1 To orrowcount Step 1
			datatable.SetCurrentRow(m)
			Call openorder(datatable("orderno","Action4"))
		Next
		
		End  Select
		
		End If
		Next
	
	End If
	Next
	
End If 

Next

Services.EndTransaction "tr1"


 @@ hightlight id_;_1912616056_;_script infofile_;_ZIP::ssf12.xml_;_


 @@ hightlight id_;_1378288_;_script infofile_;_ZIP::ssf11.xml_;_


