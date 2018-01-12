'Test console availabilitiy
Dim internals_
internals_=Array("","",,Empty,InStr(1,Right(WScript.FullName,12),"\CScript.exe",1)=1) 'input line, multiline statement,result, multiline stack, console mode
Do
	WScript.StdOut.Write "> "
	internals_(0)=WScript.StdIn.ReadLine
	If StrComp(Trim(internals_(0)),"End",1)=0 Then
		Exit Do
	ElseIf internals_(0)="" Then
		WScript.Echo "stmt     execute VBScript statement"
		WScript.Echo "?expr    evaluate and dump expression"
		WScript.Echo "End      exit"
	ElseIf Instr(internals_(0), "?")=1 Then
		On Error Resume Next
		internals_(2)=Eval(Mid(internals_(0),2))
		If Err.Number Then
			WScript.Echo "Error " & Err.Number & ": " & Err.Description
		Else
			internals_(1)="["&TypeName(internals_(2))&"]"
			Select Case VarType(internals_(2))
			Case 2, 3, 4, 5, 6, 7, 8, 11, 14, 17
				internals_(1)=internals_(1)&" "&internals_(2)
			Case Else
			End Select
			WScript.Echo internals_(1)
		End If
		internal_(2)=""
		internal_(3)=Empty
	Else
		
		'Check constructions: Sub, Function, If, Select, Do, While
		'If Instr(code_stmt, ":")=0 And Instr(code_stmt, "Sub ")=1 Then code_stk.Add Split(code_stmt)(0)
		'If Instr(code_stmt, ":")=0 And Instr(code_stmt, "Function ")=1 Then code_stk.Add Split(code_stmt)(0)
		'If Instr(code_stmt, ":")=0 And Instr(code_stmt & " ", "Do ")=1 Then code_stk.Add Split(code_stmt)(0)
		'If Instr(code_stmt, ":")=0 And Instr(code_stmt, "While ")=1 Then code_stk.Add Split(code_stmt)(0)
		'If Instr(code_stmt, ":")=0 And Instr(code_stmt, "If ")=1 Then code_stk.Add Split(code_stmt)(0)
		'If Instr(code_stmt, ":")=0 And Instr(code_stmt, "Select ")=1 Then code_stk.Add Split(code_stmt)(0)
		'If instr(code_stmt, "End Sub") Then If code_stk.Count>0 Then (code_stk)
		On Error Resume Next
		Execute internals_(0)
		If Err.Number Then
			WScript.Echo "Error " & Err.Number & ": " & Err.Description
		End If
	End If
Loop