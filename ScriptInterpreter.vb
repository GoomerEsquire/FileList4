Imports System.IO
Imports System.Text

Class ScriptLine

	Shared Operator +(ary() As ScriptLine, obj As ScriptLine) As ScriptLine()

		Array.Resize(ary, ary.Count + 1)
		ary(ary.Count - 1) = obj
		Return ary

	End Operator

	Protected s As String
	Protected o, l As Integer
	Sub New(str As String, originalLine As Integer, line As Integer)
		s = str
		l = line
		o = originalLine
	End Sub

	Public Overrides Function ToString() As String
		Return s
	End Function

	ReadOnly Property Line As Integer
		Get
			Return l
		End Get
	End Property

	ReadOnly Property OriginalLine As Integer
		Get
			Return o
		End Get
	End Property

End Class

Class Argument

	Shared Operator +(ary() As Argument, obj As Argument) As Argument()

		Array.Resize(ary, ary.Count + 1)
		ary(ary.Count - 1) = obj
		Return ary

	End Operator

	Public Enum ArgType
		Var
		Str
	End Enum

	Protected val As String
	Protected t As ArgType
	Sub New(value As String, objectType As ArgType)
		val = value
		t = objectType
	End Sub

	ReadOnly Property Value As String
		Get
			Return val
		End Get
	End Property

	ReadOnly Property Type As ArgType
		Get
			Return t
		End Get
	End Property

	Function Resolve() As String
		If Not t = ArgType.Var Then
			Return val
		ElseIf ActiveScript.isFixedVar(val) Then
			t = ArgType.Str
			val = ActiveScript.GetFixedVar(val)
		Else
			Dim varobj As Variable = ActiveScript.GetVarObj(val)
			If varobj Is Nothing Then
				ActiveScript.DisplayError("Variable not declared: " + Chr(34) + val + Chr(34))
			Else
				t = ArgType.Str
				val = varobj.Value
			End If
		End If
		Return val
	End Function

	Public Overrides Function ToString() As String
		Return val
	End Function

End Class

Class Variable

	Shared Operator +(ary() As Variable, obj As Variable) As Variable()

		Array.Resize(ary, ary.Count + 1)
		ary(ary.Count - 1) = obj
		Return ary

	End Operator

	Enum varScope
		Unknown
		Priv
		Pub
	End Enum

	Protected n As String
	Protected v As String
	Protected s As varScope

	Sub New(name As String, scope As varScope)
		n = name
		s = scope
	End Sub

	Property Value As String
		Get
			Return v
		End Get
		Set(value As String)
			v = value
		End Set
	End Property

	ReadOnly Property Name As String
		Get
			Return n
		End Get
	End Property

	Public Overrides Function ToString() As String
		Return v
	End Function

End Class

Class GotoInfo

	Shared Operator +(ary() As GotoInfo, obj As GotoInfo) As GotoInfo()

		Array.Resize(ary, ary.Count + 1)
		ary(ary.Count - 1) = obj
		Return ary

	End Operator

	Protected id As String
	Protected pos As Integer

	Sub New(name As String, position As Integer)
		id = name
		pos = position
	End Sub

	ReadOnly Property Name As String
		Get
			Return id
		End Get
	End Property

	ReadOnly Property Position As Integer
		Get
			Return pos
		End Get
	End Property

	Public Overrides Function ToString() As String
		Return id
	End Function

End Class

Class SubInfo

	Shared Operator +(ary() As SubInfo, obj As SubInfo) As SubInfo()

		Array.Resize(ary, ary.Count + 1)
		ary(ary.Count - 1) = obj
		Return ary

	End Operator

	Protected scriptobj As ScriptInterpreter
	Protected vars As Variable() = {}
	Protected argsArray As Argument() = {}
	Protected n, linestr As String
	Protected pos As Integer = 0
	Protected loc(1) As Integer
	Protected gotoArray As GotoInfo() = {}
	Protected ifBlocks As IfBlock() = {}
	Sub New(script As ScriptInterpreter, section As String)
		scriptobj = script
		n = section
	End Sub

	Property Location As Integer()
		Get
			Return loc
		End Get
		Set(value As Integer())
			loc = value
		End Set
	End Property

	ReadOnly Property Name As String
		Get
			Return n
		End Get
	End Property

	ReadOnly Property Script As ScriptInterpreter
		Get
			Return scriptobj
		End Get
	End Property

	Property Arguments As Argument()
		Get
			Return argsArray
		End Get
		Set(value As Argument())
			argsArray = value
		End Set
	End Property

	Sub AddGotoObj(name As String, position As Integer)

		Array.Resize(gotoArray, gotoArray.Count + 1)
		gotoArray(gotoArray.Count - 1) = New GotoInfo(name, position)

	End Sub

	ReadOnly Property GotoObjects As GotoInfo()
		Get
			Return gotoArray
		End Get
	End Property

	ReadOnly Property Variables() As Variable()
		Get
			Return vars
		End Get
	End Property

	Function NewVar(name As String) As Variable

		Dim var As Variable = ActiveScript.GetVarObj(name)
		If var IsNot Nothing Then
			ActiveScript.DisplayError("Variable already declared: " + Chr(34) + var.Name + Chr(34) + "!")
			Return Nothing
		End If

		var = New Variable(name, Variable.varScope.Priv)
		vars += var
		Return var

	End Function

	Sub ClearVariables()
		vars = {}
	End Sub

	Property CurPos As Integer
		Get
			Return pos
		End Get
		Set(value As Integer)
			pos = value
		End Set
	End Property

	Function GetIfBlockObj(Optional lineNumber As Integer = -1) As IfBlock

		If lineNumber = -1 Then lineNumber = pos

		For Each ifBlockObj As IfBlock In ifBlocks
			If ifBlockObj.StartPoints(0) = lineNumber Then Return ifBlockObj
		Next
		Return Nothing

	End Function

	Sub AddIfBlock(ifBlockObj As IfBlock)

		Array.Resize(ifBlocks, ifBlocks.Count + 1)
		ifBlocks(ifBlocks.Count - 1) = ifBlockObj

	End Sub

	Public Overrides Function ToString() As String
		Return n
	End Function

End Class

Class IfBlock

	Shared Operator +(ary() As IfBlock, obj As IfBlock) As IfBlock()

		Array.Resize(ary, ary.Count + 1)
		ary(ary.Count - 1) = obj
		Return ary

	End Operator

	Protected startpos(), endpos As Integer

	Sub New(startPoint As Integer)
		startpos = {startPoint}
	End Sub

	Property EndPoint As Integer
		Get
			Return endpos
		End Get
		Set(value As Integer)
			endpos = value
		End Set
	End Property

	ReadOnly Property StartPoints As Integer()
		Get
			Return startpos
		End Get
	End Property

	Sub AddElse(pos As Integer)

		Array.Resize(startpos, startpos.Count + 1)
		startpos(startpos.Count - 1) = pos

	End Sub

End Class

Class ScriptInterpreter

	Shared Operator +(ary() As ScriptInterpreter, obj As ScriptInterpreter) As ScriptInterpreter()

		Array.Resize(ary, ary.Count + 1)
		ary(ary.Count - 1) = obj
		Return ary

	End Operator

	Public Overrides Function ToString() As String
		Return scriptPath
	End Function

	Public Function Compact(lines As String()) As ScriptLine()

		ScriptLines = {}
		Dim int As Integer = 1

		For i As Integer = 1 To lines.Count
			Dim l As String = lines(i - 1).Trim
			If l.Length = 0 Then Continue For
			If l(0) = Chr(ASCII.Apostrophe) Then Continue For
			ScriptLines += New ScriptLine(l, i, int)
			int += 1
		Next

		Return ScriptLines

	End Function

	Public Function TranslatePos(line As Integer) As Integer

		For Each lineObj As ScriptLine In ScriptLines
			If lineObj.Line = line Then Return lineObj.OriginalLine
		Next
		Return 0

	End Function

	Protected scriptPath As String
	Protected ScriptLines As ScriptLine() = {}
	Sub New(path As String, allowFL4 As Boolean)
		Dim newsub As New SubInfo(Me, "Init")
		callStack += newsub
		ActiveSub = newsub
		ActiveScript = Me
		If LCase(path) = "fl4" Then
			ScriptLines = Compact(Split(My.Resources.fl4, vbCrLf))
			scriptPath = path
		Else
			Dim finfo As New FileInfo(path)
			If Not allowFL4 And LCase(finfo.Extension) = ".fl4" Then
				DisplayError("No FL4 files allowed beyond this point! Call FLS files instead.")
				Return
			End If
			scriptPath = finfo.FullName
			If Not finfo.Exists Then
				DisplayError("Could not find file: " + Chr(34) + scriptPath + Chr(34) + "!")
				Return
			End If
			ScriptLines = Compact(File.ReadAllLines(scriptPath, System.Text.Encoding.Default))
		End If
		CompileInfo()
		Array.Resize(callStack, callStack.Count - 1)
	End Sub

	ReadOnly Property FilePath As String
		Get
			Return scriptPath
		End Get
	End Property

	Protected OnError As ErrorMode = ErrorMode.Bail
	Protected ActiveSub As SubInfo
	Protected changeLine As Integer
	Protected curFile As FileInfo
	Protected curDir As DirectoryInfo
	Protected sWatch As New Stopwatch
	Protected enc As System.Text.Encoding = System.Text.Encoding.Default
	Protected ReturnToParent As Boolean = False

	Enum ErrorMode
		Bail
		Pause
		KeepGoing
	End Enum

	Sub DisplayError(text As String)

		If MainScript Is Nothing Then MainScript = Me

		Dim newCallStack As String() = {}
		If callStack.Count = 0 Then
			newCallStack = {"<EMPTY>"}
		Else
			For Each _sub As SubInfo In callStack
				Array.Resize(newCallStack, newCallStack.Count + 1)
				If _sub.CurPos = 0 Then
					newCallStack(newCallStack.Count - 1) = _sub.Name + " (" + _sub.Script.scriptPath + ")"
				Else
					newCallStack(newCallStack.Count - 1) = _sub.Name + " (" + _sub.Script.scriptPath + ", Line " + CStr(TranslatePos(_sub.CurPos)) + ")"
				End If
			Next
			Array.Reverse(newCallStack)
		End If
		Dim output As String = vbCrLf + StrDup(Console.BufferWidth, "*") + text + vbCrLf + vbCrLf + "Call stack:" + vbCrLf + StrDup(11, "-") + vbCrLf + String.Join(vbCrLf, newCallStack) + vbCrLf + StrDup(Console.BufferWidth, "*")
		Console.WriteLine(output)

		ErrorsHappened = True

		If MainScript.OnError = ErrorMode.Pause Then
			Console.WriteLine("Continue? (Y: Yes, N: No, 1: Yes for all)")
			Dim kinfo As ConsoleKeyInfo = Console.ReadKey(True)
			Do While Not (kinfo.Key = ConsoleKey.Y OrElse kinfo.Key = ConsoleKey.N OrElse kinfo.Key = ConsoleKey.D1)
				kinfo = Console.ReadKey(True)
			Loop
			If kinfo.Key = ConsoleKey.Y Then
				Return
			ElseIf kinfo.Key = ConsoleKey.D1 Then
				MainScript.OnError = ErrorMode.KeepGoing
				Return
			End If
		End If

		If Not MainScript.OnError = ErrorMode.KeepGoing Then
			Environment.Exit(3)
		End If

	End Sub

	Function NewVar(name As String) As Variable

		Dim var As Variable = GetVarObj(name, Variable.varScope.Pub)
		If var IsNot Nothing Then
			DisplayError("Variable already declared: " + Chr(34) + var.Name + Chr(34) + "!")
			Return var
		End If

		var = New Variable(name, Variable.varScope.Pub)
		vars += var
		Return var

	End Function

	Sub CompileInfo()

		Dim openSub As SubInfo = Nothing
		Dim loc0 As Integer
		Dim ifLevel As Integer = 0
		Dim openIfBlock As IfBlock = Nothing

		'gather info for quick jumping and vars
		For i As Integer = 0 To ScriptLines.Count - 1
			Dim line As String = ScriptLines(i).ToString
			ActiveSub.CurPos = i + 1
			If Not openSub Is Nothing AndAlso line(0) = Chr(ASCII.Colon) Then
				openSub.AddGotoObj(line.Substring(1, line.Length - 1), i + 1)
			Else
				Dim argArray() As String = Split(line)
				Dim command As String = argArray(0)
				Dim argstring As String = String.Empty
				If argArray.Count > 1 Then
					argstring = String.Join(Chr(ASCII.Space), argArray, 1, argArray.Count - 1)
				Else
					Continue For
				End If

				Select Case LCase(command)
					Case "sub"
						If Not openSub Is Nothing Then
							DisplayError("Sub not closed: " + Chr(34) + openSub.Name + Chr(34) + "!")
							Return
						End If
						Dim args As Argument() = ParseArgs(argstring)
						If args.Count > 0 Then
							Dim s1 As SubInfo = GetSub(Me, args(0).Value)
							If s1 Is Nothing Then
								openSub = New SubInfo(Me, args(0).Value)
							ElseIf Not s1.Location(0) = Nothing Then
								DisplayError("Sub already exists: " + Chr(34) + args(0).Value + Chr(34) + "!")
								Return
							End If
							loc0 = i
							If args.Count > 1 Then
								Array.ConstrainedCopy(args, 1, args, 0, args.Count - 1)
								Array.Resize(args, args.Count - 1)
								openSub.Arguments = args
							End If
						ElseIf args.Count < 1 Then
							DisplayError("Missing arguments for Sub!")
							Return
						End If

					Case "end"
						Dim args As Argument() = ParseArgs(argstring)
						If args.Count = 1 Then
							If LCase(args(0).Value) = "sub" Then
								If openSub Is Nothing Then
									DisplayError("End Sub must follow Sub!")
									Return
								End If
								openSub.Location = {loc0, i}
								subArray += openSub
								openSub = Nothing
							ElseIf LCase(args(0).Value) = "if" Then
								If openIfBlock Is Nothing Then
									DisplayError("End If must follow If/IfNot!")
									Return
								End If

								openIfBlock = Nothing
							End If
						ElseIf openSub Is Nothing Then
							DisplayError("Invalid use of End!")
							Return
						End If

					Case "var"
						If openSub Is Nothing Then
							Dim args As Argument() = ParseArgs(argstring)
							If args.Count = 2 Then
								NewVar(args(0).Value).Value = args(1).Resolve
							ElseIf args.Count = 1 Then
								NewVar(args(0).Value).Value = String.Empty
							End If
						End If

					Case "call"
						Dim args As Argument() = ParseArgs(argstring)
						If Not LCase(args(0).Value) = "this" Then
							scriptArray += New ScriptInterpreter(args(0).Value, False)
						End If

					Case "function"
						Dim args As Argument() = ParseArgs(argstring)
						If Not LCase(args(1).Value) = "this" Then
							scriptArray += New ScriptInterpreter(args(1).Value, False)
						End If

					Case "onerror"
						Dim args As Argument() = ParseArgs(argstring)
						If args.Count < 1 Then
							DisplayError("Missing arguments for OnError!")
							Return
						ElseIf args.Count > 1 Then
							DisplayError("Too many arguments for OnError!")
							Return
						End If
						If LCase(args(0).Value) = "pause" Then
							OnError = ErrorMode.Pause
						ElseIf LCase(args(0).Value) = "end" Then
							OnError = ErrorMode.Bail
						ElseIf LCase(args(0).Value) = "continue" Then
							OnError = ErrorMode.KeepGoing
						Else
							DisplayError("Invalid argument: " + Chr(34) + args(0).Value + Chr(34) + "!")
							Return
						End If
				End Select
			End If
		Next

		If Not openSub Is Nothing Then
			DisplayError("Sub not closed: " + Chr(34) + openSub.Name + Chr(34) + "!")
		End If

	End Sub

	Overloads Function GetSub(name As String) As SubInfo

		name = LCase(name)
		For Each s As SubInfo In subArray
			If LCase(s.Name) = name Then
				Return s
			End If
		Next
		Return Nothing

	End Function

	Overloads Function GetSub(script As ScriptInterpreter, name As String) As SubInfo

		name = LCase(name)
		For Each s As SubInfo In subArray
			If Not s.Script Is script Then
				Continue For
			End If
			If LCase(s.Name) = name Then
				Return s
			End If
		Next
		Return Nothing

	End Function

	Public Function GetVarObj(name As String, Optional scope As Variable.varScope = Variable.varScope.Unknown) As Variable

		name = LCase(name)
		If scope = Variable.varScope.Priv OrElse scope = Variable.varScope.Unknown Then
			For Each var As Variable In ActiveSub.Variables
				If LCase(var.Name) = name Then Return var
			Next
		End If

		If scope = Variable.varScope.Pub OrElse scope = Variable.varScope.Unknown Then
			For Each var As Variable In vars
				If LCase(var.Name) = name Then Return var
			Next
		End If

		Return Nothing

	End Function

	Private Function getGotoPos(name As String) As Integer

		name = LCase(name)
		For Each gotoObj As GotoInfo In ActiveSub.GotoObjects
			If LCase(gotoObj.Name) = name Then Return gotoObj.Position
		Next
		Return Nothing

	End Function

	Public Enum ASCII
		Space = 32
		Plus = 43
		Quote = 34
		Apostrophe = 39
		Colon = 58
		SmallerThan = 60
		Equals = 61
		GreaterThan = 62
		Backslash = 92
	End Enum

	Private Function ParseArgs(source As String) As Argument()

		If source.Length = 0 Then Return {}

		Dim stringOpen As Boolean = False
		Dim resolve As Boolean = False
		Dim temp, tempArg As New StringBuilder
		Dim move As Integer = 0
		Dim curChar, nextChar As Char
		Dim newArg() As Argument = {}
		Dim argType As Argument.ArgType = Argument.ArgType.Var
		Dim escape As Boolean = False

		For i As Integer = 1 To source.Length
			If i > source.Length Then
				DisplayError("Unexpected end of argument!")
				Return Nothing
			End If
			curChar = source(i - 1)
			If i + 1 <= source.Length Then
				nextChar = source(i)
			Else
				nextChar = Nothing
			End If
			If curChar = Chr(ASCII.Backslash) Then
				If stringOpen Then tempArg.Append(curChar)
				escape = Not escape
				argType = Argument.ArgType.Str
			ElseIf curChar = Chr(ASCII.Quote) Then
				If Not escape Then
					stringOpen = Not stringOpen
				Else
					tempArg.Append(curChar)
					escape = False
				End If
				argType = Argument.ArgType.Str
			ElseIf stringOpen Then
				tempArg.Append(curChar)
				escape = False
			ElseIf curChar = Chr(ASCII.Space) AndAlso Not nextChar = Chr(ASCII.Plus) Then
				If argType = Argument.ArgType.Str OrElse temp.Length > 0 Then
					tempArg.Append(temp)
					temp.Clear()
				End If
				newArg += New Argument(tempArg.ToString, argType)
				argType = Argument.ArgType.Var
				tempArg.Clear()
				resolve = False
			ElseIf nextChar = Nothing OrElse (curChar = Chr(ASCII.Space) AndAlso nextChar = Chr(ASCII.Plus)) Then
				If nextChar = Nothing Then
					temp.Append(curChar)
				Else
					move = +2
					resolve = True
					argType = Argument.ArgType.Str
				End If
				If temp.Length > 0 Then
					If resolve Then
						If isFixedVar(temp.ToString) Then
							temp.Replace(temp.ToString, GetFixedVar(temp.ToString))
						Else
							Dim tempvar As Variable = GetVarObj(temp.ToString)
							If Not tempvar Is Nothing Then
								temp.Replace(temp.ToString, tempvar.Value)
							Else
								DisplayError("Variable not declared: " + Chr(34) + temp.ToString + Chr(34) + "!")
								Return Nothing
							End If
						End If
					End If
					tempArg.Append(temp)
					temp.Clear()
				End If
			Else
				temp.Append(curChar)
			End If
			i += move
			move = 0
		Next

		If argType = Argument.ArgType.Str OrElse tempArg.Length > 0 Then
			newArg += New Argument(tempArg.ToString, argType)
		End If

		Return newArg

	End Function

	Public keyWords() As String = {"set", "var", "swstop", "swstart", "if", "ifnot", "return", "nextdir", "next", "goto", "print", "write", "replace", "lcase", "ucase", "add", "subtract", "multiply", "divide", "round", "floor", "ceil", "char", "format", "padleft", "padright", "len", "substring", "call", "function", "beep", "title", "readkey", "maxcommands", "fileloop", "end", "encoding", "sleep", "=", "<", ">", "like", "continue", "onerror", "pause", "overwrite", "this", "fl4"}

	Function Run(subName As String, dInfo As DirectoryInfo, fInfo As FileInfo, Optional arguments As Argument() = Nothing) As String

		ActiveScript = Me
		ActiveSub = GetSub(Me, subName)

		If ActiveSub Is Nothing Then
			DisplayError("Sub not found: " + Chr(34) + subName + Chr(34) + "!")
			Return Nothing
		End If

		callStack += ActiveSub

		Dim returnVal As String = Nothing
		curFile = fInfo
		curDir = dInfo
		ReturnToParent = False

		If arguments Is Nothing Then arguments = {}
		If arguments.Count < ActiveSub.Arguments.Count Then
			DisplayError("Missing arguments for sub " + subName + ": Got " + CStr(arguments.Count) + " - Need " + CStr(ActiveSub.Arguments.Count))
			Return Nothing
		ElseIf arguments.Count > ActiveSub.Arguments.Count Then
			DisplayError("Too many arguments for sub " + subName + ": Got " + CStr(arguments.Count) + " - Need " + CStr(ActiveSub.Arguments.Count))
			Return Nothing
		End If
		For argIndex As Integer = 0 To ActiveSub.Arguments.Count - 1
			ActiveSub.NewVar(ActiveSub.Arguments(argIndex).Value).Value = arguments(argIndex).Resolve
		Next

		If ActiveSub.Location(1) - ActiveSub.Location(0) > 1 Then
			For i As Integer = ActiveSub.Location(0) + 1 To ActiveSub.Location(1) - 1
				changeLine = 0
				Dim line As String = ScriptLines(i).ToString
				If line(0) = Chr(ASCII.Colon) Then Continue For
				If CommandsProcessed >= MaxCommands Then
					DisplayError("Maximum number of commands reached!")
					Exit For
				End If
				ActiveSub.CurPos = i + 1
				CommandsProcessed += 1
				Dim argArray() As String = Split(line)
				Dim command As String = argArray(0)
				Dim args As String = String.Empty
				If argArray.Count > 1 Then
					args = String.Join(Chr(ASCII.Space), argArray, 1, argArray.Count - 1)
				End If

				Select Case LCase(command)
					Case "set"
						SetVar(ParseArgs(args), False)
					Case "var"
						SetVar(ParseArgs(args), True)
					Case "swstop"
						sWatch.Stop()
					Case "swstart"
						sWatch.Restart()
					Case "if"
						IfStatement(ParseArgs(args), False)
					Case "ifnot"
						IfStatement(ParseArgs(args), True)
					Case "return"
						returnVal = _Return(ParseArgs(args))
						ReturnToParent = True
					Case "nextdir"
						If Not InLoop Then DisplayError("NextDir cannot be called outside of a file loop!")
						NextDir = True
					Case "next"
						If Not InLoop Then DisplayError("Next cannot be called outside of a file loop!")
						NextFile = True
					Case "goto"
						Jump(ParseArgs(args))
					Case "print"
						_print(ParseArgs(args))
					Case "write"
						_write(ParseArgs(args))
					Case "replace"
						Rep(ParseArgs(args))
					Case "lcase"
						lcas(ParseArgs(args))
					Case "ucase"
						ucas(ParseArgs(args))
					Case "add"
						Add(ParseArgs(args))
					Case "subtract"
						Subtract(ParseArgs(args))
					Case "multiply"
						Multiply(ParseArgs(args))
					Case "divide"
						Divide(ParseArgs(args))
					Case "round"
						Round(ParseArgs(args))
					Case "floor"
						Floor(ParseArgs(args))
					Case "ceil"
						Ceil(ParseArgs(args))
					Case "char"
						GetC(ParseArgs(args))
					Case "format"
						FormatN(ParseArgs(args))
					Case "padleft"
						Pad(ParseArgs(args), True)
					Case "padright"
						Pad(ParseArgs(args), False)
					Case "len"
						Len(ParseArgs(args))
					Case "substring"
						SubString(ParseArgs(args))
					Case "call"
						CallSection(ParseArgs(args))
					Case "function"
						CallFunction(ParseArgs(args))
					Case "beep"
						Media.SystemSounds.Beep.Play()
					Case "title"
						title(ParseArgs(args))
					Case "readkey"
						ReadKey(ParseArgs(args))
					Case "maxcommands"
						SetMaxCommands(ParseArgs(args))
					Case "fileloop"
						NewLoop(ParseArgs(args))
					Case "end"
						Environment.Exit(1)
					Case "encoding"
						SetEncoding(ParseArgs(args))
					Case "sleep"
						Wait(ParseArgs(args))
					Case Else
						DisplayError("Invalid command: " + Chr(34) + command + Chr(34) + "!")
				End Select

				i += changeLine
				If NextDir OrElse NextFile Then Exit For
				If ReturnToParent Then Exit For
			Next

			ActiveSub.ClearVariables()
		End If

		Array.Resize(callStack, callStack.Count - 1)
		If callStack.Length > 0 Then
			ActiveSub = callStack(callStack.Count - 1)
			ActiveScript = ActiveSub.Script
		End If

		Return returnVal

	End Function

	Public fixedVars() As String = {"newline", "dircount", "filecount", "accessdenied", "filepath", "filename", "filenamenx", "size", "folder", "path", "initpath", "exepath", "year", "month", "day", "hour", "minute", "second", "msecond", "utcoffset", "maxcommands", "swtime", "proccommands", "bufferwidth", "bufferheight", "fileextension"}

	Function GetFixedVar(varName As String) As String

		Select Case LCase(varName)
			Case "newline"
				Return vbCrLf
			Case "dircount"
				If AccessDenied Then
					DisplayError("Access denied in " + Chr(34) + curDir.FullName + Chr(34) + "!")
				Else
					Return CStr(curDir.GetDirectories.Count)
				End If
			Case "filecount"
				If AccessDenied Then
					DisplayError("Access denied in " + Chr(34) + curDir.FullName + Chr(34) + "!")
				Else
					Return CStr(curDir.GetFiles.Count)
				End If
			Case "accessdenied"
				Return AccessDenied.ToString
			Case "filepath"
				If curDir Is Nothing Then
					DisplayError("No directory-object!")
				ElseIf curFile Is Nothing Then
					DisplayError("No file-object in " + Chr(34) + curDir.FullName + Chr(34) + "!")
				Else
					Return curFile.FullName
				End If
			Case "filename"
				If curDir Is Nothing Then
					DisplayError("No directory-object!")
				ElseIf curFile Is Nothing Then
					DisplayError("No file-object in " + Chr(34) + curDir.FullName + Chr(34) + "!")
				Else
					Return curFile.Name
				End If
			Case "filenamenx"
				If curDir Is Nothing Then
					DisplayError("No directory-object!")
				ElseIf curFile Is Nothing Then
					DisplayError("No file-object in " + Chr(34) + curDir.FullName + Chr(34) + "!")
				Else
					Return Path.GetFileNameWithoutExtension(curFile.FullName)
				End If
			Case "fileextension"
				If curDir Is Nothing Then
					DisplayError("No directory-object!")
				ElseIf curFile Is Nothing Then
					DisplayError("No file-object in " + Chr(34) + curDir.FullName + Chr(34) + "!")
				Else
					Return Path.GetExtension(curFile.FullName)
				End If

			Case "size"
				If curDir Is Nothing Then
					DisplayError("No directory-object!")
				ElseIf curFile Is Nothing Then
					DisplayError("No file-object in " + Chr(34) + curDir.FullName + Chr(34) + "!")
				Else
					Return CStr(curFile.Length)
				End If
			Case "folder"
				If curDir Is Nothing Then
					DisplayError("No directory-object!")
				Else
					Return curDir.Name
				End If
			Case "path"
				If curDir Is Nothing Then
					DisplayError("No directory-object!")
				Else
					Return curDir.FullName
				End If
			Case "initpath"
				Return InitPath(InitPath.Count - 1)
			Case "exepath"
				Return My.Application.Info.DirectoryPath
			Case "year"
				Return Format(Date.Now, "yyyy")
			Case "month"
				Return Format(Date.Now, "MM")
			Case "day"
				Return Format(Date.Now, "dd")
			Case "hour"
				Return Format(Date.Now, "HH")
			Case "minute"
				Return Format(Date.Now, "mm")
			Case "second"
				Return Format(Date.Now, "ss")
			Case "msecond"
				Return Format(Date.Now, "fffffff")
			Case "utcoffset"
				Return Format(Date.Now, "zz")
			Case "maxcommands"
				Return MaxCommands.ToString
			Case "swtime"
				Return sWatch.Elapsed.TotalMilliseconds.ToString
			Case "proccommands"
				Return CommandsProcessed.ToString
			Case "bufferwidth"
				Return Console.BufferWidth.ToString
			Case "bufferheight"
				Return Console.BufferHeight.ToString
		End Select

		Return Nothing

	End Function

	Function isKeyword(varName As String) As Boolean

		varName = LCase(varName)
		Return Array.IndexOf(keyWords, varName) > -1

	End Function

	Function isFixedVar(varName As String) As Boolean

		varName = LCase(varName)
		Return Array.IndexOf(fixedVars, varName) > -1

	End Function

	Private Function InvalidArgs(args As Argument(), KeywordsAt As Integer(), FixedArgsAt As Integer()) As String

		For Each i As Integer In KeywordsAt
			If args.Count - 1 < i Then Continue For
			If isKeyword(args(i).Value) Then Return args(i).Value
		Next

		For Each i As Integer In FixedArgsAt
			If args.Count - 1 < i Then Continue For
			If isFixedVar(args(i).Value) Then Return args(i).Value
		Next

		Return Nothing

	End Function

	Private Function TryCastDbl(source As String) As Double

		If source.Length = 0 Then
			DisplayError("Expression is zero-length string!")
			Return 0
		End If

		Try
			Return CDbl(source)
		Catch
			DisplayError("Expression is not a number or out of range: " + Chr(34) + source + Chr(34) + "!")
			Return Nothing
		End Try

	End Function

	Private Function TryCastInt(source As String) As Integer

		If source.Length = 0 Then
			DisplayError("Expression is zero-length string!")
			Return 0
		End If

		Try
			Return CInt(source)
		Catch
			DisplayError("Expression is not a number or out of range: " + Chr(34) + source + Chr(34) + "!")
			Return Nothing
		End Try

	End Function

#Region "Commands"

	Private Sub Wait(args() As Argument)

		If args.Count > 1 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 1 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0}, {0})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		args(0).Resolve()

		Dim num As Integer = TryCastInt(args(0).Value)
		Threading.Thread.Sleep(num)

	End Sub

	Private Function _Return(args() As Argument) As String

		If args.Count > 1 Then
			DisplayError("Too many arguments for function!")
			Return Nothing
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0}, {0})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return Nothing
		End If

		If args.Count > 0 Then
			Return args(0).Resolve
		Else
			Return Nothing
		End If

	End Function

	Private Sub ReadKey(args() As Argument)

		If args.Count > 1 Then
			DisplayError("Too many arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0}, {0})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim key As ConsoleKeyInfo = Console.ReadKey(True)

		If args.Count = 1 Then
			Dim var As Variable = GetVarObj(args(0).Value)
			If var Is Nothing Then
				DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
				Return
			End If
			var.Value = key.Key.ToString
		End If

	End Sub

	Private Sub SetEncoding(args() As Argument)

		If args.Count > 1 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 1 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0, 1}, {0, 1})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim val As String = LCase(args(0).Value)
		If val = "ascii" Then
			enc = System.Text.Encoding.ASCII
		ElseIf val = "utf8" Then
			enc = System.Text.Encoding.UTF8
		ElseIf val = "unicode" Then
			enc = System.Text.Encoding.Unicode
		ElseIf val = "utf32" Then
			enc = System.Text.Encoding.UTF32
		ElseIf val = "system" Then
			enc = System.Text.Encoding.Default
		Else
			DisplayError("Invalid argument for function!")
		End If

	End Sub

	Private Sub SetVar(args() As Argument, setNew As Boolean)

		Dim minargs As Integer = 2
		If setNew Then minargs = 1

		If args.Count > 2 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < minargs Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0, 1}, {0})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		If setNew Then
			If args.Count = 2 Then
				args(1).Resolve()
				ActiveSub.NewVar(args(0).Value).Value = args(1).Value
			ElseIf args.Count = 1 Then
				ActiveSub.NewVar(args(0).Value).Value = String.Empty
			End If
		Else
			Dim var As Variable = GetVarObj(args(0).Value)
			If var Is Nothing Then
				DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
				Return
			End If
			var.Value = args(1).Resolve
		End If

	End Sub

	Private Sub title(args() As Argument)

		If args.Count > 1 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 1 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0}, {})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		args(0).Resolve()

		Console.Title = args(0).Value

	End Sub

	Private Sub _print(args() As Argument)

		If args.Count > 1 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 1 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0}, {})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		args(0).Resolve()

		Console.Write(args(0).Value)

	End Sub

	Private Sub _write(args() As Argument)

		Dim minargs As Integer = 2
		Dim maxargs As Integer = 3
		If InLoop Then
			minargs = 1
			maxargs = 2
		End If

		If args.Count > maxargs Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < minargs Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String
		Dim append As Boolean = True

		If InLoop Then
			InvalidStr = InvalidArgs(args, {0}, {})
			args(0).Resolve()
		Else
			InvalidStr = InvalidArgs(args, {0, 1}, {1})
			If args.Count = 3 AndAlso LCase(args(2).Value) = "overwrite" Then
				append = False
			End If
			args(1).Resolve()
		End If

		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		If InLoop Then
			SW.Write(args(0).Value)
		Else
			Using SW As New StreamWriter(args(0).Value, append, enc)
				SW.Write(args(1).Value)
				SW.Close()
			End Using
		End If

	End Sub

	Private Sub Jump(args() As Argument)

		If args.Count > 1 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 1 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0}, {0})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim i As Integer = getGotoPos(args(0).Value)
		If i = Nothing Then
			DisplayError("The Goto-mark does not exist: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		Else
			changeLine = i - ActiveSub.CurPos
		End If

	End Sub

	Private Sub SetMaxCommands(args() As Argument)

		If args.Count > 1 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 1 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0}, {})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim num1 As Integer = TryCastInt(args(0).Resolve)
		If num1 = Nothing Then Return
		MaxCommands = num1

	End Sub

	Private ifcmdArray As String() = {"continue", "next", "end", "nextdir", "return"}

	Private Sub IfStatement(args() As Argument, negative As Boolean)

		If args.Count > 5 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 4 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0, 2}, {1, 3, 4})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim str1 As String = LCase(args(0).Resolve())
		Dim str2 As String = LCase(args(2).Resolve())
		Dim op As String = LCase(args(1).Value)
		Dim isTrue As Boolean
		Dim onTrue As String = LCase(args(3).Value)
		Dim onFalse As String = ifcmdArray(0)
		Dim action As String
		If args.Count = 5 Then onFalse = LCase(args(4).Value)

		If op = Chr(ASCII.Equals) Then
			isTrue = str1 = str2
		ElseIf op = Chr(ASCII.SmallerThan) Then
			isTrue = TryCastDbl(str1) < TryCastDbl(str2)
		ElseIf op = Chr(ASCII.GreaterThan) Then
			isTrue = TryCastDbl(str1) > TryCastDbl(str2)
		ElseIf op = "like" Then
			isTrue = str1 Like str2
		Else
			DisplayError("Invalid operator: " + Chr(34) + op + Chr(34) + "!")
			Return
		End If

		If negative Then isTrue = Not isTrue

		If isTrue Then
			action = onTrue
		Else
			action = onFalse
		End If

		If action = ifcmdArray(0) Then
			'do nothing
		ElseIf action = ifcmdArray(1) Then
			If Not InLoop Then DisplayError("Next cannot be called outside a file loop!")
			NextFile = True
		ElseIf action = ifcmdArray(2) Then
			Environment.Exit(1)
		ElseIf action = ifcmdArray(3) Then
			If Not InLoop Then DisplayError("NextDir cannot be called outside a file loop!")
			NextDir = True
		ElseIf action = ifcmdArray(4) Then
			ReturnToParent = True
		Else
			Dim gotoPos As Integer = getGotoPos(action)
			If Not gotoPos = Nothing Then
				changeLine = gotoPos - ActiveSub.CurPos
			Else
				DisplayError("Invalid argument for function: " + Chr(34) + action + Chr(34) + "!")
			End If
		End If

	End Sub

	Private Sub CallFunction(args() As Argument)

		Dim section As String
		Dim path As String = scriptPath
		Dim minargs As Integer = 3

		If args.Count < minargs Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		section = args(2).Value
		Dim calledSub As SubInfo = GetSub(section)
		If calledSub Is Nothing Then
			DisplayError("Cannot find sub " + section + "!")
			Return
		End If

		minargs += calledSub.Arguments.Count

		If args.Count > minargs Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < minargs Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim invalidArray(args.Count - 2) As Integer
		invalidArray(0) = 0
		For i As Integer = 2 To args.Count - 1
			invalidArray(i - 1) = i
		Next

		Dim InvalidStr As String = InvalidArgs(args, invalidArray, {2})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim script As ScriptInterpreter

		If LCase(args(1).Value) = "this" Then
			script = Me
		ElseIf LCase(args(1).Value) = "fl4" Then
			script = GetScript("fl4")
		Else
			path = args(1).Resolve()
			script = GetScript(path)
			If script Is Nothing Then
				DisplayError("Cannot find script " + Chr(34) + path + Chr(34) + "!")
			End If
		End If

		Dim passargs As Argument() = Nothing
		If args.Count >= 3 Then
			Array.Resize(passargs, args.Count - 3)
			Array.ConstrainedCopy(args, 3, passargs, 0, passargs.Count)
		End If

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		Dim result As String = script.Run(section, curDir, curFile, passargs)
		If result = Nothing Then
			DisplayError("Function returned NULL!")
			Return
		End If

		var.Value = result
		ReturnToParent = False

	End Sub

	Private Sub CallSection(args() As Argument)

		Dim section As String
		Dim path As String = scriptPath
		Dim result As String = Nothing
		Dim minargs As Integer = 2

		If args.Count < minargs Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		section = args(1).Value
		Dim calledSub As SubInfo = GetSub(section)
		If calledSub Is Nothing Then
			DisplayError("Cannot find sub " + Chr(34) + section + Chr(34) + "!")
			Return
		End If

		minargs += calledSub.Arguments.Count

		If args.Count > minargs Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < minargs Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim invalidArray(args.Count - 2) As Integer
		For i As Integer = 1 To args.Count - 1
			invalidArray(i - 1) = i
		Next

		Dim InvalidStr As String = InvalidArgs(args, invalidArray, {1})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim script As ScriptInterpreter

		If LCase(args(0).Value) = "this" Then
			script = Me
		ElseIf LCase(args(0).Value) = "fl4" Then
			script = GetScript("fl4")
		Else
			path = args(0).Resolve()
			script = GetScript(path)
			If script Is Nothing Then
				DisplayError("Cannot find script " + Chr(34) + path + Chr(34) + "!")
			End If
		End If

		Dim passargs As Argument() = Nothing
		If args.Count >= 3 Then
			Array.Resize(passargs, args.Count - 2)
			Array.ConstrainedCopy(args, 3, passargs, 0, passargs.Count)
		End If

		script.Run(section, curDir, curFile, passargs)

		ReturnToParent = False

	End Sub

	Private Sub NewLoop(args() As Argument)

		If args.Count > 4 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 3 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0, 1, 2}, {1, 2, 3})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		If InLoop Then
			DisplayError("NewLoop cannot be called inside another loop!")
			Return
		End If

		args(0).Resolve()
		args(1).Resolve()

		Dim path As String = scriptPath

		If Not File.Exists(path) Then
			DisplayError("The script " + Chr(34) + path + Chr(34) + " does not exist.")
			Return
		End If

		If Not Directory.Exists(args(0).Value) Then
			DisplayError("The directory " + Chr(34) + args(0).Value + Chr(34) + " does not exist.")
			Return
		End If

		Try
			Dim OverWrite As Boolean = (args.Count = 4 AndAlso LCase(args(3).Value) = "overwrite")
			SW = New StreamWriter(args(1).Value, Not OverWrite, enc)
		Catch ex As Exception
			DisplayError("Access Denied in " + Chr(34) + args(1).Value + Chr(34) + "!")
			Return
		End Try

		FileLoop(args(0).Value, args(2).Value)

		ReturnToParent = False

		If Not SW Is Nothing Then SW.Close()

	End Sub

	Private Sub Add(args() As Argument)

		If args.Count > 3 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 3 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0, 1, 2}, {0})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		args(1).Resolve()
		args(2).Resolve()

		Dim num1 As Double = TryCastDbl(args(1).Value)
		Dim num2 As Double = TryCastDbl(args(2).Value)
		Dim result As Double

		Try
			result = num1 + num2
		Catch
			DisplayError("Resulting number is out of range.")
			Return
		End Try

		var.Value = CStr(result)

	End Sub

	Private Sub Subtract(args() As Argument)

		If args.Count > 3 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 3 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0, 1, 2}, {0})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		Dim num1 As Double = TryCastDbl(args(1).Resolve)
		Dim num2 As Double = TryCastDbl(args(2).Resolve)
		var.Value = CStr(num1 - num2)

	End Sub

	Private Sub Multiply(args() As Argument)

		If args.Count > 3 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 3 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0, 1, 2}, {0})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		Dim num1 As Double = TryCastDbl(args(1).Resolve)
		Dim num2 As Double = TryCastDbl(args(2).Resolve)
		Dim result As Double

		Try
			result = num1 * num2
		Catch
			DisplayError("Resulting number is out of range.")
			Return
		End Try

		var.Value = CStr(result)

	End Sub

	Private Sub Divide(args() As Argument)

		If args.Count > 3 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 3 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0, 1, 2}, {0})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		Dim num1 As Double = TryCastDbl(args(1).Resolve)
		Dim num2 As Double = TryCastDbl(args(2).Resolve)
		Dim result As Double

		Try
			result = num1 / num2
		Catch
			DisplayError("Resulting number is out of range.")
			Return
		End Try

		var.Value = CStr(result)

	End Sub

	Private Sub Round(args() As Argument)

		If args.Count > 2 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 2 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0, 1}, {0})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		Dim num1 As Double = TryCastDbl(args(1).Resolve)
		var.Value = CStr(Math.Round(num1))

	End Sub

	Private Sub Floor(args() As Argument)

		If args.Count > 2 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 2 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0, 1}, {0})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		Dim num1 As Double = TryCastDbl(args(1).Resolve)
		var.Value = CStr(Math.Floor(num1))

	End Sub

	Private Sub Ceil(args() As Argument)

		If args.Count > 2 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 2 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0, 1}, {0})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		Dim num1 As Double = TryCastDbl(args(1).Resolve)
		var.Value = CStr(Math.Ceiling(num1))

	End Sub

	Private Sub GetC(args() As Argument)

		If args.Count > 3 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 3 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0, 1, 2}, {0})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		Dim charIndex As Integer = TryCastInt(args(2).Resolve)

		If args(1).Resolve.Length < charIndex Then
			DisplayError("Index out of range.")
			Return
		ElseIf charIndex < 1 Then
			DisplayError("Index must be larger than 0.")
			Return
		End If

		var.Value = GetChar(args(1).Value, charIndex)

	End Sub

	Private Sub FormatN(args() As Argument)

		If args.Count > 3 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 3 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0, 1, 2}, {0, 2})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		Dim num1 As Double = TryCastDbl(args(1).Resolve)
		var.Value = Format(num1, args(2).Resolve)

	End Sub

	Private Sub Pad(args() As Argument, left As Boolean)

		If args.Count > 3 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 3 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0, 1, 2}, {0})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		Dim num1 As Integer = TryCastInt(args(2).Resolve)

		args(1).Resolve()

		If left Then
			var.Value = args(1).Value.PadLeft(num1)
		Else
			var.Value = args(1).Value.PadRight(num1)
		End If

	End Sub

	Private Sub Len(args() As Argument)

		If args.Count > 2 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 2 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0, 1}, {0})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		Dim num1 As Integer = args(1).Resolve.Length
		var.Value = CStr(num1)

	End Sub

	Private Sub lcas(args() As Argument)

		If args.Count > 2 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 2 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0, 1}, {0})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		args(1).Resolve()

		Dim str As String = LCase(args(1).Value)
		var.Value = str

	End Sub

	Private Sub ucas(args() As Argument)

		If args.Count > 2 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 2 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0, 1}, {0})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		args(1).Resolve()

		Dim str As String = UCase(args(1).Value)
		var.Value = str

	End Sub

	Private Sub Rep(args() As Argument)

		If args.Count > 4 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 4 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0, 1, 2, 3}, {0})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		Dim str As String = Replace(args(1).Resolve, args(2).Resolve, args(3).Resolve)
		var.Value = str

	End Sub

	Private Sub SubString(args() As Argument)

		If args.Count > 4 Then
			DisplayError("Too many arguments for function!")
			Return
		ElseIf args.Count < 4 Then
			DisplayError("Missing arguments for function!")
			Return
		End If

		Dim InvalidStr As String = InvalidArgs(args, {0, 1, 2, 3}, {0})
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
			Return
		End If

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		Dim num1 As Integer = TryCastInt(args(2).Resolve)
		Dim num2 As Integer = TryCastInt(args(3).Resolve)

		Dim str As String = args(1).Resolve.Substring(num1, num2)
		var.Value = str

	End Sub

#End Region

End Class
