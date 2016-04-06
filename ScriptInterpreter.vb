Imports System.IO
Imports System.Text

Class ReturnData

	Protected nfile, ndir As Boolean
	Protected val As String = String.Empty

	Property Value As String
		Get
			Return val
		End Get
		Set(value As String)
			val = value
		End Set
	End Property

	Property NextFile As Boolean
		Get
			Return nfile
		End Get
		Set(value As Boolean)
			nfile = value
		End Set
	End Property

	Property NextDirectory As Boolean
		Get
			Return ndir
		End Get
		Set(value As Boolean)
			ndir = value
		End Set
	End Property

End Class

Class ScriptLine

	Protected s As String
	Protected o, l, m As Integer
	Protected ifblk As IfBlock

	Sub New(str As String, originalLine As Integer, line As Integer)
		s = str
		l = line
		o = originalLine
		m = -1
	End Sub

	Overrides Function ToString() As String
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

	Property IfBlock As IfBlock
		Get
			Return ifblk
		End Get
		Set(value As IfBlock)
			ifblk = value
		End Set
	End Property

	Property MoveTo As Integer
		Get
			Return m
		End Get
		Set(value As Integer)
			m = value
		End Set
	End Property

End Class

Class Argument

	Shared Operator +(ary() As Argument, obj As Argument) As Argument()

		Dim nextempty As Integer = -1

		For i As Integer = 0 To ary.Count - 1
			If ary(i) Is Nothing Then
				nextempty = i
				Exit For
			End If
		Next

		If nextempty = -1 Then
			If ary.Count = 0 Then
				nextempty = 0
				Array.Resize(ary, 1)
			Else
				nextempty = ary.Count - 1
				Array.Resize(ary, ary.Count * 2)
			End If
		End If

		ary(nextempty) = obj
		Return ary

	End Operator

	Friend Enum ArgType
		Var
		Str
	End Enum

	Protected val As String
	Protected t As ArgType
	Sub New(value As String, objectType As ArgType)
		val = value
		t = objectType
	End Sub

	Property Value As String
		Get
			Return val
		End Get
		Set(value As String)
			val = value
		End Set
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

	Overrides Function ToString() As String
		Return v
	End Function

End Class

Class GotoInfo

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

	Overrides Function ToString() As String
		Return id
	End Function

End Class

Class SubInfo

	Protected scriptobj As ScriptInterpreter
	Protected vars As Variable() = {}
	Protected argsArray As Argument() = {}
	Protected n, linestr As String
	Protected pos As Integer = 0
	Protected loc(1) As Integer
	Protected gotoArray As GotoInfo() = {}
	Protected ifBlockArray As IfBlock() = {}
	Protected RData As ReturnData
	Protected rtp As Boolean = False

	Sub New(script As ScriptInterpreter, section As String)
		scriptobj = script
		n = section
		RData = New ReturnData
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

	Property GotoObjects As GotoInfo()
		Get
			Return gotoArray
		End Get
		Set(value As GotoInfo())
			gotoArray = value
		End Set
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
		vars.Add(var)
		Return var

	End Function

	Sub Reset()
		vars = {}
		pos = 0
		rtp = False
		RData = New ReturnData
	End Sub

	Property CurPos As Integer
		Get
			Return pos
		End Get
		Set(value As Integer)
			pos = value
		End Set
	End Property

	Friend Property IfBlocks() As IfBlock()
		Get
			Return ifBlockArray
		End Get
		Set(value As IfBlock())
			ifBlockArray = value
		End Set
	End Property

	Friend ReadOnly Property ReturnData As ReturnData
		Get
			Return RData
		End Get
	End Property

	Friend Property ReturnToParent As Boolean
		Get
			Return rtp
		End Get
		Set(value As Boolean)
			rtp = value
		End Set
	End Property

	Overrides Function ToString() As String
		Return n
	End Function

End Class

Class IfBlock

	Protected startpos As Integer
	Protected endpos() As Integer = {}

	Sub New(startPoint As Integer)
		startpos = startPoint
	End Sub

	ReadOnly Property StartPoint As Integer
		Get
			Return startpos
		End Get
	End Property

	Property EndPoints As Integer()
		Get
			Return endpos
		End Get
		Set(value As Integer())
			endpos = value
		End Set
	End Property

	Function GetNextEndPoint(curpos As Integer) As Integer

		For i As Integer = 0 To endpos.Count - 1
			If curpos < endpos(i) Then Return endpos(i)
		Next
		Return -1

	End Function

End Class

Class ScriptInterpreter

	Overrides Function ToString() As String
		Return scriptPath
	End Function

	Private Function TranslatePos(line As Integer) As Integer

		For Each lineObj As ScriptLine In ScriptLines
			If lineObj.Line = line Then Return lineObj.OriginalLine
		Next
		Return line

	End Function

	Protected scriptPath As String
	Protected ScriptLines As ScriptLine() = {}
	Protected ScriptFile As String() = {}

	Sub New(path As String, allowFL4 As Boolean)
		Dim newsub As New SubInfo(Me, "Init")
		callStack.Add(newsub)
		ActiveSub = newsub
		ActiveScript = Me
		keyWords = MakeKeywords()
		fixedVars = MakeFixedVars()
		If LCase(path) = "fl4" Then
			ScriptFile = Split(My.Resources.fl4, vbNewLine)
			scriptPath = path
		Else
			Dim finfo As New FileInfo(path)
			If Not allowFL4 AndAlso LCase(finfo.Extension) = ".fl4" Then
				DisplayError("No FL4 files allowed beyond this point! Call FLS files instead.")
				Return
			End If
			scriptPath = finfo.FullName
			If Not finfo.Exists Then
				DisplayError("Could not find file: " + Chr(34) + scriptPath + Chr(34) + "!")
				Return
			End If
			ScriptFile = File.ReadAllLines(scriptPath, System.Text.Encoding.Default)
		End If
		ScriptLines = CompileInfo()
		If callStack.Length = 1 Then
			CheckCalls()
		End If
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

	Enum ErrorMode
		Bail
		Pause
		KeepGoing
	End Enum

	Friend Sub DisplayError(text As String)

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
		Dim output As String = vbNewLine + StrDup(Console.BufferWidth, "*") + text + vbNewLine + vbNewLine + "Call stack:" + vbNewLine + StrDup(11, "-") + vbNewLine + String.Join(vbNewLine, newCallStack) + vbNewLine + StrDup(Console.BufferWidth, "*")
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

	Private Function NewVar(name As String) As Variable

		Dim var As Variable = GetVarObj(name, Variable.varScope.Pub)
		If var IsNot Nothing Then
			DisplayError("Variable already declared: " + Chr(34) + var.Name + Chr(34) + "!")
			Return var
		End If

		var = New Variable(name, Variable.varScope.Pub)
		vars.Add(var)
		Return var

	End Function

	Private Sub CheckCmd(args() As Argument, min As Integer, max As Integer, kwords() As Integer, fvar() As Integer)

		If max > -1 AndAlso args.Count > max Then
			DisplayError("Too many arguments for function!")
		ElseIf args.Count < min Then
			DisplayError("Missing arguments for function!")
		End If
		Dim InvalidStr As String = InvalidArgs(args, kwords, fvar)
		If InvalidStr IsNot Nothing Then
			DisplayError("Invalid argument for function: " + Chr(34) + InvalidStr + Chr(34) + "!")
		End If

	End Sub

	Private Sub CheckCalls()

		For i As Integer = 0 To ScriptLines.Length - 1
			ActiveSub.CurPos = ScriptLines(i).OriginalLine
			Dim line As String = ScriptLines(i).ToString
			Dim argArray() As String = Split(line)
			Dim command As String = argArray(0)
			Dim argstring As String = String.Empty
			If argArray.Count > 1 Then
				argstring = String.Join(Chr(ASCII.Space), argArray, 1, argArray.Count - 1)
			End If
			Dim args As Argument() = ParseArgs(argstring, False)

			Select Case command.ToLower
				Case keyWords(kword._call)
					CheckCmd(args, 2, 2, {}, {})
					Dim _sub As String = args(1).Value
					Dim calledSub As SubInfo = GetSub(_sub)
					If calledSub Is Nothing Then
						DisplayError("Cannot find sub " + _sub + "!")
						Return
					End If
					Dim minargs As Integer = calledSub.Arguments.Count + 2
					Dim invalidArray(args.Count - 2) As Integer
					For a As Integer = 1 To args.Count - 1
						invalidArray(a - 1) = a
					Next
					CheckCmd(args, minargs, minargs, invalidArray, {1})

				Case keyWords(kword._function)
					CheckCmd(args, 3, -1, {}, {})
					Dim _sub As String = args(2).Value
					Dim calledSub As SubInfo = GetSub(_sub)
					If calledSub Is Nothing Then
						DisplayError("Cannot find sub " + _sub + "!")
						Return
					End If
					Dim minargs As Integer = calledSub.Arguments.Count + 3
					Dim invalidArray(args.Count - 2) As Integer
					invalidArray(0) = 0
					For a As Integer = 2 To args.Count - 1
						invalidArray(a - 1) = a
					Next
					CheckCmd(args, minargs, minargs, invalidArray, {2})

			End Select
		Next

	End Sub

	Private Function CompileInfo() As ScriptLine()

		Dim openSub As SubInfo = Nothing
		Dim loc0 As Integer
		Dim ifLevel As Integer = -1
		Dim openIfBlocks As IfBlock() = {}
		Dim newScript As ScriptLine() = {}
		Dim curLine As Integer = -1

		'gather info for quick jumping and vars
		For i As Integer = 0 To ScriptFile.Count - 1

			Dim line As String = ScriptFile(i)
			line = line.Trim
			If line.Length = 0 Then Continue For
			If GetChar(line, 1) = Chr(ASCII.Apostrophe) Then Continue For
			curLine += 1
			ActiveSub.CurPos = i + 1
			newScript.Add(New ScriptLine(line, i, curLine))

			If Not openSub Is Nothing AndAlso line(0) = Chr(ASCII.Colon) Then
				Dim gotoname As String = line.Substring(1, line.Length - 1)
				If IsNumeric(gotoname) Then
					DisplayError("GoTo-Targets cannot be numeric!")
					Return {}
				Else
					openSub.GotoObjects.Add(New GotoInfo(gotoname, curLine))
				End If
			Else
				Dim argArray() As String = Split(line)
				Dim command As String = argArray(0)
				Dim argstring As String = String.Empty
				If argArray.Count > 1 Then
					argstring = String.Join(Chr(ASCII.Space), argArray, 1, argArray.Count - 1)
				End If

				Select Case LCase(command)
					Case keyWords(kword._sub)
						If Not openSub Is Nothing Then
							DisplayError("Sub not closed: " + Chr(34) + openSub.Name + Chr(34) + "!")
							Return {}
						End If
						Dim args As Argument() = ParseArgs(argstring, False)
						If args.Count > 0 Then
							Dim s1 As SubInfo = GetSub(Me, args(0).Value)
							If s1 Is Nothing Then
								openSub = New SubInfo(Me, args(0).Value)
							ElseIf Not s1.Location(0) = Nothing Then
								DisplayError("Sub already exists: " + Chr(34) + args(0).Value + Chr(34) + "!")
								Return {}
							End If
							loc0 = curLine
							If args.Count > 1 Then
								Array.ConstrainedCopy(args, 1, args, 0, args.Count - 1)
								Array.Resize(args, args.Count - 1)
								openSub.Arguments = args
							End If
						ElseIf args.Count < 1 Then
							DisplayError("Missing arguments for Sub!")
							Return {}
						End If

					Case keyWords(kword._if), keyWords(kword.ifnot)
						If Not openSub Is Nothing Then
							Dim args As Argument() = ParseArgs(argstring)
							CheckCmd(args, 4, 5, {0, 2}, {1, 3, 4})
							If args.Count > 3 AndAlso LCase(args(3).Value) = keyWords(kword._then) Then
								ifLevel += 1
								Array.Resize(openIfBlocks, ifLevel + 1)
								openIfBlocks(ifLevel) = New IfBlock(curLine)
								newScript(newScript.Count - 1).IfBlock = openIfBlocks(ifLevel)
							End If
						Else
							DisplayError("Invalid use of If/IfNot!")
						End If

					Case keyWords(kword._else)
						If Not openSub Is Nothing Then
							If ifLevel = -1 Then DisplayError("Else must follow If/IfNot!")
							newScript(newScript.Count - 1).IfBlock = openIfBlocks(ifLevel)
							openIfBlocks(ifLevel).EndPoints.Add(curLine + 1)
						Else
							DisplayError("Invalid use of Else!")
						End If

					Case keyWords(kword._elseif), keyWords(kword.elseifnot)
						If Not openSub Is Nothing Then
							Dim args As Argument() = ParseArgs(argstring)
							CheckCmd(args, 4, 5, {0, 2}, {1, 3, 4})
							If args.Count > 3 AndAlso LCase(args(3).Value) = keyWords(kword._then) Then
								If ifLevel = -1 Then DisplayError("ElseIf/ElseIfNot must follow If/IfNot!")
								newScript(newScript.Count - 1).IfBlock = openIfBlocks(ifLevel)
								openIfBlocks(ifLevel).EndPoints.Add(curLine)
							End If
						Else
							DisplayError("Invalid use of ElseIf/ElseIfNot!")
						End If

					Case keyWords(kword._endif)
						If Not openSub Is Nothing Then
							If ifLevel = -1 Then DisplayError("EndIf must follow If/IfNot!")
							newScript(newScript.Count - 1).IfBlock = openIfBlocks(ifLevel)
							openIfBlocks(ifLevel).EndPoints.Add(curLine + 1)
							ifLevel -= 1
							If ifLevel = -1 Then
								For Each ifObj As IfBlock In openIfBlocks
									openSub.IfBlocks.Add(ifObj)
								Next
								openIfBlocks = {}
							End If
						Else
							DisplayError("Invalid use of EndIf!")
						End If

					Case keyWords(kword._end)
						Dim args As Argument() = ParseArgs(argstring, False)
						If args.Count = 1 Then
							If LCase(args(0).Value) = keyWords(kword._sub) Then
								If openSub Is Nothing Then
									DisplayError("End Sub must follow Sub!")
									Return {}
								End If
								openSub.Location = {loc0, curLine}
								subArray.Add(openSub)
								openSub = Nothing
							End If
						ElseIf openSub Is Nothing Then
							DisplayError("Invalid use of End!")
							Return {}
						End If

					Case keyWords(kword.var)
						Dim args As Argument() = ParseArgs(argstring)
						CheckCmd(args, 1, 2, {0}, {0})
						If openSub Is Nothing Then
							If args.Count = 2 Then
								NewVar(args(0).Value).Value = args(1).Resolve
							ElseIf args.Count = 1 Then
								NewVar(args(0).Value).Value = String.Empty
							End If
						End If

					Case keyWords(kword._set)
						Dim args As Argument() = ParseArgs(argstring)
						CheckCmd(args, 2, 2, {0}, {0})

					Case keyWords(kword._call)
						Dim args As Argument() = ParseArgs(argstring)
						If Not LCase(args(0).Value) = keyWords(kword.this) Then
							scriptArray.Add(New ScriptInterpreter(args(0).Value, False))
						End If

					Case keyWords(kword._function)
						Dim args As Argument() = ParseArgs(argstring)
						If Not LCase(args(1).Value) = keyWords(kword.this) Then
							scriptArray.Add(New ScriptInterpreter(args(1).Value, False))
						End If

					Case keyWords(kword.onerror)
						Dim args As Argument() = ParseArgs(argstring)
						If args.Count < 1 Then
							DisplayError("Missing arguments for OnError!")
							Return {}
						ElseIf args.Count > 1 Then
							DisplayError("Too many arguments for OnError!")
							Return {}
						End If
						If LCase(args(0).Value) = keyWords(kword.pause) Then
							OnError = ErrorMode.Pause
						ElseIf LCase(args(0).Value) = keyWords(kword._end) Then
							OnError = ErrorMode.Bail
						ElseIf LCase(args(0).Value) = keyWords(kword._continue) Then
							OnError = ErrorMode.KeepGoing
						Else
							DisplayError("Invalid argument: " + Chr(34) + args(0).Value + Chr(34) + "!")
							Return {}
						End If

					Case keyWords(kword.sleep)
						Dim args As Argument() = ParseArgs(argstring, False)
						CheckCmd(args, 1, 1, {0}, {0})

					Case keyWords(kword._return)
						Dim args As Argument() = ParseArgs(argstring, False)
						CheckCmd(args, 0, 1, {0}, {0})

					Case keyWords(kword.pause)
						Dim args As Argument() = ParseArgs(argstring, False)
						CheckCmd(args, 0, 1, {0}, {0})

					Case keyWords(kword.encoding)
						Dim args As Argument() = ParseArgs(argstring, False)
						CheckCmd(args, 1, 1, {0, 1}, {0, 1})

					Case keyWords(kword.title)
						Dim args As Argument() = ParseArgs(argstring, False)
						CheckCmd(args, 1, 1, {0}, {})

					Case keyWords(kword.print)
						Dim args As Argument() = ParseArgs(argstring, False)
						CheckCmd(args, 1, 1, {0}, {})

					Case keyWords(kword._goto)
						Dim args As Argument() = ParseArgs(argstring, False)
						CheckCmd(args, 1, 1, {0}, {0})

					Case keyWords(kword.maxcommands)
						Dim args As Argument() = ParseArgs(argstring, False)
						CheckCmd(args, 1, 1, {0}, {})

					Case keyWords(kword.add), keyWords(kword.subtract), keyWords(kword.multiply), keyWords(kword.divide)
						Dim args As Argument() = ParseArgs(argstring, False)
						CheckCmd(args, 3, 3, {0, 1, 2}, {0})

					Case keyWords(kword.round), keyWords(kword.floor), keyWords(kword.ceil)
						Dim args As Argument() = ParseArgs(argstring, False)
						CheckCmd(args, 2, 2, {0, 1}, {0})

					Case keyWords(kword._char)
						Dim args As Argument() = ParseArgs(argstring, False)
						CheckCmd(args, 3, 3, {0, 1, 2}, {0})

					Case keyWords(kword.format), keyWords(kword.padleft), keyWords(kword.padright)
						Dim args As Argument() = ParseArgs(argstring, False)
						CheckCmd(args, 3, 3, {0, 1, 2}, {0, 2})

					Case keyWords(kword.len), keyWords(kword.lcase), keyWords(kword.ucase)
						Dim args As Argument() = ParseArgs(argstring, False)
						CheckCmd(args, 2, 2, {0, 1}, {0})

					Case keyWords(kword.replace), keyWords(kword.substring)
						Dim args As Argument() = ParseArgs(argstring, False)
						CheckCmd(args, 4, 4, {0, 1, 2, 3}, {0})

				End Select
			End If
		Next

		For Each _sub As SubInfo In subArray
			For Each block As IfBlock In _sub.IfBlocks
				For Each endpoint As Integer In block.EndPoints
					newScript(endpoint - 1).MoveTo = block.EndPoints(block.EndPoints.Count - 1)
				Next
			Next
		Next

		If Not openSub Is Nothing Then
			DisplayError("Sub not closed: " + Chr(34) + openSub.Name + Chr(34) + "!")
		End If

		Return newScript

	End Function

	Private Overloads Function GetSub(name As String) As SubInfo

		name = LCase(name)
		For Each s As SubInfo In subArray
			If LCase(s.Name) = name Then
				Return s
			End If
		Next
		Return Nothing

	End Function

	Private Overloads Function GetSub(script As ScriptInterpreter, name As String) As SubInfo

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

	Friend Function GetVarObj(name As String, Optional scope As Variable.varScope = Variable.varScope.Unknown) As Variable

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

	Friend Enum ASCII
		Space = 32
		Quote = 34
		Apostrophe = 39
		Plus = 43
		Colon = 58
		SmallerThan = 60
		Equals = 61
		GreaterThan = 62
		Backslash = 92
	End Enum

	Private Function ParseArgs(source As String, Optional resolveVars As Boolean = True) As Argument()

		If source.Length = 0 Then Return {}

		Dim temp, tempArg As New StringBuilder
		Dim newArg(4) As Argument
		Dim argType As Argument.ArgType

		For i As Integer = 0 To source.Length - 1
			Dim stringOpen, mustResolve As Boolean
			Dim curChar As Char

			If i > source.Length - 1 Then
				DisplayError("Unexpected end of argument!")
				Return Nothing
			End If
			curChar = source(i)
			If curChar = Chr(ASCII.Quote) Then
				stringOpen = Not stringOpen
				argType = Argument.ArgType.Str
			ElseIf stringOpen Then
				tempArg.Append(curChar)
				argType = Argument.ArgType.Str

			ElseIf curChar = Chr(ASCII.Space) OrElse i = source.Length - 1 Then
				Dim n As Integer = GetNextCharIndex(source, i)

				If n = -1 OrElse source(n) = Chr(ASCII.Plus) Then
					If n = -1 Then
						temp.Append(curChar)
					Else
						n = GetNextCharIndex(source, n)
						If resolveVars Then mustResolve = True
						argType = Argument.ArgType.Str
					End If
					If temp.Length > 0 Then
						If mustResolve AndAlso Not Resolve(temp) Then
							Return Nothing
						End If
						tempArg.Append(temp)
						temp.Clear()
					End If

				ElseIf Not source(n) = Chr(ASCII.Plus) Then
					If argType = Argument.ArgType.Str OrElse temp.Length > 0 Then
						If mustResolve AndAlso temp.Length > 0 AndAlso Not Resolve(temp) Then
							Return Nothing
						End If
						tempArg.Append(temp)
						temp.Clear()
					End If
					newArg += New Argument(tempArg.ToString, argType)
					argType = Argument.ArgType.Var
					tempArg.Clear()
					mustResolve = False
				End If
				If Not n = -1 AndAlso Not n - 1 = i Then
					i = n - 1
				End If

			Else
				temp.Append(curChar)
			End If
		Next

		If argType = Argument.ArgType.Str OrElse tempArg.Length > 0 Then
			newArg += New Argument(tempArg.ToString, argType)
		End If

		newArg.Trim()

		Return newArg

	End Function

	Private Function GetNextCharIndex(input As String, start As Integer) As Integer

		For i As Integer = start + 1 To input.Length - 1
			If Not input(i) = Chr(ASCII.Space) Then Return i
		Next

		Return -1

	End Function

	Private Function Resolve(ByRef input As StringBuilder) As Boolean

		If isFixedVar(input.ToString) Then
			input.Replace(input.ToString, GetFixedVar(input.ToString))
		Else
			Dim tempvar As Variable = GetVarObj(input.ToString)
			If Not tempvar Is Nothing Then
				input.Replace(input.ToString, tempvar.Value)
			Else
				DisplayError("Variable not declared: " + Chr(34) + input.ToString + Chr(34) + "!")
				Return False
			End If
		End If

		Return True

	End Function

	Private keyWords() As String '= {"set", "var", "swstop", "swstart", "if", "elseif", "ifnot", "elseifnot", "return", "nextdir", "next", "goto", "print", "write", "replace", "lcase", "ucase", "add", "subtract", "multiply", "divide", "round", "floor", "ceil", "char", "format", "padleft", "padright", "len", "substring", "call", "function", "beep", "title", "readkey", "maxcommands", "fileloop", "end", "encoding", "sleep", "like", "continue", "onerror", "pause", "overwrite", "this", "fl4", "else", "then"}

	Private Enum kword
		_sub
		_set
		var
		swstop
		swstart
		_if
		_elseif
		ifnot
		elseifnot
		_endif
		_return
		nextdir
		_next
		_goto
		print
		write
		replace
		lcase
		ucase
		add
		subtract
		multiply
		divide
		round
		floor
		ceil
		_char
		format
		padleft
		padright
		len
		substring
		_call
		_function
		beep
		title
		readkey
		maxcommands
		fileloop
		_end
		encoding
		sleep
		_like
		_continue
		onerror
		pause
		overwrite
		this
		fl4
		_else
		_then
		max
	End Enum

	Private Function MakeKeywords() As String()

		Dim temp(kword.max - 1) As String

		For i As Integer = 0 To kword.max - 1
			Dim newkeyword As String = CType(i, kword).ToString
			If newkeyword(0) = "_" Then newkeyword = newkeyword.Substring(1)
			temp(i) = newkeyword
		Next

		Return temp

	End Function

	Private fixedVars() As String '= {"q", "newline", "dircount", "filecount", "accessdenied", "filepath", "filename", "filenamenx", "fileextension", "size", "folder", "path", "initpath", "exepath", "year", "month", "day", "hour", "minute", "second", "msecond", "utcoffset", "maxcommands", "swtime", "proccommands", "bufferwidth", "bufferheight"}

	Private Enum fvars
		q
		newline
		dircount
		filecount
		accessdenied
		filepath
		filename
		filenamenx
		fileextension
		size
		folder
		path
		initpath
		exepath
		year
		month
		day
		hour
		minute
		second
		msecond
		utcoffset
		maxcommands
		swtime
		proccommands
		bufferwidth
		bufferheight
		max
	End Enum

	Private Function MakeFixedVars() As String()

		Dim temp(fvars.max - 1) As String

		For i As Integer = 0 To fvars.max - 1
			Dim newkeyword As String = CType(i, fvars).ToString
			'If newkeyword(0) = "_" Then newkeyword = newkeyword.Substring(1)
			temp(i) = newkeyword
		Next

		Return temp

	End Function

	Function Run(subName As String, dInfo As DirectoryInfo, fInfo As FileInfo, Optional arguments As Argument() = Nothing) As ReturnData

		ActiveScript = Me
		ActiveSub = GetSub(Me, subName)

		If ActiveSub Is Nothing Then
			DisplayError("Sub not found: " + Chr(34) + subName + Chr(34) + "!")
			Return Nothing
		End If

		callStack.Add(ActiveSub)

		curFile = fInfo
		curDir = dInfo

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

		Dim rdata As New ReturnData

		If ActiveSub.Location(1) - ActiveSub.Location(0) > 1 Then
			For i As Integer = ActiveSub.Location(0) + 1 To ActiveSub.Location(1) - 1
				changeLine = 0
				If ScriptLines(i).ToString(0) = Chr(ASCII.Colon) Then Continue For
				If CommandsProcessed >= MaxCommands Then
					DisplayError("Maximum number of commands reached!")
					Exit For
				End If
				ActiveSub.CurPos = i + 1
				CommandsProcessed += 1
				Dim argArray() As String = Split(ScriptLines(i).ToString)
				Dim command As String = argArray(0)
				Dim args As String = String.Empty
				If argArray.Count > 1 Then
					args = String.Join(Chr(ASCII.Space), argArray, 1, argArray.Count - 1)
				End If

				Select Case LCase(command)
					Case keyWords(kword._set)
						SetVar(ParseArgs(args), False)
					Case keyWords(kword.var)
						SetVar(ParseArgs(args), True)
					Case keyWords(kword.swstop)
						sWatch.Stop()
					Case keyWords(kword.swstart)
						sWatch.Restart()
					Case keyWords(kword._if), keyWords(kword._elseif)
						IfStatement(ParseArgs(args), False)
					Case keyWords(kword.ifnot), keyWords(kword.elseifnot)
						IfStatement(ParseArgs(args), True)
					Case keyWords(kword._return)
						ActiveSub.ReturnData.Value = _Return(ParseArgs(args))
						ActiveSub.ReturnToParent = True
					Case keyWords(kword.nextdir)
						If Not InLoop Then DisplayError("NextDir cannot be called outside of a file loop!")
						ActiveSub.ReturnData.NextDirectory = True
					Case keyWords(kword._next)
						If Not InLoop Then DisplayError("Next cannot be called outside of a file loop!")
						ActiveSub.ReturnData.NextFile = True
					Case keyWords(kword._goto)
						Jump(ParseArgs(args))
					Case keyWords(kword.print)
						_print(ParseArgs(args))
					Case keyWords(kword.write)
						_write(ParseArgs(args))
					Case keyWords(kword.replace)
						Rep(ParseArgs(args))
					Case keyWords(kword.lcase)
						lcas(ParseArgs(args))
					Case keyWords(kword.ucase)
						ucas(ParseArgs(args))
					Case keyWords(kword.add)
						Add(ParseArgs(args))
					Case keyWords(kword.subtract)
						Subtract(ParseArgs(args))
					Case keyWords(kword.multiply)
						Multiply(ParseArgs(args))
					Case keyWords(kword.divide)
						Divide(ParseArgs(args))
					Case keyWords(kword.round)
						Round(ParseArgs(args))
					Case keyWords(kword.floor)
						Floor(ParseArgs(args))
					Case keyWords(kword.ceil)
						Ceil(ParseArgs(args))
					Case keyWords(kword._char)
						GetC(ParseArgs(args))
					Case keyWords(kword.format)
						FormatN(ParseArgs(args))
					Case keyWords(kword.padleft)
						Pad(ParseArgs(args), True)
					Case keyWords(kword.padright)
						Pad(ParseArgs(args), False)
					Case keyWords(kword.len)
						Len(ParseArgs(args))
					Case keyWords(kword.substring)
						SubString(ParseArgs(args))
					Case keyWords(kword._call)
						CallSub(ParseArgs(args))
					Case keyWords(kword._function)
						CallFunction(ParseArgs(args))
					Case keyWords(kword.beep)
						Media.SystemSounds.Beep.Play()
					Case keyWords(kword.title)
						title(ParseArgs(args))
					Case keyWords(kword.readkey)
						ReadKey(ParseArgs(args))
					Case keyWords(kword.maxcommands)
						SetMaxCommands(ParseArgs(args))
					Case keyWords(kword.fileloop)
						NewLoop(ParseArgs(args))
					Case keyWords(kword._end)
						Environment.Exit(1)
					Case keyWords(kword.encoding)
						SetEncoding(ParseArgs(args))
					Case keyWords(kword.sleep)
						Wait(ParseArgs(args))
					Case Else
						DisplayError("Invalid command: " + Chr(34) + command + Chr(34) + "!")
				End Select

				If Not changeLine = 0 Then
					i += changeLine
				ElseIf ScriptLines(i).MoveTo > -1 Then
					i = ScriptLines(i).MoveTo
				End If
				If ActiveSub.ReturnData.NextDirectory OrElse ActiveSub.ReturnData.NextFile Then Exit For
				If ActiveSub.ReturnToParent Then Exit For
			Next
			rdata = ActiveSub.ReturnData
			ActiveSub.Reset()

		End If

		Array.Resize(callStack, callStack.Count - 1)
		If callStack.Length > 0 Then
			ActiveSub = callStack(callStack.Count - 1)
			ActiveScript = ActiveSub.Script
		End If

		Return rdata

	End Function

	Friend Function GetFixedVar(varName As String) As String

		Select Case LCase(varName)
			Case fixedVars(fvars.q)
				Return Chr(ASCII.Quote)
			Case fixedVars(fvars.newline)
				Return vbNewLine
			Case fixedVars(fvars.dircount)
				If callStack(0).Name = "Init" Then
					Return String.Empty
				ElseIf curDir Is Nothing Then
					DisplayError("No directory-object!")
				ElseIf AccessDenied Then
					DisplayError("Access denied in " + Chr(34) + curDir.FullName + Chr(34) + "!")
				Else
					Return CStr(curDir.GetDirectories.Count)
				End If
			Case fixedVars(fvars.filecount)
				If callStack(0).Name = "Init" Then
					Return String.Empty
				ElseIf curDir Is Nothing Then
					DisplayError("No directory-object!")
				ElseIf AccessDenied Then
					DisplayError("Access denied in " + Chr(34) + curDir.FullName + Chr(34) + "!")
				Else
					Return CStr(curDir.GetFiles.Count)
				End If
			Case fixedVars(fvars.accessdenied)
				Return AccessDenied.ToString
			Case fixedVars(fvars.filepath)
				If callStack(0).Name = "Init" Then
					Return String.Empty
				ElseIf curDir Is Nothing Then
					DisplayError("No directory-object!")
				ElseIf curFile Is Nothing Then
					DisplayError("No file-object in " + Chr(34) + curDir.FullName + Chr(34) + "!")
				Else
					Return curFile.FullName
				End If
			Case fixedVars(fvars.filename)
				If callStack(0).Name = "Init" Then
					Return String.Empty
				ElseIf curDir Is Nothing Then
					DisplayError("No directory-object!")
				ElseIf curFile Is Nothing Then
					DisplayError("No file-object in " + Chr(34) + curDir.FullName + Chr(34) + "!")
				Else
					Return curFile.Name
				End If
			Case fixedVars(fvars.filenamenx)
				If callStack(0).Name = "Init" Then
					Return String.Empty
				ElseIf curDir Is Nothing Then
					DisplayError("No directory-object!")
				ElseIf curFile Is Nothing Then
					DisplayError("No file-object in " + Chr(34) + curDir.FullName + Chr(34) + "!")
				Else
					Return Path.GetFileNameWithoutExtension(curFile.FullName)
				End If
			Case fixedVars(fvars.fileextension)
				If callStack(0).Name = "Init" Then
					Return String.Empty
				ElseIf curDir Is Nothing Then
					DisplayError("No directory-object!")
				ElseIf curFile Is Nothing Then
					DisplayError("No file-object in " + Chr(34) + curDir.FullName + Chr(34) + "!")
				Else
					Return Path.GetExtension(curFile.FullName)
				End If
			Case fixedVars(fvars.size)
				If callStack(0).Name = "Init" Then
					Return String.Empty
				ElseIf curDir Is Nothing Then
					DisplayError("No directory-object!")
				ElseIf curFile Is Nothing Then
					DisplayError("No file-object in " + Chr(34) + curDir.FullName + Chr(34) + "!")
				Else
					Return CStr(curFile.Length)
				End If
			Case fixedVars(fvars.folder)
				If callStack(0).Name = "Init" Then
					Return String.Empty
				ElseIf curDir Is Nothing Then
					DisplayError("No directory-object!")
				Else
					Return curDir.Name
				End If
			Case fixedVars(fvars.path)
				If callStack(0).Name = "Init" Then
					Return String.Empty
				ElseIf curDir Is Nothing Then
					DisplayError("No directory-object!")
				Else
					Return curDir.FullName
				End If
			Case fixedVars(fvars.initpath)
				If callStack(0).Name = "Init" Then
					Return String.Empty
				Else
					Return InitPath(InitPath.Count - 1)
				End If
			Case fixedVars(fvars.exepath)
				Return My.Application.Info.DirectoryPath
			Case fixedVars(fvars.year)
				Return Format(Date.Now, "yyyy")
			Case fixedVars(fvars.month)
				Return Format(Date.Now, "MM")
			Case fixedVars(fvars.day)
				Return Format(Date.Now, "dd")
			Case fixedVars(fvars.hour)
				Return Format(Date.Now, "HH")
			Case fixedVars(fvars.minute)
				Return Format(Date.Now, "mm")
			Case fixedVars(fvars.second)
				Return Format(Date.Now, "ss")
			Case fixedVars(fvars.msecond)
				Return Format(Date.Now, "fffffff")
			Case fixedVars(fvars.utcoffset)
				Return Format(Date.Now, "zz")
			Case fixedVars(fvars.maxcommands)
				Return MaxCommands.ToString
			Case fixedVars(fvars.swtime)
				Return sWatch.Elapsed.TotalMilliseconds.ToString
			Case fixedVars(fvars.proccommands)
				Return CommandsProcessed.ToString
			Case fixedVars(fvars.bufferwidth)
				Return Console.BufferWidth.ToString
			Case fixedVars(fvars.bufferheight)
				Return Console.BufferHeight.ToString
		End Select

		Return Nothing

	End Function

	Private Function isKeyword(varName As String) As Boolean

		Dim result As Boolean
		varName = LCase(varName)
		result = Array.IndexOf(keyWords, varName) > -1
		If result = False Then
			result = Array.IndexOf(operators, varName) > -1
		End If
		Return result

	End Function

	Friend Function isFixedVar(varName As String) As Boolean

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

		args(0).Resolve()

		Dim num As Integer = TryCastInt(args(0).Value)
		Threading.Thread.Sleep(num)

	End Sub

	Private Function _Return(args() As Argument) As String

		If args.Count > 0 Then
			Return args(0).Resolve
		Else
			Return Nothing
		End If

	End Function

	Private Sub ReadKey(args() As Argument)

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
			DisplayError("Invalid argument for function: " + Chr(34) + args(0).Value + Chr(34) + "!")
		End If

	End Sub

	Private Sub SetVar(args() As Argument, setNew As Boolean)

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

		args(0).Resolve()

		Console.Title = args(0).Value

	End Sub

	Private Sub _print(args() As Argument)

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

		Dim i As Integer = getGotoPos(args(0).Value)
		If i = Nothing Then
			DisplayError("The Goto-mark does not exist: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		Else
			changeLine = i - ActiveSub.CurPos
		End If

	End Sub

	Private Sub SetMaxCommands(args() As Argument)

		Dim num1 As Integer = TryCastInt(args(0).Resolve)
		If num1 = Nothing Then Return
		MaxCommands = num1

	End Sub

	Private operators As String() = {"<", ">", "=", "+", "like"}
	Private ifcmdArray As String() = {"continue", "next", "nextdir", "return", "then", "end"}

	Private Sub IfStatement(args() As Argument, negative As Boolean)

		Dim str1 As String = LCase(args(0).Resolve())
		Dim str2 As String = LCase(args(2).Resolve())
		Dim op As String = LCase(args(1).Value)
		Dim isTrue As Boolean
		Dim onTrue As String = LCase(args(3).Value)
		Dim onFalse As String = ifcmdArray(0)
		Dim action As String
		If args.Count = 5 Then onFalse = LCase(args(4).Value)

		If onTrue = keyWords(kword._then) Then onFalse = onTrue

		If onTrue = ifcmdArray(4) AndAlso args.Count > 4 Then
			DisplayError("Too many arguments for function!")
			Return
		End If

		If op = Chr(ASCII.Equals) Then
			isTrue = str1 = str2
		ElseIf op = Chr(ASCII.SmallerThan) Then
			isTrue = TryCastDbl(str1) < TryCastDbl(str2)
		ElseIf op = Chr(ASCII.GreaterThan) Then
			isTrue = TryCastDbl(str1) > TryCastDbl(str2)
		ElseIf op = operators(4) Then
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
			ActiveSub.ReturnData.NextFile = True
		ElseIf action = ifcmdArray(2) Then
			If Not InLoop Then DisplayError("NextDir cannot be called outside a file loop!")
			ActiveSub.ReturnData.NextDirectory = True
		ElseIf action = ifcmdArray(3) Then
			ActiveSub.ReturnToParent = True
		ElseIf action = ifcmdArray(4) Then
			If Not isTrue Then
				changeLine = ActiveScript.ScriptLines(ActiveSub.CurPos - 1).IfBlock.GetNextEndPoint(ActiveSub.CurPos) - ActiveSub.CurPos
			End If
		ElseIf action = ifcmdArray(5) Then
			Environment.Exit(1)
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

		Dim _sub As String = args(2).Value
		Dim calledSub As SubInfo = GetSub(_sub)
		Dim path As String = scriptPath
		Dim script As ScriptInterpreter

		If LCase(args(1).Value) = keyWords(kword.this) Then
			script = Me
		ElseIf LCase(args(1).Value) = keyWords(kword.fl4) Then
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

		Dim result As ReturnData = script.Run(_sub, curDir, curFile, passargs)
		If result.Value = Nothing Then
			DisplayError("Function returned NULL!")
			Return
		End If

		ActiveSub.ReturnData.NextDirectory = result.NextDirectory
		ActiveSub.ReturnData.NextFile = result.NextFile
		var.Value = result.Value

	End Sub

	Private Sub CallSub(args() As Argument)

		Dim _sub As String = args(1).Value
		Dim calledSub As SubInfo = GetSub(_sub)
		Dim path As String = scriptPath
		Dim script As ScriptInterpreter

		If LCase(args(0).Value) = keyWords(kword.this) Then
			script = Me
		ElseIf LCase(args(0).Value) = keyWords(kword.fl4) Then
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

		Dim result As ReturnData = script.Run(_sub, curDir, curFile, passargs)
		ActiveSub.ReturnData.NextDirectory = result.NextDirectory
		ActiveSub.ReturnData.NextFile = result.NextFile

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

		If Not SW Is Nothing Then SW.Close()

	End Sub

	Private Sub Add(args() As Argument)

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

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		Dim num1 As Double = TryCastDbl(args(1).Resolve)
		var.Value = CStr(Math.Round(num1))

	End Sub

	Private Sub Floor(args() As Argument)

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		Dim num1 As Double = TryCastDbl(args(1).Resolve)
		var.Value = CStr(Math.Floor(num1))

	End Sub

	Private Sub Ceil(args() As Argument)

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		Dim num1 As Double = TryCastDbl(args(1).Resolve)
		var.Value = CStr(Math.Ceiling(num1))

	End Sub

	Private Sub GetC(args() As Argument)

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

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		Dim num1 As Double = TryCastDbl(args(1).Resolve)
		var.Value = Format(num1, args(2).Resolve)

	End Sub

	Private Sub Pad(args() As Argument, left As Boolean)

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

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		Dim num1 As Integer = args(1).Resolve.Length
		var.Value = CStr(num1)

	End Sub

	Private Sub lcas(args() As Argument)

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

		Dim var As Variable = GetVarObj(args(0).Value)

		If var Is Nothing Then
			DisplayError("Variable not declared: " + Chr(34) + args(0).Value + Chr(34) + "!")
			Return
		End If

		Dim str As String = Replace(args(1).Resolve, args(2).Resolve, args(3).Resolve)
		var.Value = str

	End Sub

	Private Sub SubString(args() As Argument)

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
