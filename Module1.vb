Imports System.IO
Imports System.Runtime.CompilerServices

Module Module1

	Public callStack() As SubInfo = {}
	Public scriptArray As ScriptInterpreter() = {}
	Public subArray As SubInfo() = {}

	Public InitPath As String() = {}
	Public MainScript As ScriptInterpreter
	Public ActiveScript As ScriptInterpreter
	Public vars() As Variable = {}
	Public InLoop As Boolean = False
	Public SW As StreamWriter
	Public AccessDenied As Boolean = False
	Public ErrorsHappened As Boolean = False
	Public CommandsProcessed As Integer = 0
	Public MaxCommands As Integer = 50

	Function GetScript(path As String) As ScriptInterpreter

		For Each s As ScriptInterpreter In scriptArray
			If LCase(s.FilePath) = LCase(path) Then Return s
		Next
		Return Nothing

	End Function

	Sub Main(args() As String)

		If args.Count < 1 Then
			Console.WriteLine("Missing arguments!")
			Return
		End If

		Dim filename As String = args(0)

		If Not filename Like "*.fl4" Then
			filename += ".fl4"
		End If

		Dim finfo As New FileInfo(filename)

		Console.Title = "FileList 4: " + (finfo.Name)

		Dim newargs As Argument() = {}
		If Environment.GetCommandLineArgs().Count > 2 Then
			For i As Integer = 2 To Environment.GetCommandLineArgs().Count - 1
				Array.Resize(newargs, newargs.Count + 1)
				newargs(newargs.Count - 1) = New Argument(Environment.GetCommandLineArgs(i), Argument.ArgType.Str)
			Next
		Else
			newargs = {}
		End If

		MainScript = New ScriptInterpreter(finfo.FullName, True)
		MainScript.Run("main", Nothing, Nothing, newargs)

		If ErrorsHappened AndAlso Environment.ExitCode = 0 Then
			Environment.ExitCode = 2
		End If

		If Environment.ExitCode = 2 Then
			Console.WriteLine("Finished with errors!")
		End If

	End Sub

	Sub FileLoop(path As String, func As String, Optional newLoop As Boolean = True)

		If path.Length = 0 Then Return

		Dim fileArray As FileInfo() = {}
		Dim dirArray As DirectoryInfo() = {}

		InLoop = True

		If newLoop Then
			Array.Resize(InitPath, InitPath.Count + 1)
			InitPath(InitPath.Count - 1) = path
		End If

		AccessDenied = False
		Dim dInfo As DirectoryInfo = Nothing
		dInfo = New DirectoryInfo(path)

		Try
			fileArray = dInfo.GetFiles
		Catch
			AccessDenied = True
			fileArray = {}
		End Try

		If fileArray.Count = 0 Then
			CommandsProcessed = 0
			MainScript.Run(func, dInfo, Nothing)
		Else
			For Each fInfo As FileInfo In fileArray
				CommandsProcessed = 0
				Dim data As ReturnData = MainScript.Run(func, dInfo, fInfo)
				If data.NextDirectory Then Exit For
			Next
		End If

		AccessDenied = False

		Try
			dirArray = dInfo.GetDirectories
		Catch
			AccessDenied = True
			dirArray = {}
		End Try

		For Each d As DirectoryInfo In dirArray
			FileLoop(d.FullName, func, False)
		Next

		If newLoop Then
			Array.Resize(InitPath, InitPath.Count - 1)
		End If

		InLoop = False

	End Sub

End Module

<HideModuleName> Module Extensions

	<Extension> Public Function MoveItem(Of T)(ByRef a As T(), source As Integer, destination As Integer) As Boolean

		Dim sourceItem As T = a(source)

		If source < destination Then
			Array.ConstrainedCopy(a, source + 1, a, source, destination - source)
			a(destination) = sourceItem
		ElseIf source > destination Then
			Array.ConstrainedCopy(a, destination, a, destination + 1, source - destination)
			a(destination) = sourceItem
		Else
			Return False
		End If

		Return True

	End Function

	<Extension> Public Sub Trim(Of T)(ByRef a As T())

		If a.Length = 0 Then Return
		Dim endIndex As Integer = a.Length - 1

		For i As Integer = a.Length - 1 To 0 Step -1
			If Not a(i) Is Nothing Then
				endIndex = i
				Exit For
			End If
		Next

		If endIndex = a.Length - 1 Then Return
		Array.Resize(a, endIndex + 1)

	End Sub

	<Extension> Public Sub Add(Of T)(ByRef a As T(), item As T)

		Array.Resize(a, a.Length + 1)
		a(a.Length - 1) = item

	End Sub

End Module
