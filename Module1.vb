﻿Imports System.IO

Module Module1

	Public callStack() As SubInfo = {}
	Public scriptArray As ScriptInterpreter() = {}
	Public subArray As SubInfo() = {}

	Public InitPath As String() = {}
	Public NextDir As Boolean = False
	Public NextFile As Boolean = False
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
				NextFile = False
				CommandsProcessed = 0
				MainScript.Run(func, dInfo, fInfo)
				If NextDir Then Exit For
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
			NextDir = False
			FileLoop(d.FullName, func, False)
		Next

		If newLoop Then
			Array.Resize(InitPath, InitPath.Count - 1)
		End If

		NextDir = False
		NextFile = False
		InLoop = False

	End Sub

End Module
