Option Explicit

Public gFSO  'As Scripting.FileSystemObject

Public gLastError 'message from most recent error, if any

'********************************
'** FrapsJoin: append Fraps AVI segments by generating an AviSynth script
'
'	INSTALL:
'	   1) copy the project files to any location, eg "C:\Program Files\FrapsJoin"
'	   2) in Windows Explorer, right-click the script, "Copy"
'	   3) in the Explorer navigation bar, type "shell:sendto", Enter
'	   4) right-click, "Paste Shortcut"
'	USAGE:
'	   see README or run without arguments
'
'@version 1.0  12-Jan-2012     - wrote it
'@version 1.1  13-Jan-2012     - added USAGE
'@version 1.2  13-Jan-2012     - added INI file reader & various user options
'
'copyright 2012 Lindsay Bigelow (aka raffriff aka XyKyWyKy)
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
' Fraps is a trademark of Beepa Pty Ltd
' AviSynth is free software under GNU General Public License.

    On Error Resume Next

    Set gFSO = CreateObject("Scripting.FileSystemObject")
	TestObject gFSO, "FileSys"

    Dim iniPath: iniPath = WScript.ScriptFullName
    iniPath = gFSO.GetParentFolderName(iniPath) & "\" & gFSO.GetBaseName(iniPath) & ".ini"

	Dim dict: Set dict = CreateObject("Scripting.Dictionary")
	TestObject dict, "Dictionary"

	IniReadSection iniPath, "", dict

	'** prefix for all generated Avisynth scripts
	Dim avsPrefix: avsPrefix = UnQuoteString(SafeDictItem(dict, "avsPrefix", "__")) 

	Const APP_EXE = "FrapsJoin.vbs"
	Const APP_TITLE = "Fraps Join"

	Dim USAGE: USAGE = APP_EXE & " - USAGE:"
	USAGE = USAGE & vbCrLf & "   1) Capture some video with Fraps..."
	USAGE = USAGE & vbCrLf & "   2) Go to the capture folder in Windows Explorer"
	USAGE = USAGE & vbCrLf & "   3) Sort files by Date (all files in a 'sequence' will have the same Modified Date)"
	USAGE = USAGE & vbCrLf & "   4) Select Fraps files from the 'sequence' you want to work with"
	USAGE = USAGE & vbCrLf & "        (Suggest you create a new folder for them)"
	USAGE = USAGE & vbCrLf & "   5) Right-click, 'Send To', '" & APP_EXE & "' (or whatever the shortcut is named)"
	USAGE = USAGE & vbCrLf & "        Script sorts the files by name and joins all files with an Avisynth script"
	USAGE = USAGE & vbCrLf & "        Avisynth script has same name as first file in sequence,"
	USAGE = USAGE & vbCrLf & "        except starts with an optional prefix and has an 'avs' extension"
	USAGE = USAGE & vbCrLf & "   NOTE: you can also drag the files to a shortcut on the Desktop etc"
	USAGE = USAGE & vbCrLf & "   6) Open the new script in any application that supports Avisynth:"
	USAGE = USAGE & vbCrLf & "        VirtualDub, WinFF, StaxRip, XMedia, etc"
	USAGE = USAGE & vbCrLf & "   NOTE: all files must be same resolution and frame rate or they won't join"

    Const ForWriting = 2

	'arguments = a set of AVI files;
	'copy to local array
	'
    Dim countArgs: countArgs = WScript.Arguments.Count
    If (countArgs = 0) Then
        MsgBox USAGE, vbExclamation, APP_TITLE 
        WScript.Quit
    End If

    Dim i, j, t

    Dim args(): ReDim args(countArgs - 1)
    For i = 0 To (countArgs - 1)
        args(i) = WScript.Arguments(i)
    Next

    'sort files by name (this sorts Fraps segments by time)  
    '
    For i = (UBound(args) - 1) To 0 Step (-1)
        For j= 0 to i
            If (StrComp(args(j), args(j + 1), vbTextCompare) > 0) Then 
                t = args(j + 1)
                args(j + 1) = args(j)
                args(j) = t
            End If
        Next
    Next

	Dim aviPath: aviPath = gFSO.GetAbsolutePathName(args(0))

    Dim avsPath: avsPath = gFSO.GetParentFolderName(aviPath) & "\" & avsPrefix & gFSO.GetBaseName(aviPath) & ".avs"

    'create output file
	'
    If (gFSO.FileExists(avsPath)) Then
		If (MsgBox(avsPath & vbCrLf & "exists: overwrite it?", vbYesNo Or vbQuestion, APP_TITLE) = vbNo) Then
			WScript.Quit
		End If
        gFSO.DeleteFile(avsPath)
    End If

    Dim ts: Set ts = gFSO.OpenTextFile(avsPath, ForWriting, True)
    If (Err) Then
        MsgBox "can't open '" & avsPath & "' for writing: " & Err.Description, vbExclamation, APP_TITLE
        ts.Close
		gFSO.DeleteFile(avsPath)
		WScript.Quit
    End If

	'first line (some applications need this comment line to process Avisynth correctly)
	ts.WriteLine "#avisynth" & vbCrLf

	Dim datecheck, extensions, firstfile, nextfile, postproc
	
	datecheck = UnQuoteString(SafeDictItem(dict, "datecheck", "1"))
	extensions = UnQuoteString(SafeDictItem(dict, "extensions", "|avi|"))
	firstfile = UnQuoteString(SafeDictItem(dict, "firstfile", "C = AviSource(`%1%`, pixel_type=`RGB32`)"))
	nextfile = UnQuoteString(SafeDictItem(dict, "nextfile", "C = C + AviSource(`%1%`, pixel_type=`RGB32`)"))
	postproc = UnQuoteString(SafeDictItem(dict, "postproc", ""))

	datecheck = SafeBoolean(datecheck, True)
	firstfile = Replace(firstfile, "`", """")
	nextfile = Replace(nextfile, "`", """")
	postproc = Replace(postproc, "`", """")

    'process command line args, adding a line for each file:
	'
	CheckExtension extensions, aviPath, avsPath, ts
	ts.WriteLine Replace(firstfile, "%1%", aviPath)

	Dim firstDate, nextDate
	
	firstDate = gFSO.GetFile(aviPath).DateLastModified
	CheckError 10, avsPath, ts

    For i = 1 To (countArgs - 1)

		aviPath = gFSO.GetAbsolutePathName(args(i))
		
		CheckExtension extensions, aviPath, avsPath, ts

		nextDate = gFSO.GetFile(aviPath).DateLastModified
		CheckError 20, avsPath, ts

		If (datecheck) Then
			If ((Abs(firstDate - nextDate) * 86400) > 1) Then
				MsgBox "Join halted: file date mismatch for '" & gFSO.GetBaseName(aviPath) & "'", vbInformation, APP_TITLE
				Exit For
			End If
		End If

		ts.WriteLine Replace(nextfile, "%1%", aviPath)
    Next

	'optional postprocessing
	ts.WriteLine Replace(postproc, "%1%", gFSO.GetBaseName(args(0)))

	'final line
	ts.WriteLine "return C"
	ts.Close

WScript.Quit
'END ############

'********************************
'** if error condition exists while writing to avs file, close and quit
'
Sub CheckError(ByVal code, ByVal avsPath, ts)

	If (Err) Then
		MsgBox "Error creating Avisynth script (" & code & "): " & Err.Description, vbExclamation, APP_TITLE
		ts.Close
		gFSO.DeleteFile(avsPath)
		WScript.Quit
	End If
End Sub

'********************************
'** if file doesn't have an allowed extension, close and quit
'
Sub CheckExtension(ByVal extensions, ByVal aviPath, ByVal avsPath, ts)

	If (InStr(1, extensions, "|" & gFSO.GetExtensionName(aviPath) & "|", vbTextCompare) = 0) Then
		MsgBox "Invalid file extension: " & gFSO.GetExtensionName(aviPath), vbExclamation, APP_TITLE
		ts.Close
		gFSO.DeleteFile(avsPath)
		WScript.Quit
	End If
End Sub

'********************************
'** read all name-value pairs from an .INI file section into a Dictionary object
'
'   - if "sSection" argument is empty:
'       - if config file does not have named sections, read entire file;
'       - if "name=value" lines exist before any section header,
'           read all such lines, up to first section header and exit
'       - else, read first section and exit
'   - else, if "sSection" argument supplied:
'       - if named section does not exist, read nothing
'       - else, read section named in "sSection" argument
'
'   - NOTE any valid environment variable is supported as section name (eg, %computername%)
'
'@param iniPath     - path to config file
'@param sSection    - name of section to be loaded (see notes above)
'@param dict        - a Dictionary object; values loaded from config file will be added here,
'                     overriding existing values in dictionary, if any
'
Function IniReadSection(ByVal iniPath, ByVal sSection, dict) 'As Boolean

    On Error Resume Next

    Const ForReading = 1

    Dim ts: Set ts = gFSO.OpenTextFile(iniPath, ForReading)
    If (Err) Then
        gLastError = "Error opening '" & iniPath & "' for reading"
        ts.Close
        Err.Clear
        IniReadSection = False
        Exit Function
    End If

    Dim sLine
    Dim s1
    Dim sName
    Dim sValue
    Dim sTest
    Dim pDelim
    Dim inSection
    Dim foundSection: foundSection = False

    If (Len(sSection)) Then
        sName = "[" & sSection & "]"
        inSection = False
    Else
        inSection = True
    End If

    Do While Not (ts.AtEndOfStream)

        sLine = Trim(ts.ReadLine)
        s1 = Left(sLine, 1)

        If (s1 = "") Then

            'ignore blank lines

        ElseIf (s1 = ";") Then

            'ignore comments

        ElseIf (s1 = "[") Then

            If (foundSection And inSection) Then
                Exit Do
            ElseIf (StrComp(sLine, sName, vbTextCompare) = 0) Then
                inSection = True
            End If
            foundSection = True

        ElseIf (inSection) Then

            pDelim = InStr(sLine, "=")
            If (pDelim > 1) Then
                sName = Trim(Left(sLine, pDelim - 1))
                sValue = Trim(Mid(sLine, pDelim + 1))
                dict.Remove sName
                Err.Clear
                dict(sName) = sValue
                If (Err) Then
                    'Debug.Assert (False)
                    Err.Clear
                End If
            End If
        End If
    Loop
    ts.Close
	Err.Clear
    IniReadSection = True
End Function

'********************************
'
Function SafeBoolean(ByVal v, ByVal defval) 'As Boolean
                
    On Error Resume Next
    
    Dim s: s = LCase(Trim(CStr(v)))    
    If (Err) Then
        Err.Clear
        SafeBoolean = defval
    End If
    
    If (s = "true") Then     
        SafeBoolean = True
	Else
		Dim t: t = CInt(v)
		If (Err) Then
			Err.Clear
			SafeBoolean = defval
			Exit Function
		End If
		SafeBoolean = (v <> 0)
    End If
End Function

'********************************
'
Function SafeDictItem(dict, ByVal k, ByVal defval)
    On Error Resume Next
    SafeDictItem = dict(k)
    If (IsEmpty(SafeDictItem)) Then
        SafeDictItem = defval
    End If
    Err.Clear
End Function

'********************************
'
Sub TestObject(obj, ByVal strTest)

    On Error Resume Next

	Dim t: t = TypeName(obj) 
    If (InStr(1, t, strTest, vbTextCompare) = 0) Then
		If (Err) Then
			t = t & ": " & Err.Description
		End If
        MsgBox "Error initializing " & strTest & ": " & t, vbExclamation, APP_TITLE
        WScript.Quit
    End If
End Sub

'******************************
'** tests to see if src is wrapped in quotation marks, and if so, remove them.
'   (NOTE: src like "foo"="bar" not handled or tested for)
'
Function UnQuoteString(ByVal src)

    src = Trim(src)

    If (Len(src) = 0) Then
        UnQuoteString = ""
        Exit Function
    End If
    If (Left(src, 1) = """") Then
        If (Right(src, 1) = """") Then
            src = Mid(src, 2, Len(src) - 2)
        End If
    End If
    UnQuoteString = src
End Function


