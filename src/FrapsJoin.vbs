'Attribute VB_Name = "modFrapsJoin_22"
'********************************
'** FrapsJoin.vbs - rename fraps-generated avi files to
'   virtualdub-understandable segment names
'
' latest version available at http://github.com/XyKyWyKy/FrapsJoin
' see attached README and .INI for installation & usage
'
' koala85 1-NOV-2011  http://frapsforum.com/threads/raffriffs-awesome-virtualdub-tutorial.739/#post-3175
' raffriff mod 16-Jan-2012 (get folder from command line; make "undo" bat file)
' raffriff update 18-Jan-2012 (user-specified group name; support 100+ files; limit unknown)
' raffriff update 18-Jan-2012 (user options from INI file; tweak error handling)
' raffriff update 20-Jan-2012 (add optional VirtualDub script generation)
' raffriff update 21-Jan-2012 (add optional VirtualDub output path)
' raffriff update 22-Jan-2012 (add optional output filename prefix)
' raffriff update 27-Jan-2012 (optional output path for Avidemux also; minor bugfixes)
' raffriff update 06-Feb-2012 (minor bugfix)
'
' copyright 2012 Lindsay Bigelow (aka raffriff aka XyKyWyKy)
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
' Fraps is a trademark of Beepa Pty Ltd.
' VirtualDub is free software under GNU General Public License.

Option Explicit

Dim gFSO            'As FileSystemObject
Dim gMakeUndo       'if True, generate an Undo bar file
Dim gUndoCleanup    'if True, delete any gernerated files on Undo
Dim gMakeAvidemux   'if True, generate an Avidemux script
Dim gMakeAvisynth   'if True, generate an Avidemux script
Dim gMakeVirtDub    'if True, generate a VirtualDub script

Dim gfUndo          'As File
Dim gUndopath
Dim gWorkPath
Dim gAdScriptPath
Dim gAvScriptPath
Dim gVdScriptPath
Dim gLastError      'message from most recent error, if any

Const APP_TITLE = "FrapsJoin"

'********************************
'
'[[VBA -- uncomment this block for VB/VBA use (ie, for debugging in an IDE)
'Public Sub Main(args)
']]
    On Error Resume Next

    Set gFSO = CreateObject("Scripting.FileSystemObject")
    TestObject gFSO, "FileSys"

    Dim dict: Set dict = CreateObject("Scripting.Dictionary")
    TestObject dict, "Dictionary"

    Dim dd 'As Folder
'[[VBA
'    gWorkPath = args(0)
'][VBS -- uncomment this block for VBScript
    gWorkPath = WScript.Arguments(0)
']]
    If ((Err <> 0) Or (Len(gWorkPath) = 0)) Then
        StatMsg APP_TITLE & ": group Fraps videos in a folder by recording " & vbCrLf & _
        "Usage: " & gFSO.GetBaseName(ScriptPath()) & ".vbs <folder>"
        QuitScript 1
    End If
    If (gFSO.FileExists(gWorkPath)) Then
        gWorkPath = gFSO.GetParentFolderName(gWorkPath)
    End If
    Set dd = gFSO.GetFolder(gWorkPath)
    If (Err) Then
        StatMsg "Can't access folder: '" & gWorkPath & "': " & Err.Description
        QuitScript 1
    End If

    gWorkPath = dd.path

    '** no Name property if network share?
    Dim ddName: ddName = Mid(gWorkPath, InStrRev(gWorkPath, "\") + 1)

    If (MsgBox(APP_TITLE & " will group Fraps videos in folder '" & ddName & "' by recording; " & vbCrLf & _
               "it will rename files, but there will be a batch file to undo the changes; " & vbCrLf & _
               "proceed?", vbYesNo Or vbQuestion, APP_TITLE) = vbNo) Then
        QuitScript 1
    End If

    Dim iniPath: iniPath = ScriptPath()
    iniPath = gFSO.GetParentFolderName(iniPath) & "\" & gFSO.GetBaseName(iniPath) & ".ini"

    '** read INI file to get user options
    If (gFSO.FileExists(iniPath)) Then
        If (IniReadSection(gFSO, iniPath, "", dict) = False) Then
            Err.Clear
            '** not sure which of the next 2 lines is preferable
            MsgBox gLastError & vbCrLf & "...quitting", vbExclamation, APP_TITLE
            QuitScript 1
        End If
    End If

    gMakeUndo = SafeBoolean(SafeDictItem(dict, "makeUndo", "true"), True)
    gUndoCleanup = SafeBoolean(SafeDictItem(dict, "undoCleanup", "true"), True)
    gMakeAvidemux = SafeBoolean(SafeDictItem(dict, "makeAvidemux", "false"), False)
    gMakeAvisynth = SafeBoolean(SafeDictItem(dict, "makeAvisynth", "false"), False)
    gMakeVirtDub = SafeBoolean(SafeDictItem(dict, "makeVirtDub", "false"), False)
    Dim avisextra:    avisextra = UnQuoteString(SafeDictItem(dict, "avisynth_extra_script", ""))
    Dim vdubextra:    vdubextra = UnQuoteString(SafeDictItem(dict, "virtualdub_extra_script", ""))
    Dim outpath:      outpath = UnQuoteString(SafeDictItem(dict, "output_path", gWorkPath))
    Dim outprefix:    outprefix = UnQuoteString(SafeDictItem(dict, "output_prefix", APP_TITLE & "-"))

'    StatMsg "gMakeUndo: " & gMakeUndo & vbCrLf & _
'            "gUndoCleanup: " & gUndoCleanup & vbCrLf & _
'            "gMakeAvidemux: " & gMakeAvidemux & vbCrLf & _
'            "gMakeAvisynth: " & gMakeAvisynth & vbCrLf & _
'            "gMakeVirtDub: " & gMakeVirtDub & vbCrLf & _
'            "avisextra: " & avisextra & vbCrLf & _
'            "vdubextra: " & vdubextra & vbCrLf & _
'            "outpath: " & outpath & vbCrLf & _
'            "outprefix: " & outprefix

    Dim adsourcepath
    If (gMakeAvidemux) Then
        adsourcepath = gWorkPath
        If Right(adsourcepath, 1) <> "\" Then
            adsourcepath = adsourcepath & "\"
        End If
    End If

    If (gMakeUndo) Then

        gUndopath = gWorkPath & "\" & APP_TITLE & "-undo.bat"

        If (gFSO.FileExists(gUndopath)) Then
            gFSO.DeleteFile (gUndopath)
            Err.Clear
        End If

        Set gfUndo = gFSO.OpenTextFile(gUndopath, 2, True) 'ForWriting
        If (Err) Then
            MsgBox "can't open '" & gUndopath & "' for writing: " & Err.Description, vbExclamation, APP_TITLE
            gfUndo.Close
            gFSO.DeleteFile (gUndopath)
            Err.Clear
            QuitScript 1
        End If
    End If

    Dim adscript:    Set adscript = Nothing
    Dim avscript:    Set avscript = Nothing
    Dim vdscript:    Set vdscript = Nothing
    Dim index:       index = 0
    Dim renamecount: renamecount = 0
    Dim prevdate:    prevdate = 0

    Dim matches, match
    Dim groupname, newname
    Dim adsavepath, vdsavepath
    Dim groupdate, curdate
    Dim f, prev

    Const defprompt1 = "Enter name for this group of videos, or hit Enter to accept the default:"

    '** search for "* YYYY-MM-DD HH-MM-SS-ms.avi"
    Dim regex: Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "^(.*) (\d\d\d\d-\d\d-\d\d) (\d\d-\d\d-\d\d)-\d\d\.avi$"
    If (Err) Then
        StatMsg "Error (15): " & Err.Description
        QuitScript 1
    End If

    For Each f In dd.Files 'note: ASSUMED files sorted by name or by date (doesn't matter which)

        Set matches = regex.Execute(f.Name)
        If (Err) Then
            StatMsg "Error (20): " & Err.Description
            QuitScript 1
        End If

        For Each match In matches

            curdate = CDate(match.submatches(1) & " " & Replace(match.submatches(2), "-", ":"))
            If (Err) Then
                StatMsg "Error (25): " & Err.Description
                QuitScript 1
            End If

            If (groupdate <> f.DateLastModified) Then
                index = 0
                groupdate = f.DateLastModified
            End If

            If (curdate < prevdate) Then
                index = 0
            End If

            If (index = 0) Then
                If (gMakeAvidemux) Then
                    finalAvidemux adscript, adsavepath
                End If
                If (gMakeAvisynth) Then
                    finalAvisynth avscript, avisextra
                End If
                If (gMakeVirtDub) Then
                    finalVirtualdub vdscript, vdubextra, vdsavepath
                End If
            End If

            'StatMsg "name =" & f.Name & vbCrLf & _
            '        "date = " & curdate & vbCrLf & _
            '        "DateCreated = " & f.DateCreated & vbCrLf & _
            '        "DateLastModified = " & f.DateLastModified

            If (index = 1) Then
                '
                ' rename first video to *-000.avi
                '
                Dim prompt: prompt = defprompt1
                Do
                    groupname = InputBox(prompt, APP_TITLE, _
                                         match.submatches(0) & " " & _
                                         match.submatches(1) & " " & _
                                         match.submatches(2))
                    If (Len(groupname) = 0) Then
                        '** user canceled
                        QuitScript 1
                    End If
                    newname = groupname & "-" & zstr(0, 3) & ".avi"
                    If (False = gFSO.FileExists(gWorkPath & "\" & newname)) Then
                        Exit Do
                    Else
                        prompt = "'" & newname & "' exists: " & vbCrLf & defprompt1
                    End If
                Loop

                'StatMsg "rename " & prev.Name & " to " & newname
                If (gMakeUndo) Then
                    gfUndo.WriteLine "ren """ & newname & """ """ & prev.Name & """"
                    gfUndo.WriteLine "if errorlevel 1 pause"
                    gfUndo.WriteLine "if errorlevel 1 goto :EOF"
                    If (Err) Then
                        StatMsg "Can't write to undo file: " & Err.Description
                        QuitScript 1
                    End If
                End If
                prev.Name = newname
                renamecount = renamecount + 1
                If (Err) Then
                    StatMsg "Can't rename file: " & Err.Description
                    QuitScript 1
                End If

                If (gMakeAvidemux) Then
                    '** start avidemux script
                    gAdScriptPath = PathCheck(gWorkPath, "generated Avidemux script name", outprefix, groupname, "js")

                    Set adscript = dd.CreateTextFile(gFSO.GetBaseName(gAdScriptPath) & ".js", True)
                    adscript.WriteLine "//AD" & vbCrLf
                    adscript.WriteLine "var app = new Avidemux();" & vbCrLf
                    adscript.WriteLine "app.load(""" & Replace(adsourcepath, "\", "/") & newname & """);"
                    If (Err) Then
                        MsgBox "can't write to '" & gAdScriptPath & "': " & Err.Description, _
                               vbExclamation, APP_TITLE
                        adscript.Close
                        gFSO.DeleteFile (gAdScriptPath)
                        QuitScript 1
                    End If

                    adsavepath = PathCheck(outpath, "Avidemux saved MP4 name", outprefix, groupname, "mp4")
                End If

                If (gMakeAvisynth) Then
                    '** start Avisynth script
                    gAvScriptPath = PathCheck(gWorkPath, "generated Avisynth script name", outprefix, groupname, "avs")

                    Set avscript = dd.CreateTextFile(gFSO.GetBaseName(gAvScriptPath) & ".avs", True)
                    avscript.WriteLine "#Avisynth" & vbCrLf
                    avscript.WriteLine "C = AviSource(""" & prev.path & """)"
                    If (Err) Then
                        MsgBox "can't write to '" & gAvScriptPath & "': " & Err.Description, _
                               vbExclamation, APP_TITLE
                        avscript.Close
                        gFSO.DeleteFile (gAvScriptPath)
                        QuitScript 1
                    End If
                End If

                If (gMakeVirtDub) Then
                    '** start V-dub script
                    gVdScriptPath = PathCheck(gWorkPath, "generated VirtualDub script name", outprefix, groupname, "vcf")

                    Set vdscript = dd.CreateTextFile(gFSO.GetBaseName(gVdScriptPath) & ".vcf", True)
                    vdscript.WriteLine "VirtualDub.Open(""" & Replace(prev.path, "\", "\\") & """);"
                    If (Err) Then
                        MsgBox "can't write to '" & gVdScriptPath & "': " & Err.Description, _
                               vbExclamation, APP_TITLE
                        vdscript.Close
                        gFSO.DeleteFile (gVdScriptPath)
                        QuitScript 1
                    End If

                    vdsavepath = PathCheck(outpath, "VirtualDub saved AVI name", outprefix, groupname, "avi")
                End If
            End If

            If (index > 0) Then

                newname = groupname & "-" & zstr(index, 3) & ".avi"

                'StatMsg "rename " & f.Name & " to " & newname
                If (gMakeUndo) Then
                    gfUndo.WriteLine "ren """ & newname & """ """ & f.Name & """"
                    If (Err) Then
                        StatMsg "Can't write to undo file: " & Err.Description
                        QuitScript 1
                    End If
                End If
                f.Name = newname
                renamecount = renamecount + 1
                If (Err) Then
                    StatMsg "Can't rename file: " & Err.Description
                    QuitScript 1
                End If

                If (gMakeAvidemux) Then
                    adscript.WriteLine "app.append(""" & Replace(adsourcepath, "\", "/") & newname & """);"
                    If (Err) Then
                        StatMsg "Can't write Avidemux script: " & Err.Description
                        QuitScript 1
                    End If
                End If

                If (gMakeAvisynth) Then
                    avscript.WriteLine "C = C + AviSource(""" & f.path & """)"
                    If (Err) Then
                        StatMsg "Can't write Avisynth script: " & Err.Description
                        QuitScript 1
                    End If
                End If

                If (gMakeVirtDub) Then
                    vdscript.WriteLine "VirtualDub.Append(""" & Replace(f.path, "\", "\\") & """);"
                    If (Err) Then
                        StatMsg "Can't write VirtualDub script: " & Err.Description
                        QuitScript 1
                    End If
                End If
            End If

            index = index + 1
            prevdate = curdate
            Set prev = f
        Next
    Next

    If (gMakeAvidemux) Then
        finalAvidemux adscript, adsavepath
        Err.Clear
    End If

    If (gMakeAvisynth) Then
        finalAvisynth avscript, avisextra
        Err.Clear
    End If

    If (gMakeVirtDub) Then
        finalVirtualdub vdscript, vdubextra, vdsavepath
        Err.Clear
    End If

    If (renamecount = 0) Then
        StatMsg "no files renamed"
        gfUndo.Close
        gFSO.DeleteFile (gUndopath)
        Err.Clear
    Else
        StatMsg "renamed " & renamecount & " files; use " & _
                gFSO.GetBaseName(gUndopath) & ".bat to undo"
    End If    

    If (SafeFileLength(gAdScriptPath) = 0) Then
        gFSO.DeleteFile (gAdScriptPath)
    End If
    If (SafeFileLength(gAvScriptPath) = 0) Then
        gFSO.DeleteFile (gAvScriptPath)
    End If
    If (SafeFileLength(gVdScriptPath) = 0) Then
        gFSO.DeleteFile (gVdScriptPath)
    End If
    gfUndo.Close
    If (SafeFileLength(gUndopath) = 0) Then
        gFSO.DeleteFile (gUndopath)
    End If

    QuitScript 0
'[[VBA
'End Sub
']]

'********************************
'** finalize Avidemux script
'
Sub finalAvidemux(adscript, ByVal adsavepath)

    If (adscript Is Nothing) Then
        Exit Sub
    End If
    adscript.WriteLine ""
    adscript.WriteLine "//app.save(""" & Replace(adsavepath, "\", "/") & """);"
    adscript.Close
    Set adscript = Nothing

End Sub

'********************************
'** finalize Avisynth script
'
Sub finalAvisynth(avscript, ByVal avisextra)

    If (avscript Is Nothing) Then
        Exit Sub
    End If
    avscript.WriteLine vbCrLf & "C ### set special variable 'Last' = C"

    'optional postprocessing
    'look for file in explicit path first, userscripts folder second, working folder third

    Dim scriptBasePath: scriptBasePath = gFSO.GetParentFolderName(ScriptPath())
    Dim pathAvisScript: pathAvisScript = ""

    If (gFSO.FileExists(avisextra)) Then
        pathAvisScript = avisextra
    ElseIf (gFSO.FileExists(scriptBasePath & "\userscripts\" & avisextra)) Then
        pathAvisScript = scriptBasePath & "\userscripts\" & avisextra
    ElseIf (gFSO.FileExists(gWorkPath & "\" & avisextra)) Then
        pathAvisScript = gWorkPath & "\" & avisextra
    Else
        If (Len(avisextra) > 0) Then
            StatMsg "Avisynth extra script '" & avisextra & "' not found; ignoring"
            avisextra = ""
        End If
        avscript.Close
        Set avscript = Nothing
        Err.Clear
        Exit Sub
    End If

    On Error Resume Next

    Dim ts: Set ts = gFSO.OpenTextFile(pathAvisScript, 1) 'ForReading
    If (Err) Then
        StatMsg "Error opening '" & pathAvisScript & "': " & Err.Description
        ts.Close
        avscript.Close
        Set avscript = Nothing
        Err.Clear
        Exit Sub
    End If

    Dim sLine
    Do While Not (ts.AtEndOfStream)
        sLine = ts.ReadLine
        avscript.WriteLine sLine
    Loop
    ts.Close
    avscript.Close
    Set avscript = Nothing
    Err.Clear

End Sub

'********************************
'** finalize VirtualDub script
'
Sub finalVirtualdub(vdscript, vdubextra, ByVal vdsavepath)

    If (vdscript Is Nothing) Then
        Exit Sub
    End If

    'optional postprocessing
    'look for file in explicit path first, userscripts folder second, working folder third

    Dim scriptBasePath: scriptBasePath = gFSO.GetParentFolderName(ScriptPath())
    Dim pathVdubScript: pathVdubScript = ""

    If (gFSO.FileExists(vdubextra)) Then
        pathVdubScript = vdubextra
    ElseIf (gFSO.FileExists(scriptBasePath & "\userscripts\" & vdubextra)) Then
        pathVdubScript = scriptBasePath & "\userscripts\" & vdubextra
    ElseIf (gFSO.FileExists(gWorkPath & "\" & vdubextra)) Then
        pathVdubScript = gWorkPath & "\" & vdubextra
    Else
        If (Len(vdubextra) > 0) Then
            StatMsg "VirtualDub extra script '" & vdubextra & "' not found; ignoring"
            vdubextra = ""
        End If
        vdscript.Close
        Set vdscript = Nothing
        Err.Clear
        Exit Sub
    End If

    On Error Resume Next

    Dim ts: Set ts = gFSO.OpenTextFile(pathVdubScript, 1) 'ForReading
    If (Err) Then
        StatMsg "Error opening '" & pathVdubScript & "': " & Err.Description
        ts.Close
        vdscript.Close
        Set vdscript = Nothing
        Err.Clear
        Exit Sub
    End If

    Dim joinavi: joinavi = Replace(vdsavepath, "\", "\\")
    Dim joinwav: joinwav = Left(joinavi, Len(joinavi) - 4) & ".wav"
    Dim sLine

    Do While Not (ts.AtEndOfStream)
        sLine = ts.ReadLine
        If (InStr(1, sLine, ".Save", vbTextCompare) > 0) Then
            If (InStr(1, sLine, ".SaveWAV", vbTextCompare) > 0) Then
                vdsavepath = joinwav
            Else
                vdsavepath = joinavi
            End If
            sLine = Replace(sLine, "%1%", "%1")
            sLine = Replace(sLine, "%1", """" & vdsavepath & """")
        End If
        vdscript.WriteLine sLine
    Loop
    ts.Close
    vdscript.Close
    Set vdscript = Nothing
    Err.Clear

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
'   - NOTE if duplicate "name=" keys exist, the last one wins
'
'@param fso         - Scripting.FileSystemOject
'@param iniPath     - path to config file
'@param sSection    - name of section to be loaded (see notes above)
'@param dict        - a Dictionary object; values loaded from config file will be added here,
'                     overriding existing values in dictionary, if any
'
'@author Lindsay Bigelow 2010
'
Function IniReadSection(fso, ByVal iniPath, ByVal sSection, dict) 'As Boolean

    On Error Resume Next

    Const ForReading = 1

    Dim ts: Set ts = fso.OpenTextFile(iniPath, ForReading)
    If (Err) Then
        gLastError = "Error opening '" & iniPath & "': " & Err.Description
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
Sub QuitScript(ByVal errCode)

    On Error Resume Next

    If (gMakeUndo) Then
        If (gUndoCleanup) Then
            If (gMakeAvidemux) Then
                gfUndo.WriteLine "del """ & gAdScriptPath & """"
            End If
            If (gMakeAvisynth) Then
                gfUndo.WriteLine "del """ & gAvScriptPath & """"
            End If
            If (gMakeVirtDub) Then
                gfUndo.WriteLine "del """ & gVdScriptPath & """"
            End If
            gfUndo.WriteLine "del """ & gUndopath & """"
        End If
        gfUndo.Close
        If (SafeFileLength(gUndopath) = 0) Then
            gFSO.DeleteFile (gUndopath)
        End If
        Err.Clear
    End If

'[[VBA
'    End
'][VBS
   WScript.Quit errCode
']]
End Sub

'********************************
'** return a valid, non-conflicting path for writing, or quit script
'
Function PathCheck(ByVal path, ByVal filedesc, ByVal outprefix, ByVal groupname, ByVal extension) 'As String

    On Error Resume Next

    Const defprompt2 = "File exists - please enter another file name:"

    Dim tempname: tempname = outprefix & groupname & "." & extension
    Dim temppath: temppath = path & "\" & tempname
    Do While (gFSO.FileExists(temppath) And (SafeFileLength(temppath) > 0))

        tempname = InputBox(defprompt2, APP_TITLE & ": " & filedesc, tempname)
        If (Len(tempname) = 0) Then
            '** user canceled
            QuitScript 1
        End If
        temppath = path & "\" & gFSO.GetBaseName(tempname) & "." & extension
    Loop
    PathCheck = temppath
End Function

'********************************
'** return Boolean from any argument; handle all errors;
'   if conversion cannot be performed, return default
'
'   if argument can be converted to Integer, "0" = False, else = True;
'   else if argument is a String: "true" = True; all else = False
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
'** return Dictionary value for a key; handle all errors
'   if value not found - or any other error - return default
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
'** return length of specified file; handle all errors
'   on error (does not exist? sharing violation?), return -1
'
Function SafeFileLength(ByVal filepath) 'As Long
    On Error Resume Next
    Err.Clear
    SafeFileLength = gFSO.GetFile(filepath).Size
    If ((Err <> 0) Or (IsEmpty(SafeFileLength))) Then
        SafeFileLength = -1
    End If
    Err.Clear
End Function

'********************************
'** location of running script
'
Function ScriptPath() 'As String
'[[VBA
'    'TEST PATH
'    ScriptPath = "\\Ava-17\Projects\vba\FrapsJoin\v2,1\FrapsJoin.vbs"
'][VBS
    ScriptPath = WScript.ScriptFullName
']]
End Function

'********************************
'** send a message to the user
'
Sub StatMsg(ByVal msg)
'[[VBA
'    Debug.Print msg
'][VBS
    WScript.Echo msg
']]
End Sub

'[[VBA
''********************************
''** test routine
''
'Sub TestMain()
'    Main Array("\\Ava-17\VideoProjects\raw")
'End Sub
']]

'********************************
'** verify object is valid, or quit script (for use with CreateObject)
'
Sub TestObject(obj, ByVal strTest)
    On Error Resume Next
    Dim t: t = TypeName(obj)
    If (InStr(1, t, strTest, vbTextCompare) = 0) Then
        If (Err) Then
            t = t & ": " & Err.Description
        End If
        MsgBox "Error initializing " & strTest & ": " & t, vbExclamation, APP_TITLE
        QuitScript 1
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

'********************************
'** return zero-padded number
'
Function zstr(number, digits)
    zstr = Right(String(digits, "0") & CStr(number), digits)
End Function


