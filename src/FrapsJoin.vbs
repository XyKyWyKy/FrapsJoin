'Attribute VB_Name = "modFrapsJoin"
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

Dim gFSO 'As FileSystemObject
Dim gWorkPath  
Dim gLastError 'message from most recent error, if any

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

    On Error GoTo 0
    
    gWorkPath = dd.Path

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
            'MsgBox gLastError & vbCrLf & "reverting to default settings", vbExclamation, APP_TITLE
            QuitScript 1
        End If
    End If

    Dim makeUndo:     makeUndo = SafeBoolean(SafeDictItem(dict, "makeUndo", "true"), True)
    Dim undoCleanup:  undoCleanup = SafeBoolean(SafeDictItem(dict, "undoCleanup", "true"), True)
    Dim makeAvidemux: makeAvidemux = SafeBoolean(SafeDictItem(dict, "makeAvidemux", "false"), False)
    Dim makeAviSynth: makeAviSynth = SafeBoolean(SafeDictItem(dict, "makeAviSynth", "false"), False)
    Dim makeVirtDub:  makeVirtDub = SafeBoolean(SafeDictItem(dict, "makeVirtDub", "false"), False)
    Dim avisextra:    avisextra = UnQuoteString(SafeDictItem(dict, "avisynth_extra_script", ""))
    Dim vdubextra:    vdubextra = UnQuoteString(SafeDictItem(dict, "virtualdub_extra_script", ""))
    Dim vduboutpath:  vduboutpath = UnQuoteString(SafeDictItem(dict, "virtualdub_output_path", gWorkPath))

'    StatMsg "makeUndo: " & makeUndo & vbCrLf & _
'            "undoCleanup: " & undoCleanup & vbCrLf & _
'            "makeAvidemux: " & makeAvidemux & vbCrLf & _
'            "makeAviSynth: " & makeAviSynth & vbCrLf & _
'            "makeVirtDub: " & makeVirtDub & vbCrLf & _
'            "avisextra: " & avisextra & vbCrLf & _
'            "vdubextra: " & vdubextra & vbCrLf & _
'            "vduboutpath: " & vduboutpath

    Dim dpath
    If (makeAvidemux) Then
        dpath = gWorkPath
        If Right(dpath, 1) <> "\" Then
            dpath = dpath & "\"
        End If
        dpath = Replace(dpath, "\", "/")
    End If

    Dim uu, upath

    If (makeUndo) Then

        upath = gWorkPath & "\" & APP_TITLE & "-undo.bat"

        If (gFSO.FileExists(upath)) Then
            gFSO.DeleteFile (upath)
        End If

        Set uu = gFSO.OpenTextFile(upath, 2, True) 'ForWriting
        If (Err) Then
            MsgBox "can't open '" & upath & "' for writing: " & Err.Description, vbExclamation, APP_TITLE
            uu.Close
            gFSO.DeleteFile (upath)
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
    
    Dim adscriptpath, avscriptpath, vdscriptpath
    Dim matches, match
    Dim firstname, newname
    Dim savename, joinname
    Dim groupdate, curdate
    Dim f, prev

    Const defprompt = "Enter name for this group of videos, or hit Enter to accept the default:"
    
    '** search for "* YYYY-MM-DD HH-MM-SS-ms.avi"
    Dim regex: Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "^(.*) (\d\d\d\d-\d\d-\d\d) (\d\d-\d\d-\d\d)-\d\d\.avi$"
    
    For Each f In dd.Files 'note: ASSUMED files sorted by name or by date (doesn't matter which)
    
        Set matches = regex.Execute(f.Name)
        
        For Each match In matches
            
            curdate = CDate(match.submatches(1) & " " & Replace(match.submatches(2), "-", ":"))
            
            If (groupdate <> f.DateLastModified) Then
                index = 0
                groupdate = f.DateLastModified
            End If
            
            If (curdate < prevdate) Then
                index = 0
            End If
            
            If (index = 0) Then
                If (makeAvidemux) Then
                    ad_final adscript, savename
                End If
                If (makeAviSynth) Then
                    av_final avscript, avisextra
                End If
                If (makeVirtDub) Then
                    vd_final vdscript, vdubextra, joinname
                End If
            End If
            
            'StatMsg "name =" & f.Name & vbCrLf & _
            '        "date = " & curdate & vbCrLf & _
            '        "DateCreated = " & f.DateCreated & vbCrLf & _
            '        "DateLastModified = " & f.DateLastModified
            
            If (index = 1) Then
                '
                ' rename first video to *-00.avi
                '
                Dim prompt: prompt = defprompt
                Do
                    firstname = InputBox(prompt, APP_TITLE, match.submatches(0) & " " & match.submatches(1) & " " & match.submatches(2))
                    If (Len(firstname) = 0) Then
                        If (makeUndo) Then
                          '** user canceled
                          On Error Resume Next
                          uu.Close
                          Err.Clear
                        End If
                        QuitScript 1
                    End If
                    newname = firstname & "-" & zstr(0, 3) & ".avi"
                    If (False = gFSO.FileExists(gWorkPath & "\" & newname)) Then
                        Exit Do
                    Else
                        prompt = "'" & newname & "' exists: " & vbCrLf & defprompt
                    End If
                Loop

                'StatMsg "rename " & prev.Name & " to " & newname
                If (makeUndo) Then
                    uu.WriteLine "ren """ & newname & """ """ & prev.Name & """"
                    uu.WriteLine "if errorlevel 1 pause"
                    uu.WriteLine "if errorlevel 1 goto :EOF"
                End If
                prev.Name = newname
                renamecount = renamecount + 1
                
                If (makeAvidemux) Then
                    '** start avidemux script
                    On Error Resume Next
                    adscriptpath = APP_TITLE & "-" & firstname & ".js"
                    Set adscript = dd.CreateTextFile(adscriptpath, True)
                    adscriptpath = gWorkPath & "\" & adscriptpath
                    adscript.WriteLine "//AD" & vbCrLf
                    adscript.WriteLine "var app = new Avidemux();" & vbCrLf
                    adscript.WriteLine "app.load(""" & dpath & newname & """);"
                    If (Err) Then
                        MsgBox "can't write to '" & adscriptpath & "': " & Err.Description, vbExclamation, APP_TITLE
                        adscript.Close
                        gFSO.DeleteFile (adscriptpath)
                        If (makeUndo) Then
                            On Error Resume Next
                            uu.Close
                            If (renamecount = 0) Then
                                gFSO.DeleteFile (upath)
                            End If
                            Err.Clear
                        End If
                        QuitScript 0
                    End If
                    On Error GoTo 0
                    savename = dpath & firstname & "-compressed.mp4"
                End If

                If (makeAviSynth) Then
                    '** start avisynth script
                    On Error Resume Next
                    avscriptpath = APP_TITLE & "-" & firstname & ".avs"
                    Set avscript = dd.CreateTextFile(avscriptpath, True)
                    avscriptpath = gWorkPath & "\" & avscriptpath
                    avscript.WriteLine "#avisynth" & vbCrLf
                    avscript.WriteLine "C = AviSource(""" & prev.Path & """)"
                    If (Err) Then
                        MsgBox "can't write to '" & avscriptpath & "': " & Err.Description, vbExclamation, APP_TITLE
                        avscript.Close
                        gFSO.DeleteFile (avscriptpath)
                        If (makeUndo) Then
                            On Error Resume Next
                            uu.Close
                            If (renamecount = 0) Then
                                gFSO.DeleteFile (upath)
                            End If
                            Err.Clear
                        End If
                        QuitScript 0
                    End If
                    On Error GoTo 0
                End If

                If (makeVirtDub) Then
                    '** start V-dub script
                    On Error Resume Next
                    vdscriptpath = APP_TITLE & "-" & firstname & ".vcf"
                    Set vdscript = dd.CreateTextFile(vdscriptpath, True)
                    vdscriptpath = gWorkPath & "\" & vdscriptpath
                    vdscript.WriteLine "VirtualDub.Open(""" & Replace(prev.Path, "\", "\\") & """);"
                    If (Err) Then
                        MsgBox "can't write to '" & vdscriptpath & "': " & Err.Description, _
                               vbExclamation, APP_TITLE
                        vdscript.Close
                        gFSO.DeleteFile (vdscriptpath)
                        If (makeUndo) Then
                            On Error Resume Next
                            uu.Close
                            If (renamecount = 0) Then
                                gFSO.DeleteFile (upath)
                            End If
                            Err.Clear
                        End If
                        QuitScript 0
                    End If
                    On Error GoTo 0
                    joinname = vduboutpath & "\" & firstname & "-join.avi"
                End If
            End If
            
            If (index > 0) Then
                
                newname = firstname & "-" & zstr(index, 3) & ".avi"
                
                'StatMsg "rename " & f.Name & " to " & newname
                If (makeUndo) Then
                    uu.WriteLine "ren """ & newname & """ """ & f.Name & """"
                End If
                f.Name = newname
                renamecount = renamecount + 1
                
                If (makeAvidemux) Then
                    adscript.WriteLine "app.append(""" & dpath & newname & """);"
                End If

                If (makeAviSynth) Then
                    avscript.WriteLine "C = C + AviSource(""" & f.Path & """)"
                End If

                If (makeVirtDub) Then
                    vdscript.WriteLine "VirtualDub.Append(""" & Replace(f.Path, "\", "\\") & """);"
                End If
            End If
            
            index = index + 1
            prevdate = curdate
            Set prev = f
        Next
    Next

    On Error Resume Next

    If (makeUndo) Then
        If (undoCleanup) Then
            If (makeAvidemux) Then
                uu.WriteLine "del """ & adscriptpath & """"
            End If
            If (makeAviSynth) Then
                uu.WriteLine "del """ & avscriptpath & """"
            End If
            If (makeVirtDub) Then
                uu.WriteLine "del """ & vdscriptpath & """"
            End If
            uu.WriteLine "del """ & upath & """"
        End If
        uu.Close
        Err.Clear
    End If

    If (makeAvidemux) Then
        ad_final adscript, savename
        Err.Clear
    End If

    If (makeAviSynth) Then
        av_final avscript, avisextra
        Err.Clear
    End If

    If (makeVirtDub) Then
        vd_final vdscript, vdubextra, joinname
        Err.Clear
    End If

    If (renamecount = 0) Then
        StatMsg "no files renamed"
        gFSO.DeleteFile (upath)
        Err.Clear
    Else
        StatMsg "renamed " & renamecount & " files; use " & APP_TITLE & "-undo.bat to undo"
    End If

    QuitScript 0
'[[VBA
'End Sub
']]

'********************************
'** finalize Avidemux script
'
Sub ad_final(adscript, ByVal savename)

    If (adscript Is Nothing) Then
        Exit Sub
    End If
    adscript.WriteLine ""
    adscript.WriteLine "//app.save(""" & savename & """);"
    adscript.Close
    Set adscript = Nothing

End Sub
 
'********************************
'** finalize AviSynth script
'
Sub av_final(avscript, avisextra)

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
        If (len(avisextra) > 0) Then
            StatMsg "AviSynth extra script '" & avisextra & "' not found; ignoring"
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

'[[VBA
'    Debug.Print "...done" & vbCrLf
'    End
'][VBS
   WScript.Quit errCode
']]
End Sub

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
Function ScriptPath() 'As String
'[[VBA
'    ScriptPath = "\\Ava-17\projects\vba\FrapsJoin\v2,1\FrapsJoin.vbs"
'][VBS
    ScriptPath = WScript.ScriptFullName
']]
End Function

'********************************
'
Sub StatMsg(ByVal msg)
'[[VBA
''    Debug.Assert (InStr(1, msg, "done", vbBinaryCompare) = 0)
'    Debug.Print msg
'][VBS
    WScript.Echo msg
']]
End Sub

'[[VBA
''********************************
''
'Sub TestMain()
'    Main Array("\\ava-17\raw\tmp2")
'End Sub
']]

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
'** finalize VirtualDub script
'
Sub vd_final(vdscript, vdubextra, ByVal joinname)

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
        If (len(vdubextra) > 0) Then
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

    Dim joinavi: joinavi = Replace(joinname, "\", "\\")
    Dim joinwav: joinwav = Left(joinavi, Len(joinavi) - 4) & ".wav"
    Dim sLine

    Do While Not (ts.AtEndOfStream)
        sLine = ts.ReadLine
        If (InStr(1, sLine, ".Save", vbTextCompare) > 0) Then
            If (InStr(1, sLine, ".SaveWAV", vbTextCompare) > 0) Then
                joinname = joinwav  
            Else
                joinname = joinavi  
            End If                  
            sLine = Replace(sLine, "%1%", "%1")
            sLine = Replace(sLine, "%1", """" & joinname & """")
        End If
        vdscript.WriteLine sLine
    Loop
    ts.Close
    vdscript.Close
    Set vdscript = Nothing
    Err.Clear

End Sub
 
'********************************
'** return zero-padded number
'
Function zstr(number, digits)

    zstr = Right(String(digits, "0") & CStr(number), digits)

End Function


