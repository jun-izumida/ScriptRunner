Function ReadIniFile(path)
    Dim arrReadLine
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    fullpath = objFSO.BuildPath(objFSO.GetAbsolutePathName("./"), path)
    Set objIniFile = objFSO.OpenTextFile(fullpath)
    If Err.Number <> 0 Then
        wscript.echo "INIﾌｧｲﾙ名:" & conIniFileName
        wscript.quit(1)
    End If
    Set objSectionDic = CreateObject("Scripting.Dictionary")
    strReadLine = objIniFile.ReadLine
    Do While objIniFile.AtEndofStream = False
        If (strReadLine <> " ") And (StrComp("[]", (Left(strReadLine, 1) & Right(strReadLine, 1))) = 0) Then
            strSection = Mid(strReadLine, 2, (Len(strReadLine) - 2))
            Set objKeyDic = CreateObject("Scripting.Dictionary")
            Do While objIniFile.AtEndofStream = False
                strReadLine = objIniFile.ReadLine
                If (strReadLine <> "") And (StrComp(";", Left(strReadLine, 1)) <> 0) Then
                    If StrComp("[]", (Left(strReadLine, 1) & Right(strReadLine, 1))) = 0 Then
                        Exit Do
                    End If
                    arrReadLine = Split(strReadLine, "=", 2, vbTextCompare)
                    wscript.Echo strReadLine
                    objKeyDic.Add UCase(arrReadLine(0)), arrReadLine(1)
                End If
            Loop
            objSectionDic.Add UCase(strSection), objKeyDic
        Else
            strReadLine = objIniFile.ReadLine
        End If
    Loop
    objIniFile.Close

    set ReadIniFile = objSectionDic
End Function

Function PrintLabel

End Function

set a= ReadIniFile("iniget.ini")

