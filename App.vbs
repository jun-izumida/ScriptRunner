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

Function GetIniDictionary(objDictionary, sSection, sKey)
    Dim objTemp

    strSection = UCase(sSection)
    strKey = UCase(sKey)

    If objDictionary.Exists(strSection) Then
        Set objTemp = objDictionary.Item(strSection)
        If objTemp.Exists(strKey) Then
            GetIniDictionary = objTemp.Item(strKey)
            Exit Function 
        End If
    End If
End Function

Function CleanWorkData(current)
    Dim fso
    Dim folder
    Set fso  = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.getFolder(current)
    Set fso  = CreateObject("Scripting.FileSystemObject")
    for each subfolder in folder.subfolders
        if left(subfolder.name, 1) = "_" Then
                call fso.DeleteFolder(subfolder, True)
        End If
    next
End Function

Function ConvertCsvData(current, execPath, outputPath) 
    Dim fso
    Dim shell
    Dim outExec
    Dim StdOut
    Dim StdErr
    Dim log
    Set fso  = CreateObject("Scripting.FileSystemObject")
    Set shell = CreateObject("WScript.Shell")
    Set folder = fso.getFolder(current)
    WScript.Echo execPath
    for each file in folder.files
        If fso.GetExtensionName(file) = "csv" Then
            Set outExec = shell.Exec("""" & execPath & """ --path """ & file.path & """ --output """ & outputPath & """")
            Set StdOut = outExec.StdOut
            Set StdErr = outExec.StdErr

            logging = "STDOUT" & vbCrLf
            Do While Not StdOut.AtEndofStream
                logging = log & StdOut.ReadLine() & vbCrLf
            Loop
        End If
    next

End Function

Function GetCsvDataDirectory(current)
    Dim fso
    Dim folder
    Dim dataDirs()
    Set fso  = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.getFolder(current)
    Redim dataDirs(0)
    for each subfolder in folder.subfolders
        if left(subfolder.name, 1) = "_" Then
            Redim preserve dataDirs(UBound(dataDirs) + 1)
            dataDirs(UBound(dataDirs)) = subfolder.path
        End If
    next
    GetCsvDataDirectory = dataDirs
End Function

Function PrintLabel(printerName, appPath, tpePath, dataPath)
    Dim shell
    Dim outExec
    Dim StrOut
    Dim StdErr
    Dim logging
    Set shell = CreateObject("WScript.Shell")
    sCommand = """" & appPath & """ /p """ & dataPath & """ /C -fn -h"

    'Set outExec = shell.Exec(sCommnad)
    Set outExec = shell.Exec("""" & appPath & """ /p """ & tpePath & "," & dataPath & """ /C -fn -h /TW -off")
    WScript.Echo """" & appPath & """ /p """ & tpePath & "," & dataPath & """ /C -fn -h"
    Set StdOut = outExec.StdOut
    Set StdErr = outExec.StdErr

    logging = "STDOUT" & vbCrLf
    Do While Not StdOut.AtEndofStream
        logging = logging & StdOut.ReadLine() & vbCrLf
    Loop
    WScript.Echo logging

End Function




Dim tpeType
Dim objFSO
Dim objIniDictionary
Dim objWshShell
Dim appPath
Dim tpePath

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objIniDictionary = ReadIniFile("nBlesPUB.ini")
Set objWshShell = WScript.CreateObject("WScript.Shell")
PrinterName = GetIniDictionary(objIniDictionary, "general", "PrinterName")
appPath = GetIniDictionary(objIniDictionary, "general", "SPC10")
tpePath = GetIniDictionary(objIniDictionary, "general", "TpeDirectory")
tpeName = GetIniDictionary(objIniDictionary, "settings", tpeType)
Set objArgs = WScript.Arguments

If objArgs.Count = 0 Then
    tpeType = ""
Else
    tpeType = objArgs(0)
End If


call CleanWorkData(objWshShell.CurrentDirectory & "\\" & "data")

call ConvertCsvData(objWshShell.CurrentDirectory, _
    objWshShell.CurrentDirectory & "\\" & GetIniDictionary(objIniDictionary, "general", "CsvSplit"), _
    objWshShell.CurrentDirectory & "\\Data")

'csvfiles 
dataDirs = GetCsvDataDirectory(objWshShell.CurrentDirectory & "\\" & "data")

If objFSO.FileExists("tpePath" & "\" & tpeType) Then
End If

for each dataDir in dataDirs
    if dataDir <> "" Then
        call PrintLabel(printerName, appPath, objWshShell.CurrentDirectory & "\tpe\nBlesPUB_Batch.tpe", dataDir & "\BATCH.csv" )
        call PrintLabel(printerName, appPath, objWshShell.CurrentDirectory & ".\tpe\nBlesPUB_ID.tpe", dataDir & "\ID.csv")
        WScript.Echo dataDir
    End If
next
