' Linux commands support for VBShell
' Use "Import" command in VBShell, DO NOT RUN THIS SCRIPT DIRECTLY!! 

MsgBox "此程序为VBSSHELL的模块，请不要直接运行！", 16, "错误 - 模块不可直接运行"
WScript.Quit

'==============================/Import Execution/==============================

Function ImportExecute()
    ''' 被当做模块导入时自动执行 '''
    Import "ShellMode.vbs"
    'Call ShellMain()
End Function

'================================/Script Start/================================

Function echo(ByVal Text)
    StdOutput Text, 1
End Function

Function pwd()
    echo ws.CurrentDirectory
End Function

Function ls()
    Dim Folder, File, list, i
    ReDim list(0): i = 0
    For Each Folder In FSO.GetFolder(ws.CurrentDirectory).SubFolders
        ReDim Preserve list(i)
        list(i) = Folder.Name & "\"
        i = i + 1
    Next
    For Each File In FSO.GetFolder(ws.CurrentDirectory).Files
        ReDim Preserve list(i)
        list(i) = File.Name
        i = i + 1
    Next
    ls = list
    echo vbCrLf & ws.CurrentDirectory & "下的文件夹和文件："
    echo Join(list, vbCrLf)
End Function

Function cd(ByVal NewPath)
    If FSO.FolderExists(NewPath) Then ws.CurrentDirectory = NewPath Else echo "错误：文件夹 " & NewPath & " 不存在"
End Function