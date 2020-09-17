' Linux commands support for VBShell
' Use "Import" command in VBShell, DO NOT RUN THIS SCRIPT DIRECTLY!! 

Const ModelVersion = "1.0.0.2"
MsgBox "此程序为 VBS IDLE 的模块，请不要直接运行！", 16, "错误 - 模块不可直接运行"
WScript.Quit

Dim FSO, ws, SA, ADO
Dim SelfFolderPath
Set FSO = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("Wscript.Shell")
Set SA = CreateObject("Shell.Application")
Set ADO = CreateObject("ADODB.STREAM")



Function ImportExecute()
    ''' 被当做模块导入时自动执行 '''
    Import "ShellMode.vbs"
    'Call ShellMain()
End Function

'================================/Script Start/================================

Function echo(ByVal Text)
    StdOutput CStr(Text), 1
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

Function cp(ByVal Source, ByVal Destination)
    On Error Resume Next
    If FSO.FileExists(Source) Then
        If FSO.FileExists(Destination) Then
            If UCase(StdInput("是否需要覆盖已经存在的文件？(Y|N)")) = "N" Then
                Exit Function
            Else
                FSO.GetFile(Destination).Attributes = 0
                FSO.DeleteFile Destination, True
                If FSO.FileExists(Destination) Then echo "错误：覆盖文件失败！": Exit Function
            End If
        End If
        FSO.CopyFile Source, Destination, True
        If FSO.FileExists(Destination) Then echo "复制成功。" Else echo "错误：复制失败！"
    ElseIf FSO.FolderExists(Source) Then
        If FSO.FolderExists(Destination) Then
            If UCase(StdInput("是否需要覆盖已经存在的目录？(Y|N)")) = "N" Then
                Exit Function
            Else
                FSO.GetFolder(Destination).Attributes = 0
                FSO.DeleteFolder Destination, True
                If FSO.FolderExists(Destination) Then echo "错误：覆盖目录失败！": Exit Function
            End If
        End If
        FSO.CopyFolder Source, Destination, True
        If FSO.FolderExists(Destination) Then echo "复制成功。" Else echo "错误：复制失败！"
    Else
        echo "错误：文件或目录""" & Source & """不存在！"
    End If
End Function

Function touch(ByVal FP)
    If FSO.FolderExists(FP) Then echo "错误：""" & FP & """是一个已经存在的目录": Exit Function
    If Not FSO.FileExists(FP) Then
        ADO.Type = 1
        ADO.Open
        ADO.SaveToFile FP
        ADO.Close
    Else
        Dim objFolder, F
        Set F = FSO.GetFile(FP)
        Set objFolder = SA.NameSpace(F.ParentFolder.Path)
        objFolder.Items.Item(F.Name).ModifyDate = Now
    End If
End Function

Function mkdir(ByVal FP)
    On Error Resume Next
    If FSO.FolderExists(FP) Then echo "错误：""" & FP & """是一个已经存在的目录": Exit Function
    If FSO.FileExists(FP) Then echo "错误：""" & FP & """是一个已经存在的文件": Exit Function
    FSO.CreateFolder FP
    If FSO.FolderExists(FP) Then echo "目录创建成功" Else echo "目录创建失败"
End Function

Function md(ByVal FP)
    md = mkdir(FP)
End Function

Function rm(ByVal FP)
    On Error Resume Next
    If FSO.FileExists(FP) Then
        FSO.GetFile(FP).Attributes = 0
        FSO.DeleteFile FP, True
        If Not FSO.FileExists(FP) Then echo "文件删除成功" Else echo "错误：文件删除失败"
    ElseIf FSO.FolderExists(FP) Then
        FSO.GetFolder(FP).Attributes = 0
        FSO.DeleteFolder FP, True
        If Not FSO.FolderExists(FP) Then echo "目录删除成功" Else echo "错误：目录删除失败"
    Else
        echo "错误：文件或目录不存在"
    End If
End Function