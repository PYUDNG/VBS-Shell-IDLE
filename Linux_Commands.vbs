' Linux commands support for VBShell
' Use "Import" command in VBShell, DO NOT RUN THIS SCRIPT DIRECTLY!! 

Const ModelVersion = "1.0.0.2"
MsgBox "�˳���Ϊ VBS IDLE ��ģ�飬�벻Ҫֱ�����У�", 16, "���� - ģ�鲻��ֱ������"
WScript.Quit

Dim FSO, ws, SA, ADO
Dim SelfFolderPath
Set FSO = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("Wscript.Shell")
Set SA = CreateObject("Shell.Application")
Set ADO = CreateObject("ADODB.STREAM")



Function ImportExecute()
    ''' ������ģ�鵼��ʱ�Զ�ִ�� '''
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
    echo vbCrLf & ws.CurrentDirectory & "�µ��ļ��к��ļ���"
    echo Join(list, vbCrLf)
End Function

Function cd(ByVal NewPath)
    If FSO.FolderExists(NewPath) Then ws.CurrentDirectory = NewPath Else echo "�����ļ��� " & NewPath & " ������"
End Function

Function cp(ByVal Source, ByVal Destination)
    On Error Resume Next
    If FSO.FileExists(Source) Then
        If FSO.FileExists(Destination) Then
            If UCase(StdInput("�Ƿ���Ҫ�����Ѿ����ڵ��ļ���(Y|N)")) = "N" Then
                Exit Function
            Else
                FSO.GetFile(Destination).Attributes = 0
                FSO.DeleteFile Destination, True
                If FSO.FileExists(Destination) Then echo "���󣺸����ļ�ʧ�ܣ�": Exit Function
            End If
        End If
        FSO.CopyFile Source, Destination, True
        If FSO.FileExists(Destination) Then echo "���Ƴɹ���" Else echo "���󣺸���ʧ�ܣ�"
    ElseIf FSO.FolderExists(Source) Then
        If FSO.FolderExists(Destination) Then
            If UCase(StdInput("�Ƿ���Ҫ�����Ѿ����ڵ�Ŀ¼��(Y|N)")) = "N" Then
                Exit Function
            Else
                FSO.GetFolder(Destination).Attributes = 0
                FSO.DeleteFolder Destination, True
                If FSO.FolderExists(Destination) Then echo "���󣺸���Ŀ¼ʧ�ܣ�": Exit Function
            End If
        End If
        FSO.CopyFolder Source, Destination, True
        If FSO.FolderExists(Destination) Then echo "���Ƴɹ���" Else echo "���󣺸���ʧ�ܣ�"
    Else
        echo "�����ļ���Ŀ¼""" & Source & """�����ڣ�"
    End If
End Function

Function touch(ByVal FP)
    If FSO.FolderExists(FP) Then echo "����""" & FP & """��һ���Ѿ����ڵ�Ŀ¼": Exit Function
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
    If FSO.FolderExists(FP) Then echo "����""" & FP & """��һ���Ѿ����ڵ�Ŀ¼": Exit Function
    If FSO.FileExists(FP) Then echo "����""" & FP & """��һ���Ѿ����ڵ��ļ�": Exit Function
    FSO.CreateFolder FP
    If FSO.FolderExists(FP) Then echo "Ŀ¼�����ɹ�" Else echo "Ŀ¼����ʧ��"
End Function

Function md(ByVal FP)
    md = mkdir(FP)
End Function

Function rm(ByVal FP)
    On Error Resume Next
    If FSO.FileExists(FP) Then
        FSO.GetFile(FP).Attributes = 0
        FSO.DeleteFile FP, True
        If Not FSO.FileExists(FP) Then echo "�ļ�ɾ���ɹ�" Else echo "�����ļ�ɾ��ʧ��"
    ElseIf FSO.FolderExists(FP) Then
        FSO.GetFolder(FP).Attributes = 0
        FSO.DeleteFolder FP, True
        If Not FSO.FolderExists(FP) Then echo "Ŀ¼ɾ���ɹ�" Else echo "����Ŀ¼ɾ��ʧ��"
    Else
        echo "�����ļ���Ŀ¼������"
    End If
End Function