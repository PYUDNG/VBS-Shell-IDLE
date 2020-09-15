' Linux commands support for VBShell
' Use "Import" command in VBShell, DO NOT RUN THIS SCRIPT DIRECTLY!! 

MsgBox "�˳���ΪVBSSHELL��ģ�飬�벻Ҫֱ�����У�", 16, "���� - ģ�鲻��ֱ������"
WScript.Quit

'==============================/Import Execution/==============================

Function ImportExecute()
    ''' ������ģ�鵼��ʱ�Զ�ִ�� '''
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
    echo vbCrLf & ws.CurrentDirectory & "�µ��ļ��к��ļ���"
    echo Join(list, vbCrLf)
End Function

Function cd(ByVal NewPath)
    If FSO.FolderExists(NewPath) Then ws.CurrentDirectory = NewPath Else echo "�����ļ��� " & NewPath & " ������"
End Function