' Linux commands support for VBShell
' Use "Import" command in VBShell, DO NOT RUN THIS SCRIPT DIRECTLY!! 

Const ModelVersion = "1.0.0.2"

'==============================/Import Execution/==============================

Function ImportExecute()
    Dim ModeOn
    ModeOn = False
    ShellMode True
End Function

'================================/Script Start/================================

Function ShellMain()
    ''' ������ģʽ������ '''
    If HookErrors Then On Error Resume Next
    Const Tip_Main = "VBScript ShellMode >>> "
    Dim Input, InputDealer
    WScript.Echo ""
    Do
        Input = StdInput(Tip_Main): Input = LTrim(Input)
        'If LCase(Input) = LCase("ShellMode Off") Then Exit Function
        Input = WaitForEnd(Input)
        Input = CommandToCode(Input)
        ExecuteGlobal Input
        If HookErrors Then Call ErrorDealing()
        If Not ModeOn Then Exit Function
        If Input <> "" Then WScript.Echo ""
    Loop
End Function

Function ShellMode(ONOFF)
    Select Case UCase(CStr(ONOFF))
        Case "ON", "TRUE", "1"
            If ModeOn Then Exit Function
            ModeOn = True
            Call ShellMain()
        Case "OFF", "FALSE", "0"
            If Not ModeOn Then Exit Function
            ExecuteGlobal "ModeOn = False"
    End Select
End Function

Function CommandToCode(ByVal Command)
    ''' �������ʽ�Ĵ���ת����vbs��ʽ�Ĵ��� '''
    If InStr(1, Command, Chr(13)) > 0 Or InStr(1, Command, Chr(10)) > 0 Then
        CommandToCode = Command
        Exit Function
    End If
    ' ��ʼ��������ʽ
    Dim re, CommandArr
    Set re = New RegExp
    re.Global = True
    re.IgnoreCase = True
    re.Multiline = True
    ' ȥ����������
    re.Pattern = "^ +"
    Command = re.Replace(Command, "")
    ' ����ת��Ϊ�ַ�����ʽд��
    CommandArr = Split(Command, " ")
    If UBound(CommandArr) >= 1 Then
        For i = 1 To UBound(CommandArr)
            CommandArr(i) = """" & CommandArr(i) & """"
        Next
    End If
    Command = Join(CommandArr)
    ' ����֮��Ŀո񻻳�Ӣ�Ķ���
    Command = Replace(Command, " ", ", ")
    Command = Replace(Command, ", ", " ", 1, 1)
    ' ���ؽ��
    CommandToCode = Command
End Function