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
    ''' 命令行模式主函数 '''
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
    ''' 把命令格式的代码转换成vbs格式的代码 '''
    If InStr(1, Command, Chr(13)) > 0 Or InStr(1, Command, Chr(10)) > 0 Then
        CommandToCode = Command
        Exit Function
    End If
    ' 初始化正则表达式
    Dim re, CommandArr
    Set re = New RegExp
    re.Global = True
    re.IgnoreCase = True
    re.Multiline = True
    ' 去除行首缩进
    re.Pattern = "^ +"
    Command = re.Replace(Command, "")
    ' 参数转化为字符串格式写法
    CommandArr = Split(Command, " ")
    If UBound(CommandArr) >= 1 Then
        For i = 1 To UBound(CommandArr)
            CommandArr(i) = """" & CommandArr(i) & """"
        Next
    End If
    Command = Join(CommandArr)
    ' 参数之间的空格换成英文逗号
    Command = Replace(Command, " ", ", ")
    Command = Replace(Command, ", ", " ", 1, 1)
    ' 返回结果
    CommandToCode = Command
End Function