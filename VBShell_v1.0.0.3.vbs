Option Explicit
'Call ISOLATE_RUN()

Dim FSO, ws, SA
Dim SelfFolderPath, SI, SO
Set FSO = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("Wscript.Shell")
Set SA = CreateObject("Shell.Application")

Call GetUAC(2, False)

SelfFolderPath = FormatPath(FSO.GetFile(WScript.ScriptFullName).ParentFolder.Path)
Set SI = WScript.StdIn: Set SO = WScript.StdOut
Const Tip_Main = "VBScript >>> "
Const Tip_Wait = "------------ "
Call StartOutput()
Call Main()

Function ISOLATE_RUN()
    ''' 创建隔离环境，在隔离环境下执行本程序，以免影响到用户代码的执行（注：本函数尚未实现，不要调用！） '''
    Execute Replace(CreateObject("Scripting.FileSystemObject").OpenTextFile(WScript.ScriptFullName).ReadAll(), vbCrLf & "Call ISOLATE_RUN()", "")
    WScript.Quit
End Function

Function StdInput(ByVal Text)
    ''' 输出一段文字然后接受用户输入，类似Python的input '''
    On Error Resume Next
    SO.Write Text
    StdInput = SI.ReadLine()
End Function

Function StdOutput(ByVal Content, ByVal WithCrLfs)
    ''' 输出多行文本，Content既可以是文本型数组也可以是文本，如果是数组就用vbCrLf（换行符组合）连接其成员后输出 '''
    ''' WithCrLfs指定Content输出完毕后输出几个换行符组合 '''
    On Error Resume Next
    Dim All(), i
    If IsArray(Content) Then Content = Join(Content, vbCrLf)
    ReDim All(WithCrLfs)
    All(0) = Content
    For i = 1 To WithCrLfs
        All(i) = vbCrLf
    Next
    Content = Join(All, "")
    SO.Write Content
    StdOutput = Content
End Function

Function GetIfWaiting(ByVal Code)
    ''' 用于判断是否为一个新代码块的开始，如果是，就返回结束标志，否则返回空字符串 '''
    ' 初始化[开始标志-结束标志]字典
    Dim Dict
    Set Dict = CreateObject("Scripting.Dictionary")
    Dict.Add "if", "end if": Dict.Add "do", "loop": Dict.Add "for", "next": Dict.Add "while", "wend": Dict.Add "select case", "end select"
    Dict.Add "function", "end function": Dict.Add "sub", "end sub": Dict.Add "class", "end class"
    ' 判断开始标志，返回结果
    GetIfWaiting = ""
    Dim StartFlag
    For Each StartFlag In Dict.Keys()
        If LCase(Left(Code, Len(StartFlag))) = StartFlag Then 
            GetIfWaiting = Dict.Item(StartFlag)
            Exit For
        End If
    Next
End Function

Function WaitForEnd(ByVal Code)
    ''' 读入语句块，直到当前语句块层结束；返回语句块全部内容 '''
    On Error Resume Next
    ' 是否需要开启新语句块？
    Dim NowWaiting
    NowWaiting = GetIfWaiting(Code)
    If NowWaiting = "" Then 
        WaitForEnd = Code
        Exit Function
    End If
    ' 需要等待语句块结束
    Dim Input, Block
    Block = Code
    Do
        Input = StdInput(Tip_Wait): Input = LCase(LTrim(Input))
        Block = Block & vbCrLf & WaitForEnd(Input)
        If Left(Input, Len(NowWaiting)) = NowWaiting Then Exit Do
    Loop
    WaitForEnd = Block
End Function

Function ErrorDealing()
    ''' 处理错误信息，简单地说就是向标准输出流输出错误代号和错误文本信息 '''
    If Err.number <> 0 Then Stdoutput Array("错误: " & CStr(Err.number), Err.Description), 1: Err.Clear
End Function

Function Main()
    ''' 程序主功能 '''
    On Error Resume Next
    Dim Input
    Do
        Input = StdInput(Tip_Main): Input = LTrim(Input)
        Input = WaitForEnd(Input)
        ExecuteGlobal Input
        Call ErrorDealing()
        If Input <> "" Then WScript.Echo ""
    Loop
End Function

Function StartOutput()
    On Error Resume Next
    Dim SOPText, UIL
    UIL = GetUILanguage()
    Select Case UIL
        Case &H409 'English
            SOPText = Array("VBS Shell Written By PY-DNG(R)",_ 
                    """AN IDLE All About VBScript""",_ 
                    """Copyright(C) PY-DNG. All Rights Reserved.""",_
                    "Enter ""Help"" to get help. Enter ""Tips"" to get a tip. ")
        Case &H804 '中文
            SOPText = Array("VBS Shell|作者：PY-DNG(R)",_ 
                    """属于VBScript的IDLE""",_ 
                    "版权所有(C) PY-DNG。保留一切权利。",_
                    "输入""Help""以获取帮助。输入""Tips""以获取提示。")
    End Select
    StdOutput SOPText, 2
End Function

Function Import(ByVal FP)
    ''' 用于引用模块，类似Python的from FP import *，不同的是，本函数只会导入sub和function，变量、对象均不会导入 '''
    On Error Resume Next
    If Not FSO.FileExists(FP) Then
        Import = -1
        Exit Function
    End If
    Dim Code
    Dim Funcs(), Line, InFuc, FuncCode, Count
    InFuc = False: Count = -1: Code = Split(FSO.OpenTextFile(FP).ReadAll(), vbCrLf)
    For Each Line In Code
        InFuc = InFuc Or (LCase(Left(LTrim(Line), 9)) = "function " Or LCase(Left(LTrim(Line), 4)) = "sub ")
        If InFuc Then 
            Count = Count + 1
            If Count = 0 Then ReDim Funcs(Count) Else ReDim Preserve Funcs(Count)
            Funcs(Count) = Line
        End If
        InFuc = InFuc Xor (LCase(Left(LTrim(Line), 12)) = "end function" Or LCase(Left(LTrim(Line), 7)) = "end sub")
    Next
    FuncCode = Join(Funcs, vbCrLf)
    ExecuteGlobal FuncCode
    Import = 0
End Function



Function Tips()
    ''' 为用户提供提示 '''
    Dim Tips_
    Tips_ = Array("试试Import？使用""Import xxx.vbs""语句导入你自己的函数！",_ 
                  "要不要读读我的源代码？",_ 
                  "有些变量名和函数名是""VBS Shell""本身已经使用了的(输入ShowUsed显示这些这些名称)，请尽量不要重新定义这些名称哦~",_ 
                  "不会用？尝试输入一些VBScript语句！",_ 
                  "提示不止一条哦~ 每次输入""Tips""都会随机输出一条哦！",_ 
                  "输入Help以获取帮助。")
    Randomize
    StdOutput Tips_(Int(Rnd * UBound(Tips_))), 1
End Function

Function ShowUsed()
    ''' 获取所有本程序使用的（全局）变量名、函数方法名、class类名的函数 '''
    ' 使用本函数需要SplitVBSLines的支持
    Dim Self, Lines, Line, All, Name
    Dim Names(2), Names_Count(2), Names_Count_Old(2), Variable
    Dim i, NameType, Depth
    Dim NumDisplay
    Self = FSO.OpenTextFile(WScript.ScriptFullName).ReadAll
    Lines = SplitVBSLines(ClearREM(ClearStrings(Self)))
    Names_Count(0) = -1: Names_Count(1) = -1: Names_Count(2) = -1: Depth = 0
    For Each Line In Lines
        Line = LTrim(Line)
        NameType = -1
        If LCase(Left(Line, 17)) = "private function " Then
            Line = Right(Line, Len(Line) - 17)
            Depth = Depth + 1
            NameType = 1
        ElseIf LCase(Left(Line, 12)) = "private sub" Then
            Line = Right(Line, Len(Line) - 12)
            Depth = Depth + 1
            NameType = 1
        ElseIf LCase(Left(Line, 16)) = "public function " Then
            Line = Right(Line, Len(Line) - 16)
            Depth = Depth + 1
            NameType = 1
        ElseIf LCase(Left(Line, 11)) = "public sub" Then
            Line = Right(Line, Len(Line) - 11)
            Depth = Depth + 1
            NameType = 1
        ElseIf LCase(Left(Line, 4)) = "dim " Then
            Line = Right(Line, Len(Line) - 4)
            NameType = 0
        ElseIf LCase(Left(Line, 6)) = "redim " Then
            Line = Right(Line, Len(Line) - 6)
            NameType = 0
        ElseIf LCase(Left(Line, 8)) = "private " Then
            Line = Right(Line, Len(Line) - 8)
            NameType = 0
        ElseIf LCase(Left(Line, 7)) = "public " Then
            Line = Right(Line, Len(Line) - 7)
            NameType = 0
        ElseIf LCase(Left(Line, 14)) = "private const " Then
            Line = Right(Line, Len(Line) - 14)
            NameType = 0
        ElseIf LCase(Left(Line, 13)) = "public const " Then
            Line = Right(Line, Len(Line) - 13)
            NameType = 0
        ElseIf LCase(Left(Line, 9)) = "function " Then
            Line = Right(Line, Len(Line) - 9)
            Depth = Depth + 1
            NameType = 1
        ElseIf LCase(Left(Line, 4)) = "sub" Then
            Line = Right(Line, Len(Line) - 4)
            Depth = Depth + 1
            NameType = 1
        ElseIf LCase(Left(Line, 6)) = "class " Then
            Line = Right(Line, Len(Line) - 6)
            Depth = Depth + 1
            NameType = 2
        ElseIf LCase(Left(Line, 9)) = "end class" Then
            Depth = Depth - 1
        ElseIf LCase(Left(Line, 7)) = "end sub" Then
            Depth = Depth - 1
        ElseIf LCase(Left(Line, 12)) = "end function" Then
            Depth = Depth - 1
        End If
        If NameType >= 0 And (Depth = 0 Or NameType <> 0) Then 
            If NameType = 0 Then Line = Replace(Line, " ", "")
            If NameType = 0 Then All = Split(Line, ",") Else All = Array(Line)
            Names_Count_Old(NameType) = Names_Count(NameType)
            Names_Count(NameType) = Names_Count(NameType) + UBound(All) + 1
            If Names_Count_Old(NameType) = -1 Then
                Names(NameType) = All
            Else
                Name = Names(NameType): ReDim Preserve Name(Names_Count(NameType))
                Names(NameType) = Name
                For i = Names_Count_Old(NameType) + 1 To Names_Count(NameType)
                    Names(NameType)(i) = All(i - Names_Count_Old(NameType) - 1)
                Next
            End If
        End If
    Next
    StdOutput "本程序定义的全局变量名(" & CStr(Names_Count(0) + 1) & "个)：", 1
    StdOutput Names(0), 2
    StdOutput "本程序定义的函数方法名(" & CStr(Names_Count(1) + 1) & "个)：", 1
    StdOutput Names(1), 2
    StdOutput "本程序定义的class名(" & CStr(Names_Count(2) + 1) & "个)：", 1
    StdOutput Names(2), 2
    ShowUsed = Names
End Function


Function Help()
    ''' 用户帮助 '''
    StdOutput Array("◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆",_
                    "◇※ Help| 用户帮助					◇",_ 
                    "◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆",_ 
                    "◇   VBSShell是一个用VBS编写的VBSyuyan本身的IDLE，类似于◇",_ 
                    "◆Python自带的的IDLE，旨在于让VBS编程更加简单。 在这里你◆",_ 
                    "◇可以直观地执行VBScript语句，逐行设计你的程序并实时查看◇",_ 
                    "◆结果。						◆",_ 
                    "◇   上面扯了那么多，就是要装个13， 实际上它并没有那么牛◇",_ 
                    "◆比。废话不多说，下面说用法：				◆",_ 
                    "◇   一般情况，输入你想执行的VBScript语句，按下回车； 这◇",_ 
                    "◆时，你刚才输入的语句就会被立即执行，并产生相应的效果。◆",_ 
                    "◇然后，你可以输入下一条语句。VBSShell会记住你每次定义的◇",_ 
                    "◆变量、对象、函数、类，直到VBSShell退出为止。也就是说，◆",_ 
                    "◇你可以在两次语句输入中引用同一个对象而不用担心该对象被◇",_ 
                    "◆销毁。比如，你可以试试分多次输入以下语句：		◆",_ 
                    "◇      Dim text					◇",_ 
                    "◆      Text = ""Hello, VBSShell""			◆",_ 
                    "◇      Wscript.Echo Text				◇",_ 
                    "◆    那如果我要定义函数方法怎么办呢？不要担心，VBSShell◆",_ 
                    "◇可以自动识别函数定义语句，在你输入第一句定义起始句后不◇",_ 
                    "◆会立即执行，而是会等到你的整个函数方法代码完成后才会被◆",_ 
                    "◇执行。 同理，定义class、循环、判断等语句也会被自动识别◇",_ 
                    "◆并等待你最终输入完毕再执行。 				◆",_ 
                    "◇   那么，希望这个程序能够在你的VBS编程中助你一臂之力，◇",_ 
                    "◆祝你使用顺利~						◆",_ 
                    "◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇",_ 
                    "◆※ 注意事项：						◆",_ 
                    "◇   ●有一些变量|函数名称是VBSShell本身使用的，不要重新◇",_ 
                    "◆     定义使用这些名称， 否则VBSShell可能运行不正常甚至◆",_
                    "◇     直接崩溃。如果想要知道哪些名称已经被使用 ，请输入◇",_ 
                    "◆     ShowUsed。					◆",_ 
                    "◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆"), 1
End Function

Function SplitVBSLines(ByVal Code)
    ''' 分割VBScript逻辑行函数 '''
    Dim L, Le, Char, Line_Count, Char_Count
    Dim IsREM, IsStr, Bool
    Dim i
    Dim Final(), ThisLine()
    Code = Code & vbCrLf
    Le = Len(Code): Char_Count = 0: Line_Count = 0
    IsREM = False: IsStr = False
    For L = 1 To Le 
        Char = Mid(Code, L, 1)
        IsStr = IsStr Xor (Char = """" And Not(IsREM)) ' 判断是不是字符串：第二个判断条件决定IsStr要不要Not一下，即：True+True=False,True+False=True,False+True=True,False+False=False
        IsREM = Not(IsStr) And Char = "'" Or IsREM '判断是不是注释
        If Mid(Code, L, 2) = vbCrLf Or (Char = ":" And Not(IsREM Or IsStr)) Then
            IsStr = False: IsREM = False
            If Line_Count = 0 Then
                ReDim Final(Line_Count)
            Else
                ReDim Preserve Final(Line_Count)
            End If
            Final(Line_Count) = Join(ThisLine, ""): Erase ThisLine
            Line_Count = Line_Count + 1
        Else
            If L > 1 Then Bool = Mid(Code, L - 1, 2) <> vbCrLf Else Bool = True
            If Bool Then
                If Char_Count = 0 Then ReDim ThisLine(Char_Count) Else ReDim Preserve ThisLine(Char_Count)
                ThisLine(Char_Count) = Char
                Char_Count = Char_Count + 1
            End If
        End If
    Next
    SplitVBSLines = Final
End Function

Function ClearREM(ByVal Code) 
    ''' 去除所有注释函数 '''
    Dim Line, Char
    Dim IsREM, IsStr
    Dim i
    Code = Split(Code, vbCrLf)
    For Each Line In Code
        If Len(Line) >= 4 Then IsREM = UCase(Left(Line, 4)) = "REM "
        IsREM = UCase(Line) = "REM" Or IsREM
        IsStr = False
        For i = 1 To Len(Line)
            Char = Mid(Line,i,1)
            IsStr = IsStr Xor (Char = """" And Not(IsREM)) ' 判断是不是字符串：第二个判断条件决定IsStr要不要Not一下，即：True+True=False,True+False=True,False+True=True,False+False=False
            IsREM = Not(IsStr) And Char = "'" Or IsREM '判断是不是注释
            If Not IsREM Then ClearREM = ClearREM & Char
        Next
        ClearREM = ClearREM & vbCrLf
    Next
End Function

Function ClearStrings(ByVal Code) 
    ''' 去除所有字符串函数 '''
    Dim Line, Deal(), Count, Final()
    Dim i, j
    Code = Split(Code, vbCrLf)
    ReDim Final(UBound(Code))
    For i = 0 To UBound(Code)
        Line = Split(Code(i), """")
        Count = UBound(Line)
        ReDim Deal(Count \ 2 + 1)
        For j = 0 To UBound(Line) Step 2
            Deal(j / 2) = Line(j)
        Next
        Code(i) = Join(Deal, "")
    Next
    ClearStrings = Join(Code, vbCrLf)
End Function

Function GetUAC(ByVal Host, ByVal Hide)
    Dim HostName, Hidden, Args, i
    If Not Hide Then Hidden = 1
    If Host = 1 Then HostName = "wscript.exe"
    If Host = 2 Then HostName = "cscript.exe"
    If wscript.Arguments.Count > 0 Then
        For i = 0 To wscript.Arguments.Count - 1
            If Not(i = 0 And (wscript.Arguments(i) = "uac" Or wscript.Arguments(i) = "uacHidden")) Then Args = Args & " " & Chr(34) & wscript.Arguments(i) & Chr(34)
        Next
    End If
    If wscript.Arguments.Count = 0 Then
        SA.ShellExecute "wscript.exe", Chr(34) & wscript.ScriptFullName & chr(34) & " uac" & Args, "", "runas", 1
        wscript.Quit
    ElseIf LCase(Right(WScript.FullName,12)) <> "\" & HostName Or wscript.Arguments(0) <> "uacHidden" Then
        ws.Run HostName & " //NoLogo """ & WScript.ScriptFullName & """ uacHidden" & Args,Hidden,False
        WScript.Quit
    End If
End Function

Function FormatPath(ByVal Path)
    If Not Right(Path,1) = "\" Then
        Path = Path & "\"
    End If
    FormatPath = Path
End Function

Function CreateTempPath(ByVal IsFolder)
    Dim TempPath
    TempPath = FSO.GetSpecialFolder(2) & "\" & FSO.GetTempName()
    If IsFolder Then TempPath = FormatPath(TempPath)
    CreateTempPath = TempPath
End Function