Option Explicit
'Call ISOLATE_RUN()

'===================================/System Info & Settings/===================================
Const Version = "1.0.0.6"
Const HookErrors = True

'===================================/ User Info & Settings /===================================
Const DefaultModels = "Linux_Commands"

'===================================/     UpDate__Logs     /===================================
' 尚未完成ShellMode模块的开关功能。因其需要读取自身模块文件代码，故需要在ImportExecute内获知自身模块文件路径，所以需要Import函数配合传参（可以的话）

'===================================/      Code Start      /===================================
Dim FSO, ws, SA, ADO
Dim SelfFolderPath
Set FSO = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("Wscript.Shell")
Set SA = CreateObject("Shell.Application")
Set ADO = CreateObject("ADODB.STREAM")

Call GetUAC(2, False)
'Dim SI, SO: Set SI = WScript.StdIn: Set SO = WScript.StdOut

' 初始化模块信息
Dim ImportInfos
Set ImportInfos = New ImportInfoVariant
ImportInfos.ModelFullPath = WScript.ScriptFullName
ImportInfos.MainScriptFullPath = WScript.ScriptFullName

SelfFolderPath = FormatPath(FSO.GetFile(WScript.ScriptFullName).ParentFolder.Path): ws.CurrentDirectory = SelfFolderPath
Const Tip_Main = "VBScript >>> "
Const Tip_Wait = "------------ "
Call StartOutput()
Call DefaultImports()
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
    'On Error Resume Next
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

Function ExpandCode(ByRef Code)
    ''' 展开压缩在一行的If代码，如If BOOL Then Statement: Statement Else Statement '''
    ' 初始化正则表达式
    Dim re, Resault: Set re = New RegExp
    re.Global = True
    re.IgnoreCase = True
    ' 匹配压缩的If语句
    re.Pattern = " +Then +\w+"
    If re.Execute(Code).Count > 0 Then
        re.Pattern = " +Then +"
        ExpandCode = True 
        Code = re.Replace(Code, " Then" & vbCrLf)
        Code = Code & vbCrLf & "End If"
        re.Pattern = " +Else +"
        Code = re.Replace(Code, vbCrLf & "Else" & vbCrLf)
    Else
        ExpandCode = False
    End If
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
    If ExpandCode(Code) Then
        WaitForEnd = Code
        Exit Function
    End If
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
    If HookErrors Then On Error Resume Next
    Dim Input, InputDealer
    Do
        Input = StdInput(Tip_Main): Input = LTrim(Input)
        If VariantPreview(Input) = 0 Then
            Input = WaitForEnd(Input)
            ExecuteGlobal Input
            If HookErrors Then Call ErrorDealing()
            If Input <> "" Then WScript.Echo ""
        End If
    Loop
End Function

Function VariantPreview(ByVal Input)
    ''' 如果输入是变量名，显示其值 '''
    Input = Trim(Input)
    If InStr(1, Input, " ") + InStr(1, Input, "(") + InStr(1, Input, "=") > 0 Then
        VariantPreview = 0
        Exit Function
    End If
    Select Case VarType(eval(Input))
        Case 2, 8, 11
            StdOutput eval(Input), 2
            VariantPreview = 1
        Case Else
            StdOutput "变量类型: " & CStr(VarType(eval(Input)))
            VariantPreview = 2
    End Select
End Function

Function StartOutput()
    On Error Resume Next
    Dim SOPText, UIL
    UIL = GetUILanguage()
    Select Case UIL
        Case &H409 ' English
            SOPText = Array("VBS Shell Written By PY-DNG(R)",_ 
                            """AN IDLE All About VBScript"", Version " & Version,_ 
                            """Copyright(C) PY-DNG. All Rights Reserved.""",_
                            "Enter ""Help"" to get help. Enter ""Tips"" to get a tip. ")
        Case &H804 ' 中文
            SOPText = Array("VBS Shell|作者：PY-DNG(R)",_ 
                            """属于VBScript的IDLE"", 版本 " & Version,_ 
                            "版权所有(C) PY-DNG。保留一切权利。",_
                            "输入""Help""以获取帮助。输入""Tips""以获取提示。")
        Case Else 'Else: Use English
            SOPText = Array("VBS Shell Written By PY-DNG(R)",_ 
                            """AN IDLE All About VBScript"", Version " & Version,_ 
                            """Copyright(C) PY-DNG. All Rights Reserved.""",_
                            "Enter ""Help"" to get help. Enter ""Tips"" to get a tip. ")
    End Select
    StdOutput SOPText, 2
End Function

Class ImportInfoVariant
    Private Sub Class_Initialize()
        ' Do Nothing
    End Sub
    
    Private Sub Class_Terminate()
		' Do Nothing
	End Sub
	
	Private MFP, MSFP
	
	Property Get ModelFullPath()
	   ''' 属性：模块完整路径 '''
	   ModelFullPath = MFP
    End Property
    
    Property Let ModelFullPath(Path)
        MFP = Path
    End Property
    
    Property Get MainScriptFullPath()
	   ''' 属性：模块完整路径 '''
	   ModelFullPath = MSFP
    End Property
    
    Property Let MainScriptFullPath(Path)
        MSFP = Path
    End Property
End Class

Function Import(ByVal FP)
    ''' 用于引用模块，类似Python的from FP import *，不同的是，本函数只会导入sub、function和class，变量、对象均不会导入 '''
    ' 引用模块式自动执行模块中的ImportExecute函数（如果有的话），所有的ImportExecute函数自身均不会被引入
    'On Error Resume Next
    If Not FSO.FileExists(FP) Then FP = FP & ".vbs"
    If Not FSO.FileExists(FP) Then
        Import = -1
        Exit Function
    End If
    Dim CodeAll
    Dim re
    Dim Funcs, IECode, Code
    CodeAll = FSO.OpenTextFile(FP).ReadAll()
    ' 初始化正则表达式
    Set re = New RegExp
    re.Global = True
    re.IgnoreCase = True
    re.Multiline = True
    ' 匹配函数语句
    re.Pattern = "^(function|sub|class) +.+\(.*\)(.*\s*)*?^end +(function|sub|class)"
    Set Funcs = re.Execute(CodeAll)
    ' 匹配ImportExecute
    re.Pattern = "^(function|sub) +ImportExecute\(.*\)(.*\s*)*?^end +(function|sub)"
    Set IECode = re.Execute(CodeAll)
    If IECode.Count > 0 Then IECode = IECode(IECode.Count-1) Else IECode = ""
    ' 去除IECode的Function/Sub定义
    re.Pattern = "^(function|sub) +ImportExecute\(.*\)"
    IECode = re.Replace(IECode, "")
    re.Pattern = "^End +(function|sub)"
    IECode = re.Replace(IECode, "")
    ' 执行所有导入函数
    For Each Code In Funcs
        If Code <> IECode Then ExecuteGlobal Code
    Next
    ' 传递给IE的环境信息
    ImportInfos.ModelFullPath = FSO.GetFile(FP).Path
    ' 执行ImportExecute
    ExecuteGlobal IECode
    ' 恢复原信息
    ImportInfos.ModelFullPath = WScript.ScriptFullName
    Import = 0
End Function

Function DefaultImports()
    ''' 启动时自动引用这些模块 '''
    Dim Importance
    DefaultImports = Split(DefaultModels, "|")
    For Each Importance In DefaultImports
        SO.Write "自动引用模块: [" & Importance & "]..."
        Select Case Import(Importance)
            Case 0 ' 引用成功
                SO.WriteLine "成功"
            Case -1 ' 文件不存在
                SO.WriteLine "失败：模块文件不存在"
            Case Else '其他错误
                SO.WriteLine "失败：未知错误"
        End Select
    Next
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
    ''' GetUAC By PY-DNG; Version 1.7 '''
    ' 最近更新：更换了UAC判断方式，不再占用命令行参数，兼容了没有UAC机制的更老版本Windows系统（如XP，2003）；简化了代码的表示
    On Error Resume Next: Err.Clear
    Dim HostName, Args, Arg, i, Argv, TFPath, HaveUAC
    If Host = 1 Then HostName = "wscript.exe"
    If Host = 2 Then HostName = "cscript.exe"
    ' Get All Arguments
    Set Argv = WScript.Arguments
    For Each Arg in Argv
        Args = Args & " " & Chr(34) & Arg & Chr(34)
    Next
    ' Test If We Have UAC
    TFPath = FSO.GetSpecialFolder(0) & "\system32\UACTestFile"
    FSO.CreateTextFile TFPath, True
    HaveUAC = FSO.FileExists(TFPath) And Err.number <> 70
    If HaveUAC Then FSO.DeleteFile TFPath, True
    ' If No UAC Then Get It Else Check & Correct The Host
    If Not HaveUAC Then
        SA.ShellExecute "wscript.exe", "//NOLOGO //e:VBScript " & Chr(34) & WScript.ScriptFullName & chr(34) & Args, "", "runas", 1
        WScript.Quit
    ElseIf LCase(Right(WScript.FullName,12)) <> "\" & HostName Then
        ws.Run HostName & " //NOLOGO //e:VBScript """ & WScript.ScriptFullName & """" & Args, Int(Hide)+1, False
        WScript.Quit
    End If
    If Host = 2 Then ExecuteGlobal "Dim SI, SO: Set SI = WScript.StdIn: Set SO = WScript.StdOut"
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



