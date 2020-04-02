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
    ''' �������뻷�����ڸ��뻷����ִ�б���������Ӱ�쵽�û������ִ�У�ע����������δʵ�֣���Ҫ���ã��� '''
    Execute Replace(CreateObject("Scripting.FileSystemObject").OpenTextFile(WScript.ScriptFullName).ReadAll(), vbCrLf & "Call ISOLATE_RUN()", "")
    WScript.Quit
End Function

Function StdInput(ByVal Text)
    ''' ���һ������Ȼ������û����룬����Python��input '''
    On Error Resume Next
    SO.Write Text
    StdInput = SI.ReadLine()
End Function

Function StdOutput(ByVal Content, ByVal WithCrLfs)
    ''' ��������ı���Content�ȿ������ı�������Ҳ�������ı���������������vbCrLf�����з���ϣ��������Ա����� '''
    ''' WithCrLfsָ��Content�����Ϻ�����������з���� '''
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
    ''' �����ж��Ƿ�Ϊһ���´����Ŀ�ʼ������ǣ��ͷ��ؽ�����־�����򷵻ؿ��ַ��� '''
    ' ��ʼ��[��ʼ��־-������־]�ֵ�
    Dim Dict
    Set Dict = CreateObject("Scripting.Dictionary")
    Dict.Add "if", "end if": Dict.Add "do", "loop": Dict.Add "for", "next": Dict.Add "while", "wend": Dict.Add "select case", "end select"
    Dict.Add "function", "end function": Dict.Add "sub", "end sub": Dict.Add "class", "end class"
    ' �жϿ�ʼ��־�����ؽ��
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
    ''' �������飬ֱ����ǰ������������������ȫ������ '''
    On Error Resume Next
    ' �Ƿ���Ҫ���������飿
    Dim NowWaiting
    NowWaiting = GetIfWaiting(Code)
    If NowWaiting = "" Then 
        WaitForEnd = Code
        Exit Function
    End If
    ' ��Ҫ�ȴ��������
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
    ''' ���������Ϣ���򵥵�˵�������׼��������������źʹ����ı���Ϣ '''
    If Err.number <> 0 Then Stdoutput Array("����: " & CStr(Err.number), Err.Description), 1: Err.Clear
End Function

Function Main()
    ''' ���������� '''
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
        Case &H804 '����
            SOPText = Array("VBS Shell|���ߣ�PY-DNG(R)",_ 
                    """����VBScript��IDLE""",_ 
                    "��Ȩ����(C) PY-DNG������һ��Ȩ����",_
                    "����""Help""�Ի�ȡ����������""Tips""�Ի�ȡ��ʾ��")
    End Select
    StdOutput SOPText, 2
End Function

Function Import(ByVal FP)
    ''' ��������ģ�飬����Python��from FP import *����ͬ���ǣ�������ֻ�ᵼ��sub��function����������������ᵼ�� '''
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
    ''' Ϊ�û��ṩ��ʾ '''
    Dim Tips_
    Tips_ = Array("����Import��ʹ��""Import xxx.vbs""��䵼�����Լ��ĺ�����",_ 
                  "Ҫ��Ҫ�����ҵ�Դ���룿",_ 
                  "��Щ�������ͺ�������""VBS Shell""�����Ѿ�ʹ���˵�(����ShowUsed��ʾ��Щ��Щ����)���뾡����Ҫ���¶�����Щ����Ŷ~",_ 
                  "�����ã���������һЩVBScript��䣡",_ 
                  "��ʾ��ֹһ��Ŷ~ ÿ������""Tips""����������һ��Ŷ��",_ 
                  "����Help�Ի�ȡ������")
    Randomize
    StdOutput Tips_(Int(Rnd * UBound(Tips_))), 1
End Function

Function ShowUsed()
    ''' ��ȡ���б�����ʹ�õģ�ȫ�֣���������������������class�����ĺ��� '''
    ' ʹ�ñ�������ҪSplitVBSLines��֧��
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
    StdOutput "���������ȫ�ֱ�����(" & CStr(Names_Count(0) + 1) & "��)��", 1
    StdOutput Names(0), 2
    StdOutput "��������ĺ���������(" & CStr(Names_Count(1) + 1) & "��)��", 1
    StdOutput Names(1), 2
    StdOutput "���������class��(" & CStr(Names_Count(2) + 1) & "��)��", 1
    StdOutput Names(2), 2
    ShowUsed = Names
End Function


Function Help()
    ''' �û����� '''
    StdOutput Array("��������������������������������������������",_
                    "��� Help| �û�����					��",_ 
                    "��������������������������������������������",_ 
                    "��   VBSShell��һ����VBS��д��VBSyuyan�����IDLE�������ڡ�",_ 
                    "��Python�Դ��ĵ�IDLE��ּ������VBS��̸��Ӽ򵥡� ���������",_ 
                    "�����ֱ�۵�ִ��VBScript��䣬���������ĳ���ʵʱ�鿴��",_ 
                    "�������						��",_ 
                    "��   ���泶����ô�࣬����Ҫװ��13�� ʵ��������û����ôţ��",_ 
                    "���ȡ��ϻ�����˵������˵�÷���				��",_ 
                    "��   һ���������������ִ�е�VBScript��䣬���»س��� ���",_ 
                    "��ʱ����ղ���������ͻᱻ����ִ�У���������Ӧ��Ч������",_ 
                    "��Ȼ�������������һ����䡣VBSShell���ס��ÿ�ζ���ġ�",_ 
                    "�����������󡢺������ֱ࣬��VBSShell�˳�Ϊֹ��Ҳ����˵����",_ 
                    "��������������������������ͬһ����������õ��ĸö��󱻡�",_ 
                    "�����١����磬��������Էֶ������������䣺		��",_ 
                    "��      Dim text					��",_ 
                    "��      Text = ""Hello, VBSShell""			��",_ 
                    "��      Wscript.Echo Text				��",_ 
                    "��    �������Ҫ���庯��������ô���أ���Ҫ���ģ�VBSShell��",_ 
                    "������Զ�ʶ����������䣬���������һ�䶨����ʼ��󲻡�",_ 
                    "��������ִ�У����ǻ�ȵ����������������������ɺ�Żᱻ��",_ 
                    "��ִ�С� ͬ������class��ѭ�����жϵ����Ҳ�ᱻ�Զ�ʶ���",_ 
                    "�����ȴ����������������ִ�С� 				��",_ 
                    "��   ��ô��ϣ����������ܹ������VBS���������һ��֮������",_ 
                    "��ף��ʹ��˳��~						��",_ 
                    "��������������������������������������������",_ 
                    "���� ע�����						��",_ 
                    "��   ����һЩ����|����������VBSShell����ʹ�õģ���Ҫ���¡�",_ 
                    "��     ����ʹ����Щ���ƣ� ����VBSShell�������в�����������",_
                    "��     ֱ�ӱ����������Ҫ֪����Щ�����Ѿ���ʹ�� ���������",_ 
                    "��     ShowUsed��					��",_ 
                    "��������������������������������������������"), 1
End Function

Function SplitVBSLines(ByVal Code)
    ''' �ָ�VBScript�߼��к��� '''
    Dim L, Le, Char, Line_Count, Char_Count
    Dim IsREM, IsStr, Bool
    Dim i
    Dim Final(), ThisLine()
    Code = Code & vbCrLf
    Le = Len(Code): Char_Count = 0: Line_Count = 0
    IsREM = False: IsStr = False
    For L = 1 To Le 
        Char = Mid(Code, L, 1)
        IsStr = IsStr Xor (Char = """" And Not(IsREM)) ' �ж��ǲ����ַ������ڶ����ж���������IsStrҪ��ҪNotһ�£�����True+True=False,True+False=True,False+True=True,False+False=False
        IsREM = Not(IsStr) And Char = "'" Or IsREM '�ж��ǲ���ע��
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
    ''' ȥ������ע�ͺ��� '''
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
            IsStr = IsStr Xor (Char = """" And Not(IsREM)) ' �ж��ǲ����ַ������ڶ����ж���������IsStrҪ��ҪNotһ�£�����True+True=False,True+False=True,False+True=True,False+False=False
            IsREM = Not(IsStr) And Char = "'" Or IsREM '�ж��ǲ���ע��
            If Not IsREM Then ClearREM = ClearREM & Char
        Next
        ClearREM = ClearREM & vbCrLf
    Next
End Function

Function ClearStrings(ByVal Code) 
    ''' ȥ�������ַ������� '''
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