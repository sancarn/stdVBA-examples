VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Console 
   Caption         =   "Console"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765.001
   OleObjectBlob   =   "Console.frx":0000
   ShowModal       =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' Private Variables
    
    Private       CurrentLineIndex      As Long
    Private Const Recognizer            As String = "\>>>"
    Private Const ArgSeperator          As String = ", " 
    Private Const AsgOperator           As String = " = "
    Private Const LineSeperator         As String = "LINEBREAK/()\"
    Private       PasteStarter          As Boolean
    Private       UserInput             As Variant
    Private       LastError             As Variant

    Private       Intellisense_Index    As Long
    Private       ConsVarIndex          As Long

    Private       PreviousCommands()    As Variant
    Private       PreviousCommandsIndex As Long

    Private       p_Password            As String
    Private       PasswordActive        As Boolean
    Private       PasswordMode          As Boolean

    Private       DimVariables(255)     As Long
    Private       DimIndex              As Long

    Private WorkMode As Long
    Private Enum WorkModeEnum
        Logging = 0
        UserInputt = 1
        MultilineMode = 2
        ScriptMode = 3
    End Enum

    Private Const Intellisense_Active As Boolean = True
    Private Const Extensebility_Active As Boolean = True
    Private Const stdLambda_Active    As Boolean = True

    Private in_Basic       As Long
    Private in_System      As Long
    Private in_Procedure   As Long
    Private in_Operator    As Long
    Private in_Datatype    As Long
    Private in_Value       As Long
    Private in_String      As Long
    Private in_Statement   As Long
    Private in_Keyword     As Long
    Private in_Parantheses As Long
    Private in_Variable    As Long
    Private in_Script      As Long
    Private in_Lambda      As Long
'


' Tree
    Private Type Node
        Value As Variant
        Branches() As Long
    End Type

    Private Type cCollection
        Nodes() As Node
    End Type

    Private Trees() As cCollection

    Private Enum NodeType
        e_ReturnType = 0
        e_Arguments = 1
        e_Value = 2
        e_Branches = 3
    End Enum

    Private Function GetNode(TreeIndex As Long, Positions() As Long) As Long
        Dim i As Long
        Dim Index As Long
        Index = Positions(0)
        For i = 1 To UboundK(Positions)
            Index = Trees(TreeIndex).Nodes(Index).Branches(Positions(i))
        Next i
        GetNode = Index
    End Function

    Private Function FindNode(TreeIndex As Long, Positions() As Long, Value As Variant, Optional StartBranch As Long = 0, Optional EndBranch As Long = -1) As Long
        
        Dim i As Long
        Dim Index As Long
        Dim CurrentNode As Long
        Index = GetNode(TreeIndex, Positions)
        If Index = -1 Then FindNode = Index: Exit Function
        
        CurrentNode = Index
        FindNode = -1
        If EndBranch = -1 Then EndBranch = UboundK(Trees(TreeIndex).Nodes(CurrentNode).Branches)
        If Value = Empty Then Exit Function
        For i = StartBranch To EndBranch
            Index = Trees(TreeIndex).Nodes(CurrentNode).Branches(i)
            If IsObject(Value) Then
                If Trees(TreeIndex).Nodes(Index).Value Is Value Then
                    FindNode = Index
                    Exit Function
                End If
            Else
                If Trees(TreeIndex).Nodes(Index).Value = Value Then
                    FindNode = Index
                    Exit Function
                End If
            End If
        Next i
    End Function

    Private Sub FindDepth(ByVal TreeIndex As Long, ByVal CurrentNode As Long, ByVal CurrentDepth As Long, ByRef MaxDepth As Long)
        Dim i As Long
        Dim NewNode As Long
        If UboundK(Trees(TreeIndex).Nodes(CurrentNode).Branches) >= NodeType.e_ReturnType Then CurrentDepth = CurrentDepth + 1
        For i = 0 To UboundK(Trees(TreeIndex).Nodes(CurrentNode).Branches)
            NewNode = Trees(TreeIndex).Nodes(CurrentNode).Branches(i)
            Call FindDepth(TreeIndex, NewNode, CurrentDepth, MaxDepth)
        Next i
        If CurrentDepth > MaxDepth Then
            MaxDepth = CurrentDepth
        Else
            CurrentDepth = MaxDepth
        End If
    End Sub

    Private Function AddNode(TreeIndex As Long, NodeIndex As Long, Value As Variant) As Long

        Dim n_Node As Node
        Dim NewSize As Long

        If IsObject(Value) Then
            Set n_Node.Value = Value 
        Else
            n_Node.Value = Value
        End If

        NewSize = UboundN(Trees(TreeIndex)) + 1
        ReDim Preserve Trees(TreeIndex).Nodes(NewSize)
        Trees(TreeIndex).Nodes(NewSize) = n_Node

        If NodeIndex <> -1 Then
            NewSize = UboundK(Trees(TreeIndex).Nodes(NodeIndex).Branches) + 1
            ReDim Preserve Trees(TreeIndex).Nodes(NodeIndex).Branches(NewSize)
            Trees(TreeIndex).Nodes(NodeIndex).Branches(NewSize) = UboundN(Trees(TreeIndex))
        End If

        AddNode = UboundN(Trees(TreeIndex))
    End Function

    Private Function DeleteNode(TreeIndex As Long, NodeIndex As Long) As Long
        Dim Temp As cCollection
        Dim Size As Long
        Dim Difference As Long
        Dim NodesToDelete() As Long
        Dim i As Long
        Dim j As Long
        Temp = Trees(TreeIndex)

        ReDim NodesToDelete(0)
        NodesToDelete(0) = NodeIndex
        For i = 0 To UboundK(Trees(TreeIndex).Nodes(NodeIndex).Branches)
            ReDim Preserve NodesToDelete(i + 1)
            NodesToDelete(i + 1) = Trees(TreeIndex).Nodes(NodeIndex).Branches(i)
        Next i


        Difference = UboundK(NodesToDelete)
        Size = UboundN(Trees(TreeIndex)) - (Difference + 1)
        ReDim Trees(TreeIndex).Nodes(Size)
        For i = 0 To NodesToDelete(0) - 1
            Trees(TreeIndex).Nodes(i) = Temp.Nodes(i)
        Next i
        For i = NodesToDelete(Difference) + 1 To UboundN(Temp)
            Trees(TreeIndex).Nodes(i - (Difference + 1)) = Temp.Nodes(i)
        Next i
        
        Size = UboundK(Temp.Nodes(0).Branches)
        ReDim Trees(TreeIndex).Nodes(0).Branches(Size - 1)
        For i = 0 To Size
            If Temp.Nodes(0).Branches(i) <> NodeIndex Then
                Trees(TreeIndex).Nodes(0).Branches(j) = Temp.Nodes(0).Branches(i)
                j = j + 1
            End If
        Next
        DeleteNode = NodeIndex
    End Function

    Private Function SetVariable(TreeIndex As Long, SearchIndex As Long, Name As String, ReturnType As String, Arguments As String, Value As Variant) As Long
        
        Dim Temp(0) As Long
        Dim Index As Long
        Temp(0) = SearchIndex

        Index = GetNode(TreeIndex, Temp)
        If Index <> -1 Then
            If CStr(Trees(TreeIndex).Nodes(Index).Value) = Name Then
                Repeat:
                SetVariable = Index
                Trees(TreeIndex).Nodes(Index).Value = Name
                If ReturnType <> "NOCHANGE" Then
                    Index = Trees(TreeIndex).Nodes(SetVariable).Branches(NodeType.e_ReturnType)
                    Trees(TreeIndex).Nodes(Index).Value = ReturnType
                End If
                If Arguments  <> "NOCHANGE" Then
                    Index = Trees(TreeIndex).Nodes(SetVariable).Branches(NodeType.e_Arguments)
                    Trees(TreeIndex).Nodes(Index).Value = Arguments
                End If
                If IsObject(Value) Then
                    Index = Trees(TreeIndex).Nodes(SetVariable).Branches(NodeType.e_Value)
                    Set Trees(TreeIndex).Nodes(Index).Value = Value
                Else
                    If Value <> "NOCHANGE" Then
                        Index = Trees(TreeIndex).Nodes(SetVariable).Branches(NodeType.e_Value)
                        Trees(TreeIndex).Nodes(Index).Value = Value
                    End If
                End If
                Exit Function
            End If
        End If
        Index = FindNode(TreeIndex, Temp, Name)
        If Index = -1 Then
            Index = AddNode(TreeIndex, SearchIndex, Name)
            If ReturnType <> "NOCHANGE" Then Call AddNode(TreeIndex, Index, ReturnType)
            If Arguments  <> "NOCHANGE" Then Call AddNode(TreeIndex, Index, Arguments)
            If IsObject(Value) Then
                Call AddNode(TreeIndex, Index, Value)
            Else
                If Arguments <> "NOCHANGE" Then Call AddNode(TreeIndex, Index, Value)
            End If
            SetVariable = Index
        Else
            GoTo Repeat
        End If
    End Function

    Private Function ReturnVariable(TreeIndex As Long, Positions() As Long, Optional ReturnValue As Boolean = False, Optional Arguments As Variant) As Variant
        
        Dim ReturnType As String
        Dim Value As Variant
        Dim Name As Variant
        Dim ScriptArguments As Variant

        ReturnType = ReturnVariableValue(TreeIndex, Positions, 0)
        ScriptArguments = ReturnVariableValue(TreeIndex, Positions, 1)
        If IsObject(ReturnVariableValue(TreeIndex, Positions, 2)) Then
            Set Value = ReturnVariableValue(TreeIndex, Positions, 2)
        Else
            Value = ReturnVariableValue(TreeIndex, Positions, 2)
        End If
        Name = ReturnVariableValue(TreeIndex, Positions, -1)
        If Returnvalue Then
            Select Case True
                Case ReturnType Like "* ConsoleVariable As stdLambda"      : ReturnVariable = RunLambda(Value, Arguments)
                Case ReturnType Like "* ConsoleVariable As ConsoleScript"  : ReturnVariable = RunScript(Value, ScriptArguments, Arguments)
                Case ReturnType Like "* Procedure *"                       : ReturnVariable = RunApplication(CStr(Name), Arguments)
                Case ReturnType Like "* ConsoleVariable As *"              : ReturnVariable = GetVariableByType(Value, ReturnType)
                Case Else                                                  : ReturnVariable = "Could not return Value"
            End Select
        Else
            ReturnVariable = Value
        End If
        
    End Function

    Private Sub InitializeTree()
        Dim i As Long
        Dim WB As Workbook

        ReDim Trees((Workbooks.Count - 1) + 1)
        ConsVarIndex = Ubound(Trees)
        i = 0
        For Each WB In Workbooks
            Call SetVariable(i, -1, WB.VBProject.Name, "As VBProject", "No Arguments", "No Value")
            i = i + 1
        Next
        Call SetVariable(i, -1, "ConsoleVariables", "Private ConsoleVariable As ", "No Arguments", "No Value")

    End Sub

    Private Function ReturnVariableValue(TreeIndex As Long, Positions() As Long, ReturnWhat As Long) As Variant
        Dim Index As Long
        Index = GetNode(TreeIndex, Positions)
        If ReturnWhat <> -1 Then Index = Trees(TreeIndex).Nodes(Index).Branches(ReturnWhat)
        If IsObject(Trees(TreeIndex).Nodes(Index).Value) Then
            Set ReturnVariableValue = Trees(TreeIndex).Nodes(Index).Value
        Else
            ReturnVariableValue = Trees(TreeIndex).Nodes(Index).Value
        End If
    End Function

    Private Function RetVarSin(TreeIndex As Long, Position As Long, ReturnWhat As Long) As Variant
        Dim Pos(0) As Long
        Pos(0) = Position
        RetVarSin = ReturnVariableValue(TreeIndex, Pos, ReturnWhat)
    End Function

    Private Function FindNodeAll(TreeIndex As Long, Value As Variant, ByRef Positions() As Long) As Long
        
        Dim i As Long
        Dim NewIndex(0) As Long

        FindNodeAll = -1
        FindNodeAll = FindNode(TreeIndex, Positions, Value, 3)
        If FindNodeAll <> -1 Then Exit Function

        For i = 3 To UboundK(Trees(TreeIndex).Nodes(Positions(0)).Branches)
            NewIndex(0) = Trees(TreeIndex).Nodes(Positions(0)).Branches(i)
            FindNodeAll = FindNodeAll(TreeIndex, Value, NewIndex)
        Next

    End Function

    
'

' Public Console Functions

    Public Function Execute(Command As String) As Variant
        Execute = HandleCode(Command)
    End Function

    Public Function GetUserInput(Message As Variant, Optional InputType As Long = 12) As Variant

        If HandlePassword = False Then Set GetUserInput = Nothing: Exit Function
        Call PrintConsole(Message, in_System)
        WorkMode = WorkModeEnum.UserInputt
        PasteStarter = False
        Do While WorkMode = WorkModeEnum.UserInputt
            DoEvents
            If UserInput <> "" Then
                UserInput = Replace(UserInput, Message, "")
                If VarType(InterpretVariable(UserInput)) = InputType Then
                    GetUserInput = UserInput
                    WorkMode = WorkModeEnum.Logging
                Else
                    Call PrintEnter("Wrong Datatype", in_System)
                End If
                UserInput = ""
            End If
        Loop
        PasteStarter = True

    End Function

    Public Function CheckPredeclaredAnswer(Message As Variant, AllowedValues As Variant, Optional Answers As Variant = Empty) As Variant

        Dim i As Long
        Dim Found As Boolean
        Dim Index As Long

        If HandlePassword = False Then Set CheckPredeclaredAnswer = Nothing: Exit Function
        Message = Message & "("
        For i = 0 To UBoundK(AllowedValues)
            Message = Message & AllowedValues(i) & "|"
        Next i
        Message = Message & ") "
        Call PrintConsole(Message)

        WorkMode = WorkModeEnum.UserInputt
        PasteStarter = False
        Do While WorkMode = WorkModeEnum.UserInputt
            Index = 0
            DoEvents
            If UserInput <> "" Then
                UserInput = Replace(UserInput, Message, "")
                For i = 0 To UBoundK(AllowedValues)
                    If AllowedValues(i) = UserInput Then
                        CheckPredeclaredAnswer = i
                        Found = True
                        WorkMode = WorkModeEnum.Logging
                        Exit For
                    End If
                    Index = Index + 1
                Next i
                If Found <> True Then
                    Call PrintEnter("Value not Valid", in_System)
                    Call PrintConsole(Message)
                End If
                UserInput = ""
            End If
        Loop
        PasteStarter = True
        Call PrintEnter(Answers(Index))
        Call PrintConsole(PrintStarter)

    End Function

    Public Sub PrintEnter(Text As Variant, Optional Colors As Variant, Optional ColorLength As Variant)
        Call PrintConsole(Text & vbCrLf, Colors, ColorLength)
    End Sub

    Public Sub PrintConsole(Text As Variant, Optional Colors As Variant, Optional ColorLength As Variant)
        
        Dim i As Long
        Dim StartPoint As Long
        Dim RealColors() As Long
        Dim Offset As Long

        If IsMissing(ColorLength) Then ReDim ColorLength(0): ColorLength(0) = Len(Text)
        If IsMissing(Colors)      Then ReDim Colors(0): Colors(0) = in_Basic
        If UboundK(Colors) = -1   Then
            ReDim RealColors(0)
            RealColors(0) = Colors
        Else
            ReDim RealColors(UboundK(Colors))
            For i = 0 To UboundK(Colors)
                RealColors(i) = Colors(i)
            Next i
        End If

        If UboundK(RealColors) <> UboundK(ColorLength) Then
            LastError = 4
            Call PrintEnter(HandleLastError, in_System)
            Exit Sub
        End If


        StartPoint = Len(ConsoleText.Text)
        Offset = 1
        For i = 0 To UboundK(ColorLength)
            ConsoleText.SelStart = StartPoint
            ConsoleText.SelLength = 0
            ConsoleText.SelColor = RealColors(i)
            ConsoleText.SelText = Mid(Text, Offset, ColorLength(i))
            Offset = Offset + ColorLength(i)
            StartPoint = StartPoint + ColorLength(i)
        Next
        Call SetUpNewLine
        ConsoleText.SelStart = StartPoint

    End Sub

    Public Sub AddScript(Name As String, Arguments As String, Script As String)
        If Arguments = "" Then Arguments = "No Arguments"
        Call SetVariable(ConsVarIndex, 0, Name, "Public ConsoleVariable As ConsoleScript", Arguments, Script)
    End Sub

    Public Function Password(Old_Password As String, New_PassWord As String) As Boolean
        If Old_PassWord = p_Password Then
            p_Password = New_PassWord
            PasswordActive = True
        End If
    End Function

    Public Function GetPublicVariable(Name As String) As Variant
        Dim Pos() As Long
        Dim Temp(0) As Variant
        Pos = GetFuncTreePosition(Name)
        If UboundK(Pos) <> -1 Then
            Temp(0) = Pos(1)
            If UCase(ReturnVariableValue(Pos(0), Temp, 0)) Like "PUBLIC*" Then
                GetPublicVariable = ReturnVariableValue(Pos(0), Temp, 2)
            Else
                Set GetPublicVariable = Nothing
            End If
        Else
            Set GetPublicVariable = Nothing
        End If
    End Function
'

' Initialization

    Private Sub UserForm_Initialize()
        Call AssignColor
        PasteStarter = True

        ConsoleText.Text = GetStartText
        ConsoleText.SelStart = 0
        ConsoleText.SelLength = Len(ConsoleText.Text)
        ConsoleText.SelColor = in_System
        ConsoleText.SelStart = Len(ConsoleText.Text)

        CurrentLineIndex = UBoundK(Split(ConsoleText.Text, vbCrLf))
        ScrollHeight = 5000
        ScrollWidth = 3000
        p_Password = ""
        PasswordMode = True


        InitializeTree
        ' Dependenant on Microsoft Visual Studio Extensebility 5.3
        If Extensebility_Active = True Then GetAllProcedures
    End Sub

    Private Sub UserForm_Terminate()
    End Sub
'

' Get/Set Values

    Private Function GetMaxSelStart() As Long
        Dim Temp(1) As Long
        Temp(0) = ConsoleText.SelStart
        Temp(1) = ConsoleText.SelLength
        ConsoleText.SelStart = Len(ConsoleText.Text)
        GetMaxSelStart = ConsoleText.SelStart
        ConsoleText.SelStart = Temp(0)
        ConsoleText.SelLength = Temp(1)
    End Function

    Private Function PrintStarter() As Variant
        PrintStarter = ThisWorkbook.Path & Recognizer
    End Function

    Private Function GetStartText() As String
        GetStartText =                   _
        "VBA Console [Version 1.0]" & vbCrLf & _
        "No Rights reserved"        & vbCrLf & _
        vbCrLf                               & _
        "Enter Password"            & vbCrLf
    End Function

    Private Function GetTextLength(Text As String, Seperator As String, Optional IndexBreakPoint As Long = -2) As Long
        Dim i As Long
        Dim Lines() As String
        Lines = Split(Text, Seperator)
        If IndexBreakPoint = -2 Then IndexBreakPoint = UboundK(Lines)
        For i = 0 To IndexBreakPoint
            GetTextLength = GetTextLength + Len(Lines(i)) + 1
        Next i
    End Function

    Private Function GetLine(Text As String, Index As Long) As String
        Dim Lines() As String
        Dim SearchString As String
        Dim ReplaceString As String
        Dim i As Variant
        Lines = Split(Text, vbCrLf)
        If Index > 0 And Index <= UBoundK(Lines) + 1 Then
            SearchString = Lines(Index)
            If InStr(1, SearchString, Recognizer) = 0 Then
                ReplaceString = ""
            Else
                ReplaceString = Mid(SearchString, 1, InStr(1, SearchString, Recognizer) - 1 + Len(Recognizer))
            End If
            GetLine = Replace(SearchString, ReplaceString, "")
        Else
            GetLine = "Line number out of range"
        End If
    End Function

    Private Function GetWord(Text As String, Optional Index As Long = -1) As String
        Dim Words() As String
        Words = Split(Text, " ")
        If Index = -1 Then Index = UboundK(Words)
        If UboundK(Words) > -1 Then GetWord = Words(Index)
    End Function

    Private Function SplitString(Text As String, SplitText As String) As String()
        Dim Temp() As String
        Dim ReturnArray() As String
        Dim i As Long
        
        If Text = "" Then
        ElseIf InStr(1, Text, SplitText) <> 0 Then
            Temp = Split(Text, SplitText)
            ReDim ReturnArray(UboundK(Temp))
            For i = 0 To UboundK(ReturnArray)
                ReturnArray(i) = Temp(i)
            Next i
        Else
        End If
        SplitString = ReturnArray
    End Function

    Private Sub SetUpNewLine()
        CurrentLineIndex = UboundK(Split(ConsoleText.Text, vbCrLf))
    End Sub

    Private Function InStrAll(Text As String, SearchText As String, Optional StartIndex As Long = 1, Optional EndIndex As Long = 0, Optional StartFinding As Long = 0, Optional ReturnCount As Long = 255, Optional Line As Long = 0, Optional BreakText As String = Empty) As Long()
        
        Dim ReturnArray() As Long: ReDim ReturnArray(0)
        Dim Lines() As String
        Dim EndLine As Long
        Dim CurrentValue As Long : CurrentValue = 0
        Dim Found As Long        : Found = 0
        Dim Saved As Long        : Saved = -1
        Dim j As Long            : j = 0
        Dim i As Long            : i = StartIndex
        
        If EndIndex = 0 Then EndIndex = Len(Text)
        If BreakText <> Empty Then
            Lines = Split(Text, BreakText)
            EndLine = Line
        Else
            ReDim Lines(0)
            Lines(0) = Text
            EndLine = UboundK(Lines)
        End If
        
        For j = Line To EndLine
            Do Until i > EndIndex
                CurrentValue = 0
                CurrentValue = InStr(i, Lines(j), SearchText)
                If CurrentValue <> 0 Then 
                    Found = Found + 1
                Else
                    Exit Do
                End If
                If Found >= StartFinding Then
                    Saved = Saved + 1
                    ReDim Preserve ReturnArray(Saved)
                    ReturnArray(Saved) = CurrentValue
                End If
                If Saved = ReturnCount Then Exit For
                
                i = CurrentValue + Len(SearchText)
                If i > Len(Lines(j)) Then Exit Do
            Loop
            i = 1
        Next j
        InStrAll = ReturnArray

    End Function

    Private Function GetFunctionArgs(Line As String) As Variant()
        Dim Temp() As String
        Dim Tempp() As Variant
        Dim i As Long
        Temp = SplitString(GetParanthesesText(Line), ", ")
        If UboundK(Temp) <> -1 Then
            ReDim Tempp(UboundK(Temp))
            For i = 0 To UBoundK(Temp)
                Tempp(i) = Temp(i)
            Next
        End If
        GetFunctionArgs = Tempp
    End Function

    Private Function GetFuncTreePosition(Line As String) As Long()

        Dim CurrentWord As String
        Dim i As Long
        Dim j As Long
        Dim ReturnArray() As Long

        CurrentWord = GetFunctionName(Line)
        For i = 0 To ConsVarIndex
            Dim Temp() As Long
            ReDim Temp(0)
            If i = ConsVarIndex Then
                j = FindNode(i, Temp, CurrentWord, 3)
            Else
                j = FindNodeAll(i, CurrentWord, Temp)
            End If
            If j <> -1 Then
                Redim ReturnArray(1)
                ReturnArray(0) = i
                ReturnArray(1) = j
                GetFuncTreePosition = ReturnArray
                Exit Function
            End If
        Next i

    End Function

    Private Function GetFunctionName(Line As String) As Variant
        Dim i As Long
        i = InStr(1, Line, "(")
        If i = 0 Then
            GetFunctionName = Line
        Else
            GetFunctionName = MidP(Line, 1, i -1)
        End If
    End Function

    Private Function MidP(Text As String, StartPoint As Long, EndPoint As Long) As String
        MidP = Mid(Text, StartPoint, (EndPoint - StartPoint) + 1)
    End Function

    Private Function GetParanthesesText(Line As String) As String

        Dim OpenPos() As Long: OpenPos = InStrAll(Line, "(") 
        Dim ClosePos() As Long: ClosePos = InStrAll(Line, ")")
        Dim StartPoint As Long
        Dim EndPoint As Long

        If UboundK(OpenPos) = UboundK(ClosePos) Then
            StartPoint = OpenPos(0) + 1
            EndPoint = ClosePos(UboundK(ClosePos)) - 1
        End If
        If UboundK(OpenPos) = 0 And OpenPos(0) = 0 Then StartPoint = Len(Line) + 1
        If UboundK(ClosePos) = 0 And ClosePos(0) = 0 Then EndPoint = Len(Line)
        GetParanthesesText = MidP(Line, StartPoint, EndPoint)

    End Function

    Private Function InString(Text As String, StartPoint As Long, EndPoint As Long) As Boolean
        Dim Quotes() As Long
        Dim i As Long
        Dim EndIndex As Long
        Quotes = InStrAll(Text, Chr(34))
        Select Case UboundK(Quotes)
            Case 0
                If Quotes(0) = 0 Then
                    InString = False
                ElseIf Quotes(0) =< StartPoint Then
                    InString = True
                End If
            Case >0
                If (UboundK(Quotes) + 1) Mod 2 = 0 Then
                    EndIndex = UboundK(Quotes)
                Else
                    EndIndex = UboundK(Quotes) - 1
                    If Quotes(UboundK(Quotes)) =< StartPoint Then InString = True: Exit Function
                End If
                For i = 0 To EndIndex Step+2
                    If Quotes(i) =< StartPoint And Quotes(i + 1) >= EndPoint Then InString = True
                Next i
            Case Else
        End Select
    End Function

    Private Function GetAllOperators(Variable() As Variant) As Variant()

        Dim Operators() As Variant
        Dim FoundOperators() As Variant
        Dim Temp() As Variant
        Dim TempStr() As String
        Dim i As Long
        Dim j As Long
        Dim k As Long
        Operators = Array(" IS ", "==", "<>", "=>", "=<", "<=", ">=", "<", ">", " NOT ", " AND ", " OR ", " XOR ", "!=", "||", "&&", "//", "**", "++", "--", "^^",  "+", "-", "*", "/", "^", "?", ":", ";", "!", "|", "?", ",", "&")
        For i = 0 To UboundK(Operators)
            For j = 0 To UboundK(Variable)
                If InStr(1, Variable(j), Operators(i)) > 1 Then
                    For k = 0 To UBoundK(FoundOperators)
                        If InStr(1, FoundOperators(k), Operators(i)) Then GoTo Skip
                    Next
                    Call PushArray(FoundOperators, Operators(i))
                    TempStr = Split(CStr(Variable(j)), CStr(Operators(i)))
                    ReDim Temp(UboundK(TempStr))
                    For k = 0 To UboundK(TempStr)
                        Temp(k) = Replace(TempStr(k), " ", "")
                    Next
                    Call InsertElements(Temp, Operators(i))
                    Call ReplaceArrayPoint(Variable, Temp, j)
                End If
            Next j
            Skip:
        Next i
        GetAllOperators = Variable

    End Function

'

' Array
    Private Sub MergeArray(ByRef Goal() As Variant, Adder() As Variant, Position As Long)

        Dim Temp() As Variant
        Dim i As Long
        Dim j As Long

        Temp = Goal
        If UboundK(Temp) = -1 Then
            Redim Goal(UboundK(Adder))
        Else
            Redim Goal((UboundK(Temp) + 1) + (UboundK(Adder) + 1) - 1)
        End If
        For i = 0 To Position - 1
            Goal(i) = Temp(i)
        Next
        For j = 0 To UboundK(Adder)
            Goal(i + j) = Adder(j)
        Next
        j = i + j
        For i = j To UboundK(Goal)
            Goal(j + i) = Temp(i)
        Next

    End Sub

    Private Sub ReplaceArrayPoint(ByRef Goal() As Variant, Adder() As Variant, Position As Long)
        Dim Temp() As Variant
        Dim i As Long
        Temp = Goal
        If (UboundK(Goal) - 1) <> -1 Then
            ReDim Goal(UboundK(Goal) - 1)
            For i = 0 To Position - 1
                Goal(i) = Temp(i)
            Next i
            For i = Position To UBoundK(Goal)
                Goal(i) = Temp(i + 1)
            Next
            Call MergeArray(Goal, Adder, Position)
        Else
            Dim TempArray() As Variant
            Call MergeArray(TempArray, Adder, Position)
            Goal = TempArray
        End If
    End Sub

    Private Sub StitchArray(ByRef Arr() As Variant, StartPosition As Long, EndPosition As Long)
        
        Dim Temp() As Variant
        Dim i As Long, j As Long

        Temp = Arr

        ReDim Arr(UboundK(Arr) - (EndPosition - StartPosition + 1))
        For i = 0 To StartPosition - 1
            Arr(i) = Temp(i)
        Next i
        For j = EndPosition + 1 To UboundK(Temp)
            Arr(i) = Temp(j)
            i = i + 1
        Next j
    End Sub

    Private Sub InsertElements(ByRef Goal() As Variant, Value As Variant)

        Dim Temp() As Variant
        Dim i As Long

        Temp = Goal
        ReDim Goal((2 * (1 + UboundK(Temp))) - 2)
        For i = 0 To UboundK(Temp)
            Goal(i * 2 + 0) = Temp(i)
            If i * 2 + 1 < UboundK(Goal) Then Goal(i * 2 + 1) = Value
        Next

    End Sub

    Private Function UboundK(Arr As Variant) As Long
        On Error Resume Next
        UBoundK = -1
        UBoundK = Ubound(Arr)
    End Function

    Private Function UboundN(Arr As cCollection) As Long
        On Error Resume Next
        UboundN = -1
        UboundN = Ubound(Arr.Nodes)
    End Function

    Private Sub PushArray(Byref Arr As Variant, Value As Variant)
        ReDim Preserve Arr(UboundK(Arr) + 1)
        Arr(UboundK(Arr)) = Value
    End Sub
'

' Handle Input

    Private Sub ConsoleText_KeyDown(pKey As Long, ByVal ShiftKey As Integer)
        If pKey = 13 Then
            ConsoleText.SelStart = GetMaxSelStart
            ConsoleText.Sellength = 0
            ConsoleText.SelColor = in_Basic
        End If
    End Sub
    
    Private Sub ConsoleText_KeyPress(Char As Long)
        If PasswordMode Then 
            Select Case Char
                Case 8
                    UserInput = Mid(UserInput, 1, Len(UserInput) - 1)
                Case 32 To 126, 128 To 255
                    UserInput = UserInput & Chr(Char)
                    Char = 42 '"*""
                Case Else
            End Select
        End If
    End Sub

    Private Sub ConsoleText_KeyUp(pKey As Long, ByVal ShiftKey As Integer)
        
        Dim Lines As Variant
        Lines = Split(ConsoleText.Text, vbCrLf)
        Call SetUpNewLine
        Select Case pKey
            Case vbKeyReturn
                If PasswordMode = False Then
                    Call PushArray(PreviousCommands, GetLine(ConsoleText.Text, UboundK(Lines) - 1))
                    PreviousCommandsIndex = UboundK(PreviousCommands) + 1
                End If
                If HandleEnter <> "EXIT()/\" Then
                    Call SetPositions
                    ConsoleText.SelStart = GetMaxSelStart
                    ConsoleText.Sellength = 0
                    ConsoleText.SelColor = in_Basic
                Else
                    Unload Console
                End If
            Case vbKeyUp
                If PasswordMode Then Exit Sub
                If Workmode = WorkModeEnum.Logging Then
                    PreviousCommandsIndex = PreviousCommandsIndex - 1
                    If PreviousCommandsIndex < 0 Then PreviousCommandsIndex = UboundK(PreviousCommands)
                    ConsoleText.SelStart = GetTextLength(ConsoleText.Text, vbCrLf, UboundK(Lines) - 1)
                    ConsoleText.SelLength = Len(ConsoleText.Text)
                    ConsoleText.SelText = PrintStarter & PreviousCommands(PreviousCommandsIndex)
                End If
            Case vbKeyDown
                If PasswordMode Then Exit Sub
                If Workmode = WorkModeEnum.Logging Then
                    PreviousCommandsIndex = PreviousCommandsIndex + 1
                    If PreviousCommandsIndex > UboundK(PreviousCommands) Then PreviousCommandsIndex = 0
                    ConsoleText.SelStart = GetTextLength(ConsoleText.Text, vbCrLf, UboundK(Lines) - 1)
                    ConsoleText.SelLength = Len(ConsoleText.Text)
                    ConsoleText.SelText = PrintStarter & PreviousCommands(PreviousCommandsIndex)
                End If
            Case Else
                Call HandleOtherKeys(pKey, ShiftKey)
        End Select

    End Sub

    Private Function HandleEnter() As Variant

        Dim i As Long
        Dim Line As String
        Dim Value As Variant
        If PasswordActive = False Then
            Call HandlePassword
            Exit Function
        End If
        Line = GetLine(ConsoleText.Text, CurrentLineIndex - 1)
        MulitlineEnd:
        Select Case WorkMode
            Case WorkModeEnum.Logging
                Value = HandleCode(Line)
                If CStr(Value) = "EXIT()/\" Then HandleEnter = Value: Exit Function
                If Value Like "DIMVARIABLE*" And Not Value Like "DIMVARIABLE-1" Then
                    Call HandleDimVariable(-1, CLng(Replace(CStr(Value), "DIMVARIABLE", "")))
                End If
                Call PrintEnter(Value, in_System)
            Case WorkModeEnum.UserInputt
                UserInput = Replace(Line, vbCrLf, "")
            Case WorkModeEnum.MultilineMode, WorkModeEnum.ScriptMode
                If UCase(Line) = "ENDSCRIPT" Or UCase(Line) = "ENDMULTILINE" Then
                    Dim Temp As String
                    Dim TempCount As Long

                    TempCount = CurrentLineIndex - 2
                    Line = GetLine(ConsoleText.Text, TempCount)
                    Do Until UCase(GetLine(ConsoleText.Text, TempCount)) = "MULTILINE" Or UCase(GetLine(ConsoleText.Text, TempCount)) = "SCRIPT"
                        Line = GetLine(ConsoleText.Text, TempCount)
                        TempCount = TempCount - 1
                        Temp = Line & vbCrLf & Temp
                    Loop
                    Line = Mid(Temp, 3, Len(Temp))
                    

                    If WorkMode = WorkModeEnum.ScriptMode Then
                        Dim Name As String
                        Dim Arguments As String
                        Dim Script As String
                        Line = Replace(Line, "_" & vbCrLf, "")
                        Line = Replace(Line, vbCrLf, LineSeperator)
                        TempCount = InStr(1, Line, LineSeperator)
                        Name = GetFunctionName(Mid(Line, 1, TempCount - 1))
                        Arguments = GetParanthesesText((Mid(Line, 1, TempCount - 1)))
                        Script = Mid(Line, TempCount + Len(LineSeperator), Len(Line))
                        Call AddScript(Name, Arguments, Script)
                        Call PrintEnter("New Script with Name " & Name & " was created", in_System)
                        WorkMode = WorkModeEnum.Logging
                    Else
                        Line = Replace(Line, vbCrLf, "")
                        WorkMode = WorkModeEnum.Logging
                        GoTo MulitlineEnd
                    End If
                End If
        End Select
        If PasteStarter = False Or Workmode = WorkModeEnum.MultilineMode Or WorkMode = WorkModeEnum.ScriptMode Then
        Else
            Call PrintConsole(PrintStarter, in_System)
        End If

    End Function

    Private Function HandleCode(Line As String) As Variant

        Dim AssignOperator  As Long
        Dim LeftSide        As String
        Dim LeftSidePos()   As Long
        Dim RightSide       As String
        Dim RightSidePos()  As Long
        Dim Value           As Variant

        AssignOperator = InStr(1, Line, AsgOperator)
        If AssignOperator <> 0 Then LeftSide = Mid(Line, 1, AssignOperator - 1) Else LeftSide = Empty
        LeftSidePos   = GetFuncTreePosition(LeftSide)
        If AssignOperator <> 0 Then RightSide = Mid(Line, AssignOperator + Len(AsgOperator), Len(Line)) Else RightSide = Mid(Line, 1, Len(Line))
        RightSidePos  = GetFuncTreePosition(RightSide)

        Select Case True
            Case UboundK(RightSidePos) <> -1
                Dim Args() As Variant
                Dim Pos(0) As Long
                Args = GetFunctionArgs(RightSide)
                If UboundK(Args) = -1 And GetParanthesesText(RightSide) <> "" Then
                    ReDim Args(0)
                    Args(0) = GetParanthesesText(RightSide)
                End If
                Call RecursiveReturnVariable(Args)
                Pos(0) = RightSidePos(1)
                If UboundK(Args) <> -1 Then
                    Value = ReturnVariable(RightSidePos(0), Pos, True, Args)
                Else
                    Value = ReturnVariable(RightSidePos(0), Pos, True)
                End If
            Case IsNumeric(RightSide), IsDate(RightSide)
                Value = InterpretVariable(RightSide)
            Case InString(RightSide, 1, Len(RightSide))
                Value = MidP(CStr(RightSide), 2, Len(CStr(RightSide)) - 1)
            Case Mid(RightSide, 1, 1) Like "[?]*"
                If stdLambda_Active Then
                    If CreateLambda(LeftSide, Mid(RightSide, 2, Len(RightSide)), False, False) Then
                        HandleCode = "Lambda created successfully"
                        Exit Function
                    Else
                        HandleCode = "Could not create Lambda"
                        Exit Function
                    End If
                Else
                    HandleCode = "stdLambda is deactivated"
                    Exit Function
                End If
            Case Else
                Value = HandleSpecial(RightSide)
                If CStr(Value) = "NOSPECIAL" Then
                    Value = HandleReturnOperator(RightSide)
                    If Value = "could not handle operator" Then
                        Dim Temp As Variant
                        Value = RunApplication(RightSide, Temp)
                    End If
                ElseIf CStr(Value) = "EXIT()/\" Then
                    HandleCode = Value
                    Exit Function
                Else
                End If
        End Select

        If LeftSide = Empty Then
            HandleCode = Value
        Else
            Call SetVariable(ConsVarIndex, 0, LeftSide, "Public ConsoleVariable As Variant", "No Arguments", Value)
            If UboundK(LeftSidePos) = -1 Then
                HandleCode = "New Variable " & LeftSide & " was created with Value: " & Value
            Else
                HandleCode = "New Value assigned to Variable "
            End If
        End If

    End Function

    Private Function HandleSpecial(Line As String) As Variant
        Select Case True
            Case UCase(Line) Like "HELP"           : HandleSpecial = HandleHelp
            Case UCase(Line) Like "CLEAR"          : HandleSpecial = HandleClear
            Case UCase(Line) Like "MULTILINE"      : HandleSpecial = "": Workmode = WorkModeEnum.MultilineMode
            Case UCase(Line) Like "INFO"           : HandleSpecial = HandleLastError
            Case UCase(Line) Like "EXIT"           : HandleSpecial = "EXIT()/\": Call HandleClear
            Case UCase(Line) Like "SCRIPT"         : HandleSpecial = "": Workmode = WorkModeEnum.ScriptMode
            Case UCase(Line) Like "FOR(*)"         : HandleSpecial = HandleLoop(Line)
            Case UCase(Line) Like "UNTIL(*)"       : HandleSpecial = HandleLoop(Line)
            Case UCase(Line) Like "WHILE(*)"       : HandleSpecial = HandleLoop(Line)
            Case UCase(Line) Like "IF(*)"          : HandleSpecial = HandleCondition(Line)
            Case UCase(Line) Like "SELECT(*)"      : HandleSpecial = HandleCondition(Line)
            Case UCase(Line) Like "DIM * AS *"     : HandleSpecial = HandleNewVariable(Line)
            Case UCase(Line) Like "PUBLIC * AS *"  : HandleSpecial = HandleNewVariable(Line)
            Case UCase(Line) Like "PRIVATE * AS *" : HandleSpecial = HandleNewVariable(Line)
            Case Else                              : HandleSpecial = "NOSPECIAL"
        End Select
    End Function

    Private Function HandleLastError() As String
        Select Case LastError
        Case Empty:
            HandleLastError = "No previous Error detected"
        Case 1
            HandleLastError = "The last run line was executed without problems"
        Case 2
            HandleLastError = "The last run line couldnt be executed. Some Problems could be:" & vbCrLf & _
                      "    1. Line wasnt written correctly"                                    & vbCrLf & _
                      "    2. Code doesnt exist"                                               & vbCrLf & _
                      "    3. There exists more than one publ1c procedure with the same name"  & vbCrLf & _
                      "    4. The Procedure has the same name as the component it sits in"     & vbCrLf & _
                      "    5. The Workbook with its VBProject isnt open"                       & vbCrLf & _
                      "    6. The parameters were passed wrong"                                & vbCrLf '1 in publ1c to not mess with GetAllProcedures
        Case 3
            HandleLastError = " You passed too many arguments, VBA limits ParamArray Arguments to 30"
        End Select
    End Function

    Private Function HandleClear() As String
        ConsoleText.Text = ""
        ConsoleText.SelStart = 0
        ConsoleText.SelLength = Len(ConsoleText.Text)
        ConsoleText.SelColor = in_Basic
        Call SetUpNewLine
        HandleClear = " "
    End Function

    Private Function HandleHelp() As String
        HandleHelp = _
        "--------------------------------------------------"                                           & vbCrLf & _
        "This Console can do the following:"                                                           & vbCrLf & _
        "1. It can be used as a form to show messages, ask questions to the user or get a user input"  & vbCrLf & _
        "2. It can be used to show and log errors and handle them by user input"                       & vbCrLf & _
        "3. It can run Procedures with up to 29 arguments"                                             & vbCrLf & _
        ""                                                                                             & vbCrLf & _
        "HOW TO USE IT:"                                                                               & vbCrLf & _
        "   Run a Procedure:"                                                                          & vbCrLf & _
        "       To run a procedure you have to write the name of said procedure (Case sensitive)"      & vbCrLf & _
        "       If you want to pass parameters you have to write    |; | between every parameter"      & vbCrLf & _
        "       Example:"                                                                              & vbCrLf & _
        "           Say; THIS IS A PARAMETER; THIS IS ANOTHER PARAMETER"                               & vbCrLf & _
        ""                                                                                             & vbCrLf & _
        "   Ask a question:"                                                                           & vbCrLf & _
        "       Use CheckPredeclaredAnswer"                                                            & vbCrLf & _
        "           Param1 = Message to be showwn"                                                     & vbCrLf & _
        "           Param2 = Array of Values, which are acceptable answers"                            & vbCrLf & _
        "           Param3 = Array of Messages, which show a text according to answer in Param2"       & vbCrLf & _
        "       The Function will loop until one of the acceptable answers is typed"                   & vbCrLf & _
        "--------------------------------------------------"                                           & vbCrLf
    End Function

    Private Function RunApplication(Name As String, Arguments As Variant) As Variant

        On Error GoTo Error
        Select Case UBoundK(Arguments)
            Case -1:   RunApplication = Application.Run(Name)
            Case 00:   RunApplication = Application.Run(Name, Arguments(0))
            Case 01:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1))
            Case 02:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2))
            Case 03:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3))
            Case 04:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4))
            Case 05:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5))
            Case 06:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6))
            Case 07:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7))
            Case 08:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8))
            Case 09:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9))
            Case 10:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10))
            Case 11:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11))
            Case 12:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12))
            Case 13:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13))
            Case 14:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14))
            Case 15:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15))
            Case 16:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16))
            Case 17:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17))
            Case 18:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18))
            Case 19:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19))
            Case 20:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20))
            Case 21:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21))
            Case 22:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22))
            Case 23:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23))
            Case 24:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24))
            Case 25:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25))
            Case 26:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26))
            Case 27:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26), Arguments(27))
            Case 28:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26), Arguments(27), Arguments(28))
            Case 29:   RunApplication = Application.Run(Name, Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26), Arguments(27), Arguments(28), Arguments(29))
            Case Else: RunApplication = "Too many Arguments": LastError = 3
        End Select
        If IsError(RunApplication) Then
            GoTo Error:
        Else
            Exit Function
        End If
        Error:
        RunApplication = "Could not run Procedure. Procedure might not exist"
        LastError = 2

    End Function

    Private Static Function HandleOtherKeys(pKey As Long, ByVal ShiftKey As Integer) As String

        Static CapitalKey As Boolean
        Dim AsciiChar As String
        Dim CurrentWord As String
        Dim CurrentLine As String
        Dim CurrentSelection(1) As Long
        
        ' Adjust for Shift key (Uppercase letters, special characters)
        CurrentLine = GetLine(ConsoleText.Text, CurrentLineIndex)
        CurrentWord = GetWord(CurrentLine)
        If pKey = vbKeyCapital Then
            CapitalKey = CapitalKey Xor True
            GoTo SkipKey
        End If
        If CapitalKey = True Then ShiftKey = 1
        Select Case ShiftKey
            Case 0
                ' Base character
                Select Case pKey
                    Case vbKeyA To vbKeyZ:      AsciiChar = LCase(Chr(pKey))
                    Case vbKey0 To vbKey9:      AsciiChar = Chr(pKey)
                    Case vbKeySpace:            AsciiChar = " "
                    Case vbKeyBack:             AsciiChar = Chr(8) ' Backspace
                    Case vbKeyReturn:           AsciiChar = Chr(13) ' Carriage Return
                    Case vbKeyTab:              AsciiChar = Chr(9) ' Tab
                    Case vbKeyMultiply:         AsciiChar = "*"
                    Case vbKeyAdd, 187:         AsciiChar = "+"
                    Case vbKeySubtract, 189:    AsciiChar = "-"
                    Case vbKeyDecimal, 190:     AsciiChar = "."
                    Case vbKeyDivide:           AsciiChar = "/"
                    Case 188:                   AsciiChar = ","
                    Case 191:                   AsciiChar = "#"
                    Case 226:                   AsciiChar = "<"
                    Case vbKeyRight:            AsciiChar = "RIGHT"
                    Case vbKeyLeft:             AsciiChar = "LEFT"
                    Case vbKeyUp:               AsciiChar = "UP"
                    Case vbKeyDown:             AsciiChar = "DOWN"
                End Select
            Case = 1
                Select Case pKey
                    Case vbKeyA To vbKeyZ:      AsciiChar = UCase(AsciiChar)
                    Case vbKey1:                AsciiChar = "!"
                    Case vbKey2:                AsciiChar = Chr(34) ' """
                    Case vbKey3:                AsciiChar = "§"
                    Case vbKey4:                AsciiChar = "$"
                    Case vbKey5:                AsciiChar = "%"
                    Case vbKey6:                AsciiChar = "&"
                    Case vbKey7:                AsciiChar = "/"
                    Case vbKey8:                AsciiChar = "("
                    Case vbKey9:                AsciiChar = ")"
                    Case vbKey0:                AsciiChar = "="
                    Case 187:                   AsciiChar = "*"
                    Case 188:                   AsciiChar = ";"
                    Case 189:                   AsciiChar = "_"
                    Case 190:                   AsciiChar = ":"
                    Case 191:                   AsciiChar = "'"
                    Case 226:                   AsciiChar = ">"
                End Select
            Case 2

            Case 3
                    Case 226:                   AsciiChar = "|"
        End Select
        Select Case AsciiChar
            Case "RIGHT"
                If PasswordMode Then Exit Function
                If ConsoleText.SelStart = GetMaxSelStart Then
                    If Intellisense_Active = True Then
                        If IntellisenseList.ListCount > 0 Then IntellisenseList.SetFocus
                    End If
                End If
            Case Else
                Call SetUp_IntelliSenseList(CurrentWord)
        End Select
        SkipKey:
        Call SetPositions
        CurrentSelection(0) = ConsoleText.SelStart
        CurrentSelection(1) = ConsoleText.Sellength
        Call ColorWord
        ConsoleText.SelStart = CurrentSelection(0)
        ConsoleText.SelLength = CurrentSelection(1)
        ConsoleText.SelColor = in_Basic
        HandleOtherKeys = AsciiChar

    End Function

    Private Sub RecursiveReturnVariable(ByRef Arguments() As Variant)
        
        Dim i As Long

        Dim CurrentArgPos() As Long
        Dim CurrentArguments() As Variant
        Dim Quotes() As Long
        Dim TempVar() As Variant
        ReDim TempVar(0)
        If UboundK(Arguments) = -1 Then Exit Sub
        For i = 0 To UboundK(Arguments)
            Quotes = InStrAll(CStr(Arguments(i)), Chr(34))
            TempVar(0) = CStr(Arguments(i))
            Select Case True
                Case IsNumeric(Arguments(i))
                    Arguments(i) = CLng(Arguments(i))
                Case IsDate(Arguments(i))
                    Arguments(i) = CDate(Arguments(i))
                Case UboundK(Quotes) >= 1
                    Arguments(i) = MidP(CStr(Arguments(i)), 2, Len(CStr(Arguments(i))) - 1)
                Case UboundK(GetAllOperators(TempVar)) <> 0 And TempVar(UboundK(TempVar)) <> Arguments(i)
                    Arguments(i) = HandleReturnOperator(CStr(Arguments(i)))
                Case Else
                        CurrentArgPos = GetFuncTreePosition(CStr(Arguments(i)))
                        If UboundK(CurrentArgPos) <> -1 Then
                            Dim Tempp(0) As Long
                            Tempp(0) = CurrentArgPos(1)
                            CurrentArguments = GetFunctionArgs(CStr(Arguments(i)))
                            Call RecursiveReturnVariable(CurrentArguments)
                            If UboundK(CurrentArguments) > 0 Then
                                Arguments(i) = ReturnVariable(CurrentArgPos(0), Tempp, True, CurrentArguments)
                            Else
                                Arguments(i) = ReturnVariable(CurrentArgPos(0), Tempp, True)
                            End If
                        End If
                    If UboundK(CurrentArgPos) = -1 Then Arguments(i) = "Could not handle Argument"
            End Select
        Next i
    End Sub

    Private Function HandleReturnOperator(Line As String) As Variant

        Dim Values() As Variant
        Dim i As Long
        Dim PassValue1 As Variant
        Dim PassValue2 As Variant

        ReDim Values(0)
        Values(0) = Line
        Values = GetAllOperators(Values)

        On Error GoTo Error
        If UboundK(Values) <> -1 Then
            Do Until UboundK(Values) = 0
                PassValue1 = InterpretVariableTEMP(Values(0))
                PassValue2 = InterpretVariableTEMP(Values(2))
                Select Case UCase(Values(1))
                    Case "==", " IS "        : If(PassValue1   =   PassValue2) Then Values(0) = True Else Values(0) = False
                    Case "<>", "!=", " NOT " : If(PassValue1   <>  PassValue2) Then Values(0) = True Else Values(0) = False
                    Case "||", " OR "        : If(PassValue1   Or  PassValue2) Then Values(0) = True Else Values(0) = False
                    Case "&&", " AND "       : If(PassValue1   And PassValue2) Then Values(0) = True Else Values(0) = False
                    Case "", "", " XOR "     : If(PassValue1   Xor PassValue2) Then Values(0) = True Else Values(0) = False
                    Case "<"                 : If(PassValue1    <  PassValue2) Then Values(0) = True Else Values(0) = False
                    Case ">"                 : If(PassValue1    >  PassValue2) Then Values(0) = True Else Values(0) = False
                    Case "=<", "<="          : If(PassValue1    =< PassValue2) Then Values(0) = True Else Values(0) = False
                    Case ">=", "=>"          : If(PassValue1    >= PassValue2) Then Values(0) = True Else Values(0) = False
                    Case Else                : Values(0) = HandleCalcOperator(Values(0), Values(1), Values(2))
                End Select
                Call StitchArray(Values, 1, 2)
            Loop
        Else
            GoTo Error
        End If
        HandleReturnOperator = Values(0)
        Exit Function

        Error:
        HandleReturnOperator = "could not handle operator"
    End Function

    Private Function HandleCalcOperator(ByRef Value1 As Variant, Operator As Variant, Value2 As Variant) As Variant
        Dim VariablePos() As Long
        Dim Name As String
        VariablePos = GetFuncTreePosition(CStr(Value1))
        If UboundK(VariablePos) <> -1 Then Name = Value1
        Value1 = InterpretVariableTEMP(Value1)
        Value2 = InterpretVariableTEMP(Value2)
        Select Case UCase(Operator)
            Case "++"                : HandleCalcOperator = CLng(Value1)  +   CLng(Value2): Value1 = CLng(Value1) + CLng(Value2)
            Case "--"                : HandleCalcOperator = CLng(Value1)  -   CLng(Value2): Value1 = CLng(Value1) - CLng(Value2)
            Case "**"                : HandleCalcOperator = CLng(Value1)  *   CLng(Value2): Value1 = CLng(Value1) * CLng(Value2)
            Case "//"                : HandleCalcOperator = CLng(Value1)  /   CLng(Value2): Value1 = CLng(Value1) / CLng(Value2)
            Case "^^"                : HandleCalcOperator = CLng(Value1)  ^   CLng(Value2): Value1 = CLng(Value1) ^ CLng(Value2)
            Case "+="                : HandleCalcOperator = CLng(Value1)  +   CLng(Value2): Value1 = CLng(Value1) + CLng(Value2)
            Case "-="                : HandleCalcOperator = CLng(Value1)  -   CLng(Value2): Value1 = CLng(Value1) - CLng(Value2)
            Case "*="                : HandleCalcOperator = CLng(Value1)  *   CLng(Value2): Value1 = CLng(Value1) * CLng(Value2)
            Case "/="                : HandleCalcOperator = CLng(Value1)  /   CLng(Value2): Value1 = CLng(Value1) / CLng(Value2)
            Case "+"                 : HandleCalcOperator = CLng(Value1)  +   CLng(Value2)
            Case "-"                 : HandleCalcOperator = CLng(Value1)  -   CLng(Value2)
            Case "*"                 : HandleCalcOperator = CLng(Value1)  *   CLng(Value2)
            Case "/"                 : HandleCalcOperator = CLng(Value1)  /   CLng(Value2)
            Case "^"                 : HandleCalcOperator = CLng(Value1)  ^   CLng(Value2)
            Case "&"                 : HandleCalcOperator = CStr(Value1)  &   CStr(Value2)
            Case Else                : HandleCalcOperator = "No valid Operator"
        End Select
        If UboundK(VariablePos) <> -1 Then Call SetVariable(VariablePos(0), VariablePos(1), Name, "NOCHANGE", "NOCHANGE", HandleCalcOperator)
    End Function

    Private Function HandlePassword() As Boolean
        If PasswordActive Then
            HandlePassword = True
        Else
            If UserInput = p_Password Then
                HandlePassword = True
                PasswordActive = True
                UserInput = ""
                PasswordMode = False
                Call PrintEnter("Password accepted", in_System)
                Call PrintConsole(PrintStarter, in_System)
            Else
                UserInput = ""
                PasswordMode = True
                Call PrintEnter("Enter Password", in_System)
            End If
        End If
    End Function

    Private Function RunScript(Script As Variant, ScriptArgs As Variant, Optional Arguments As Variant) As Variant
        Dim Lines() As String
        Dim i As Long
        Dim Value As Variant
        Dim ScriptArguments() As String
        Dim CurrentArg As String
        Dim Scope(255) As Long
        Dim ScopeIndex As Long

        Call InitScope(Scope)
        ScriptArguments = SplitString(CStr(ScriptArgs), ArgSeperator)
        For i = 0 To UboundK(ScriptArguments)
            Scope(ScopeIndex) = CLng(Replace(CStr(HandleCode(CStr("DIM " & ScriptArguments(i)))), "DIMVARIABLE", ""))
            ScopeIndex = ScopeIndex + 1
        Next i
        If UBoundK(ScriptArguments) = UboundK(Arguments) And UboundK(Arguments) <> -1 Then
            For i = 0 To UboundK(Arguments)
                CurrentArg = MidP(CStr(ScriptArguments(i)), 1, InStr(1, CStr(ScriptArguments(i)), " ") - 1)
                Call HandleCode(CurrentArg & AsgOperator & CStr(Arguments(i)))
            Next
        End If

        Lines = SplitString(CStr(Script), LineSeperator)
        For i = 0 To UboundK(Lines) - 1
            Value = HandleCode(Lines(i))
            If CStr(Value) = "EXIT()/\" Then RunScript = Value: Exit Function
            If Value Like "DIMVARIABLE*" And Not Value Like "DIMVARIABLE-1" Then
                Scope(ScopeIndex) = CLng(Replace(CStr(Value), "DIMVARIABLE", ""))
                ScopeIndex = ScopeIndex + 1
            End If
            Call PrintEnter(Value, in_System)
        Next
        Call DeleteScope(Scope)
        If UboundK(Lines) <> -1 Then
            RunScript = "Script ran successfully"
        Else
            RunScript = "Script did not run successfully"
        End If
    End Function

    Private Function InterpretVariable(Value As Variant) As Variant
        Dim Arr() As Variant
        Select Case True
            Case IsNumeric(Value)
                If CLng(Value) = Round(CLng(Value), 0) Then
                   InterpretVariable = CLng(Value)
                Else 
                    InterpretVariable = CDbl(Value)
                End If
            Case IsDate(Value)                      : InterpretVariable = Cdate(Value)
            Case VarType(Value) = vbEmpty           : InterpretVariable = Empty
            Case VarType(Value) = vbNull            : InterpretVariable = Null
            Case VarType(Value) = vbInteger         : InterpretVariable = CInt(Value)
            Case VarType(Value) = vbLong            : InterpretVariable = CLng(Value)
            Case VarType(Value) = vbSingle          : InterpretVariable = CSng(Value)
            Case VarType(Value) = vbDouble          : InterpretVariable = CDbl(Value)
            Case VarType(Value) = vbCurrency        : InterpretVariable = Ccur(Value)
            Case VarType(Value) = vbString          : InterpretVariable = CStr(Value)
            Case VarType(Value) = vbObject          : InterpretVariable = "Object"
            Case VarType(Value) = vbError           : InterpretVariable = "Error"
            Case VarType(Value) = vbBoolean         : InterpretVariable = CBool(Value)
            Case VarType(Value) = vbVariant         : InterpretVariable = CVar(Value)
            Case VarType(Value) = vbDataObject      : InterpretVariable = "DataObject"
            Case VarType(Value) = vbDecimal         : InterpretVariable = CDec(Value)
            Case VarType(Value) = vbByte            : InterpretVariable = CByte(Value)
            Case VarType(Value) = vbLongLong        : InterpretVariable = CLngLng(Value)
            Case VarType(Value) = vbUserDefinedType : InterpretVariable = "User defined Type"
            Case VarType(Value) = vbArray           : InterpretVariable = Arr
        End Select
    End Function

    Private Function InterpretVariableTEMP(Value As Variant) As Variant
        Dim VariablePos() As Long
        Dim Temp(0) As Long
        VariablePos = GetFuncTreePosition(CStr(Value))
        If UboundK(VariablePos) <> -1 Then
            Temp(0) = CLng(VariablePos(1))
            InterpretVariableTEMP = ReturnVariable(VariablePos(0), Temp, True)
        Else
            InterpretVariableTEMP = InterpretVariable(CStr(Value))
        End If
    End Function

    Private Function GetVariableByType(Value As Variant, DataType As String) As Variant
        Dim Arr() As Variant
        Select Case True
            Case UCase(DataType) Like "* AS EMPTY"      : GetVariableByType = Empty
            Case UCase(DataType) Like "* AS NULL"       : GetVariableByType = Null
            Case UCase(DataType) Like "* AS INTEGER"    : GetVariableByType = CInt(Value)
            Case UCase(DataType) Like "* AS LONG"       : GetVariableByType = CLng(Value)
            Case UCase(DataType) Like "* AS SINGLE"     : GetVariableByType = CSng(Value)
            Case UCase(DataType) Like "* AS DOUBLE"     : GetVariableByType = CDbl(Value)
            Case UCase(DataType) Like "* AS CURRENCY"   : GetVariableByType = Ccur(Value)
            Case UCase(DataType) Like "* AS DATE"       : GetVariableByType = Cdate(Value)
            Case UCase(DataType) Like "* AS STRING"     : GetVariableByType = CStr(Value)
            Case UCase(DataType) Like "* AS OBJECT"     : GetVariableByType = "Object"
            Case UCase(DataType) Like "* AS ERROR"      : GetVariableByType = "Error"
            Case UCase(DataType) Like "* AS BOOLEAN"    : GetVariableByType = CBool(Value)
            Case UCase(DataType) Like "* AS VARIANT"    : GetVariableByType = CVar(Value)
            Case UCase(DataType) Like "* AS DATAOBJECT" : GetVariableByType = "DataObject"
            Case UCase(DataType) Like "* AS DECIMAL"    : GetVariableByType = CDec(Value)
            Case UCase(DataType) Like "* AS BYTE"       : GetVariableByType = CByte(Value)
            Case UCase(DataType) Like "* AS LONGLONG"   : GetVariableByType = CLngLng(Value)
            Case UCase(DataType) Like "* AS UDT"        : GetVariableByType = "User defined Type"
            Case UCase(DataType) Like "* AS ARRAY"      : GetVariableByType = Arr
        End Select
    End Function

    Private Function HandleLoop(Line As String) As Variant
        Dim Arguments() As String
        Dim i As Long
        Dim Value As Variant
        Dim Scope(255) As Long
        Dim ScopeIndex As Long

        Call InitScope(Scope)
        Arguments = SplitString(GetParanthesesText(Line), ArgSeperator)
        Select Case True
            Case UCase(Line) Like "FOR(*)"
                Call PrintEnter(HandleCode(Arguments(0)), in_System)
                Do Until HandleReturnOperator(Arguments(1))
                    For i = 3 To UboundK(Arguments)
                        Value = HandleCode(Arguments(i))
                        If CStr(Value) = "EXIT()/\" Then HandleLoop = Value: Exit Function
                        If Value Like "DIMVARIABLE*" And Not Value Like "DIMVARIABLE-1" Then
                            Scope(ScopeIndex) = CLng(Replace(CStr(Value), "DIMVARIABLE", ""))
                            ScopeIndex = ScopeIndex + 1
                        End If
                        Call PrintEnter(Value, in_System)
                        Call PrintEnter(HandleReturnOperator(Arguments(2)), in_System)
                    Next
                Loop
            Case UCase(Line) Like "UNTIL(*)", UCase(Line) Like "WHILE(*)"
                Do Until HandleReturnOperator(Arguments(0))
                    For i = 1 To UboundK(Arguments)
                        Value = HandleCode(Arguments(i))
                        If CStr(Value) = "EXIT()/\" Then HandleLoop = Value: Exit Function
                        If Value Like "DIMVARIABLE*" And Not Value Like "DIMVARIABLE-1" Then
                            Scope(ScopeIndex) = CLng(Replace(CStr(Value), "DIMVARIABLE", ""))
                            ScopeIndex = ScopeIndex + 1
                        End If
                        Call PrintEnter(Value, in_System)
                    Next
                Loop
        End Select
        Call DeleteScope(Scope)
    End Function

    Private Function HandleCondition(Line As String) As Variant
        Dim i As Long
        Dim j As Long
        Dim Condition As String
        Dim Arguments() As String
        Dim Value As Variant
        Dim Scope(255) As Long
        Dim ScopeIndex As Long

        Call InitScope(Scope)
        If UCase(Line) Like "IF(*)" Then
            Dim ArgumentsThen() As String
            Dim ArgumentsElse() As String

            Arguments = Split(GetParanthesesText(Line), " Then ")
            Condition = Arguments(0)
            Arguments = SplitString(Arguments(1), " Else ")
            ArgumentsThen = Split(Arguments(0), ArgSeperator)
            ArgumentsElse = Split(Arguments(1), ArgSeperator)
            If HandleReturnOperator(Condition) Then
                For i = 0 To UboundK(ArgumentsThen)
                    Value = HandleCode(ArgumentsThen(i))
                    If CStr(Value) = "EXIT()/\" Then HandleCondition = Value: Exit Function
                    If Value Like "DIMVARIABLE*" And Not Value Like "DIMVARIABLE-1" Then
                        Scope(ScopeIndex) = CLng(Replace(CStr(Value), "DIMVARIABLE", ""))
                        ScopeIndex = ScopeIndex + 1
                    End If
                    Call PrintEnter(Value, in_System)
                Next i
                HandleCondition = True
            Else
                For i = 0 To UboundK(ArgumentsElse)
                    Value = HandleCode(ArgumentsElse(i))
                    If CStr(Value) = "EXIT()/\" Then HandleCondition = Value: Exit Function
                    If Value Like "DIMVARIABLE*" And Not Value Like "DIMVARIABLE-1" Then
                        Scope(ScopeIndex) = CLng(Replace(CStr(Value), "DIMVARIABLE", ""))
                        ScopeIndex = ScopeIndex + 1
                    End If
                    Call PrintEnter(Value, in_System)
                Next i
                HandleCondition = False
            End If
        Else
            Dim ArgumentsCase() As String
            Arguments = Split(GetParanthesesText(Line), " Case ")
            Condition = Arguments(0)
            For i = 1 To UboundK(Arguments)
                ArgumentsCase = Split(Arguments(i), ArgSeperator)
                If UCase(ArgumentsCase(0)) = "ELSE" Or HandleReturnOperator(Condition & ArgumentsCase(0)) Then
                    For j = 1 To UboundK(ArgumentsCase)
                        Value = HandleCode(ArgumentsCase(j))
                        If CStr(Value) = "EXIT()/\" Then HandleCondition = Value: Exit Function
                        If Value Like "DIMVARIABLE*" And Not Value Like "DIMVARIABLE-1" Then
                            Scope(ScopeIndex) = CLng(Replace(CStr(Value), "DIMVARIABLE", ""))
                            ScopeIndex = ScopeIndex + 1
                        End If
                        Call PrintEnter(Value, in_System)
                    Next
                    HandleCondition = ArgumentsCase(0)
                    Exit Function
                Else
                    HandleCondition = Empty
                End If
            Next
        End If
        Call DeleteScope(Scope)
    End Function

    Private Function HandleNewVariable(Line As String) As Variant
        Dim Words() As String
        Dim Pos As Long

        Words = Split(Line, " ")
        Pos = SetVariable(ConsVarIndex, 0, Words(1), Words(0) & " ConsoleVariable As " & Words(3), "No Arguments", Empty)
        If UCase(Words(0)) Like "DIM" Then
            HandleNewVariable = "DIMVARIABLE" & HandleDimVariable(Pos)
        Else
            HandleNewVariable = "DIMVARIABLE-1"
        End If
    End Function

    Private Function HandleDimVariable(Optional NodeIndex As Long = -1, Optional DeleteNodeIndex As Long = -1) As Long
        If DeleteNodeIndex <> -1 Then
            HandleDimVariable = DeleteNode(ConsVarIndex, DimVariables(DeleteNodeIndex))
            DimVariables(DeleteNodeIndex) = 0
            DimIndex = DimIndex - 1
        Else
            DimVariables(DimIndex) = NodeIndex
            HandleDimVariable = DimIndex
            DimIndex = DimIndex + 1
        End If
    End Function

    Private Sub DeleteScope(Arr() As Long)
        Dim i As Long
        For i = UBoundK(Arr) To 0 Step -1
            If Arr(i) <> -1 Then Call HandleDimVariable(-1, Arr(i))
        Next
    End Sub

    Private Sub InitScope(ByRef Arr() As Long)
        Dim i As Long
        For i = 0 To UBoundK(Arr)
            Arr(i) = -1
        Next
    End Sub

    
'

' Coloring

    Private Sub AssignColor()
        in_Basic        = RGB(255, 255, 255)
        in_System       = RGB(170, 170, 170)
        in_Procedure    = RGB(255, 255, 000)
        in_Operator     = RGB(255, 000, 000)
        in_Datatype     = RGB(000, 170, 000)
        in_Value        = RGB(000, 255, 000)
        in_String       = RGB(255, 165, 000)
        in_Statement    = RGB(255, 000, 255)
        in_Keyword      = RGB(000, 000, 255)
        in_Parantheses  = RGB(170, 170, 000)
        in_Variable     = RGB(000, 255, 255)
        in_Script       = RGB(255, 128, 128)
        in_Lambda       = RGB(170, 000, 170)
    End Sub

    Private Sub ColorWord()

        Dim Lines()            As String: Lines              = Split(ConsoleText.Text, vbCrLf)
        Dim CurrentLine        As String: CurrentLine        = RemoveOperators(GetLine(ConsoleText.Text, CurrentLineIndex))
        Dim RecognizerPosition As Long  : RecognizerPosition = InStr(1, Lines(UboundK(Lines)), Recognizer)
        Dim CurrentLinePoint   As Long  : CurrentLinePoint   = GetTextLength(ConsoleText.Text, vbCrLf, UboundK(Lines) - 1)
        Dim Words()            As String: Words              = Split(CurrentLine, " ")
        Dim i                  As Long
        Dim j                  As Long
        Dim PreviousWord       As String
        Dim Color              As Long

        If RecognizerPosition <> 0 Then CurrentLinePoint = CurrentLinePoint + (RecognizerPosition - 1) + Len(Recognizer)
        For i = 0 To UBoundK(Words)
            ConsoleText.SelStart = CurrentLinePoint + GetTextLength(CurrentLine, " ", i - 1)
            ConsoleText.SelLength = Len(Words(i))
            Select Case UCase(Words(i))
                Case "IF", "THEN", "ELSE", "END", "FOR", "EACH", "NEXT", "DO", "WHILE", "LOOP", "SELECT", "CASE", "EXIT", "CONTINUE"
                    Color = in_Statement
                Case "DIM", "PUBLIC", "PRIVATE", "GLOBAL", "TRUE", "FALSE", "FUNCTION", "SUB", "REDIM", "PRESERVE"
                    Color = in_Keyword
                Case "NOT", "AND", "OR", "XOR"
                    Color = in_Operator
                Case Else
                    If Ucase(PreviousWord) = "AS" Then
                        Color = in_Datatype                 
                    ElseIf IsNumeric(UCase(Words(i))) Then
                        Color = in_Value
                    Else
                        Dim Temp() As Long
                        Dim TreeIndex As Long
                        Dim Position As Long
                        Dim VariableType As Variant

                        Temp = GetFuncTreePosition(Words(i))
                        If UboundK(Temp) <> -1 Then
                            TreeIndex = Temp(0)
                            Position = Temp(1)
                            VariableType = RetVarSin(TreeIndex, Position, 0)
                            Select Case True
                                Case VariableType Like "* Procedure *"                      : Color = in_Procedure
                                Case VariableType Like "* ConsoleVariable As stdLambda"     : Color = in_Lambda
                                Case VariableType Like "* ConsoleVariable As ConsoleScript" : Color = in_Script
                                Case VariableType Like "* ConsoleVariable As *"             : Color = in_Variable
                                Case VariableType Like "As *"                               : Color = in_Variable
                                Case Else                                                   : Color = in_Basic
                            End Select
                            If InStr(1, Words(i), "(") <> 0 Then ConsoleText.SelLength = Len(Words(i)) - (Len(Words(i)) - InStr(1, Words(i), "(") + 1)
                        Else
                            Color = in_Basic
                        End If
                    End If
            End Select
            ConsoleText.SelColor = Color
            PreviousWord = Words(i)
        Next i

        Dim CurrentLineChar() As Long
        CurrentLine = GetLine(ConsoleText.Text, CurrentLineIndex)
        CurrentLineChar = InStrAll(CurrentLine, "(")
        For i = 0 To UboundK(CurrentLineChar)
            If CurrentLineChar(i) <> 0 Then
                ConsoleText.SelStart  = CurrentLinePoint + (CurrentLineChar(i) - 1)
                ConsoleText.SelLength = 1
                ConsoleText.SelColor = in_Parantheses + ((i Mod 2) * RGB(050, 050, 000))
            End If
        Next i
        CurrentLineChar = InStrAll(CurrentLine, ")")
        For i = 0 To UboundK(CurrentLineChar)
            If CurrentLineChar(i) <> 0 Then
                ConsoleText.SelStart  = CurrentLinePoint + (CurrentLineChar(i) - 1)
                ConsoleText.SelLength = 1
                ConsoleText.SelColor = in_Parantheses + (((i + 1) Mod 2) * RGB(050, 050, 000))
            End If
        Next i
        Dim Operators() As Variant
        Operators = Array("+", "*", "/", "-", "^", ":", ";", "<", ">", "=", "!", "|", "?", ",")
        For i = 0 To UboundK(Operators)
            CurrentLineChar = InStrAll(CurrentLine, CStr(Operators(i)))
            For j = 0 To UboundK(CurrentLineChar)
                If CurrentLineChar(j) <> 0 Then
                    ConsoleText.SelStart  = CurrentLinePoint + (CurrentLineChar(j) - 1)
                    ConsoleText.SelLength = 1
                    ConsoleText.SelColor = in_Operator
                End If
            Next
        Next i
        CurrentLineChar = InStrAll(CurrentLine, Chr(34))
        For i = 0 To UboundK(CurrentLineChar) Step 2
            If CurrentLineChar(i) <> 0 Then
                ConsoleText.SelStart  = CurrentLinePoint + (CurrentLineChar(i) - 1)
                If UboundK(CurrentLineChar) < (i + 1) Then ' Check for odd array
                    ConsoleText.SelLength = Len(ConsoleText.Text)
                Else
                    ConsoleText.SelLength = CurrentLineChar(i + 1)
                End If
                ConsoleText.SelColor = in_String
            End If
        Next i

    End Sub

    Private Function RemoveOperators(Text As String) As String
        Dim Operators() As Variant
        Dim i As Long
        Operators = Array("+", "*", "/", "-", "^", ":", ";", "<", ">", "=", "!", "|", "?", ",", "(", ")", "[", "]", "{", "}")
        For i = 0 To UboundK(Operators)
            Text = Replace(Text, CStr(Operators(i)), " ")
        Next
        RemoveOperators = Text
    End Function
'

' Intellisense

    Private Sub SetPositions()
        Dim Temp()       As String: Temp = Split(ConsoleText.Text, vbCrLf)
        Dim FactorHeight As Double: FactorHeight = Height / 4
        Dim FactorWidth  As Double: FactorWidth = Width / 8
        Dim ListOffset   As Double: ListOffset = 1.45
        Dim ListFactor   As Double: ListFactor = 1.35
        Dim CurrentLine  As String: CurrentLine = GetLine(ConsoleText.Text, UBoundK(Temp))
        ScrollTop = UBoundK(Temp) * 10 * ListFactor - FactorHeight - 100
        If Len(CurrentLine) * 10 >= ConsoleText.Left + 200 Then
            ScrollLeft = Len(CurrentLine) * 10 - 200
        Else
            ScrollLeft = ConsoleText.Left
        End If
        If Intellisense_Active = True Then
            IntellisenseList.Top = ScrollTop + (FactorWidth * ListOffset) + 130
            IntellisenseList.Left = ScrollLeft
            IntellisenseList.ColumnWidths = "400;1600"
        End If
    End Sub

    Private Sub GetAllProcedures()

        Dim WB As Workbook
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        Dim CodeMod As VBIDE.CodeModule
        
        Dim CurrentRow As String
        Dim StartPoint As Long
        Dim EndPoint As Long
        Dim TempArg() As String
        Dim AsProcedure As String
        Dim CurrentComponent As Long

        Dim CurrentProcedurePos As Long
        Dim ReturnType As String
        Dim Name As String
        Dim Arguments As String

        Dim i As Integer
        Dim j As Integer
        Dim k As Long
        Dim Temp As Long

        ' This is to get the second last space, which indicates, that its the startpoint for the returntype
        Dim TempArray() As String
        k = -1

        For Each WB In Workbooks
            Set VBProj = WB.VBProject
            k = k + 1
            For Each VBComp In VBProj.VBComponents
                Set CodeMod = VBComp.CodeModule
                CurrentComponent = AddNode(k, 0, VBComp.Name)
                Select Case VBComp.Type
                    Case vbext_ComponentType.vbext_ct_StdModule   : Call AddNode(k, CurrentComponent, "As Module")
                    Case vbext_ComponentType.vbext_ct_ClassModule : Call AddNode(k, CurrentComponent, "As Class")
                    Case vbext_ComponentType.vbext_ct_MSForm      : Call AddNode(k, CurrentComponent, "As Form")
                End Select
                Call AddNode(k, CurrentComponent, "No Arguments")
                Call AddNode(k, CurrentComponent, "No Value")
                For i = 1 To CodeMod.CountOfLines
                    CurrentRow = CodeMod.Lines(i, 1)
                    AsProcedure = ""
                    If UCase(CurrentRow) Like "*PUBLIC *" And InStr(1, CurrentRow, "'") = 0 And Not UCase(CurrentRow) Like "*" & Chr(34) & "*PUBLIC*" & Chr(34) & "*" Then
                        If (UCase(CurrentRow) Like "* FUNCTION *" Or UCase(CurrentRow) Like "* SUB *" Or UCase(CurrentRow) Like "* SET *" Or UCase(CurrentRow) Like "* LET *" Or UCase(CurrentRow) Like "* GET *") Then
                            ' A Procedure
                            '                          |----------|
                            '   Public Static Function VariableName(Arg1 As Variant, Arg2 As Variant) As Variant
                            AsProcedure = "Public Procedure "
                            Select Case True
                                Case UCase(CurrentRow) Like "*PUBLIC STATIC SUB *(*)*"      : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC STATIC SUB ")      + Len("PUBLIC STATIC SUB ")      : EndPoint = InStr(1, UCase(CurrentRow), "("): Name = Mid(CurrentRow, StartPoint, EndPoint - StartPoint)
                                Case UCase(CurrentRow) Like "*PUBLIC SUB *(*)*"             : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC SUB ")             + Len("PUBLIC SUB ")             : EndPoint = InStr(1, UCase(CurrentRow), "("): Name = Mid(CurrentRow, StartPoint, EndPoint - StartPoint)
                                Case UCase(CurrentRow) Like "*PUBLIC STATIC FUNCTION *(*)*" : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC STATIC FUNCTION ") + Len("PUBLIC STATIC FUNCTION ") : EndPoint = InStr(1, UCase(CurrentRow), "("): Name = Mid(CurrentRow, StartPoint, EndPoint - StartPoint)
                                Case UCase(CurrentRow) Like "*PUBLIC FUNCTION *(*)*"        : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC FUNCTION ")        + Len("PUBLIC FUNCTION ")        : EndPoint = InStr(1, UCase(CurrentRow), "("): Name = Mid(CurrentRow, StartPoint, EndPoint - StartPoint)
                                Case UCase(CurrentRow) Like "*PUBLIC PROPERTY GET *(*)*"    : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC PROPERTY GET ")    + Len("PUBLIC PROPERTY GET ")    : EndPoint = InStr(1, UCase(CurrentRow), "("): Name = Mid(CurrentRow, StartPoint, EndPoint - StartPoint)
                                Case UCase(CurrentRow) Like "*PUBLIC PROPERTY SET *(*)*"    : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC PROPERTY SET ")    + Len("PUBLIC PROPERTY SET ")    : EndPoint = InStr(1, UCase(CurrentRow), "("): Name = Mid(CurrentRow, StartPoint, EndPoint - StartPoint)
                                Case UCase(CurrentRow) Like "*PUBLIC PROPERTY LET *(*)*"    : StartPoint = InStr(1, UCase(CurrentRow), "PUBLIC PROPERTY LET ")    + Len("PUBLIC PROPERTY LET ")    : EndPoint = InStr(1, UCase(CurrentRow), "("): Name = Mid(CurrentRow, StartPoint, EndPoint - StartPoint)
                                Case Else
                            End Select
                            '                                       |------------------------------|
                            '   Public Static Function VariableName(Arg1 As Variant, Arg2 As Variant) As Variant
                                StartPoint = InStr(1, CurrentRow, "(")
                                EndPoint   = InStr(1, CurrentRow, ")")
                                If StartPoint + 1 <> EndPoint Then
                                    TempArg = Split(Mid(CurrentRow, StartPoint + 1, EndPoint - StartPoint - 1), ",")
                                    Arguments = TempArg(0)
                                    For j = 1 To UboundK(TempArg)
                                        Arguments =  Arguments & ArgSeperator & TempArg(j)
                                    Next
                                End If
                        Else
                            ' A Variable
                            '          |----------|
                            '   Public VariableName As Variant
                            Select Case True
                                Case UCase(CurrentRow) Like "*PUBLIC CONST *": StartPoint = InStr(1, UCase(CurrentRow), "*PUBLIC CONST *") + Len("PUBLIC CONST "): EndPoint = InStr(1, UCase(CurrentRow), " AS "): Name = Mid(CurrentRow, StartPoint + 1, EndPoint - StartPoint - 1)
                                Case UCase(CurrentRow) Like "*PUBLIC *":       StartPoint = InStr(1, UCase(CurrentRow), "*PUBLIC *")       + Len("PUBLIC "):       EndPoint = InStr(1, UCase(CurrentRow), " AS "): Name = Mid(CurrentRow, StartPoint + 1, EndPoint - StartPoint - 1)
                                Case Else
                            End Select
                        End If
                            '                                                                        |---------|
                            '   Public Static Function VariableName(Arg1 As Variant, Arg2 As Variant) As Variant
                            TempArray = Split(CurrentRow, " ")
                            ReturnType = AsProcedure & TempArray(UboundK(TempArray) - 1) & " " & TempArray(UboundK(TempArray))
                            ' If last character is ")", then it returns void
                            If Mid(TempArray(UboundK(TempArray)), Len(TempArray(UboundK(TempArray))), 1) = ")" Then ReturnType = AsProcedure & "As Void"
                            CurrentProcedurePos = AddNode(k,  CurrentComponent, Name)
                            Call AddNode(k,  CurrentProcedurePos, ReturnType)
                            Call AddNode(k,  CurrentProcedurePos, Arguments)
                            Call AddNode(k,  CurrentProcedurePos, "No Value")
                    End If
                    NoLines:
                Next
            Next
        Next
                
    End Sub
    
    Private Sub Close_IntelliSenseList()
        IntelliSenseList.Clear
        IntelliSenseList.Visible = False
        ConsoleText.SetFocus
        ConsoleText.SelStart = Len(ConsoleText.Text)
    End Sub

    Private Sub SetUp_IntelliSenseList(Text As String)

        Dim Words() As String
        Dim Index As Long
        Dim i As Long, j As Long, k As Long
        Dim NameSpaces() As Variant
        Dim CurrentComponent As Long


        Words = Split(Text, ".")
        IntelliSenseList.Clear
        If UboundK(Words) = -1 Then Exit Sub
        For i = 3 To UboundK(Trees(ConsVarIndex).Nodes(0).Branches)
            Index = Trees(ConsVarIndex).Nodes(0).Branches(i)
            Call AddArray(NameSpaces, RetVarSin(ConsVarIndex, Index, -1))
            Call AddArray(NameSpaces, RetVarSin(ConsVarIndex, Index, 0))
            Call AddArray(NameSpaces, RetVarSin(ConsVarIndex, Index, 1))
        Next i

        Select Case True
            Case UboundK(Words) = 0
                For i = 0 To ConsVarIndex - 1
                    Call AddArray(NameSpaces, RetVarSin(i, 0, -1))
                    Call AddArray(NameSpaces, RetVarSin(i, 0, 0))
                    Call AddArray(NameSpaces, RetVarSin(i, 0, 1))
                Next i
                For i = 0 To ConsVarIndex - 1
                    For j = 3 To UboundK(Trees(i).Nodes(0).Branches)
                        Index = Trees(i).Nodes(0).Branches(j)
                        If RetVarSin(i, Index, 0) = "As Module" Then
                            For k = 3 To UboundK(Trees(i).Nodes(Index).Branches)
                                CurrentComponent = Trees(i).Nodes(Index).Branches(k)
                                Call AddArray(NameSpaces, RetVarSin(i, CurrentComponent, -1))
                                Call AddArray(NameSpaces, RetVarSin(i, CurrentComponent, 0))
                                Call AddArray(NameSpaces, RetVarSin(i, CurrentComponent, 1))
                            Next k
                        End If
                    Next j
                Next i

            Case UboundK(Words) = 1
                For i = 0 To ConsVarIndex - 1
                    If Words(0) = CStr(Trees(i).Nodes(0).Value) Then Exit For
                Next i
                If i >= ConsVarIndex Then
                    Exit Sub
                Else
                    For j = 3 To UboundK(Trees(i).Nodes(0).Branches)
                        Index = Trees(i).Nodes(0).Branches(j)
                        Call AddArray(NameSpaces, RetVarSin(i, Index, -1))
                        Call AddArray(NameSpaces, RetVarSin(i, Index, 0))
                        Call AddArray(NameSpaces, RetVarSin(i, Index, 1))
                    Next j
                End If

            Case UboundK(Words) = 2
                For i = 0 To ConsVarIndex - 1
                    If Words(0) = CStr(Trees(i).Nodes(0).Value) Then Exit For
                Next i
                If i > ConsVarIndex Then
                    Exit Sub
                Else
                    For j = 3 To UboundK(Trees(i).Nodes(0).Branches)
                        Index = Trees(i).Nodes(0).Branches(j)
                        If RetVarSin(i, Index, -1) = Words(1) Then Exit For
                    Next j
                    If j > UboundK(Trees(i).Nodes(Index).Branches) Then
                        Exit Sub
                    Else
                        For k = 3 To UboundK(Trees(i).Nodes(Index).Branches)
                            CurrentComponent = Trees(i).Nodes(Index).Branches(k)
                            Call AddArray(NameSpaces, RetVarSin(i, CurrentComponent, -1))
                            Call AddArray(NameSpaces, RetVarSin(i, CurrentComponent, 0))
                            Call AddArray(NameSpaces, RetVarSin(i, CurrentComponent, 1))
                        Next k
                    End If
                End If

            Case Else
        End Select

        For i = 0 To UboundK(NameSpaces) Step 3
            If UboundK(Words) = -1 Then
                IntelliSenseList.AddItem
                IntelliSenseList.List(IntelliSenseList.ListCount - 1, 0) = CStr(NameSpaces(i))
                IntelliSenseList.List(IntelliSenseList.ListCount - 1, 1) = CStr(NameSpaces(i + 1)) & ", " & CStr(NameSpaces(i + 2))
            Else
                If CStr(NameSpaces(i)) Like Words(UboundK(Words)) & "*" Then
                    IntelliSenseList.AddItem
                    IntelliSenseList.List(IntelliSenseList.ListCount - 1, 0) = CStr(NameSpaces(i))
                    IntelliSenseList.List(IntelliSenseList.ListCount - 1, 1) = CStr(NameSpaces(i + 1)) & ", " & CStr(NameSpaces(i + 2))
                End If
            End If
        Next i
        IntelliSenseList.Visible = (IntelliSenseList.ListCount > 0)

    End Sub

    Private Sub AddArray(ByRef Arr As Variant, Value As Variant)
        ReDim Preserve Arr(UboundK(Arr) + 1)
        Arr(UboundK(Arr)) = Value
    End Sub

    Private Sub IntelliSenseList_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
        
        Dim Line As String
        Dim Word As String
        Dim Words() As String
        Dim Start As Long
        Dim ReturnString As String

        Line = GetLine(ConsoleText.Text, CurrentLineIndex)
        Words = Split(Line, ".")
        If UboundK(Words) = -1 Then Word = Line
        Word = Words(UboundK(Words))
        Words = Split(Word, " ")
        If UboundK(Words) <> -1 Then Word = Words(UboundK(Words))
        Select Case KeyCode
            Case vbKeyLeft
                Close_IntelliSenseList
                Exit Sub
            Case vbKeyRight
                If IntellisenseList.ListCount > 0 Then
                    ReturnString = IntelliSenseList.List(IntelliSense_Index, 0)
                    Start = InStr(1, Ucase(ReturnString), Ucase(Word))
                    If Start = 0 Then Start = 1
                    Call PrintConsole(Mid(ReturnString, Start + Len(Word), Len(ReturnString) - Len(Word)))
                    Close_IntelliSenseList
                    Exit Sub
                End If
            Case vbKeyUp
                Intellisense_Index = Intellisense_Index - 1
            Case vbKeyDown
                Intellisense_Index = Intellisense_Index + 1
        End Select
        If Intellisense_Index > IntelliSenseList.ListCount - 1 Then
            Intellisense_Index = 0
        ElseIf Intellisense_Index < 0 Then
            Intellisense_Index = IntelliSenseList.ListCount - 1
        Else

        End If
        If IntelliSenseList.ListCount > 0 Then IntelliSenseList.ListIndex = IntelliSense_Index

    End Sub
'

' stdLambda

    Private Function RunLambda(Lambda As Variant, Arguments As Variant) As Variant
        On Error GoTo Error
        Select Case UBoundK(Arguments)
            Case -1:   RunLambda = Lambda.Run()
            Case 00:   RunLambda = Lambda.Run(Arguments(0))
            Case 01:   RunLambda = Lambda.Run(Arguments(0), Arguments(1))
            Case 02:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2))
            Case 03:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3))
            Case 04:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4))
            Case 05:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5))
            Case 06:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6))
            Case 07:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7))
            Case 08:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8))
            Case 09:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9))
            Case 10:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10))
            Case 11:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11))
            Case 12:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12))
            Case 13:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13))
            Case 14:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14))
            Case 15:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15))
            Case 16:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16))
            Case 17:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17))
            Case 18:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18))
            Case 19:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19))
            Case 20:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20))
            Case 21:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21))
            Case 22:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22))
            Case 23:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23))
            Case 24:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24))
            Case 25:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25))
            Case 26:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26))
            Case 27:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26), Arguments(27))
            Case 28:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26), Arguments(27), Arguments(28))
            Case 29:   RunLambda = Lambda.Run(Arguments(0), Arguments(1), Arguments(2), Arguments(3), Arguments(4), Arguments(5), Arguments(6), Arguments(7), Arguments(8), Arguments(9), Arguments(10), Arguments(11), Arguments(12), Arguments(13), Arguments(14), Arguments(15), Arguments(16), Arguments(17), Arguments(18), Arguments(19), Arguments(20), Arguments(21), Arguments(22), Arguments(23), Arguments(24), Arguments(25), Arguments(26), Arguments(27), Arguments(28), Arguments(29))
            Case Else: RunLambda = "Too many Arguments": LastError = 3
        End Select
        Exit Function
        Error:
        RunLambda = "could not run lambda successfully"
    End Function

    Private Function CreateLambda(Name As String, Optional Equation As Variant = "", Optional UsePerformanceCache As Boolean = False, Optional SandboxExtras As Boolean = False) As Boolean
        Dim Temp As stdLambda
        Set Temp = stdLambda.Create(Equation, UsePerformanceCache, SandboxExtras)
        Call SetVariable(ConsVarIndex, 0, Name, "Public ConsoleVariable As stdLambda", "Arguments might exist", Temp)
        CreateLambda = True
    End Function

    ' Lambda needs to be of Type stdLambda. It wanst specified, so that Console works without stdLambda dependency
    Public Function LoadLambda(Name As String, Lambda As Variant) As Boolean
        Dim NewNode As Node
        If stdLambda_Active = False Then Call PrintEnter("stdLambda is deactivated", in_System): LoadLambda = False: Exit Function
        Call SetVariable(ConsVarIndex, 0, Name, "Public ConsoleVariable As stdLambda", "Arguments might exist", Lambda)
        LoadLambda = True
    End Function

    ' Load a custom array to Lambdas 2D (1,n)
    Public Function LoadLambdas(Arr() As Variant) As Boolean

        Dim i As Integer
        Dim NoOfDimenions As Integer
        Dim Temp As Integer
        On Error GoTo Error

        If stdLambda_Active = False Then Call PrintEnter("stdLambda is deactivated", in_System): LoadLambdas = False: Exit Function
        For i = 1 To 255
            Temp = UboundK(Arr, i)
            NoOfDimenions = i
        Next
        Error:
        If NoOfDimenions = 2 Then
            If UboundK(Arr, 1) = 1 Then
                For i = 0 To UboundK(Arr, 2)
                    Call SetVariable(ConsVarIndex, 0, Arr(0, i), "Public ConsoleVariable As stdLambda", "Arguments might exist", Arr(1, i))
                Next i
                LoadLambdas = True
            Else
                LoadLambdas = False ' Wrong Number of Elements for 1st dimension
            End If        
        Else
            LoadLambdas = False ' Wrong Number of dimensions
        End If

    End Function

    Public Function BindGlobal(LambdaName As String, Name As String, Variable As Variant) As Boolean
        Dim Index As Long
        Dim Pos(0) As Long
        If stdLambda_Active = False Then Call PrintEnter("stdLambda is deactivated", in_System): Exit Function
        Index = FindNode(ConsVarIndex, Pos, LambdaName)
        If Index <> -1 Then
            VBProject(ConsVarIndex).Nodes(Index).Value.BindGlobal Name, Variable 
            BindGlobal = True
        Else
            Call PrintEnter("Couldnt Bind Variable globally to " & LambdaName)
        End If
    End Function

    Public Function Bind(LambdaName As Variant, ParamArray Arguments() As Variant) As Boolean
        Dim Args() As Variant
        Dim Index As Long
        Dim Pos(0) As Long
        If stdLambda_Active = False Then Call PrintEnter("stdLambda is deactivated", in_System): Exit Function
        Args = Arguments
        Index = FindNode(ConsVarIndex, Pos, LambdaName)
        If Index <> -1 Then
            Set Lambdas(1, Index) = VBProject(ConsVarIndex).Nodes(Index).Value.BindEx(Args)
            Bind = True
        Else
            Call PrintEnter("Couldnt Bind Arguments to " & LambdaName)
        End If
    End Function

    ' Lambda needs to be of Type stdLambda. It wanst specified, so that Console works without stdLambda dependency
    Public Function GetLambda(Name As Variant) As Variant
        Dim Index As Long
        Dim Pos(0) As Long
        If stdLambda_Active = False Then Call PrintEnter("stdLambda is deactivated", in_System): GetLambda = Nothing: Exit Function
        Index = FindNode(ConsVarIndex, Pos, Name)
        If Index <> -1 Then Set GetLambda = VBProject(ConsVarIndex).Nodes(Index).Value
    End Function

    Public Function GetLambdas() As Variant
        If stdLambda_Active = False Then Call PrintEnter("stdLambda is deactivated", in_System): GetLambdas = Nothing: Exit Function
        Set GetLambdas = Lambdas
    End Function
'