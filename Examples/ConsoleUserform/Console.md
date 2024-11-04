# `Console.frm`

## Introduction

### What is the VBA Console?
The main idea behind the console is to let the user decide which procedures to run and to work with basic script at runtime.


### Preparation
To Prepare the console you have to do the following:  
If you want to activate Intellisense then go to the variable `Intellisense_Active` and set it to `true`.  
To further Improve Intellisense enable `Microsoft Visual Studio Extensebility 5.3` as reference.    
If you want to activate stdLambda function go to the variable `stdLambda_Active` and set it to `true`. It is dependant sancarns stdVBA class `stdLambda`  
https://github.com/sancarn/stdVBA/tree/master  

Run `Console.Show`  
    This will initialize the console  

Now the Console can be used in process  

### How to use is
The Console sees the current Line as the code that needs to be run.  
To write several line see below.  
To pass arguments you have to write it like in vba: FunctionName(Arg1, Arg2, ArgN)
The Console interprets the line as a Leftside and a Rightside. Leftside is the destination-Location, like a variable or the Console itself. The Rightside is the Code that will be interpreted.  


1. Password Protection
The Console is Password protected:  
If no Password is assigned then it is ""
Changing it only works from outside the Console through its Public Function

2. Run Macros with and without arguments:  
```vb
    HelloWorld
    HelloWorld()
    PrintToDebug("TextAsString")
    Add(Arg1, Arg2)
```

3. Create and run variables:  
```vb
    Dim x As Long       'Will remove this Function at the End of its Scope. More on that furher down
    Private x As Long   'Can only be used by the Console
    Public x As Long    'This can be returned outside of the Console
    x = 3               ' Assumes Variant
    x = "3"             ' Assumes Variant, but only gets the 3
    x = ?1+2            ' Creates Lambda, that when run returns 1+2
    x = ?$1+$2          ' Creates Lambda, that adds two arguments together. More on Lambdas see: https://github.com/sancarn/stdVBA/tree/master
    x = y               ' Assigns Value of another Variable
    x = Add(Arg1, Arg2) ' Assigns Value of Procedure/Lambda
    x = y==z            ' Compare 2 Values
    x = y+z             ' Adds two Values and returns it to x
    x = y++2            ' Increments y by 2 and returns its value to x
```

4. Loops and Conditions
```vb
    ' First arguments will run once, second is the Break-Condition, third will be run at the end of each cycle. All following arguments will be run for each cycle
    for(Dim x as Long, x==30, x++1, HelloWorld)
    ' First argument is the breakcondition for both until and while loops. All following Arguments will be repeated every cycle
    until(x==3, HelloWorld, x++1)
    while(x<>3, HelloWorld, x++1)
    ' if and select statements need their case-sensitive keywords
    if(x==y Then x++1, x Else y--1, y)
    select(x Case ==0, x++1, Case >=2, x--1, Case <>2, x Case Else, x**2)
```
5. Multiline
Those long statements will get pretty wild pretty soon. For that there is the `Multiline` and `EndMultiline` statements.  
As long as `Endmultiline` is not written every enter will insert a new line without executing it.  
Once `Endmultiline` is written and executed all previous lines up to `Multiline` will be combined into 1 Line and executed.  
6. Script
Scripts are written with `Script` and `EndScript`.  
Same as `Multiline`, but here the Lines will not get saved as one, but every line is a command for the Console.  
To still write several lines, that get compressed to one you have to write the `_` Character last and then press Enter.  
```vb
ScriptName(Arg1 As Long, Arg2 As Long)           ' Needs to be first line
Dim x As Long                                    ' Exists till the End of the Script
Dim Lam As stdLambda
Lam = ?$1*$2
Until(x==30, _
Lam(Arg1+x, Arg2+x)_
x++1)
If(Arg1==1 Then _
x = 2 _
Else _
x = 3)
EndScript
```
As of now, Intendation is not supported
7. Dim
Dim Variables will run out of Scope and be deleted. That means:  
When Pressing Enter: It will be deleted immediatly  
When in a Loop     : Once the Loop is finished
When in a Condition: Once the Condition is finished
When in a Script   : Once the Script is finished

8. Special Keyword
Help      : Prints the Help Text
Clear     : Deletes Text 
Multiline : See Point 5
Info      : Prints the Last Error
Exit      : Hides the Console and Clears the Text (Will currently kill Excel)
Script    : See Point 6

9. Public Methods
```vb
    ' Passes Command to HandleCode and executes it
    Public Function Execute(Command As String) As Variant

    ' Gets UserInput of specific Type. In Normal Case it takes Variant
    Public Function GetUserInput(Message As Variant, Optional InputType As Long = 12) As Variant

    ' Gets UserInput, but only allows the PreDeclared Answers of AllowedValues, Answers will run if defined.
    ' Answers, when defined need to be of same size as AllowedValues
    Public Function CheckPredeclaredAnswer(Message As Variant, AllowedValues As Variant, Optional Answers As Variant = Empty) As Variant

    ' Adds vbCrLf to Text and Calls PrintConsole
    Public Sub PrintEnter(Text As Variant, Optional Colors As Variant, Optional ColorLength As Variant)

    ' Prints Text to Control
    ' Color may be 1 Value or an Array of Values
    ' ColorLength is an Array of Long, that will color ColorLength(i) of Characters in Color(i) 
    Public Sub PrintConsole(Text As Variant, Optional Colors As Variant, Optional ColorLength As Variant)

    ' Will add a Script to ConsVarIndex
    Public Sub AddScript(Name As String, Arguments As String, Script As String)

    ' Will set a new Password if you know the old one
    Public Function Password(Old_Password As String, New_PassWord As String) As Boolean

    ' Returns the Value of a Public Variable from the Console
    ' If that Variable is NOT Public it will return Nothing
    Public Function GetPublicVariable(Name As String) As Variant
```

## Code Documentation


### Privately declared Variables

```vb
' Private Variables
    
    Private       CurrentLineIndex      As Long                     ' Index of the CurrentLine
    Private Const Recognizer            As String = "\>>>"          ' Used to recognize when CurrentLine should start
    Private Const ArgSeperator          As String = ", "            ' Used to seperate Arguments
    Private Const AsgOperator           As String = " = "           ' Used to assign Values to a Variable
    Private Const LineSeperator         As String = "LINEBREAK/()\" ' Used by Script to recognize Lines, because vbcrlf didnt work for some reason
    Private       PasteStarter          As Boolean                  ' Determines if starter should be printed or not
    Private       UserInput             As Variant                  ' Value the user put in. Needs to be private for DoEvents
    Private       LastError             As Variant                  ' Saves the last recognized Error, gets printed with "Info"

    Private       Intellisense_Index    As Long                     ' Used for to search the Intellisense-Listbox with up and down buttons
    Private       ConsVarIndex          As Long                     ' Will always be the highset number for trees, as it will be defined by the amount of open Workbooks for their VBAProjects For better visibility than UboundK(Trees)

    Private       PreviousCommands()    As Variant                  ' Saves all lines in an array entered
    Private       PreviousCommandsIndex As Long                     ' Used to traverse PreviousCommands() with up and down button

    Private       p_Password            As String                   ' Password
    Private       PasswordActive        As Boolean                  ' Used to check if password was already entered
    Private       PasswordMode          As Boolean                  ' When true only "*" will be printed, until the password was entered

    Private       DimVariables(255)     As Long                     ' Maximum amount of dim-Variables. Fixed Size because i was too lazy to implement a dynamic solution :)
    Private       DimIndex              As Long                     ' Currently highest Number for DimVariables


    
    Private WorkMode As Long                                        ' Used to determine what the Console awaits from the user
    Private Enum WorkModeEnum
        Logging = 0                                                 ' Logging is the basic one, where the console only recieves information / runs code
        UserInputt = 1                                              ' UserInputt is variable input of the user  
        MultilineMode = 2                                           ' MultiLineMode gets activated when running MultiLine
        ScriptMode = 3                                              ' Script gets activated when running Script 
    End Enum

    
    
    
    



    Private Const Intellisense_Active As Boolean = True              ' Will Activate Intellisense

    Private Const Extensebility_Active As Boolean = True             ' Dependenant on Microsoft Visual Studio Extensebility 5.3     Will include all Public Variables and Procedures defined in all Projects

    Private Const stdLambda_Active    As Boolean = True              ' Dependant on sancarn´s stdLambda class, will allows to defined stdLambdas in the Console

    'Intellisense Colors
    Private in_Basic       As Long
    Private in_System      As Long
    Private in_Procedure   As Long
    Private in_Operator    As Long 'smooooooth operatooooor
    Private in_Datatype    As Long
    Private in_Value       As Long
    Private in_String      As Long
    Private in_Statement   As Long
    Private in_Keyword     As Long
    Private in_Parantheses As Long
    Private in_Variable    As Long
    Private in_Script      As Long
    Private in_Lambda      As Long
```


### Intellisense
```vb

    ' Will Adjust ListBox And ScrollValues of the Console According to the Currentline and Character in said line
    Private Sub SetPositions()

    ' Will Get all Public´s in all Components of All open VBAProjets
        ' Runs once by initialization
    Private Sub GetAllProcedures()
    
    ' Closes the Listbox and Clears it
    Private Sub Close_IntelliSenseList()

    ' Will Show all hits according to current depth.
        ' Runs every KeyUp event that is not Enter or Arrowkeys
        ' When there are no "." in the Text than it will assume VBAProjects, Public Module-Procedures and ConsoleVariables
        ' one "." will assume as VBAProject beforehand and search that VBProjects for Components
        ' Seconde "." will assume a Component beforehand and search that Component for a Public
    Private Sub SetUp_IntelliSenseList(Text As String)

    ' Only used by SetUp_IntelliSenseList to add a new Value to the Array that will be searched of the current Depth
    Private Sub AddArray(ByRef Arr As Variant, Value As Variant)

    ' Will be entered when pressing [RIGHT] key when ConsoleText.SelStart is at the last Character. 
        ' Once entered there are basically only 4 buttons:
        ' [RIGHT] will get the current selection to the Console
        ' [LEFT]  will cancel intellisense
        ' [UP]    will increment IntellisenseIndex down(which will be shown as going one up) (will loop over when too low)
        ' [DOWN]  will increment IntellisenseIndex up(which will be shown as going one down) (will loop over when too high)
    Private Sub IntelliSenseList_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

```

### Color

```vb
' Coloring

    ' Cannot Assign the Colors as Const, because it would result in a negative number 
    Private Sub AssignColor()

    ' Will look into Words and Characters
        ' Depending on What the Word is it will be colored differently (also checks those words in this Sub)
    Private Sub ColorWord()

        in_Basic       = ' everytime when nothing is defined
        in_System      = ' used to show what text came from the Console
        in_Procedure   = ' when it recognizes it in a Component
        in_Operator    = ' when it is a operator
        in_Datatype    = ' when the previous word is "AS"
        in_Value       = ' when it is a number
        in_String      = ' when it is inbetween "
        in_Statement   = ' when it is a vba statement
        in_Keyword     = ' when it is a vba keyword
        in_Parantheses = ' when it is ( or )
        in_Variable    = ' when it is a ConsoleVariable or Public Variable from a Component
        in_Script      = ' when it is a defined script
        in_Lambda      = ' when it is a defined stdLambda

    ' For the Save of Coloring all Operators will be removed from the Line to better Color Values and Procedures
    Private Function RemoveOperators(Text As String) As String
```


### Tree

```vb

    'Basic Design of Trees()


    '             [Trees(0)]                                         #                  [Trees(ConsVarIndex)]
    '             VBAProject                                         #                     ConsoleVariables
    '  _________/    |  \_________ _____________                     #     ________________/      \
    '  /             |            \             \                    #    /       /       /        \
    '  As VBProject  No Arguments  No Values   [Components]          #    [Type]  [Args]  [Value]  [Variables]
    '                                            Console             #                                  Line
    '         _________________________________ /  \                 #     ____________________________/    \____
    '        /             /             /          \                #    /          /             /             \
    '        As Component  No Arguments  No Values  [Public´s]       #    As String  No Arguments  "Hello World"  []
    '                                                 Execute        #
    '            ______________________________________/     \___    #
    '           /                     /               /          \   #
    '           Procedure As Variant  Line As String  No Values  []  #
    



    ' A Tree is defined as an Array of Nodes, whereas Node is defined as a Value with a "Pointer" to an Index of the Tree 
    Private Type Node
        Value As Variant
        Branches() As Long
    End Type

    Private Type cCollection
        Nodes() As Node
    End Type

    ' This is an Array of Trees
    Private Trees() As cCollection

    ' Which Subtree is what at any given depth
    Private Enum NodeType
        e_ReturnType = 0
        e_Arguments = 1
        e_Value = 2
        e_Branches = 3
    End Enum

    '######################################################
    ' TreeIndex is always the Index for Trees(TreeIndex)
    '######################################################



    ' Get the Index of the Array defined by its Path.
        ' Like (4-0) will get the ReturnType of the 4th SubTree of the TreeIndex
        ' Returns the Index, else -1
    Private Function GetNode(TreeIndex As Long, Positions() As Long) As Long

    ' Searches a TreeIndex for a Value by specified Position
        ' Returns the Index, else -1
        ' Searches inbetween the Branches, else all Branches
    Private Function FindNode(TreeIndex As Long, Positions() As Long, Value As Variant, Optional StartBranch As Long = 0, Optional EndBranch As Long = -1) As Long

    ' Finds the Depth of a Tree recursively
    Private Sub FindDepth(ByVal TreeIndex As Long, ByVal CurrentNode As Long, ByVal CurrentDepth As Long, ByRef MaxDepth As Long)

    ' adds new Node to defined Index as Subtree
    Private Function AddNode(TreeIndex As Long, NodeIndex As Long, Value As Variant) As Long

    ' Deletes Node as defined Index and all its SubTrees
    Private Function DeleteNode(TreeIndex As Long, NodeIndex As Long) As Long

    ' Adds/Changes a Variable to a Tree at defined Index.
    Private Function SetVariable(TreeIndex As Long, SearchIndex As Long, Name As String, ReturnType As String, Arguments As String, Value As Variant) As Long

    ' Will Return a Variable, depending on Arguments and ReturnValue it might return its Value or the Object
    Private Function ReturnVariable(TreeIndex As Long, Positions() As Long, Optional ReturnValue As Boolean = False, Optional Arguments As Variant) As Variant

    ' Will Initialize Trees() to all VBAProjects and ConsVarIndex
    Private Sub InitializeTree()

    ' Returns either Name, Type, Arguments, Value or Branches of a Position
    Private Function ReturnVariableValue(TreeIndex As Long, Positions() As Long, ReturnWhat As Long) As Variant

    ' Calls ReturnVariableValue, just with a single Index as Position
    Private Function RetVarSin(TreeIndex As Long, Position As Long, ReturnWhat As Long) As Variant

    ' Will find All Values in a Tree recursively
    Private Function FindNodeAll(TreeIndex As Long, Value As Variant, ByRef Positions() As Long) As Long

```

### Interpretation

```vb

    ' Only used to put the Cursor to the last Character before the Line is run
    Private Sub ConsoleText_KeyDown(pKey As Long, ByVal ShiftKey As Integer)
    
    ' Only used for PasswordInput to get the real characters before the become "*"
    Private Sub ConsoleText_KeyPress(Char As Long)

    ' Used to check what to do with the Currentbutton
        ' Enter, Up, Down and Else
    Private Sub ConsoleText_KeyUp(pKey As Long, ByVal ShiftKey As Integer)

    ' Interprets the Enter Key depending on Workmode
    Private Function HandleEnter() As Variant

    ' Splits the Line in Leftside and Rightside and tries to interpret it.
        ' Checks if Leftside exists and changes the Value to it/ creates new Variable
        ' Rightside could be Script, Lambda, Number, Data, String, Procedure, Variable, Condition
            ' Will check for what it is and run its code
    Private Function HandleCode(Line As String) As Variant

    ' Handles special keywords and runs their Code
        ' If, Select, For, Until, While, Help, Info etc.
    Private Function HandleSpecial(Line As String) As Variant

    ' Returns a Text depending on LastError
    Private Function HandleLastError() As String

    ' Clears Text of Console and sets it up, return " "
    Private Function HandleClear() As String

    ' Returns Help-text
    Private Function HandleHelp() As String

    ' Tries to run a Procedure with Passed arguments, if failed it will return an Error-text
    Private Function RunApplication(Name As String, Arguments As Variant) As Variant

    ' Checks what Character it is.
        ' Needs to be static to catch uppercase key
        ' SetFocus on ListBox when right and last character 
    Private Static Function HandleOtherKeys(pKey As Long, ByVal ShiftKey As Integer) As String

    ' Will look at every Arguments and try to interpret it
        ' Does this recursively to catch procedures and Functions in Functions
    Private Sub RecursiveReturnVariable(ByRef Arguments() As Variant)

    ' Will look for all Operations in a Line and handle the Operations
    Private Function HandleReturnOperator(Line As String) As Variant

    ' will try to calculate the operation of 2 Values
    Private Function HandleCalcOperator(ByRef Value1 As Variant, Operator As Variant, Value2 As Variant) As Variant

    ' Will check if Password is activated or if the UserInput is the Password
    Private Function HandlePassword() As Boolean

    ' Will Run a Script with defined Arguments line by line (handles DimVariables)
    Private Function RunScript(Script As Variant, ScriptArgs As Variant, Optional Arguments As Variant) As Variant

    ' Will try to change the Variant datatype to a interpretable DataType like Numeric
    Private Function InterpretVariable(Value As Variant) As Variant

    ' Checks if Variable exists
        ' If yes it will return the Variable, else it will return InterpretVariable
    Private Function InterpretVariableTEMP(Value As Variant) As Variant

    ' Converts Value to defined Type
    Private Function GetVariableByType(Value As Variant, DataType As String) As Variant

    ' Will Handle For, Until and While (handles DimVariables)
    Private Function HandleLoop(Line As String) As Variant

    ' Will Handle If and Select (handles DimVariables)
    Private Function HandleCondition(Line As String) As Variant

    ' Creates a new Variable without a Value
    Private Function HandleNewVariable(Line As String) As Variant

    ' When HandleNewVariable is a Dim Variable, it will add it to DimVariables
        ' If DeleteNodeIndex = -1 Then it will Delete the defined Node, else it will at the NodeIndex
    Private Function HandleDimVariable(Optional NodeIndex As Long = -1, Optional DeleteNodeIndex As Long = -1) As Long

    ' Deletes all Variables from DimVariables defined by the Input-Array
    Private Sub DeleteScope(Arr() As Long)

    ' Sets all Indices to -1.
        ' -1 means no Index for DimVariables
    Private Sub InitScope(ByRef Arr() As Long)
```


### Functions
```vb
' Get/Set Values

    ' The highest ConsoleText.SelStart <> Len(ConsoleText.Text)
    ' For that reason this Function will return the highest ConsoleText.SelStart 
    Private Function GetMaxSelStart() As Long

    ' Will Print Workbook directory (if existing) and the recognizer
    Private Function PrintStarter() As Variant

    ' The Text printed at Initialization
    Private Function GetStartText() As String

    ' Will get Charlength of a specified amount of lines.
    ' Mainly used to find the startpoint for coloring and PreviousCommands
    Private Function GetTextLength(Text As String, Seperator As String, Optional IndexBreakPoint As Long = -2) As Long

    ' Gets Text to the right side of the recognizer at any given Line of ConsoleText.text
    Private Function GetLine(Text As String, Index As Long) As String

    ' Gets Words of a Line with Index
    ' Mainly used for coloring and Intellisense
    Private Function GetWord(Text As String, Optional Index As Long = -1) As String

    ' Basically just Split(), but will return an empty array if it cant split the Text
    ' Used when a Split is needed for example see GetFunctionArgs
    Private Function SplitString(Text As String, SplitText As String) As String()

    ' Used to Keep CurrentLineIndex up to date
    Private Sub SetUpNewLine()

    ' Will find all Positions of all Occurences of a String in a Text between defined Points
    ' One Element with Value 0 means nothing was found
    Private Function InStrAll(Text As String, SearchText As String, Optional StartIndex As Long = 1, Optional EndIndex As Long = 0, Optional StartFinding As Long = 0, Optional ReturnCount As Long = 255, Optional Line As Long = 0, Optional BreakText As String = Empty) As Long()

    '                        |--|  |--------|                  Returns Array of Length -1                    
    ' ExampleString=Function(Arg1, Arg2(Arg3))                 ExampleString=Function()
    Private Function GetFunctionArgs(Line As String) As Variant()

    '               |------|                        AS TREE POSITION LIKE eg. 1, 3
    ' ExampleString=Function(Arg1, Arg2(Arg3))
    Private Function GetFuncTreePosition(Line As String) As Long()

    '               |------|
    ' ExampleString=Function(Arg1, Arg2(Arg3))
    Private Function GetFunctionName(Line As String) As Variant

    ' Mid, but with EndPoint instead of Length
    Private Function MidP(Text As String, StartPoint As Long, EndPoint As Long) As String

    '                        |--------------|
    ' ExampleString=Function(Arg1, Arg2(Arg3))
    Private Function GetParanthesesText(Line As String) As String

    ' Searches Text for """ at the defined Points and return true if it is a String
    Private Function InString(Text As String, StartPoint As Long, EndPoint As Long) As Boolean

    ' Will Split all Elements down to a big array with the pattern:    Variable1, Operator2, Variable2, Operator2  ...
    ' Used to pass it to HandleReturnOperators
    Private Function GetAllOperators(Variable() As Variant) As Variant()
```


### Array

```vb

    '        ___________________________    
    '        |                          |__________|
    ' (1, 4, [], 3, 7, 4, 7, 2)         (1, 2, 7, 3)
    Private Sub MergeArray(ByRef Goal() As Variant, Adder() As Variant, Position As Long)

    '      ___________________________    
    '      |                          |__________|
    ' (1, [4], 3, 7, 4, 7, 2)         (1, 2, 7, 3)
    Private Sub ReplaceArrayPoint(ByRef Goal() As Variant, Adder() As Variant, Position As Long)

    '      _________________________
    '     |                         |______|
    ' (1, 4, 3, 7, 4, 7, 2)         (remove)
    Private Sub StitchArray(ByRef Arr() As Variant, StartPosition As Long, EndPosition As Long)

    '      _________________________________________________
    '     |      |      |      |      |      |              |_|
    ' (1, [], 4, [], 3, [], 7, [], 4, [], 7, [], 2)         (+)
    Private Sub InsertElements(ByRef Goal() As Variant, Value As Variant)

    ' Returns Ubound, if not possible returns -1
    ' This is needed, because some Variant may be an Array, but dont need to be. In that Case Ubound would throw an Error
    Private Function UboundK(Arr As Variant) As Long

    ' Same reason as UbounK, but just for cCollection, because UDT dont work with Variant
    Private Function UboundN(Arr As cCollection) As Long

    '                        ____________    
    '                       |           |_|
    ' (1, 4, 3, 7, 4, 7, 2, [])         (6)
    Private Sub PushArray(Byref Arr As Variant, Value As Variant)
```

The other Code as of now should be understandable