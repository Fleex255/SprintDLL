Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Runtime.CompilerServices
Imports System.Management
Public Class Workspace
    Public Slots As New Dictionary(Of String, Slot)
    Public BuildingAssembly As AssemblyBuilder
    Public LastError As Integer
    Public RunningScripts As New Stack(Of Script)()
End Class
Public Class Script
    Public ScriptSource As String
    Public CurrentInstructionText As String
    Public CurrentInstructionId As Integer
End Class
Module Main
    Sub Main()
        Dim parse As New Parser(Environment.CommandLine)
        parse.GetPossiblyQuotedText() ' Skip the program name
        parse.SkipWhitespace()
        Dim state As New Workspace
        state.BuildingAssembly = AssemblyBuilder.DefineDynamicAssembly(New AssemblyName(Guid.NewGuid().ToString), AssemblyBuilderAccess.Run)
        Dim zeroSlot As New Slot With {.Kind = SlotKinds.LongKind, .Pointer = Marshal.AllocHGlobal(8)}
        Marshal.WriteInt64(zeroSlot.Pointer, 0)
        state.Slots.Add("zero", zeroSlot)
        Dim createStandardIntSlot = Sub(Name As String, Data As Integer)
                                        Dim slot As New Slot With {.Kind = SlotKinds.IntKind, .Pointer = Marshal.AllocHGlobal(4)}
                                        Marshal.WriteInt32(slot.Pointer, Data)
                                        state.Slots.Add(Name, slot)
                                    End Sub
        createStandardIntSlot("lasterror", 0)
        createStandardIntSlot("pid", Process.GetCurrentProcess().Id)
        Dim parentQuery As New ManagementObjectSearcher("root\cimv2", "select ParentProcessId from Win32_Process where ProcessId = " & Process.GetCurrentProcess().Id)
        Dim resultsEnum = parentQuery.Get().GetEnumerator()
        resultsEnum.MoveNext()
        createStandardIntSlot("parentpid", resultsEnum.Current("ParentProcessId"))
        state.RunningScripts.Push(New Script With {.ScriptSource = "command line"})
        Try
            Do Until parse.AtEnd()
                InvokeOneLine(parse.GetLine(), state)
            Loop
        Catch ex As Exception
            Dim curScript = state.RunningScripts.Peek()
            Console.WriteLine("Error running script from " & curScript.ScriptSource & " at instruction " & curScript.CurrentInstructionId & ":")
            Console.WriteLine(curScript.CurrentInstructionText)
            Console.WriteLine()
            If TypeOf ex Is ParserException Then
                CType(ex, ParserException).PrintParseErrorReport()
            Else
                PrintExceptionReport(ex)
            End If
            Console.ReadKey(True)
        End Try
    End Sub
    Declare Function GetLastError Lib "kernel32.dll" () As Integer
    Function CreateStructTypeFromParamTypes(ParamTypes As IEnumerable(Of Type), State As Workspace) As Type
        Dim dynModule = State.BuildingAssembly.DefineDynamicModule(Guid.NewGuid().ToString)
        Dim dynStruct = dynModule.DefineType(Guid.NewGuid().ToString, TypeAttributes.SequentialLayout)
        For n = 0 To ParamTypes.Count - 1
            dynStruct.DefineField("Member" & n, ParamTypes(n), FieldAttributes.Public)
        Next
        dynModule.CreateGlobalFunctions()
        Return dynStruct.CreateType()
    End Function
    Function CreateStructFromParamValues(StructType As Type, Values As IEnumerable(Of Object)) As Object
        Dim structObject = StructType.GetConstructor(Type.EmptyTypes).Invoke(Nothing)
        For n = 0 To Values.Count - 1
            Dim finalField = StructType.GetField("Member" & n)
            finalField.SetValue(structObject, Values(n))
        Next
        Return structObject
    End Function
    Function GetSlotData(Slot As Slot, State As Workspace) As Object
        Dim smallStructType = CreateStructTypeFromParamTypes({Slot.Kind.MarshalAsType}, State)
        Dim smallStruct = Marshal.PtrToStructure(Slot.Pointer, smallStructType)
        Return smallStructType.GetField("Member0").GetValue(smallStruct)
    End Function
    Sub SetSlotData(Slot As Slot, State As Workspace, Value As Object)
        Dim smallStructType = CreateStructTypeFromParamTypes({Slot.Kind.MarshalAsType}, State)
        Dim smallStruct = smallStructType.GetConstructor(Type.EmptyTypes).Invoke({})
        smallStructType.GetField("Member0").SetValue(smallStruct, Value)
        Marshal.StructureToPtr(smallStruct, Slot.Pointer, False)
    End Sub
    Function ParseUnitFactor(Parser As Parser, State As Workspace) As Integer
        Dim unit = Parser.GetTextToDelimiter().ToLowerInvariant
        Select Case unit
            Case "bytes"
                Return 1
            Case "chars"
                Return 2
            Case "sizesof"
                Parser.SkipWhitespace()
                Return SlotKinds.StandardKinds(Parser.GetTextToDelimiter().ToLowerInvariant()).Length
            Case "blockssizedlike"
                Parser.SkipWhitespace()
                Return State.Slots(Parser.GetTextToDelimiter()).Kind.Length
            Case Else
                Throw New InvalidOperationException("Invalid unit: " & unit)
        End Select
    End Function
    Function ParseConvertSizeTypeAndUnit(OriginalLength As Integer, Parser As Parser, State As Workspace) As Object
        Dim divisor As Integer = 1
        Dim numType As Type = GetType(Integer)
        Do
            Dim nextText = Parser.GetTextToDelimiter().ToLowerInvariant
            Parser.SkipWhitespace()
            Select Case nextText
                Case "in"
                    divisor = ParseUnitFactor(Parser, State)
                Case "as"
                    numType = SlotKinds.StandardKinds(Parser.GetTextToDelimiter()).NumericParseType
                Case ""
                    Exit Do
                Case Else
                    Throw New InvalidOperationException("Invalid size adjustment: " & nextText)
            End Select
            Parser.SkipWhitespace()
        Loop
        Return CTypeDynamic(CInt(OriginalLength / divisor), numType)
    End Function
    Function ParseCopyPointerAndLength(Parser As Parser, State As Workspace) As Tuple(Of IntPtr, Integer)
        Dim slot = State.Slots(Parser.GetTextToDelimiter())
        Parser.SkipWhitespace()
        Dim finalPtr = slot.Pointer
        Dim fieldSize As Integer = slot.Kind.Length
        Dim canUseSlotInfo As Boolean = True
        Do
            If Parser.AtEnd() Then Exit Do
            If Parser.PeekCharacter() = "="c Then Exit Do
            Dim adjustKeyword = Parser.GetTextToDelimiter().ToLowerInvariant()
            Parser.SkipWhitespace()
            Select Case adjustKeyword
                Case "offset"
                    Dim offset = ParseGetNumberOrNumericSlotValue(Parser, State)
                    Parser.SkipWhitespace()
                    If Not Parser.AtEnd() AndAlso Parser.PeekCharacter() <> "="c Then offset *= ParseUnitFactor(Parser, State)
                    finalPtr += offset
                    fieldSize = 0
                Case "field"
                    If Not canUseSlotInfo Then Throw New InvalidOperationException("Field offsets are not available for inner blocks")
                    Dim field = Parser.GetNumber(Of Integer)()
                    finalPtr += Marshal.OffsetOf(slot.Kind.MarshalAsType, "Member" & field)
                    fieldSize = Marshal.SizeOf(slot.Kind.MarshalAsType.GetField("Member" & field).FieldType)
                    canUseSlotInfo = False
                Case "dereferenced"
                    finalPtr = Marshal.ReadIntPtr(finalPtr)
                    canUseSlotInfo = False
                    fieldSize = 0
                Case Else
                    Throw New InvalidOperationException("Invalid pointer adjustment keyword: " & adjustKeyword)
            End Select
            Parser.SkipWhitespace()
        Loop
        Return Tuple.Create(finalPtr, fieldSize)
    End Function
    Function ParseGetNumberOrNumericSlotValue(Parser As Parser, State As Workspace) As Integer
        If Parser.AtNumber() Then
            Return Parser.GetNumber(Of Integer)()
        Else
            Return GetSlotData(State.Slots(Parser.GetTextToDelimiter()), State)
        End If
    End Function
    Function ProcessParamList(Text As String, State As Workspace) As Object()
        Dim parse As New Parser(Text)
        Dim sizeIndex As Integer?
        Dim sizeSpecifier As String = ""
        Dim paramValues As New List(Of Object)
        Do Until parse.AtEnd()
            Dim typeName = parse.GetTextToDelimiter().ToLowerInvariant
            parse.SkipWhitespace()
            If SlotKinds.StandardKinds.ContainsKey(typeName) Then
                paramValues.Add(SlotKinds.StandardKinds(typeName).ParseFunc(parse))
            ElseIf typeName = "nullptr" Then
                paramValues.Add(IntPtr.Zero)
            ElseIf typeName = "blockptr" Then
                Dim structMembers = ProcessParamList(parse.GetUntilBalanced("("c, ")"c), State)
                Dim finalType = CreateStructTypeFromParamTypes(structMembers.Select(Function(o) o.GetType()), State)
                Dim structObject = CreateStructFromParamValues(finalType, structMembers)
                Dim blockPtr = Marshal.AllocHGlobal(Marshal.SizeOf(structObject))
                Marshal.StructureToPtr(structObject, blockPtr, False)
                paramValues.Add(blockPtr)
            ElseIf typeName = "slotdata" Then
                Dim slot = State.Slots(parse.GetTextToDelimiter())
                paramValues.Add(GetSlotData(slot, State))
            ElseIf typeName = "slotptr" Then
                Dim slot = State.Slots(parse.GetTextToDelimiter())
                paramValues.Add(slot.Pointer)
            ElseIf typeName = "allocsize" Then
                Dim slot As AllocatedSlot = State.Slots(parse.GetTextToDelimiter())
                parse.SkipWhitespace()
                paramValues.Add(ParseConvertSizeTypeAndUnit(slot.Length, parse, State))
            ElseIf typeName = "slotsize" Then
                Dim slot = State.Slots(parse.GetTextToDelimiter())
                parse.SkipWhitespace()
                paramValues.Add(ParseConvertSizeTypeAndUnit(slot.Kind.Length, parse, State))
            ElseIf typeName = "blocksize" Then
                If sizeIndex.HasValue Then Throw New InvalidOperationException("Block size already used at index " & sizeIndex)
                sizeIndex = paramValues.Count
                sizeSpecifier = parse.GetAllWordsToDelimiter()
                Dim tempParser As New Parser(sizeSpecifier)
                paramValues.Add(ParseConvertSizeTypeAndUnit(0, tempParser, State))
            Else
                Throw New Exception("Invalid type name: " & typeName)
            End If
            parse.SkipWhitespace()
            If Not parse.AtEnd() Then
                parse.AssumeCharacter(",")
                parse.SkipWhitespace()
            End If
        Loop
        If sizeIndex.HasValue Then
            Dim tempParser As New Parser(sizeSpecifier)
            paramValues(sizeIndex) = ParseConvertSizeTypeAndUnit(Marshal.SizeOf(CreateStructTypeFromParamTypes(paramValues.Select(Function(o) o.GetType()), State)), tempParser, State)
        End If
        Return paramValues.ToArray
    End Function
    Sub InvokeOneLine(Line As String, State As Workspace)
        Dim curScript = State.RunningScripts.Peek()
        curScript.CurrentInstructionText = Line
        curScript.CurrentInstructionId += 1
        Dim parse As New Parser(Line)
        parse.SkipWhitespace()
        Dim instruction = parse.GetTextToDelimiter().ToLowerInvariant()
        parse.SkipWhitespace()
        Select Case instruction
            Case ""
                Exit Sub ' Blank line
            Case "//"
                Exit Sub ' Comment
            Case "call"
                DoCall(parse.PeekToEnd(), State)
            Case "newslot"
                Dim slot As New Slot
                Dim kindName = parse.GetTextToDelimiter().ToLowerInvariant()
                parse.SkipWhitespace()
                Dim slotName = parse.GetTextToDelimiter()
                parse.SkipWhitespace()
                If kindName = "block" Then
                    parse.AssumeCharacter("="c)
                    parse.SkipWhitespace()
                    Dim paramValues = ProcessParamList(parse.PeekToEnd(), State)
                    Dim structType = CreateStructTypeFromParamTypes(paramValues.Select(Function(o) o.GetType()), State)
                    Dim structObject = CreateStructFromParamValues(structType, paramValues)
                    Dim customKind As New SlotKind(Nothing, structType, Marshal.SizeOf(structType), Nothing, Nothing)
                    slot.Kind = customKind
                    slot.Pointer = Marshal.AllocHGlobal(slot.Kind.Length)
                    Marshal.StructureToPtr(structObject, slot.Pointer, False)
                Else
                    slot.Kind = SlotKinds.StandardKinds(kindName)
                    slot.Pointer = Marshal.AllocHGlobal(slot.Kind.Length)
                    If Not parse.AtEnd() Then
                        parse.AssumeCharacter("="c)
                        parse.SkipWhitespace()
                        Dim value = slot.Kind.ParseFunc(parse)
                        SetSlotData(slot, State, value)
                    End If
                End If
                State.Slots.Add(slotName, slot)
            Case "readslot"
                If parse.AtEnd Then Throw New Exception("Slot not specified")
                Dim printSlotName As Boolean = True
                Do
                    If parse.PeekCharacter() <> "/"c Then Exit Do
                    Select Case parse.GetTextToDelimiter()
                        Case "/raw"
                            printSlotName = False
                    End Select
                    parse.SkipWhitespace()
                Loop
                Dim slotName = parse.GetTextToDelimiter()
                Dim slot = State.Slots(slotName)
                If SlotKinds.StandardKinds.Values.Contains(slot.Kind) Then
                    Dim value = GetSlotData(slot, State)
                    If printSlotName Then
                        Console.WriteLine("The value in slot " & slotName & " is " & slot.Kind.StringFunc(value))
                    Else
                        Console.WriteLine(slot.Kind.StringFunc(value))
                    End If
                Else
                    Console.WriteLine("The value in slot " & slotName & " is a block of length " & slot.Kind.Length)
                End If
            Case "lasterror"
                Console.WriteLine("The last Win32 error was " & State.LastError)
            Case "allocslot"
                Dim slot As New AllocatedSlot
                slot.Kind = SlotKinds.StandardKinds(parse.GetTextToDelimiter())
                If Not {GetType(IntPtr), GetType(UIntPtr)}.Contains(slot.Kind.MarshalAsType) Then Throw New InvalidOperationException("Allocated slots must be pointers")
                slot.Pointer = Marshal.AllocHGlobal(IntPtr.Size)
                parse.SkipWhitespace()
                State.Slots.Add(parse.GetTextToDelimiter(), slot)
                parse.SkipWhitespace()
                parse.AssumeCharacter(":"c)
                parse.SkipWhitespace()
                Dim length As Integer = ParseGetNumberOrNumericSlotValue(parse, State)
                parse.SkipWhitespace()
                If Not parse.AtEnd() Then length *= ParseUnitFactor(parse, State)
                slot.Length = length
                SetSlotData(slot, State, Marshal.AllocHGlobal(length))
                parse.SkipWhitespace()
            Case "copyslot"
                parse.AssumeNotEnd()
                Dim copyLength As Integer = 0
                If parse.PeekCharacter() = "/"c Then
                    Dim switch = parse.GetTextToDelimiter().ToLowerInvariant()
                    parse.SkipWhitespace()
                    Select Case switch
                        Case "/bytes"
                            copyLength = ParseGetNumberOrNumericSlotValue(parse, State)
                        Case "/length"
                            Dim rawNum = ParseGetNumberOrNumericSlotValue(parse, State)
                            parse.SkipWhitespace()
                            copyLength = rawNum * ParseUnitFactor(parse, State)
                        Case "/lengthof"
                            copyLength = SlotKinds.StandardKinds(parse.GetTextToDelimiter().ToLowerInvariant()).Length
                        Case Else
                            Throw New InvalidOperationException("Invalid copyslot length switch: " & switch)
                    End Select
                    parse.SkipWhitespace()
                End If
                Dim destInfo = ParseCopyPointerAndLength(parse, State)
                parse.SkipWhitespace()
                parse.AssumeCharacter("="c)
                parse.SkipWhitespace()
                Dim srcInfo = ParseCopyPointerAndLength(parse, State)
                parse.SkipWhitespace()
                parse.AssumeEnd()
                If copyLength = 0 Then copyLength = destInfo.Item2
                If copyLength = 0 Then copyLength = srcInfo.Item2
                If copyLength = 0 Then Throw New InvalidOperationException("Copy length could not be automatically determined")
                Dim valueBytes(copyLength - 1) As Byte
                Marshal.Copy(srcInfo.Item1, valueBytes, 0, copyLength)
                Marshal.Copy(valueBytes, 0, destInfo.Item1, copyLength)
            Case "run"
                Dim filename = parse.GetPossiblyQuotedText
                Using fScript As New IO.StreamReader(filename)
                    State.RunningScripts.Push(New Script With {.ScriptSource = filename})
                    Do Until fScript.EndOfStream
                        InvokeOneLine(Trim(fScript.ReadLine), State)
                    Loop
                    State.RunningScripts.Pop()
                End Using
            Case "deleteslot"
                Dim slotName = parse.GetTextToDelimiter
                Dim slot = State.Slots(slotName)
                Marshal.FreeHGlobal(slot.Pointer)
                State.Slots.Remove(slotName)
            Case "zeroslot"
                Dim slotName = parse.GetTextToDelimiter
                Dim slot = State.Slots(slotName)
                Dim zeros(slot.Kind.Length - 1) As Byte
                Marshal.Copy(zeros, 0, slot.Pointer, zeros.Length)
            Case "pause"
                Console.ReadKey(True)
            Case "interactive"
                State.RunningScripts.Push(New Script With {.ScriptSource = "console"})
                Console.WriteLine("Entered interactive mode - type ""exit"" to quit/continue")
                Do
                    Console.Write("interactive> ")
                    Dim text = Console.ReadLine
                    If text = "exit" Or text = "" Then Exit Do
                    Try
                        InvokeOneLine(text, State)
                    Catch ex As ParserException
                        ex.PrintParseErrorReport()
                    Catch ex As Exception
                        PrintExceptionReport(ex)
                    End Try
                Loop
                State.RunningScripts.Pop()
            Case "about"
                Console.WriteLine("SprintDLL by Ben Nordick")
                Console.WriteLine("Available on GitHub: https://github.com/Fleex255/SprintDLL")
            Case Else
                Throw New InvalidOperationException("Invalid instruction: " & instruction)
        End Select
    End Sub
    Sub DoCall(Line As String, State As Workspace)
        Dim parse As New Parser(Line)
        Dim dllName = parse.GetPossiblyQuotedText()
        Dim funcName As String = ""
        Dim address As IntPtr
        If dllName = "funcat" Then
            parse.SkipWhitespace()
            Dim ptrSlot = State.Slots(parse.GetTextToDelimiter())
            address = GetSlotData(ptrSlot, State)
        Else
            parse.AssumeCharacter("!")
            funcName = parse.GetTextToDelimiter()
        End If
        parse.SkipWhitespace()
        Dim callConv As CallingConvention = CallingConvention.Winapi
        Dim callConvType As Type = GetType(CallConvStdcall)
        Dim retKind As SlotKind = Nothing
        Dim returnIntoSlotName As String = ""
        Do While parse.PeekCharacter() = "/"c
            Dim instruction = parse.GetTextToDelimiter().ToLowerInvariant
            parse.SkipWhitespace()
            Select Case instruction
                Case "/call"
                    Select Case parse.GetTextToDelimiter().ToLowerInvariant
                        Case "stdcall"
                            callConv = CallingConvention.StdCall
                        Case "cdecl"
                            callConv = CallingConvention.Cdecl
                            callConvType = GetType(CallConvCdecl)
                        Case "thiscall"
                            callConv = CallingConvention.ThisCall
                            callConvType = GetType(CallConvThiscall)
                        Case "fastcall"
                            callConv = CallingConvention.FastCall
                            callConvType = GetType(CallConvFastcall)
                    End Select
                Case "/return"
                    Dim retTypeName = parse.GetTextToDelimiter().ToLowerInvariant
                    If SlotKinds.StandardKinds.ContainsKey(retTypeName) Then
                        retKind = SlotKinds.StandardKinds(retTypeName)
                    Else
                        Throw New InvalidOperationException("Invalid type name: " & retTypeName)
                    End If
                Case "/into"
                    returnIntoSlotName = parse.GetTextToDelimiter()
            End Select
            parse.SkipWhitespace()
        Loop
        Dim returnIntoSlot As Slot = Nothing
        If returnIntoSlotName <> "" Then
            If retKind Is Nothing Then
                returnIntoSlot = State.Slots(returnIntoSlotName)
                retKind = returnIntoSlot.Kind
            ElseIf State.Slots.ContainsKey(returnIntoSlotName) Then
                If State.Slots(returnIntoSlotName).Kind.MarshalAsType IsNot retKind.MarshalAsType Then Throw New InvalidOperationException("Return type doesn't match type of slot: " & returnIntoSlotName)
            End If
        End If
        Dim dynModule = State.BuildingAssembly.DefineDynamicModule(Guid.NewGuid().ToString)
        Dim paramValues As Object() = Nothing
        Dim paramTypes As Type() = Nothing
        If Not parse.AtEnd() Then
            paramValues = ProcessParamList(parse.GetUntilBalanced("("c, ")"c), State)
            paramTypes = paramValues.Select(Function(o) o.GetType()).ToArray
        End If
        Dim result As Object
        If funcName = "" Then
            Dim delegateType = dynModule.DefineType(Guid.NewGuid.ToString(), TypeAttributes.Public Or TypeAttributes.Sealed, GetType(MulticastDelegate))
            delegateType.DefineConstructor(MethodAttributes.RTSpecialName Or MethodAttributes.SpecialName Or MethodAttributes.Public Or
                                           MethodAttributes.HideBySig, CallingConventions.Standard, {GetType(Object), GetType(IntPtr)}).SetImplementationFlags(MethodImplAttributes.Runtime)
            Dim invoker = delegateType.DefineMethod("Invoke", MethodAttributes.Public Or MethodAttributes.Virtual Or MethodAttributes.NewSlot Or MethodAttributes.HideBySig,
                                                    CallingConventions.Standard, retKind?.MarshalAsType, Nothing, {callConvType}, paramTypes, Nothing, Nothing)
            invoker.SetImplementationFlags(MethodImplAttributes.Runtime)
            Dim finishedType = delegateType.CreateType()
            Dim nativeMethod = Marshal.GetDelegateForFunctionPointer(address, finishedType)
            result = nativeMethod.DynamicInvoke(paramValues)
            funcName = "Function at " & address.ToInt64.ToString
        Else
            Dim dynMethod = dynModule.DefinePInvokeMethod("PFunc", dllName, funcName, MethodAttributes.Public Or MethodAttributes.Static,
                                                          CallingConventions.Standard, retKind?.MarshalAsType, paramTypes, callConv, CharSet.Auto)
            dynMethod.SetImplementationFlags(dynMethod.GetMethodImplementationFlags() Or MethodImplAttributes.PreserveSig)
            dynModule.CreateGlobalFunctions()
            Dim finalMethod = dynModule.GetMethod("PFunc")
            result = finalMethod.Invoke(Nothing, paramValues)
        End If
        State.LastError = GetLastError
        Marshal.WriteInt32(State.Slots("lasterror").Pointer, State.LastError)
        If result IsNot Nothing Then
            If returnIntoSlot Is Nothing And returnIntoSlotName <> "" Then
                returnIntoSlot = New Slot With {.Kind = retKind, .Pointer = Marshal.AllocHGlobal(retKind.Length)}
                State.Slots.Add(returnIntoSlotName, returnIntoSlot)
            End If
            If returnIntoSlot IsNot Nothing Then
                SetSlotData(returnIntoSlot, State, result)
            Else
                Console.WriteLine(funcName & " returned " & retKind.StringFunc(result))
            End If
        End If
    End Sub
    Sub PrintExceptionReport(Ex As Exception)
        If TypeOf Ex Is TargetInvocationException Then
            Console.WriteLine(Ex.InnerException.Message)
        Else
            Console.WriteLine(Ex.Message)
        End If
    End Sub
End Module
