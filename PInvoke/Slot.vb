Imports System.Runtime.InteropServices
Public Class SlotKind
    Public NumericParseType As Type
    Public MarshalAsType As Type
    Public Length As Integer
    Public StringFunc As Func(Of Object, String)
    Public ParseFunc As Func(Of Parser, Object)
    Public Sub New(NumParseAs As Type, MarshalType As Type, ByteLength As Integer, Stringifier As Func(Of Object, String), Parser As Func(Of Parser, Object))
        NumericParseType = NumParseAs
        MarshalAsType = MarshalType
        Length = ByteLength
        StringFunc = Stringifier
        If StringFunc Is Nothing Then StringFunc = Function(o) o.ToString()
        ParseFunc = Parser
        If NumParseAs IsNot Nothing Then ParseFunc = Function(p As Parser) p.GetNumber(NumParseAs)
    End Sub
End Class
Public Class SlotKinds
    Public Shared IntKind As New SlotKind(GetType(Integer), GetType(Integer), 4, Nothing, Nothing)
    Public Shared UintKind As New SlotKind(GetType(UInteger), GetType(UInteger), 4, Nothing, Nothing)
    Public Shared ShortKind As New SlotKind(GetType(Short), GetType(Short), 2, Nothing, Nothing)
    Public Shared UshortKind As New SlotKind(GetType(UShort), GetType(UShort), 2, Nothing, Nothing)
    Public Shared LongKind As New SlotKind(GetType(Long), GetType(Long), 8, Nothing, Nothing)
    Public Shared UlongKind As New SlotKind(GetType(ULong), GetType(ULong), 8, Nothing, Nothing)
    Public Shared NativeKind As New SlotKind(If(IntPtr.Size = 4, GetType(Integer), GetType(Long)), GetType(IntPtr), IntPtr.Size, Nothing, Nothing)
    Public Shared UnativeKind As New SlotKind(If(UIntPtr.Size = 4, GetType(UInteger), GetType(ULong)), GetType(UIntPtr), UIntPtr.Size, Nothing, Nothing)
    Public Shared LpstrKind As New SlotKind(Nothing, GetType(IntPtr), IntPtr.Size, Function(o) Marshal.PtrToStringAnsi(o), Function(p) Marshal.StringToHGlobalAnsi(p.GetQuotedString()))
    Public Shared LpwstrKind As New SlotKind(Nothing, GetType(IntPtr), IntPtr.Size, Function(o) Marshal.PtrToStringUni(o), Function(p) Marshal.StringToHGlobalUni(p.GetQuotedString()))
    Public Shared ByteKind As New SlotKind(GetType(Byte), GetType(Byte), 1, Nothing, Nothing)
    Public Shared GuidKind As New SlotKind(Nothing, GetType(Guid), Marshal.SizeOf(GetType(Guid)), Nothing, Function(p) Guid.Parse(p.GetTextToDelimiter()))
    Public Shared SingleKind As New SlotKind(GetType(Single), GetType(Single), 4, Nothing, Nothing)
    Public Shared DoubleKind As New SlotKind(GetType(Double), GetType(Double), 8, Nothing, Nothing)
    Public Shared BstrKind As New SlotKind(Nothing, GetType(IntPtr), IntPtr.Size, Function(o) Marshal.PtrToStringBSTR(o), Function(p) Marshal.StringToBSTR(p.GetQuotedString()))
    Public Shared StandardKinds As New Dictionary(Of String, SlotKind)
    Shared Sub New()
        StandardKinds.Add("int", IntKind)
        StandardKinds.Add("uint", UintKind)
        StandardKinds.Add("short", ShortKind)
        StandardKinds.Add("ushort", UshortKind)
        StandardKinds.Add("long", LongKind)
        StandardKinds.Add("ulong", UlongKind)
        StandardKinds.Add("native", NativeKind)
        StandardKinds.Add("unative", UnativeKind)
        StandardKinds.Add("lpstr", LpstrKind)
        StandardKinds.Add("lpwstr", LpwstrKind)
        StandardKinds.Add("byte", ByteKind)
        StandardKinds.Add("guid", GuidKind)
        StandardKinds.Add("single", SingleKind)
        StandardKinds.Add("double", DoubleKind)
        StandardKinds.Add("bstr", BstrKind)
    End Sub
End Class
Public Class Slot
    Public Kind As SlotKind
    Public Pointer As IntPtr
End Class
Public Class AllocatedSlot
    Inherits Slot
    Public Length As Integer
End Class