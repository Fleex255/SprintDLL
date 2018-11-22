Imports System.Text
Public Class Parser
    Shared Delimiters As Char() = {","c, "!"c, "("c, ")"c, """"c, ";"c, ":"c, vbCr(0)}
    Dim Text As String
    Dim Position As Integer
    Public Sub New(Input As String)
        Text = Input
    End Sub
    Public Function AtEnd() As Boolean
        Return Position >= Text.Length
    End Function
    Public Function PeekCharacter() As Char?
        If AtEnd() Then Return Nothing Else Return Text(Position)
    End Function
    Public Function GetCharacter() As Char
        AssumeNotEnd()
        Dim nextChar = Text(Position)
        Position += 1
        Return nextChar
    End Function
    Public Sub SkipWhitespace()
        Do While PeekCharacter().HasValue AndAlso Char.IsWhiteSpace(PeekCharacter().Value)
            Position += 1
        Loop
    End Sub
    Public Sub AssumeCharacter(Character As Char)
        If GetCharacter() <> Character Then ThrowParseException("Expected character was " & Character, Position - 1)
    End Sub
    Public Sub AssumeEnd()
        If Not AtEnd() Then ThrowParseException("Expected end of text fragment")
    End Sub
    Public Sub AssumeNotEnd()
        If AtEnd() Then ThrowParseException("Unexpected end of text fragment")
    End Sub
    Public Function GetQuotedString() As String
        AssumeCharacter("""")
        Dim sb As New StringBuilder
        Do
            Dim curChar = GetCharacter()
            If curChar = """" Then
                Exit Do
            ElseIf curChar = "\" Then
                Dim maybeEscapedChar = GetCharacter()
                If Not {""""c, "\"c, "N"c, "n"c, "r"c}.Contains(maybeEscapedChar) Then sb.Append("\")
                Select Case maybeEscapedChar
                    Case """"c, "\"c
                        sb.Append(maybeEscapedChar)
                    Case "N"c
                        sb.Append(vbCrLf)
                    Case "n"c
                        sb.Append(vbLf)
                    Case "r"c
                        sb.Append(vbCr)
                End Select
            Else
                sb.Append(curChar)
            End If
        Loop
        Return sb.ToString
    End Function
    Public Function GetTextToDelimiter() As String
        Dim sb As New StringBuilder
        Do Until AtEnd() OrElse Char.IsWhiteSpace(PeekCharacter()) OrElse Delimiters.Contains(PeekCharacter())
            sb.Append(GetCharacter())
        Loop
        Return sb.ToString
    End Function
    Public Function GetAllWordsToDelimiter() As String
        Dim sb As New StringBuilder
        Do Until AtEnd() OrElse Delimiters.Contains(PeekCharacter())
            sb.Append(GetCharacter())
        Loop
        Return sb.ToString
    End Function
    Public Function GetPossiblyQuotedText() As String
        If Not PeekCharacter().HasValue Then ThrowParseException("Unexpected end of text fragment")
        If PeekCharacter() = """" Then
            Return GetQuotedString()
        Else
            Return GetTextToDelimiter()
        End If
    End Function
    Public Function GetUntilBalanced(Opener As Char, Closer As Char) As String
        AssumeCharacter(Opener)
        Dim depth As Integer = 1
        Dim inQuote As Boolean = False
        Dim escaping As Boolean = False
        Dim startPos = Position
        Do
            Dim thisChar = GetCharacter()
            If thisChar = """" And Not escaping Then
                inQuote = Not inQuote
            ElseIf inQuote And thisChar = "\" And Not escaping Then
                escaping = True
            ElseIf thisChar = Opener Then
                depth += 1
            ElseIf thisChar = Closer Then
                depth -= 1
            End If
            If depth = 0 Then Exit Do
            escaping = False
        Loop
        Return Text.Substring(startPos, Position - startPos - 1)
    End Function
    Public Function GetLine() As String
        Dim sb As New StringBuilder
        Do While Not AtEnd()
            Dim curChar = GetCharacter()
            If curChar = vbCr Or curChar = ";"c Then
                Do Until AtEnd() OrElse Not {vbCr, vbLf}.Contains(PeekCharacter())
                    Position += 1
                Loop
                Exit Do
            Else
                sb.Append(curChar)
            End If
        Loop
        SkipWhitespace()
        Return sb.ToString
    End Function
    Public Function PeekToEnd() As String
        Return Text.Substring(Position)
    End Function
    Public Function AtNumber() As Boolean
        If AtEnd() Then Return False
        Return Char.IsNumber(PeekCharacter())
    End Function
    Public Function GetNumber(OfType As Type) As Object
        Dim isHex As Boolean = False
        If PeekCharacter() = "0"c Then
            Position += 1
            If PeekCharacter().HasValue AndAlso PeekCharacter() = "x"c Then
                isHex = True
                Position += 1
            Else
                Position -= 1
            End If
        End If
        Dim numberText = GetTextToDelimiter()
        Dim parseMethod = OfType.GetMethod("Parse", {GetType(String), GetType(Globalization.NumberStyles)})
        Return parseMethod.Invoke(Nothing, {numberText, If(isHex, Globalization.NumberStyles.HexNumber, Globalization.NumberStyles.Number)})
    End Function
    Public Function GetNumber(Of T)() As T
        Return GetNumber(GetType(T))
    End Function
    Public Sub SkipCharacters(Count As Integer)
        Position += Count
    End Sub
    Private Sub ThrowParseException(ErrorText As String, Optional Position As Integer = -1)
        If Position = -1 Then Position = Me.Position
        Throw New ParserException(ErrorText, Text, Position)
    End Sub
End Class
Public Class ParserException
    Inherits Exception
    Dim ErrorText As String
    Dim Text As String
    Dim Position As Integer
    Public Sub New(ErrorText As String, Text As String, Position As Integer)
        Me.ErrorText = ErrorText
        Me.Text = Text
        Me.Position = Position
    End Sub
    Public Sub PrintParseErrorReport()
        Console.WriteLine(ErrorText)
        Console.WriteLine(Text)
        Console.Write(Space(Position))
        Console.WriteLine("^")
    End Sub
End Class