Attribute VB_Name = "modCalculate"
Option Explicit

Dim AngleMode As Integer
Dim InError As Boolean
Dim LogBase As Double

Dim Char As String
Dim CurrentEntryIndex As Integer
Dim InputString As String
Dim OutputString As String
Dim OutputValue As Double
Dim Value As Double
Dim ValueString As String

Const Pi = 3.14159265358979

Public Function CalculateString(IString As String, AMode As Integer, BaseMode As Integer, Decimals As Integer, LBase As Double)
On Error GoTo ErrorHandler:
Dim Answer As String
Dim BinAnswer As String
Dim DecimalCheck As Long
Dim i As Integer
Dim LenAfterDecimal As Long
Dim Remainder As String

    'If nothing was entered, exit
    If IString = "" Then
        CalculateString = "Error: Nothing entered"
        Exit Function
    End If

    'Set values
    AngleMode = AMode
    CurrentEntryIndex = 1
    InError = False
    InputString = IString
    LogBase = LBase

    'Start calculation routine
    ExtractToken
    Answer = GetE()

    'Load error into returned variable
    If InError Then
        CalculateString = OutputString
        Exit Function
    End If

    Select Case BaseMode
        Case 0 'Decimal

            '14 decimals and above are floating
            If Decimals < 14 Then

                'Check for decimal
                DecimalCheck = InStr(1, CStr(Answer), ".")

                'If decimal does not exist, tag on the number
                'of zeroes that the user specified
                If DecimalCheck = 0 Then
                    If Decimals <> "0" Then
                        Answer = Answer + "."
                        For i = 1 To Decimals
                            Answer = Answer + "0"
                        Next i
                    End If

                'If decimal does exist, adjust the answer to
                'the number of decimal places that the user
                'specified
                Else
                    LenAfterDecimal = Len(Answer) - DecimalCheck
                    If LenAfterDecimal > Decimals Then
                        If Decimals = "0" Then
                            DecimalCheck = DecimalCheck - 1
                        End If
                        Answer = Mid(Answer, 1, DecimalCheck + Decimals)
                    Else
                        For i = 1 To (Decimals - LenAfterDecimal)
                            Answer = Answer + "0"
                        Next i
                    End If
                End If
            End If

        Case 1 'Binary

            If CDbl(Answer) <= 32767 Then
                BinAnswer = ""
                DecimalCheck = InStr(1, CStr(Answer), ".")
                If DecimalCheck <> 0 Then
                    If CInt(Mid(CStr(Answer), DecimalCheck + 1, 1)) < 5 Then
                        Answer = CDbl(Left(Answer, DecimalCheck - 1))
                    Else
                        Answer = CDbl(Left(Answer, DecimalCheck - 1)) + 1
                    End If
                End If
                Do
                    Answer = Answer / 2
                    DecimalCheck = InStr(1, CStr(Answer), ".")
                    If DecimalCheck = 0 Then
                        Remainder = "0"
                    Else
                        Answer = CDbl(Left(Answer, DecimalCheck - 1))
                        Remainder = "1"
                    End If
                    BinAnswer = Remainder + BinAnswer
                Loop Until Answer < 1
                Answer = CDbl(BinAnswer)
            End If

        Case 2 'Hexadecimal

            Answer = Hex(Answer)

        Case 3 'Octal

            Answer = Oct(Answer)

    End Select

    'Display final answer
    CalculateString = Answer

    Exit Function

ErrorHandler:

    'Trap errors
    TrapErrors Err.Number

End Function

Private Sub ExtractToken()
Dim i As Integer

    'Set default values
    OutputString = ""
    OutputValue = 0
    ValueString = ""

    'If at the end of string, return EOS
    If CurrentEntryIndex > Len(InputString) Then
        OutputString = "EOS"
        Exit Sub
    End If

    'Get character to be examined
    Char = Mid(InputString, CurrentEntryIndex, 1)

    'Space
    If Char = " " Then
        CurrentEntryIndex = CurrentEntryIndex + 1
        ExtractToken
        Exit Sub
    End If

    'Operator or parenthesis
    If Char = "+" Or Char = "-" Or Char = "*" Or Char = "/" Or Char = "^" Or Char = "(" Or Char = ")" Or Char = "!" Then
        CurrentEntryIndex = CurrentEntryIndex + 1

        'Set return value
        OutputString = Char
        Exit Sub
    End If

    'Number
    If (Char >= "0" And Char <= "9") Or Char = "." Then

        'Digits before decimal
        While Char >= "0" And Char <= "9"
            ValueString = ValueString + Char
            CurrentEntryIndex = CurrentEntryIndex + 1
            If CurrentEntryIndex <= Len(InputString) Then
                Char = Mid(InputString, CurrentEntryIndex, 1)
            Else
                Char = ""
            End If
        Wend

        'Decimal
        While Char = "."
            ValueString = ValueString + Char
            CurrentEntryIndex = CurrentEntryIndex + 1
            If CurrentEntryIndex <= Len(InputString) Then
                Char = Mid(InputString, CurrentEntryIndex, 1)
            Else
                Char = ""
            End If
        Wend

        'Digits after decimal
        While Char >= "0" And Char <= "9"
            ValueString = ValueString + Char
            CurrentEntryIndex = CurrentEntryIndex + 1
            If CurrentEntryIndex <= Len(InputString) Then
                Char = Mid(InputString, CurrentEntryIndex, 1)
            Else
                Char = ""
            End If
        Wend

        'Set return values
        OutputString = "Number"
        OutputValue = CDbl(ValueString)
        Exit Sub
    End If

    'Return text language identifiers
    If LCase(Char) >= "a" And LCase(Char) <= "z" Then
        While (LCase(Char) >= "a" And LCase(Char) <= "z")
            ValueString = ValueString + Char
            CurrentEntryIndex = CurrentEntryIndex + 1
            If CurrentEntryIndex <= Len(InputString) Then
                Char = Mid(InputString, CurrentEntryIndex, 1)
            Else
                Char = ""
            End If
        Wend

        'Pi or e
        If LCase(ValueString) = "pi" Or LCase(ValueString) = "e" Then
            OutputString = "Number"
            If LCase(ValueString) = "pi" Then
                OutputValue = Pi
            Else
                OutputValue = Exp(1)
            End If
            Exit Sub
        End If

        'Set return value
        OutputString = LCase(ValueString)
        Exit Sub
    End If

End Sub

Private Function GetE()
On Error GoTo ErrorHandler

    'Get the lower value (T)
    Value = GetT()

    'Exit function if error call returned
    If InError Then
        Exit Function
    End If

    'Allow for multiple operators of the same precedence
    'level occuring immediately after each other
    While OutputString = "+" Or OutputString = "-"

        Select Case OutputString
    
            'Addition operator
            Case "+"
                ExtractToken
                Value = Value + GetT()
    
            'Subraction operator
            Case "-"
                ExtractToken
                Value = Value - GetT()

        End Select

    Wend

    'Return value for E
    GetE = Value

    'Exit function before error handler
    Exit Function

ErrorHandler:

    'Trap errors
    TrapErrors Err.Number

End Function

Private Function GetT()
On Error GoTo ErrorHandler

    'Get the lower value (F)
    Value = GetF

    'Exit function if error call returned
    If InError Then
        Exit Function
    End If

    'Allow for multiple operators of the same precedence
    'level occuring immediately after each other
    While OutputString = "*" Or OutputString = "/"

        Select Case OutputString
    
            'Multiplication operator
            Case "*"
                ExtractToken
                Value = Value * GetF()
    
            'Division operator
            Case "/"
                ExtractToken
                Value = Value / GetF()
    
        End Select

    Wend

    'Return value for T
    GetT = Value

    'Exit function before error handler
    Exit Function

ErrorHandler:

    'Trap errors
    TrapErrors Err.Number

End Function

Private Function GetF()
On Error GoTo ErrorHandler

    'Handle the low level calculations
    Select Case OutputString

        '***************
        'Basic Functions
        '***************

        'Number
        Case "Number"
            Value = OutputValue
            ExtractToken
            GetF = PostToken

        'Negative
        Case "-"
            ExtractToken
            GetF = -(GetF())

        'Random number
        Case "rnd"
            Randomize
            Value = Rnd
            ExtractToken
            GetF = PostToken

        'Parenthesis
        Case "("
            ExtractToken
            Value = GetE
            If OutputString <> ")" And OutputString <> "EOS" Then
                TrapErrors 0
                Exit Function
            End If
            If OutputString = "EOS" Then
                GetF = Value
            Else
                ExtractToken
                GetF = PostToken
            End If

        '*************
        'Miscellaneous
        '*************

        'Absolute value
        Case "abs"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                GetF = Abs(Value)
            End If

        'Square Root
        Case "sr"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                GetF = Sqr(Value)
            End If

        '**********
        'Logarithms
        '**********

        'Logarithm (to a base)
        Case "log"

            'Get logarithm base
            If Not IsNumeric(LogBase) Then
                TrapErrors (-5)
                Exit Function
            End If

            'Get number
            ExtractToken
            Value = GetF()
            GetF = Log(Value) / Log(LogBase)

        'Natural logarithm
        Case "ln"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                GetF = Log(Value)
            End If

        '***********************
        'Trigonometric Functions
        '***********************

        'Cosine
        Case "cos"
            ExtractToken
            Value = GetF()
            ConvertToDegrees
            If InError Then
                Exit Function
            Else
                GetF = Cos(Value)
            End If

        'Cotangent
        Case "cot"
            ExtractToken
            Value = GetF()
            ConvertToDegrees
            If InError Then
                Exit Function
            Else
                GetF = 1 / Tan(Value)
            End If

        'Cosecant
        Case "csc"
            ExtractToken
            Value = GetF()
            ConvertToDegrees
            If InError Then
                Exit Function
            Else
                GetF = 1 / Sin(Value)
            End If

        'Hyperbolic cosecant
        Case "hcsc"
            ExtractToken
            Value = GetF()
            ConvertToDegrees
            If InError Then
                Exit Function
            Else
                GetF = 2 / (Exp(Value) - Exp(-Value))
            End If
            Exit Function

        'Hyperbolic cosine
        Case "hcos"
            ExtractToken
            Value = GetF()
            ConvertToDegrees
            If InError Then
                Exit Function
            Else
                GetF = (Exp(Value) + Exp(-Value)) / 2
            End If

        'Hyperbolic cotangent
        Case "hcot"
            ExtractToken
            Value = GetF()
            ConvertToDegrees
            If InError Then
                Exit Function
            Else
                GetF = (Exp(Value) + Exp(-Value)) / (Exp(Value) - Exp(-Value))
            End If

        'Hyperbolic secant
        Case "hsec"
            ExtractToken
            Value = GetF()
            ConvertToDegrees
            If InError Then
                Exit Function
            Else
                GetF = 2 / (Exp(Value) + Exp(-Value))
            End If

        'Hyperbolic sine
        Case "hsin"
            ExtractToken
            Value = GetF()
            ConvertToDegrees
            If InError Then
                Exit Function
            Else
                GetF = (Exp(Value) - Exp(-Value)) / 2
            End If

        'Hyperbolic tangent
        Case "htan"
            ExtractToken
            Value = GetF()
            ConvertToDegrees
            If InError Then
                Exit Function
            Else
                GetF = (Exp(Value) - Exp(-Value)) / (Exp(Value) + Exp(-Value))
            End If

        'Inverse hyperbolic cosine
        Case "ihcos"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Log(Value + Sqr(Value * Value - 1))
                ConvertToRadians
                GetF = Value
            End If

        'Inverse hyperbolic cosecant
        Case "ihcsc"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Log((Sgn(Value) * Sqr(Value * Value + 1) + 1) / Value)
                ConvertToRadians
                GetF = Value
            End If

        'Inverse hyperbolic cotangent
        Case "ihcot"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Log((Value + 1) / (Value - 1)) / 2
                ConvertToRadians
                GetF = Value
            End If

        'Inverse hyperbolic sine
        Case "ihsin"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Log(Value + Sqr(Value * Value + 1))
                ConvertToRadians
                GetF = Value
            End If

        'Inverse hyperbolic secant
        Case "ihsec"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Log((Sqr(-Value * Value + 1) + 1) / Value)
                ConvertToRadians
                GetF = Value
            End If

        'Inverse hyperbolic tangent
        Case "ihtan"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Log((1 + Value) / (1 - Value)) / 2
                ConvertToRadians
                GetF = Value
            End If

        'Inverse cosecant
        Case "icsc"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Atn(Value / Sqr(Value * Value - 1)) + (Sgn(Value) - 1) * (2 * Atn(1))
                ConvertToRadians
                GetF = Value
            End If

        'Inverse cosine
        Case "icos"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Atn(-Value / Sqr(-Value * Value + 1)) + 2 * Atn(1)
                ConvertToRadians
                GetF = Value
            End If

        'Inverse cotangent
        Case "icot"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Atn(Value) + 2 * Atn(1)
                ConvertToRadians
                GetF = Value
            End If

        'Inverse secant
        Case "isec"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Atn(Value / Sqr(Value * Value - 1)) + Sgn((Value) - 1) * (2 * Atn(1))
                ConvertToRadians
                GetF = Value
            End If

        'Inverse sine
        Case "isin"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Atn(Value / Sqr(-Value * Value + 1))
                ConvertToRadians
                GetF = Value
            End If

        'Inverse tangent
        Case "itan"
            ExtractToken
            Value = GetF()
            If InError Then
                Exit Function
            Else
                Value = Atn(Value)
                ConvertToRadians
                GetF = Value
            End If

        'Secant
        Case "sec"
            ExtractToken
            Value = GetF()
            ConvertToDegrees
            If InError Then
                Exit Function
            Else
                GetF = 1 / Cos(Value)
            End If

        'Sine
        Case "sin"
            ExtractToken
            Value = GetF()
            ConvertToDegrees
            If InError Then
                Exit Function
            Else
                GetF = Sin(Value)
            End If

        'Tangent
        Case "tan"
            ExtractToken
            Value = GetF()
            ConvertToDegrees
            If InError Then
                Exit Function
            Else
                GetF = Tan(Value)
            End If

        'Everything not handled is an error
        Case Else
            TrapErrors 0

    End Select

    'Exit function before error handler
    Exit Function

ErrorHandler:

    'Trap errors
    TrapErrors Err.Number

End Function

Private Function PostToken()
On Error GoTo ErrorHandler
Dim Factorial As Double
Dim i As Integer

    'Ignore operators, EOS strings, right parentheses, and
    'equals signs
    If OutputString = "+" Or OutputString = "-" Or OutputString = "*" Or OutputString = "/" Or OutputString = "EOS" Or OutputString = ")" Then
        PostToken = Value

    'Handle special tokens that come after the value
    Else
        Select Case OutputString

            'Factorial
            Case "!"
                If (CDbl(Value) <> CLng(Value)) Or Value < 0 Then
                    TrapErrors 0
                    Exit Function
                End If
                Factorial = 1
                For i = Value To 1 Step -1
                    Factorial = Factorial * i
                Next i
                ExtractToken

                'Ignore operators, EOS strings, right
                'parentheses, and equals signs
                If OutputString = "+" Or OutputString = "-" Or OutputString = "*" Or OutputString = "/" Or OutputString = "EOS" Or OutputString = ")" Then
                    PostToken = Factorial
                    ExtractToken

                'Handle special tokens that come after a
                'factorial
                Else

                    Select Case OutputString

                        'Factorial
                        Case "!"
                            TrapErrors 0
                            Exit Function

                        'Exponent
                        Case "^"
                            ExtractToken
                            PostToken = Factorial ^ GetF

                        'Other "post" tokens multiply
                        Case Else
                            PostToken = Factorial * GetF
                    End Select
                End If

            'Exponent
            Case "^"
                ExtractToken
                PostToken = Value ^ GetF

            'Left parenthesis
            Case "("
                PostToken = Value * GetF

            'Other "post" tokens multiply
            Case Else
                PostToken = Value * GetF
        End Select
    End If

    'Exit function before error handler
    Exit Function

ErrorHandler:

    TrapErrors Err.Number

End Function

Private Sub ConvertToDegrees()

    'Convert to degrees
    If AngleMode = 0 Then
        Value = Value * (Pi / 180)
    End If

End Sub

Private Sub ConvertToRadians()

    'Convert to degrees
    If AngleMode = 0 Then
        Value = Value * (180 / Pi)
    End If

End Sub

Private Sub TrapErrors(ErrNumber As Long)

    'Set trapped error message
    Select Case ErrNumber

        'VB Runtime Error
        Case Is > 0
            OutputString = "Error " & Err.Number & ": " & Err.Description

        Case (-5)
            OutputString = "Error: Invalid logarithm base"

        'Trapped runtime calculation error
        Case Else
            OutputString = "Error: General calculation error"

    End Select

    'Set return value
    InError = True

End Sub
