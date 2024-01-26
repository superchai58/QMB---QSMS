Attribute VB_Name = "mdlConvertDec2Base"
Option Explicit
 
Public Const ERROR_NUMBER = 13&
 
Public Const HexChars = "0123456789ABCDEF"
Public Const AlphaChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Const Apple34Chars = "0123456789ABCDEFGHJKLMNPQRSTUVWXYZ"

 
Public Enum Bases
    ebBinary = 2&
    ebOctal = 8&
    ebDecimal = 10&
    ebHexadecimal = 16&
    ebAlphabet = 26&
    ebApple34 = 34&
    ebSexagesimal = 60&     'Base 60, e.g. time
End Enum
Public Function ConvertBase2Dec(ByVal Number As String, _
    ByVal Base As Bases) As Double
'Convert the number from the specified base to decimal
Dim dblTemp As Double
Dim strDigit As String, lngDigit As Long, i As Long
Dim lngPwr As Long, lngSign As Long, lngDigitSize
    
    If Base < 2 Then
        Err.Raise ERROR_NUMBER, "ConvertBase2Dec", "Invalid Base"
    Else
        lngDigitSize = DigitLength(Base)
        lngPwr = 0
        lngSign = 1
        i = 1
        Do Until i > Len(Number)
            strDigit = Mid$(Number, i, lngDigitSize)
            If Left$(strDigit, 1) = "." Then
                i = i + 1
                If lngPwr = 0 Then
                    lngPwr = 1
                Else
                    Err.Raise ERROR_NUMBER, "ConvertBase2Dec", _
                        "More than one decimal point"
                End If
            ElseIf Left$(strDigit, 1) = "-" Then
                i = i + 1
                If lngPwr = 0 And dblTemp = 0 Then
                    lngSign = -lngSign
                Else
                    Err.Raise ERROR_NUMBER, "ConvertBase2Dec", _
                        "Invalid negation"
                End If
            Else
                i = i + lngDigitSize
                lngDigit = DeconvertDigit(strDigit, Base)
                dblTemp = dblTemp * Base + lngDigit
                lngPwr = lngPwr * Base
            End If
        Loop
        If lngPwr > 1 Then
            ConvertBase2Dec = CDbl(lngSign) * (dblTemp / CDbl(lngPwr))
        Else
            ConvertBase2Dec = CDbl(lngSign) * dblTemp
        End If
    End If
End Function
Public Function DigitLength(Base As Bases) As Long
'Return the length of a digit in a given base
    Select Case Base
        Case ebSexagesimal:
            DigitLength = 3
        
        'Add other special cases here
        
        Case Else   'ebBinary, ebOctal, ebDecimal, ebHexadecimal
            DigitLength = 1
    End Select
End Function
Public Function DeconvertDigit(strDigit As String, Base As Bases) As Long
'Convert a single digit from the specified base to decimal
Dim lngTemp As Long
    Select Case Base
        Case ebBinary, ebOctal, ebDecimal:
            If IsNumeric(strDigit) Then
                lngTemp = CLng(strDigit)
                If lngTemp < Base Then
                    DeconvertDigit = lngTemp
                Else
                    Err.Raise ERROR_NUMBER, "DeconvertDigit", _
                        "Invalid digit for base"
                End If
            Else
                Err.Raise ERROR_NUMBER, "DeconvertDigit", "Invalid character"
            End If
            
        Case ebHexadecimal:
            lngTemp = InStr(1, HexChars, UCase$(strDigit))
            If lngTemp = 0 Then
                Err.Raise ERROR_NUMBER, "DeconvertDigit", _
                    "Invalid digit for base"
            Else
                DeconvertDigit = lngTemp - 1
            End If
            
       Case ebAlphabet:
           lngTemp = InStr(1, AlphaChars, UCase$(strDigit))
           If lngTemp = 0 Then
               Err.Raise ERROR_NUMBER, "DeconvertDigit", _
                   "Invalid Alpha Character"
           Else
               DeconvertDigit = lngTemp - 1
           End If
            
       Case ebApple34:
           lngTemp = InStr(1, Apple34Chars, UCase$(strDigit))
           If lngTemp = 0 Then
               Err.Raise ERROR_NUMBER, "DeconvertDigit", _
                   "Invalid Apple34Chars Character"
           Else
               DeconvertDigit = lngTemp - 1
           End If
           
        Case ebSexagesimal:
            If Len(strDigit) = 3 Then
                If Right$(strDigit, 1) = ":" And IsNumeric(Left$(strDigit, _
                    2)) Then
                    lngTemp = CLng(Left$(strDigit, 2))
                    If lngTemp < Base Then
                        DeconvertDigit = lngTemp
                    Else
                        Err.Raise ERROR_NUMBER, "DeconvertDigit", _
                            "Invalid digit for base"
                    End If
                Else
                    Err.Raise ERROR_NUMBER, "DeconvertDigit", _
                        "Invalid digit for base"
                End If
            Else
                Err.Raise ERROR_NUMBER, "DeconvertDigit", _
                    "Invalid digit for base"
            End If
        
        'Add other bases here
        
        Case Else:
            Err.Raise ERROR_NUMBER, "DeconvertDigit", "Unknown base"
    End Select
End Function
Public Function Floor(ByVal Number As Double) As Double
If Int(Number) > Number Then
Floor = Int(Number) - 1
Else
Floor = Int(Number)
End If
End Function
