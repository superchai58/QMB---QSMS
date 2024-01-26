Attribute VB_Name = "mdlDec2Base"
Option Explicit


  Public Function Base_B_EquivOf_A(Dec_A, Base_B, Optional Digits As String = "0123456789ABCDEFGHJKLMNPQRSTUVWXYZ")
' Compute (Base_B) equivalent of (Dec_A) value with
' up to 28 digits

  Dim DecVal, Accum, Radix, Qi As String
  Dim Q, R
  Dim i As Integer
    
' Define valid digit symbols up to base 36.  For the
' letter "O", the lower case letter "o" was used to avoid
' confusion with the digit zero "0".  All of the other
' letters are rendered in upper case.

'mark by Alex Wang 2003/6/5, instead of the input argument from caller
  'Dim Digits As String
   '   Digits = "0123456789ABCDEFGHIJKLMNoPQRSTUVWXYZ"
    
' Check for non-numeric decimal argument
  If IsNumeric(Dec_A) = False Then
     Base_B_EquivOf_A = "ERROR: Argument is not an integer."
     Exit Function
  End If
    
' Check if either argument has a zero value
  If Val(Dec_A) = 0 Then Dec_A = 0
     DecVal = CDec(Dec_A)
  If Val(Base_B) = 0 Then Base_B = 0
     Radix = CDec(Base_B)
                  
' Check for non-integer decimal argument
  If IsNumeric(DecVal) = True Then
     If InStr(DecVal, ".") > 0 Then
        Base_B_EquivOf_A = "ERROR: Argument is not an integer."
        Exit Function
     End If
  End If
     
' Check for valid base (radix) argument
  If Base_B < 2 Or Base_B > 36 Then
     Base_B_EquivOf_A = "ERROR: Base (Radix) must be in the range from 2 to 36."
     Exit Function
  Else
   
  End If

   
' Compute and accumulate the digits of the (Base_B)
' equivalent of (Dec_A) one at a time.
        Q = 1
  While Q > 0

      Q = DecVal / Radix
     Qi = Trim(Q)
      i = InStr(Qi, ".")
   If i > 0 Then Qi = Left(Qi, i - 1)
   If Qi = "" Then Qi = "0"
      R = DecVal - Radix * CDec(Qi)
  Accum = Mid(Digits, R + 1, 1) & Accum
 DecVal = CDec(Qi)
   If Val(Qi) = 0 Then Q = 0 ' Check if done yet
   
  Wend

' Return the computed (Base_B) equivalent of (Dec_A)
  Base_B_EquivOf_A = Accum

  End Function

