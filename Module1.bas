Attribute VB_Name = "Module1"
'Option Compare Database
Option Explicit
Global smonth As Integer
Global sday As Integer
Global myear As Long
Global mmonth As Integer
Global mday As Integer
Global myrest As Integer
Global months(12) As Integer
Global hmonth As Integer
Global hyear As Long
Global hday As Integer
Global adjust As Integer
Global but As Integer
Global cl As Boolean
Public Function MYCalendar(myd As String) As Long
On Error GoTo err_cal
Dim totday As Long
Dim monthtot As Integer
Dim i As Integer
Dim mleap As Integer
Dim fleap As Integer
Dim hundleap As Integer
Dim fhundleap As Integer
sday = CInt(Mid(myd, 1, 2))
mday = 1
'mmonth = 1
smonth = CInt(Mid(myd, 4, 2))
mmonth = smonth
myear = CInt(Mid(myd, 7, 4))
totday = 0
totday = totday + myear * 365
For i = 1 To mmonth - 1
monthtot = monthtot + months(i)
Next i
monthtot = monthtot + mday
mleap = (myear - 1) \ 4
hundleap = (myear - 1) \ 100
fhundleap = (myear - 1) \ 400
mleap = mleap - (hundleap - fhundleap)
If mmonth > 2 And (myear Mod 4) = 0 Then
 If (myear Mod 100) <> 0 Then
   mleap = mleap + 1
 ElseIf (myear Mod 400) = 0 Then
   mleap = mleap + 1
 End If
ElseIf mmonth = 2 And mday > 28 Then
  mleap = mleap + 1
'Else
'mleap = mleap - 1
End If
'MsgBox mleap
 totday = totday + monthtot + mleap
' MsgBox totday
 myrest = totday - (totday \ 7) * 7
 MYCalendar = totday
 'MsgBox myrest
err_cal:
End Function

Public Function HijriCalendar(X As Long)
Dim hyear1 As Long
Dim hday1 As Integer
Dim hmonth1 As Integer
Dim HijDiff  As Long
Dim Hijday  As Long
Dim c As Long
HijDiff = 227544
Hijday = X - HijDiff
hyear1 = Hijday \ 354
hyear = hyear1
c = Hijday - (hyear1 * 354)

Select Case c
 Case Is < 31
  hmonth = 1
  hday = c
 Case Is < 60
  hmonth = 2
  hday = c - 30
 Case Is < 90
 
 hmonth = 3
  hday = c - 59
 Case Is < 119
  hmonth = 4
  hday = c - 89
 Case Is < 149
  hmonth = 5
  hday = c - 118
 Case Is < 178
  hmonth = 6
  hday = c - 148
 Case Is < 208
  hmonth = 7
  hday = c - 177
 Case Is < 237
  hmonth = 8
  hday = c - 207
 Case Is < 267
  hmonth = 9
  hday = c - 236
 Case Is < 295
  hmonth = 10
  hday = c - 267
 Case Is < 325
  hmonth = 11
  hday = c - 295
 Case Is < 355
  hmonth = 12
  hday = c - 325
 End Select
 hday = hday + adjust
End Function

''Public Function HijOnly(myd As String) As Long
'On Error GoTo err_cal
''Dim totday As Long
''Dim monthtot As Integer
''Dim i As Integer
''sday = CInt(Mid(myd, 1, 2))
''smonth = CInt(Mid(myd, 4, 2))
''hday = 1
''hmonth = 1
'mday = CInt(Mid(myd, 1, 2))
'smonth = CInt(Mid(myd, 4, 2))
''hyear = CInt(Mid(myd, 7, 4))
''totday = hday
''totday = totday + hyear * 354
''myrest = totday - (totday \ 7) * 7 + 3
''myrest = myrest Mod 7
'MsgBox myrest
''HijOnly = totday + adjust
''End Function

Public Function HijriToEnglish(X As Long)
Dim myear400 As Long
Dim myear100 As Long
Dim myear4 As Long
Dim myear1 As Long
Dim mday1 As Integer
Dim mmonth1 As Integer
Dim Engday As Long
Const HijDiff = 227544
Const year400 = 146097
Const year100 = 36524
Const year4 = 1461
Const year1 = 365
Dim monthday  As Long
Dim temp As Integer
Dim c As Integer
Engday = X + HijDiff
myear400 = Engday \ year400
Engday = Engday - (myear400 * year400)

myear100 = Engday \ year100
Engday = Engday - myear100 * year100
myear4 = Engday \ year4
Engday = Engday - myear4 * year4

myear1 = Engday \ year1
Engday = Engday - myear1 * year1
myear1 = myear400 * 400 + myear100 * 100 + myear4 * 4 + myear1
myear = myear1
For c = 1 To 12
temp = Engday - months(c)
'MsgBox temp
If (c = 2) And (myear Mod 4) = 0 Then
 If (myear Mod 100) <> 0 Then
   temp = temp - 1
 ElseIf (myear Mod 400) = 0 Then
   temp = temp - 1
 End If
End If

 If temp <= 0 Then
  mmonth = c
  mday = Engday
  Exit For
 Else
  Engday = temp
 End If
Next c
'MsgBox myear
'MsgBox mmonth
'MsgBox mday
mday = mday
End Function

Public Function initmonth()
months(1) = 31
months(2) = 28
months(3) = 31
months(4) = 30
months(5) = 31
months(6) = 30
months(7) = 31
months(8) = 31
months(9) = 30
months(10) = 31
months(11) = 30
months(12) = 31
End Function

