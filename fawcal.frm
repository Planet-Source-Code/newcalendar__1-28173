VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Calendar "
   ClientHeight    =   6480
   ClientLeft      =   -75
   ClientTop       =   -225
   ClientWidth     =   9720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF80&
      Caption         =   "Add Picture"
      Height          =   375
      Left            =   7920
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6240
      TabIndex        =   18
      Text            =   "0"
      Top             =   6120
      Width           =   375
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Gregorian With Hijri"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   17
      Top             =   6120
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Gregorian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   6120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6120
      Width           =   735
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      Max             =   2200
      Min             =   1900
      TabIndex        =   14
      Top             =   5760
      Value           =   1900
      Width           =   9735
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "December"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   8040
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "November"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   6480
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "October"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   4800
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "September"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   3120
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "August"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1560
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "July"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   0
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "June"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   8040
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "May"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6480
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "April"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4800
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "February"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1560
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "March"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3120
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "January"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   0
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   0
      ScaleHeight     =   4995
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   360
      Width           =   9735
      Begin MSComDlg.CommonDialog commdlg 
         Left            =   840
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "Pictures (*.bmp;*.jpg)|*.bmp;*.jpg"
      End
      Begin VB.Image Image1 
         Height          =   6600
         Left            =   0
         Top             =   0
         Width           =   9675
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Hijri Adjust"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   19
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   6000
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CX, CY As Single ' Declare variables
Dim mname(12) As String
Dim hname(12) As String
Dim oldbut As Integer

Private Sub print_calendar11()
      Dim strMessage As String
      Dim strmessage1 As String
      Dim intHorSize As Integer, intVerSize As Integer
      Dim mymon, Hmon As Integer
      Dim i, j, k, X, Y, spset As Integer
      Dim prX, prY, cury As Integer
      Dim inX, inY As Integer
      Dim a, b, c As Integer
      Dim ex As Boolean
      Dim hyear1 As Integer
      Dim hmonth1 As Integer
      Dim hday1 As Integer
      Dim trest As Integer
      Dim tmessage As String
      Dim hmess As String
      Picture1.Cls
    ' shuffle
      Set Rpt = Me
      strMessage = "    Sat     Sun     Mon     Tue     Wed      Thu     Fri"
      strmessage1 = "   ______________________________________________________"
      With Picture1
      .ScaleMode = 4
          .FontName = "Courier"
          .FontSize = 12
          .ForeColor = RGB(255, 0, 0)
          .FontBold = True
         ' .ScaleMode = 4
         ' .ForeColor = RGB(255, 0, 0)
           
      End With
      intHorSize = Picture1.TextWidth(strMessage)
      intVerSize = Picture1.TextHeight(strMessage)
      Picture1.CurrentX = 1
      Picture1.CurrentY = 1
      trest = myrest
      hmonth1 = hmonth
      hday1 = hday
      hyear1 = hyear
      If (hmonth Mod 2) = 0 Then
       Hmon = 29
      Else
       Hmon = 30
      End If
      c = mmonth
     ' For b = 1 To 6
     ' For a = 1 To 2
      mymon = months(c)
      ex = False
      inX = Picture1.CurrentX
      inY = Picture1.CurrentY
      Picture1.FontBold = True
      Picture1.ForeColor = RGB(0, 0, 255)
      hmess = "                        " & hname(hmonth) & "/" & hname((hmonth Mod 12) + 1)
      If hmonth = 12 And hday + (Hmon - 1) > 30 Then
        hmess = hmess & "-" & hyear & "/" & hyear + 1 & "-"
      Else
       hmess = hmess & "-" & hyear & "-"
      'tmessage = mname(c) & hname(hmonth) & "/" & hname((hmonth Mod 12) + 1) & Chr(10) & strMessage
      End If
     ' tmessage = hmess & Chr(13) & strMessage
      Picture1.Print mname(c);
      Picture1.ForeColor = RGB(255, 0, 0)
      Picture1.Print hmess
            
      Picture1.CurrentX = inX
      Picture1.CurrentY = inY + 1
      Picture1.ForeColor = RGB(0, 0, 0)
      Picture1.Print strMessage
      'picture1.Print
      Picture1.CurrentX = inX
      Picture1.CurrentY = inY + 1
     ' rpt.ForeColor = RGB(255, 0, 0)
     Picture1.Print strmessage1
       Picture1.Print
      ''  picture1.Print
      Picture1.FontBold = False
      ''rpt.ForeColor = 187
      prX = Picture1.CurrentX
     '' prY = rpt.CurrentY - 1
      prY = Picture1.CurrentY - 2
      Picture1.CurrentX = inX + myrest
      If c = 2 Then
       If (myear Mod 100) = 0 Then
        If (myear Mod 400) = 0 Then
         mymon = 29
        End If
       ElseIf (myear Mod 4) = 0 Then
         mymon = 29
       End If
      End If
      c = c + 1
      k = 1
      X = myrest
      Select Case X
      Case 0
         Picture1.CurrentX = myrest + 1
      Case 1
       Picture1.CurrentX = myrest + 1.5
      Case 2
        Picture1.CurrentX = myrest + 2.8
      Case 3
        Picture1.CurrentX = myrest + 3.8
      Case 4
      Picture1.CurrentX = myrest + 5
       Case 5
       Picture1.CurrentX = myrest + 6.2
      Case 6
       Picture1.CurrentX = myrest + 6.7
      End Select
      Y = myrest
      spset = 4.2
      a = Day(Now)
      'MsgBox a
   For j = 1 To 6
    For i = 1 To 7
      If k = a Then
       Picture1.ForeColor = RGB(0, 255, 0)
      ' Picture1.FontBold = True
     End If
      Picture1.Print Spc(X * 7.2 + spset); Format(k, "00");
      Picture1.ForeColor = RGB(255, 0, 0)
      'Picture1.FontBold = False
      ' rpt.NumeralShapes = 2
      If hday = 1 Then
       Picture1.ForeColor = RGB(0, 0, 0)
       'Rpt.FontBold = True
      End If
      Picture1.Print Spc(1); Format(hday, "00");
       
      Picture1.FontBold = False
      Picture1.ForeColor = RGB(0, 0, 255)
     k = k + 1
     X = 0
     hday = hday + 1
     If hday > Hmon Then
      hday = 1
      hmonth = hmonth + 1
      If hmonth > 12 Then
       hmonth = 1
       hyear = hyear + 1
      End If
      If hmonth Mod 2 = 0 Then
       Hmon = 29
      Else
       Hmon = 30
      End If
     End If
     If (i + myrest) >= 7 Then Exit For
     If k > mymon Then
       ex = True
       Exit For
     End If
    Next i
 
    If k > mymon Then Exit For
    spset = 4.2
   myrest = 0
    Picture1.Print
        Picture1.Print
  Picture1.Print
    Picture1.CurrentX = inX - 0.3
    Next j
   
End Sub
Private Sub Print_calendar()
     ' Dim Rpt As Report
      Dim strMessage As String
      Dim intHorSize As Integer, intVerSize As Integer
      Dim mymon As Integer
      Dim i, a, j, k, X, spset As Integer
     ' Set Rpt = Me
     Picture1.Cls
     'shuffle
    ' Picture1.Cls
      strMessage = "      Sat     Sun      Mon      Tue      Wed      Thu      Fri"
      
      With Picture1
          'Set scale to pixels, and set FontName and FontSize properties.
          .ScaleMode = 4
          .FontName = "Courier"
          .FontSize = 12
          .ForeColor = RGB(225, 0, 0)
          .FontBold = True
      End With
      ' Horizontal width.
      intHorSize = Picture1.TextWidth(strMessage)
      ' Vertical height.
      intVerSize = Picture1.TextHeight(strMessage)
  
      ' Calculate location of text to be displayed.
   
      Picture1.CurrentX = 0
      Picture1.CurrentY = 0
     
      Picture1.Print
      Picture1.Print strMessage
       Picture1.Print
        Picture1.Print
       '  Picture1.Print
       Picture1.ForeColor = RGB(0, 0, 255)
      Picture1.CurrentX = myrest
      mymon = months(mmonth)
      If mmonth = 2 Then
       If (myear Mod 100) = 0 Then
        If (myear Mod 400) = 0 Then
         mymon = 29
        Else
         mymon = 28
        End If
       ElseIf (myear Mod 4) = 0 Then
         mymon = 29
       End If
      End If
      k = 1
      X = myrest
      Select Case X
      Case 0
         Picture1.CurrentX = myrest
      Case 1
       Picture1.CurrentX = myrest + 1.2
      Case 2
        Picture1.CurrentX = myrest + 2.2
      Case 3
        Picture1.CurrentX = myrest + 3.8
      Case 4
      Picture1.CurrentX = myrest + 5
      Case 5
       Picture1.CurrentX = myrest + 6.7
      Case 6
       Picture1.CurrentX = myrest + 7.9
      End Select
      spset = 5.8
      'MsgBox myrest
      a = Day(Now)
   For j = 1 To 6
    For i = 1 To 7
    If k > mymon Then Exit Sub
     If k = a Then
       Picture1.ForeColor = RGB(255, 0, 0)
     End If
     Picture1.Print Spc(X * 7 + spset); Format(k, "00");
      Picture1.ForeColor = RGB(0, 0, 255)
     k = k + 1
     X = 0
     If (i + myrest) >= 7 Then Exit For
    Next i
    spset = 5.8
   myrest = 0
    Picture1.Print
    Picture1.Print
     Picture1.Print
    Picture1.CurrentX = 0
    'rpt.CurrentY = j
    Next j
End Sub



Private Sub Command1_Click(Index As Integer)
Dim mx As String
Dim i As Integer
Dim totday As Long
For i = 1 To 12
command1(i).BackColor = &HFFFFC0
command1(i).FontBold = False
command1(i).FontSize = 9

Next i
command1(Index).BackColor = &HFFFFFF
command1(Index).FontBold = True
command1(Index).FontSize = 12

If Index < 10 Then
 mx = "01/" & "0" & CStr(Index) & "/" & CStr(HScroll1.Value)
Else
 mx = "01/" & CStr(Index) & "/" & CStr(HScroll1.Value)

End If
adjust = CInt(Text1)
'Form1.Caption = Space((ScaleWidth \ 2) - 8) & "Calendar- " & CStr(HScroll1.Value) & "----     By M.Saleem   "
Form1.Caption = Space(40) & "                   Calendar  " & CStr(HScroll1.Value) & "                                                           By M.Saleem   "

but = Index
'MsgBox mx
If Option1.Value = True Then
totday = MYCalendar(mx)
Print_calendar
Else
totday = MYCalendar(mx)
HijriCalendar (totday)
print_calendar11
End If
End Sub

Private Sub command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = but Then Exit Sub
Command1_Click (Index)
but = Index
End Sub

Private Sub Command2_Click()
End
End Sub


Private Sub Command3_Click()
If Command3.Caption = "Clear" Then
Command3.Caption = "AddPhoto"
 Image1.Picture = LoadPicture("")
 command1(but).Value = True
Else
 Command3.Caption = "Clear"
 commdlg.ShowOpen
 Image1.Picture = LoadPicture(commdlg.FileName)
 command1(but).Value = True
End If
End Sub

Private Sub Form_Activate()
cl = True
Dim X As Integer
X = Month(Now)
command1(X).Value = True
End Sub

Private Sub Form_Load()
initmonth
AssignMonths
but = 1
HScroll1.Value = Year(Now)
Label1.Caption = HScroll1.Value

End Sub

Private Sub HScroll1_Change()
Label1.Caption = CStr(HScroll1.Value)
command1(but).Value = True
End Sub



Private Sub Option1_Click()
command1(but).Value = True

End Sub

Private Sub Option2_Click()
command1(but).Value = True

End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If Text1 = "" Then
Text1 = 0
End If
If Abs(CInt(Text1)) > 1 Then
Text1 = 0
End If
command1(but).Value = True

End Sub



Public Sub AssignMonths()
      mname(1) = "    January  "
      mname(2) = "    February  "
      mname(3) = "    March  "
      mname(4) = "    April  "
      mname(5) = "     May  "
      mname(6) = "    June  "
      mname(7) = "    July  "
      mname(8) = "    August  "
      mname(9) = "    September  "
      mname(10) = "    October  "
      mname(11) = "    November  "
      mname(12) = "    December  "
      '***************
       hname(1) = "ãÍÑã"
      hname(2) = " ÕÝÑ"
      hname(3) = "ÑÈíÚ ÇáÇæá"
      hname(4) = "ÑÈíÚ ÇÎÑ"
      hname(5) = "ÌãÇÏ Çæá"
      hname(6) = "ÌãÇÏ ÇÎÑ"
      hname(7) = "ÑÌÈ"
      hname(8) = "ÔÚÈÇä"
      hname(9) = "ÑãÖÇä"
      hname(10) = "ÔæÇá"
      hname(11) = "Ðæ ÇáÞÚÏÉ"
      hname(12) = "Ðæ ÇáÍÌÉ"
      '**************
End Sub
