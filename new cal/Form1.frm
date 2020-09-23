VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Save this"
   ClientHeight    =   12555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17040
   LinkTopic       =   "Form1"
   ScaleHeight     =   837
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1136
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check6 
      Caption         =   "Week No"
      Height          =   255
      Left            =   9120
      TabIndex        =   32
      Top             =   840
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   11760
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Cutting Line"
      Height          =   255
      Left            =   9120
      TabIndex        =   31
      Top             =   480
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Sun rise/set"
      Height          =   255
      Left            =   10320
      TabIndex        =   30
      Top             =   120
      Width           =   1455
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   240
      TabIndex        =   29
      Text            =   "Combo5"
      Top             =   960
      Width           =   855
   End
   Begin VB.ComboBox Text1 
      Height          =   315
      Left            =   2040
      TabIndex        =   28
      Text            =   "Combo5"
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   3120
      Sorted          =   -1  'True
      TabIndex        =   27
      Text            =   "Combo2"
      Top             =   600
      Width           =   2775
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   6000
      Sorted          =   -1  'True
      TabIndex        =   26
      Text            =   "Combo2"
      Top             =   600
      Width           =   2775
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   25
      Text            =   "Combo2"
      Top             =   600
      Width           =   2775
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Include diary"
      Height          =   255
      Left            =   8880
      TabIndex        =   24
      Top             =   120
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Outer ring"
      Height          =   255
      Left            =   7440
      TabIndex        =   23
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Moon phases"
      Height          =   255
      Left            =   5760
      TabIndex        =   22
      Top             =   120
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin Project1.IoxContainer Iox1 
      Height          =   11055
      Left            =   120
      TabIndex        =   20
      Top             =   1440
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   19500
      Begin VB.PictureBox Pic1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   11895
         Left            =   0
         ScaleHeight     =   791
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1120
         TabIndex        =   21
         Top             =   0
         Width           =   16830
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save year"
      Height          =   375
      Left            =   4560
      TabIndex        =   19
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save this"
      Height          =   375
      Left            =   3360
      TabIndex        =   18
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H000000FF&
      Height          =   255
      Index           =   15
      Left            =   16560
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   16
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H009504FF&
      Height          =   255
      Index           =   14
      Left            =   16320
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   15
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00C004FF&
      Height          =   255
      Index           =   13
      Left            =   16080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   14
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00FF00FF&
      Height          =   255
      Index           =   12
      Left            =   15840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   13
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00FF04B4&
      Height          =   255
      Index           =   11
      Left            =   15600
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   12
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00FF0000&
      Height          =   255
      Index           =   10
      Left            =   15360
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   11
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00FF9B04&
      Height          =   255
      Index           =   9
      Left            =   15120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H00FFFF00&
      Height          =   255
      Index           =   8
      Left            =   14880
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0000FF00&
      Height          =   255
      Index           =   7
      Left            =   14640
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0004FF8E&
      Height          =   255
      Index           =   6
      Left            =   14400
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0004FFBA&
      Height          =   255
      Index           =   5
      Left            =   14160
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Index           =   4
      Left            =   13920
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0004DAFF&
      Height          =   255
      Index           =   3
      Left            =   13680
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H0004A7FF&
      Height          =   255
      Index           =   2
      Left            =   13440
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H000482FF&
      Height          =   255
      Index           =   1
      Left            =   13200
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picColors 
      BackColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   12960
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim HasLoaded As Boolean
Dim PI As Double
Dim agDay As Double
Dim xx As Long
Dim yy As Long
Dim lg As Long
Dim trAg As Double
Dim stLg As Double
Dim xad As Long
Dim ndYear As Long
Dim ndMonth As Long
Dim ndfMonth As Long
Dim segP() As POINTAPI
Dim lMoon() As POINTAPI
Dim dMoon() As POINTAPI
Dim gcol As Long
Dim rd As Long, gd As Long, bd As Long
Dim r2 As Long, g2 As Long, b2 As Long
Dim r3 As Long, g3 As Long, b3 As Long
Dim h As Double, h2 As Long, c As Long
Dim lgPD As Long
Dim cc As Long
Dim cc2 As Long
Dim WkN As Long

Private Type RGBset
    Angle As Integer
    R(0 To 15)
    G(0 To 15)
    b(0 To 15)
    Count As Integer
End Type
Dim gradtemp As RGBset

Private Type POINTAPI
        x As Long
        y As Long
End Type


Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function PolyDraw Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, lpbTypes As Byte, ByVal cCount As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function PolylineTo Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
Private Declare Function PolyPolygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
Dim AA2 As New LineGS

Private Type tAppoint
    ID As Long
    Hol_Date As Date
    Comments As String
    Task As Boolean
    Status As Long
End Type

Private hList() As tAppoint

Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Const BS_NULL = 1
Private Const BS_HOLLOW = BS_NULL

Dim FontsAdded As Boolean
Private cSun As clsSunrise

Private Sub DrawWkNo(tYear As Long, tMonth As Long)
Dim cc As Long
Dim cc2 As Long
Dim wkC As Long
Dim nWks As Long
Dim wkP() As POINTAPI
Dim tmpN As Long
Dim MonSun As Single
Dim tLine As Long

For cc = 1 To ndMonth
    
    If nWks = 0 Then
        tmpN = 8 - Weekday(DateSerial(tYear, tMonth, cc), vbMonday)
        If cc = 1 And tmpN < 7 Then tLine = 1
        'Debug.Print Weekday(DateSerial(tYear, tMonth, cc), vbMonday)
        If cc + tmpN - 1 > ndMonth Then
            tLine = 2
            tmpN = tmpN - (cc + tmpN - ndMonth) + 1
        End If
        tmpN = tmpN * 2 + 1
        ReDim wkP(tmpN)
        wkC = 0
        
    End If
        
        If Weekday(DateSerial(tYear, tMonth, cc), vbMonday) = 1 Then
            MonSun = -0.5
        ElseIf Weekday(DateSerial(tYear, tMonth, cc), vbMonday) = 7 Then
            MonSun = -1.5
        Else
            MonSun = -1
        End If
        
        trAg = ((cc + MonSun) * agDay) * PI / 180
        wkP(tmpN - wkC).x = (lg * stLg + 10) * Cos(trAg) - xad
        wkP(tmpN - wkC).y = (lg * stLg + 10) * Sin(trAg)
        
        trAg = ((cc + MonSun) * agDay) * PI / 180
        wkP(wkC).x = (lg * stLg + 10 + WkN) * Cos(trAg) - xad
        wkP(wkC).y = (lg * stLg + 10 + WkN) * Sin(trAg)
               
        wkC = wkC + 1
        nWks = nWks + 1
    
    If Weekday(DateSerial(tYear, tMonth, cc), vbMonday) = 7 Or cc = ndMonth Then
        
        If Weekday(DateSerial(tYear, tMonth, cc), vbMonday) = 1 Then
            MonSun = -1
        'ElseIf Weekday(DateSerial(tYear, tMonth, cc), vbMonday) = 7 Then
        '    MonSun = -1.5
        'Else
        '    MonSun = -1
        End If
        
        trAg = ((cc + 1 + MonSun) * agDay) * PI / 180
        wkP(tmpN - wkC).x = (lg * stLg + 10) * Cos(trAg) - xad
        wkP(tmpN - wkC).y = (lg * stLg + 10) * Sin(trAg)
        
        trAg = ((cc + 1 + MonSun) * agDay) * PI / 180
        wkP(wkC).x = (lg * stLg + 10 + WkN) * Cos(trAg) - xad
        wkP(wkC).y = (lg * stLg + 10 + WkN) * Sin(trAg)
        
    'Pic1.ForeColor = RGB(250, 230, 210)
    Pic1.ForeColor = RGB(250, 230, 210)
    Pic1.FillColor = RGB(250, 230, 210)
    Pic1.DrawWidth = 1
    Pic1.DrawStyle = vbTransparent
    Call Polygon(Pic1.hdc, wkP(0), tmpN + 1)
    Pic1.DrawStyle = vbSolid
    
    For cc2 = 0 To tmpN - 1
        If cc2 = (tmpN - 1) / 2 And tLine = 2 Then
        Else
        AA2.LineGP Pic1.hdc, wkP(cc2).x, wkP(cc2).y, wkP(cc2 + 1).x, wkP(cc2 + 1).y, 0
        End If
    Next
    If tLine <> 1 Then AA2.LineGP Pic1.hdc, wkP(0).x, wkP(0).y, wkP(tmpN).x, wkP(tmpN).y, 0
    tLine = 0
    'Pic1.DrawWidth = 1
    
'    If Weekday(DateSerial(tYear, tMonth, cc + 1), vbMonday) = 4 Then
'        trAg = ((cc + 0.7) * agDay) * PI / 180
'        xx = (lg * stLg + 13 + WkN / 2) * Cos(trAg) - xad
'        yy = (lg * stLg + 13 + WkN / 2) * Sin(trAg)
'        'Print wk no
'        Pic1.Font = Combo4
'        Pic1.ForeColor = 0
'        Call cFont(Pic1.hdc, Format(DateSerial(tYear, tMonth, cc + 1), "ww", vbMonday, vbFirstJan1), xx, yy - 7, 12, -90 + (ndfMonth * agDay), True)
'    End If
    nWks = 0
    wkC = 0
    Do
    'Debug.Print wkC, wkP(wkC).x, wkP(wkC).y, UBound(wkP), tmpN
    wkC = wkC + 1
    Loop Until wkC = UBound(wkP) + 1
    wkC = 0
    'ReDim wkP(0)
    End If
Next

Pic1.Font = Combo4
Pic1.ForeColor = 0
For cc = 1 To ndMonth
    If Weekday(DateSerial(tYear, tMonth, cc), vbMonday) = 4 Then
        trAg = ((cc + 0.7 - 1) * agDay) * PI / 180
        xx = (lg * stLg + 13 + WkN / 2) * Cos(trAg) - xad
        yy = (lg * stLg + 13 + WkN / 2) * Sin(trAg)
        'Print wk no
        Call cFont(Pic1.hdc, Format(DateSerial(tYear, tMonth, cc), "ww", vbMonday, vbFirstJan1), xx, yy - 7, 12, -90 + (ndfMonth * agDay), True)
    End If
Next
End Sub

Function MoonPhase(dInDate As Date) As Integer
  Dim lD As Long
  Dim dd As Double
 
  lD = DateDiff("d", "January 1, 2001", dInDate)
  dd = 0.20439731 + lD * 0.03386319269
  lD = Int(dd)
  dd = dd - lD
  lD = 360 * dd
  If lD < 0 Then lD = lD + 360
  lD = lD \ 2
  'lD = lD * 2
  'Debug.Print lD
  'If lD > 179 Then lD = 179 - lD
  MoonPhase = lD * 2
  
  'Debug.Print 179 - (179 - MoonPhase), MoonPhase
  If MoonPhase > 179 Then MoonPhase = 179 + (179 - MoonPhase)
'  Debug.Print MoonPhase
  'MoonPhase = lD
  
End Function

Sub FindInfo(TheDate As Date, NumDays As Long)
    'Dim s$, p&, X&, Y As Long, xOff&, yOff&, n&, h&
    Dim tYear As Long
    'Dim NumDays As Long
    Dim tmpDate As Date
    Dim tmpDate2 As Date
    Dim tmpString As String
    Dim TmpEnd As Date
    Dim TmpInter As Long
    Dim TmpNum As String
    Dim oitemTmp As String
    Dim hcc As Long
    
    hcc = 0
    Erase hList
    ReDim hList(hcc)
        
    'CanDraw = False
    'NumDays = 60
          
    Dim oApp As Outlook.Application
    Dim oNspc As NameSpace
    Dim oItm As TaskItem
    Dim myItem As TaskItem
    Dim oItm2 As AppointmentItem
    Dim oItm3 As ContactItem
    
    Set oApp = CreateObject("Outlook.Application")
    Set oNspc = oApp.GetNamespace("MAPI")

    
    For Each oItm2 In oNspc.GetDefaultFolder(olFolderCalendar).Items
        If oItm2.Categories <> "Personal" Then
        'Debug.Print oItm2, oItm2.Start, oItm2.GetRecurrencePattern, oItm2.IsRecurring, oItm2.End
        'Debug.Print oItm2, oItm2.Categories
        If oItm2.IsRecurring Then
            If oItm2.GetRecurrencePattern = olRecursYearly Then
                If Month(TheDate) > Month(oItm2.Start) Then
                    tmpDate = DateSerial(Year(TheDate) + 1, Month(oItm2.Start), Day(oItm2.Start))
                    'Debug.Print "HERE"
                Else
                    tmpDate = DateSerial(Year(TheDate), Month(oItm2.Start), Day(oItm2.Start))
                    'Debug.Print "HERE2"
                End If
                
                'Debug.Print tmpDate
                If tmpDate >= TheDate And tmpDate <= (TheDate + NumDays) Then
                    If oItm2.GetRecurrencePattern = olRecursYearly Then
                    oitemTmp = oItm2
                    If Left(oItm2, 9) = "Birthday:" Then
                        'Debug.Print "BIRTHDAY"
                        oitemTmp = Mid(oitemTmp, 11, Len(oitemTmp) - 10)
                    End If
                    TmpNum = Right(oitemTmp, 6)
                    ReDim Preserve hList(hcc)
                    If Mid(TmpNum, 1, 1) = "(" And Mid(TmpNum, 6, 1) = ")" Then
                        hList(hcc).Comments = Mid(oitemTmp, 1, Len(oitemTmp) - 6) & "(" & Year(tmpDate) - CLng(Mid(TmpNum, 2, 4)) & ")"
                    Else
                        hList(hcc).Comments = oitemTmp
                    End If
                    hList(hcc).Hol_Date = DateSerial(Year(tmpDate), Month(oItm2.Start), Day(oItm2.Start))
                    hcc = hcc + 1
                    End If
                End If
    
            ElseIf oItm2.GetRecurrencePattern = olRecursDaily Then
                'Debug.Print oItm2.GetRecurrencePattern.Interval
                If oItm2.GetRecurrencePattern.NoEndDate Then
                    TmpEnd = TheDate + NumDays
                Else
                    If oItm2.GetRecurrencePattern.PatternEndDate <= TheDate + NumDays Then
                        TmpEnd = oItm2.GetRecurrencePattern.PatternEndDate
                    Else
                        TmpEnd = TheDate + NumDays
                    End If
                End If
                
                TmpInter = oItm2.GetRecurrencePattern.Interval
                For tmpDate = oItm2.GetRecurrencePattern.PatternStartDate To TmpEnd
                    If TmpInter = oItm2.GetRecurrencePattern.Interval Then
                        ReDim Preserve hList(hcc)
                        hList(hcc).Hol_Date = tmpDate
                        hList(hcc).Comments = oItm2
                        hcc = hcc + 1
                        TmpInter = 0
                    End If
                    TmpInter = TmpInter + 1
                Next
            ElseIf oItm2.GetRecurrencePattern = olRecursMonthNth Then
                        ReDim Preserve hList(hcc)
                        hList(hcc).Hol_Date = tmpDate
                        hList(hcc).Comments = oItm2
                        hcc = hcc + 1
                'Debug.Print "nth Monthly"
            ElseIf oItm2.GetRecurrencePattern = olRecursMonthly Then
                'Debug.Print "Monthly", oItm2.GetRecurrencePattern.DayOfMonth, oItm2.GetRecurrencePattern.Interval, oItm2.GetRecurrencePattern.PatternStartDate, oItm2.GetRecurrencePattern.PatternEndDate
                tmpDate = DateSerial(Year(oItm2.GetRecurrencePattern.PatternStartDate), Month((oItm2.GetRecurrencePattern.PatternStartDate)), oItm2.GetRecurrencePattern.DayOfMonth)
                Do
                    If DateSerial(Year(tmpDate), Month(tmpDate), oItm2.GetRecurrencePattern.DayOfMonth) >= TheDate And DateSerial(tYear, Month(tmpDate), oItm2.GetRecurrencePattern.DayOfMonth) <= TheDate + NumDays Then
                        ReDim Preserve hList(hcc)
                        hList(hcc).Hol_Date = tmpDate
                        hList(hcc).Comments = oItm2
                        hcc = hcc + 1
                    End If
                    tmpDate = DateAdd("m", 1, tmpDate)
                Loop Until tmpDate > TheDate + NumDays Or tmpDate >= oItm2.GetRecurrencePattern.PatternEndDate
            ElseIf oItm2.GetRecurrencePattern = olRecursWeekly Then
                'Debug.Print oItm2, oItm2.Start, oItm2.GetRecurrencePattern.DayOfWeekMask, oItm2.GetRecurrencePattern.PatternStartDate
                'Debug.Print oItm2.GetRecurrencePattern.PatternStartDate, oItm2.GetRecurrencePattern.PatternEndDate
                If oItm2.GetRecurrencePattern.NoEndDate Then
                    TmpEnd = TheDate + NumDays
                Else
                    If oItm2.GetRecurrencePattern.PatternEndDate <= TheDate + NumDays Then
                        TmpEnd = oItm2.GetRecurrencePattern.PatternEndDate
                    Else
                        TmpEnd = TheDate + NumDays
                    End If
                End If
                
                'Debug.Print ((theDate - oItm2.Start) Mod 7)
                For tmpDate2 = TheDate - ((TheDate - oItm2.Start) Mod 7) To TmpEnd Step 7
                'Debug.Print tmpDate2
                'Debug.Print tmpDate2, oItm2.GetRecurrencePattern.Interval, oItm2.Start - tmpDate, (tmpDate2 Mod oItm2.Start) Mod oItm2.GetRecurrencePattern.Interval
                    'Debug.Print (2 ^ (Weekday(tmpDate) - 1)), oItm2.GetRecurrencePattern.DayOfWeekMask Or (2 ^ (Weekday(tmpDate) - 1)), oItm2.GetRecurrencePattern.DayOfWeekMask
                Debug.Print tmpDate2, oItm2.Start, (tmpDate2 Mod oItm2.Start) Mod oItm2.GetRecurrencePattern.Interval
                If (tmpDate2 Mod oItm2.Start) Mod oItm2.GetRecurrencePattern.Interval = 0 Then
                'Debug.Print tmpDate2, (tmpDate2 Mod oItm2.Start) Mod oItm2.GetRecurrencePattern.Interval
                'End If
                For tmpDate = tmpDate2 To tmpDate2 + 6
                    If (oItm2.GetRecurrencePattern.DayOfWeekMask Or (2 ^ (Weekday(tmpDate) - 1))) = oItm2.GetRecurrencePattern.DayOfWeekMask Then
                        ReDim Preserve hList(hcc)
                        hList(hcc).Hol_Date = tmpDate
                        hList(hcc).Comments = oItm2
                        hcc = hcc + 1
                    End If
                Next
                End If
                Next
            Else
                        ReDim Preserve hList(hcc)
                        hList(hcc).Hol_Date = tmpDate
                        hList(hcc).Comments = "ERROR " & oItm2
                        hcc = hcc + 1
            End If
        Else
            If oItm2.Start >= TheDate And oItm2.Start <= TheDate + NumDays Then
                    ReDim Preserve hList(hcc)
                    hList(hcc).Hol_Date = DateSerial(Year(oItm2.Start), Month((oItm2.Start)), Day((oItm2.Start)))
                    hList(hcc).Comments = oItm2
                    hcc = hcc + 1
            End If
        End If
        End If
    Next oItm2

    
    Set myItem = Nothing
    Set oItm = Nothing
    Set oItm2 = Nothing
    Set oNspc = Nothing
    Set oApp = Nothing
End Sub

Private Function Red(ByVal Color As Long) As Integer
    Red = Color Mod &H100
End Function
Private Function Green(ByVal Color As Long) As Integer
    Green = (Color \ &H100) Mod &H100
End Function
Private Function Blue(ByVal Color As Long) As Integer
    Blue = (Color \ &H10000) Mod &H100
End Function

Function cFont(thdc As Long, tmpPrint As String, xxx As Long, yyy As Long, FontSize As Long, rAng As Long, Optional IsCentered As Boolean = False)
Dim F As LOGFONT, hPrevFont As Long, hFont As Long, fontname As String
Dim PI As Double
Dim trAg As Double
Dim tmpLength As Double
Dim tmpDir As Double
Dim xx As Long
Dim yy As Long
Dim txtHeight As Long
Dim txtWidth As Long

PI = 3.14159265358979
Pic1.FontSize = FontSize
    FontSize = Val(FontSize)
    F.lfEscapement = 10 * Val(rAng) 'rotation angle, in tenths
    F.lfOrientation = 900
    fontname = Pic1.fontname + Chr$(0)
    F.lfFacename = fontname
    F.lfHeight = (FontSize * -20) / Screen.TwipsPerPixelY
    hFont = CreateFontIndirect(F)
    hPrevFont = SelectObject(thdc, hFont)
    txtHeight = Pic1.TextHeight(tmpPrint)
    txtWidth = Pic1.TextWidth(tmpPrint)
    tmpLength = Sqr(txtWidth ^ 2 + txtHeight ^ 2)
    tmpDir = Atn(txtHeight / IIf(txtWidth = 0, 1, txtWidth)) * 180 / PI
    trAg = (tmpDir - rAng) * PI / 180
    xx = (tmpLength) * Cos(trAg)
    yy = (tmpLength) * Sin(trAg)
    Pic1.CurrentX = xxx - IIf(IsCentered = True, (xx / 2), 0)
    Pic1.CurrentY = yyy - IIf(IsCentered = True, (yy / 2), 0)
    Pic1.Print tmpPrint
    hFont = SelectObject(thdc, hPrevFont)
    DeleteObject hFont
    cFont = tmpLength
End Function

Sub DoWorkOuts(tYear As Long, tMonth As Long)
Dim tmpDate As Date

PI = 3.14159265358979
ndYear = (DateSerial(tYear + 1, 1, 1) - 1) - (DateSerial(tYear, 1, 1))
ndMonth = Day((DateSerial(tYear, tMonth + 1, 1) - 1))
agDay = 360 / (ndYear + 1)
stLg = 0.6
lg = (Pic1.ScaleWidth * 1.35) - 70 - 200
xad = Pic1.ScaleWidth * 0.4
trAg = ((1) * agDay) * PI / 180
If tMonth > 1 Then
    ndfMonth = 1
    tmpDate = DateSerial(tYear, 1, 1)
    Do
        ndfMonth = ndfMonth + 1
        tmpDate = tmpDate + 1
    Loop Until Month(tmpDate) = tMonth
End If
End Sub

Sub Draw1(tMonth As Long, tYear As Long)
For cc = 0 To ndMonth - 1
    'outer ring
    If Check2.Value = vbChecked Then
        trAg = (cc * agDay) * PI / 180
        segP(1).x = (lg + 50) * Cos(trAg) - xad
        segP(1).y = (lg + 50) * Sin(trAg)
        trAg = ((cc + 1) * agDay) * PI / 180
        segP(2).x = (lg + 50) * Cos(trAg) - xad
        segP(2).y = (lg + 50) * Sin(trAg)
        AA2.LineGP Pic1.hdc, segP(1).x, segP(1).y, segP(2).x, segP(2).y, 0
    End If
       
For cc2 = lg - 80 To lg
    trAg = (cc * agDay) * PI / 180
    segP(1).x = (cc2) * Cos(trAg) - xad
    segP(1).y = (cc2) * Sin(trAg)
    trAg = ((cc + 1) * agDay) * PI / 180
    segP(2).x = (cc2) * Cos(trAg) - xad
    segP(2).y = (cc2) * Sin(trAg)
    gcol = (255 / 80) * (lg - cc2)
    If tMonth Mod 2 = 0 Then
        AA2.LineGP Pic1.hdc, segP(1).x, segP(1).y, segP(2).x, segP(2).y, RGB(255, 255 - gcol, 255 - gcol)
    Else
        AA2.LineGP Pic1.hdc, segP(1).x, segP(1).y, segP(2).x, segP(2).y, RGB(255, 255, 255 - gcol)
    End If
Next
    
    trAg = (cc * agDay) * PI / 180
    segP(0).x = (lg * stLg) * Cos(trAg) - xad
    segP(0).y = (lg * stLg) * Sin(trAg)
    trAg = (cc * agDay) * PI / 180
    segP(1).x = (lg) * Cos(trAg) - xad
    segP(1).y = (lg) * Sin(trAg)
    trAg = ((cc + 1) * agDay) * PI / 180
    segP(2).x = (lg) * Cos(trAg) - xad
    segP(2).y = (lg) * Sin(trAg)
    trAg = ((cc + 1) * agDay) * PI / 180
    segP(3).x = (lg * stLg) * Cos(trAg) - xad
    segP(3).y = (lg * stLg) * Sin(trAg)
    Pic1.FillStyle = vbSolid
    Pic1.FillColor = QBColor(14)
    Pic1.ForeColor = QBColor(14)
    trAg = ((cc + 0.9) * agDay) * PI / 180
    lgPD = lg - (((lg - (lg * stLg) - 40) / ndMonth) * cc) - 40
    xx = (lgPD) * Cos(trAg) - xad
    yy = (lgPD) * Sin(trAg)
    trAg = ((cc + 0.5) * agDay) * PI / 180
    lgPD = lg - 10
    xx = (lgPD) * Cos(trAg) - xad
    yy = (lgPD) * Sin(trAg)
Next
End Sub

Sub Draw2(tMonth As Long, tYear As Long)
Dim tmp1 As Long
Dim tmp2 As Long
Dim ocr As Long
Dim tmpcol As Long

For cc = 0 To ndMonth - 1
    'outer ring
    If Check2.Value = vbChecked Then
        For ocr = 20 To 30
            trAg = (cc * agDay) * PI / 180
            segP(1).x = (lg + ocr) * Cos(trAg) - xad
            segP(1).y = (lg + ocr) * Sin(trAg)
            trAg = ((cc + 1) * agDay) * PI / 180
            segP(2).x = (lg + ocr) * Cos(trAg) - xad
            segP(2).y = (lg + ocr) * Sin(trAg)
            tmpcol = (255 / 10) * (ocr - 20)
            AA2.LineGP Pic1.hdc, segP(1).x, segP(1).y, segP(2).x, segP(2).y, RGB(tmpcol, tmpcol, tmpcol)
        Next
    End If
    
    If tMonth Mod 2 = 0 Then
        Pic1.FillStyle = vbSolid
        Pic1.FillColor = QBColor(14)
        Pic1.ForeColor = QBColor(14)
        tmp1 = 40
        tmp2 = 0
    Else
        Pic1.FillStyle = vbSolid
        Pic1.FillColor = RGB(255, 0, 0)
        Pic1.ForeColor = RGB(255, 0, 0)
        tmp1 = 80
        tmp2 = 40
    End If
    
    trAg = (cc * agDay) * PI / 180
    segP(0).x = (lg - tmp1) * Cos(trAg) - xad
    segP(0).y = (lg - tmp1) * Sin(trAg)
    trAg = (cc * agDay) * PI / 180
    segP(1).x = (lg - tmp2) * Cos(trAg) - xad
    segP(1).y = (lg - tmp2) * Sin(trAg)
    trAg = ((cc + 1) * agDay) * PI / 180
    segP(2).x = (lg - tmp2) * Cos(trAg) - xad
    segP(2).y = (lg - tmp2) * Sin(trAg)
    trAg = ((cc + 1) * agDay) * PI / 180
    segP(3).x = (lg - tmp1) * Cos(trAg) - xad
    segP(3).y = (lg - tmp1) * Sin(trAg)
   
    Call Polygon(Pic1.hdc, segP(0), 4)
    AA2.LineGP Pic1.hdc, segP(1).x, segP(1).y, segP(2).x, segP(2).y, Pic1.ForeColor
Next
End Sub


Sub DrawMonth(tYear As Long, tMonth As Long)
'Dim cc As Long
Dim tmpDate As Date
'Dim cc2 As Long
Dim tmpLength As Double
Dim tmpDir As Double
Dim DirAdd As Double
Dim ttTextwidth As Long
Dim nc As Long
Dim tmptxt As String
Dim txtLength As Long
ReDim lMoon(3)
ReDim dMoon(3)
ReDim segP(3)
Dim ocr As Long
Dim tmpcol As Long

WkN = 34
h2 = 0
h = 0
rd = 0
gd = 0
bd = 0
r2 = 0
g2 = 0
b2 = 0
r3 = 0
g3 = 0
b3 = 0
c = 0

Call DoWorkOuts(tYear, tMonth)

    r2 = gradtemp.R(0)
    g2 = gradtemp.G(0)
    b2 = gradtemp.b(0)
    h = Int(ndYear / gradtemp.Count)
    h2 = h - 1

lg = lg + 80

Call Draw2(tMonth, tYear)

lg = lg - 80

tmpDate = DateSerial(tYear, 1, 1)
If Month(tmpDate) <> tMonth Then
Do
    If h2 >= h - 1 Then
        h2 = 0
        If c = 15 Then c = 14
        rd = (gradtemp.R(c + 1) - r2) / (ndYear / gradtemp.Count)
        gd = (gradtemp.G(c + 1) - g2) / (ndYear / gradtemp.Count)
        bd = (gradtemp.b(c + 1) - b2) / (ndYear / gradtemp.Count)
        c = c + 1
    End If
    r2 = r2 + rd
    g2 = g2 + gd
    b2 = b2 + bd
    h2 = h2 + 1
    tmpDate = tmpDate + 1
Loop Until Month(tmpDate) = tMonth
End If

For cc = 0 To ndMonth - 1
    If h2 >= h - 1 Then
        h2 = 0
        If c = 15 Then c = 14
        rd = (gradtemp.R(c + 1) - r2) / (ndYear / gradtemp.Count)
        gd = (gradtemp.G(c + 1) - g2) / (ndYear / gradtemp.Count)
        bd = (gradtemp.b(c + 1) - b2) / (ndYear / gradtemp.Count)
        c = c + 1
    End If
    'Debug.Print c
                        
    r2 = r2 + rd
    g2 = g2 + gd
    b2 = b2 + bd
    
    r3 = r2 + ((255 - r2) / 7 * Weekday(DateSerial(tYear, tMonth, cc + 1), vbMonday))
    g3 = g2 + ((255 - g2) / 7 * Weekday(DateSerial(tYear, tMonth, cc + 1), vbMonday))
    b3 = b2 + ((255 - b2) / 7 * Weekday(DateSerial(tYear, tMonth, cc + 1), vbMonday))
    
    'trAg = cc * agDay
    trAg = (cc * agDay) * PI / 180
    segP(0).x = (lg * stLg) * Cos(trAg) - xad
    segP(0).y = (lg * stLg) * Sin(trAg)
    trAg = (cc * agDay) * PI / 180
    segP(1).x = (lg) * Cos(trAg) - xad
    segP(1).y = (lg) * Sin(trAg)
    trAg = ((cc + 1) * agDay) * PI / 180
    segP(2).x = (lg) * Cos(trAg) - xad
    segP(2).y = (lg) * Sin(trAg)
    trAg = ((cc + 1) * agDay) * PI / 180
    segP(3).x = (lg * stLg) * Cos(trAg) - xad
    segP(3).y = (lg * stLg) * Sin(trAg)
    
    'Draw day segment
    Pic1.FillStyle = vbSolid
    Pic1.FillColor = RGB(r3, g3, b3)
    Pic1.ForeColor = QBColor(0)
    Call Polygon(Pic1.hdc, segP(0), 4)
    
    
'Moon phases
If Check1.Value = vbChecked Then
    trAg = ((cc) * agDay - 0) * PI / 180
    lgPD = lg - 350
    lMoon(0).x = (lgPD) * Cos(trAg) - xad
    lMoon(0).y = (lgPD) * Sin(trAg)
    trAg = ((cc) * agDay - 0) * PI / 180
    lgPD = lg - 350 - MoonPhase(DateSerial(tYear, tMonth, cc + 1))
    lMoon(1).x = (lgPD) * Cos(trAg) - xad
    lMoon(1).y = (lgPD) * Sin(trAg)
    trAg = ((cc + 1) * agDay) * PI / 180
    lgPD = lg - 350 - MoonPhase(DateSerial(tYear, tMonth, cc + 2))
    lMoon(2).x = (lgPD) * Cos(trAg) - xad
    lMoon(2).y = (lgPD) * Sin(trAg)
    trAg = ((cc + 1) * agDay) * PI / 180
    lgPD = lg - 350
    lMoon(3).x = (lgPD) * Cos(trAg) - xad
    lMoon(3).y = (lgPD) * Sin(trAg)
    AA2.LineGP Pic1.hdc, lMoon(0).x, lMoon(0).y, lMoon(3).x, lMoon(3).y, RGB(200, 200, 200)
    
    trAg = ((cc) * agDay - 0) * PI / 180
    lgPD = lg - 350 - (179)
    dMoon(0).x = (lgPD) * Cos(trAg) - xad
    dMoon(0).y = (lgPD) * Sin(trAg)
    dMoon(1).x = lMoon(1).x
    dMoon(1).y = lMoon(1).y
    dMoon(2).x = lMoon(2).x
    dMoon(2).y = lMoon(2).y
    trAg = ((cc + 1) * agDay) * PI / 180
    lgPD = lg - 350 - (179)
    dMoon(3).x = (lgPD) * Cos(trAg) - xad
    dMoon(3).y = (lgPD) * Sin(trAg)
    AA2.LineGP Pic1.hdc, dMoon(0).x, dMoon(0).y, dMoon(3).x, dMoon(3).y, RGB(100, 100, 100)
    
    
    Pic1.FillStyle = vbSolid
    Pic1.FillColor = RGB(IIf(r3 >= 205, 255, r3 + 50), IIf(g3 >= 205, 255, g3 + 50), IIf(b3 >= 205, 255, b3 + 50))
    Pic1.ForeColor = RGB(IIf(r3 >= 205, 255, r3 + 50), IIf(g3 >= 205, 255, g3 + 50), IIf(b3 >= 205, 255, b3 + 50))
    Call Polygon(Pic1.hdc, lMoon(0), 4)
    Pic1.FillColor = RGB(IIf(r3 <= 50, 0, r3 - 50), IIf(g3 <= 50, 0, g3 - 50), IIf(b3 <= 50, 0, b3 - 50))
    Pic1.ForeColor = RGB(IIf(r3 <= 50, 0, r3 - 50), IIf(g3 <= 50, 0, g3 - 50), IIf(b3 <= 50, 0, b3 - 50))
    Call Polygon(Pic1.hdc, dMoon(0), 4)
End If

    Pic1.ForeColor = QBColor(0)
    

    AA2.LineGP Pic1.hdc, segP(0).x, segP(0).y, segP(1).x, segP(1).y, RGB(IIf(r3 > 50, r3 - 50, r3), IIf(g3 > 50, g3 - 50, g3), IIf(b3 > 50, b3 - 50, b3))
    'Draw outer black line
    AA2.LineGP Pic1.hdc, segP(1).x, segP(1).y, segP(2).x, segP(2).y, RGB(0, 0, 0)
    AA2.LineGP Pic1.hdc, segP(2).x, segP(2).y, segP(3).x, segP(3).y, RGB(IIf(r3 > 50, r3 - 50, r3), IIf(g3 > 50, g3 - 50, g3), IIf(b3 > 50, b3 - 50, b3))
    
    'Draw inner black line
    For ocr = 0 To 10
    trAg = (cc * agDay) * PI / 180
    segP(0).x = (lg * stLg - ocr) * Cos(trAg) - xad
    segP(0).y = (lg * stLg - ocr) * Sin(trAg)
    trAg = ((cc + 1) * agDay) * PI / 180
    segP(3).x = (lg * stLg - ocr) * Cos(trAg) - xad
    segP(3).y = (lg * stLg - ocr) * Sin(trAg)
    tmpcol = 255 / 10 * ocr
    AA2.LineGP Pic1.hdc, segP(3).x, segP(3).y, segP(0).x, segP(0).y, RGB(tmpcol, tmpcol, tmpcol)
    Next
    
    trAg = (cc * agDay) * PI / 180
    segP(1).x = (lg - 40) * Cos(trAg) - xad
    segP(1).y = (lg - 40) * Sin(trAg)
    trAg = ((cc + 1) * agDay) * PI / 180
    segP(2).x = (lg - 40) * Cos(trAg) - xad
    segP(2).y = (lg - 40) * Sin(trAg)
    AA2.LineGP Pic1.hdc, segP(1).x, segP(1).y, segP(2).x, segP(2).y, RGB(IIf(r3 > 50, r3 - 50, r3), IIf(g3 > 50, g3 - 50, g3), IIf(b3 > 50, b3 - 50, b3))

    trAg = ((cc + 0.7) * agDay) * PI / 180
    lgPD = lg - (((lg - (lg * stLg + IIf(Check6.Value = vbChecked, WkN, 0)) - 80) / ndMonth) * cc) - 80
    xx = (lgPD) * Cos(trAg) - xad
    yy = (lgPD) * Sin(trAg)
    
    'Print day of the month
    Pic1.Font = Combo4
    Call cFont(Pic1.hdc, cc + 1, xx, yy - 7, 12, -90 + (ndfMonth * agDay), True)
                      
'Print sunrise/sunset times
If Check4.Value = vbChecked Then
    cSun.DateDay = DateSerial(tYear, tMonth, cc + 1)
    'cSun.TimeZone = 0
    cSun.DaySavings = DaylightSavingsTime(DateSerial(tYear, tMonth, cc + 1))
    
    cSun.CalculateSun
    Pic1.Font = 12
    
    'Debug.Print Format((NextFullMoon(DateSerial(tYear, tMonth, cc))), "dd/mm/yyyy"), DateSerial(tYear, tMonth, cc + 1)
    tmptxt = ""
    If Format((NextFullMoon(DateSerial(tYear, tMonth, cc))), "dd/mm/yyyy") = Format(DateSerial(tYear, tMonth, cc + 1), "dd/mm/yyyy") Then
        tmptxt = tmptxt & Format(NextFullMoon(DateSerial(tYear, tMonth, 1)), "hh:mm:ss") & "     "
    End If
        
    tmptxt = tmptxt & Format(cSun.Sunrise, "hh:mm") & " < " & Format(cSun.SolarNoon, "hh:mm") & " > " & Format(cSun.Sunset, "hh:mm")
    
    If Format((NextNewMoon(DateSerial(tYear, tMonth, cc))), "dd/mm/yyyy") = Format(DateSerial(tYear, tMonth, cc + 1), "dd/mm/yyyy") Then
        tmptxt = tmptxt & "     " & Format(NextNewMoon(DateSerial(tYear, tMonth, 1)), "hh:mm:ss")
    End If
    
    txtLength = Pic1.TextWidth(tmptxt) + 20
    Pic1.ForeColor = RGB(100, 100, 100)
    If (ndfMonth + cc) * agDay > 180 Then
        trAg = ((cc + 0.6) * agDay) * PI / 180
        If cc + 1 < 16 Then
            xx = (lg * stLg + 15 + txtLength / 2 + IIf(Check6.Value = vbChecked, WkN, 0)) * Cos(trAg) - xad
            yy = (lg * stLg + 15 + txtLength / 2 + IIf(Check6.Value = vbChecked, WkN, 0)) * Sin(trAg)
        Else
            xx = (lg - 50 - txtLength / 2) * Cos(trAg) - xad
            yy = (lg - 50 - txtLength / 2) * Sin(trAg)
        End If
        Call cFont(Pic1.hdc, tmptxt, xx, yy, 12, -CLng(trAg * 180 / PI) + 180, True)
        Else
        trAg = ((cc + 0.5) * agDay) * PI / 180
        If cc + 1 < 16 Then
            xx = (lg * stLg + 15 + txtLength / 2 + IIf(Check6.Value = vbChecked, WkN, 0)) * Cos(trAg) - xad
            yy = (lg * stLg + 15 + txtLength / 2 + IIf(Check6.Value = vbChecked, WkN, 0)) * Sin(trAg)
        Else
            xx = (lg - 50 - txtLength / 2) * Cos(trAg) - xad
            yy = (lg - 50 - txtLength / 2) * Sin(trAg)
        End If
        Call cFont(Pic1.hdc, tmptxt, xx, yy, 12, -CLng(trAg * 180 / PI), True)
    End If
    'Debug.Print cSun.Sunrise, cSun.SolarNoon, cSun.Sunset
End If
               
    trAg = ((cc + 0.5) * agDay) * PI / 180
    lgPD = lg - 20
    xx = (lgPD) * Cos(trAg) - xad
    yy = (lgPD) * Sin(trAg)
    Pic1.ForeColor = RGB(0, 0, 0)
    Call cFont(Pic1.hdc, Left(Format(DateSerial(tYear, tMonth, cc + 1), "ddd", vbMonday), 1), xx, yy, 10, -90 + (ndfMonth * agDay), True)
    'Call cFont(Pic1.hdc, Left(Format(DateSerial(tYear, tMonth, cc + 1), "ddd", vbMonday), 1), xx, yy, 7, 0)
    
    h2 = h2 + 1
Next

Pic1.Font = Combo3
Dim tmpMonth As String
tmpMonth = UCase(Format(DateSerial(tYear, tMonth, 1), "mmmm", vbMonday))
tmpMonth = StrConv(tmpMonth, vbProperCase)
Pic1.FontSize = 50
ttTextwidth = Pic1.TextWidth(tmpMonth) + (10 * Len(tmpMonth))
trAg = ((ndMonth * agDay) / 2 - (0)) * PI / 180
xx = (lg + 40) * Cos(trAg) '- xad
yy = (lg + 40) * Sin(trAg)
tmpLength = Sqr(ttTextwidth ^ 2 + (Sqr(xx ^ 2 + yy ^ 2)) ^ 2)
tmpDir = Atn(ttTextwidth / tmpLength) * 180 / PI

For cc = 1 To Len(tmpMonth)
    trAg = ((ndMonth * agDay) / 2 - (tmpDir / 2) + DirAdd) * PI / 180
    xx = (lg + 40) * Cos(trAg) - xad
    yy = (lg + 40) * Sin(trAg)
    Call cFont(Pic1.hdc, Mid(tmpMonth, cc, 1), xx, yy, 50, -CLng(trAg * 180 / PI) - 90, True)

    trAg = ((ndMonth * agDay) / 2 - (0)) * PI / 180
    xx = (lg + 40) * Cos(trAg) '- xad
    yy = (lg + 40) * Sin(trAg)
    tmpLength = Sqr((Pic1.TextWidth(Mid(tmpMonth, cc, 1)) + 10) ^ 2 + (Sqr(xx ^ 2 + yy ^ 2)) ^ 2)
    DirAdd = DirAdd + Atn((Pic1.TextWidth(Mid(tmpMonth, cc, 1)) + 10) / tmpLength) * 180 / PI

    'If Day(tdate) = 21 And Month(tdate) = 6 Then pb1.Line (xx + xmarg, yy + ymarg)-(xmarg + xx2, ymarg + yy2), 0
    'If Day(tdate) = 22 And Month(tdate) = 9 Then pb1.Line (xx + xmarg, yy + ymarg)-(xmarg + xx2, ymarg + yy2), 0
    'If Day(tdate) = 22 And Month(tdate) = 12 Then pb1.Line (xx + xmarg, yy + ymarg)-(xmarg + xx2, ymarg + yy2), 0
    'If Day(tdate) = 21 And Month(tdate) = 3 Then pb1.Line (xx + xmarg, yy + ymarg)-(xmarg + xx2, ymarg + yy2), 0
Next

Pic1.Font = Combo2

If Check6.Value = vbChecked Then
    Call DrawWkNo(tYear, tMonth)
End If

Pic1.ForeColor = 0

If Check3.Value = vbChecked Then

For cc = 0 To ndMonth - 1
    'trAg = cc * agDay
    tmptxt = ""
    For nc = LBound(hList) To UBound(hList)
        If hList(nc).Hol_Date = DateSerial(tYear, tMonth, cc + 1) Then
            tmptxt = tmptxt & hList(nc).Comments & " - "
        End If
    Next
    
    
    If tmptxt <> "" Then
    tmptxt = Mid(tmptxt, 1, Len(tmptxt) - 3)
    trAg = ((cc + 0.5) * agDay) * PI / 180
    
    'Pic1.Font = "Arial"
    Pic1.Font = Combo2
    Pic1.FontSize = Val(Combo5)
    
    'tmpTxt = Trim(tmpTxt)
    txtLength = Pic1.TextWidth(tmptxt) + 30
    
    segP(0).x = (lg * stLg - 20) * Cos(trAg) - xad
    segP(0).y = (lg * stLg - 20) * Sin(trAg)
    trAg = ((cc + 0.5) * agDay) * PI / 180
    
    
    segP(1).x = (lg * stLg) * Cos(trAg) - xad
    segP(1).y = (lg * stLg) * Sin(trAg)
    AA2.LineGP Pic1.hdc, segP(0).x, segP(0).y, segP(1).x, segP(1).y, QBColor(0)
    
    If (ndfMonth + cc) * agDay > 180 Then
        trAg = ((cc + 0.5) * agDay) * PI / 180
        xx = (lg * stLg - txtLength / 2 - 15) * Cos(trAg) - xad
        yy = (lg * stLg - txtLength / 2 - 15) * Sin(trAg)
        Call cFont(Pic1.hdc, tmptxt, xx, yy, Val(Combo5), -CLng(trAg * 180 / PI) + 180, True)
    Else
        trAg = ((cc + 0) * agDay) * PI / 180
        xx = (lg * stLg - txtLength) * Cos(trAg) - xad
        yy = (lg * stLg - txtLength) * Sin(trAg)
        Call cFont(Pic1.hdc, tmptxt, xx, yy, Val(Combo5), -CLng(trAg * 180 / PI))
    End If
    
    
    End If
Next
End If
    
    
' Draw cutting line
If Check5.Value = vbChecked Then
    trAg = (ndMonth * agDay + 0.03) * PI / 180
    segP(0).x = -xad
    segP(0).y = 0
    segP(1).x = (lg * 2) * Cos(trAg) - xad
    segP(1).y = (lg * 2) * Sin(trAg)
    Pic1.ForeColor = RGB(200, 200, 200)
    Pic1.DrawStyle = 2
    Pic1.Line (segP(0).x, segP(0).y)-(segP(1).x, segP(1).y)
    Pic1.DrawStyle = 0
    Pic1.ForeColor = QBColor(0)
End If
    

Pic1.Font = "Arial"
Pic1.FontSize = 12
Pic1.CurrentX = 20
Pic1.CurrentY = Pic1.ScaleHeight - 300
Pic1.Print Format(DateSerial(tYear, tMonth, 1), "mmmm yyyy")
Pic1.CurrentX = 20
Pic1.Print ((DateSerial(tYear, tMonth, 1) - DateSerial(tYear, 1, 1)) * agDay) - 90
Pic1.CurrentX = 20
Pic1.Print "No. days in month: " & ndMonth
Pic1.CurrentX = 20
Pic1.Print DateSerial(tYear, tMonth, 1) - DateSerial(tYear, 1, 1) + ndMonth & " of " & ndYear + 1

Pic1.Refresh
End Sub

Private Sub Command1_Click()
Dim i As Long
    gradtemp.Count = 15
    For i = 0 To 15
        gradtemp.R(i) = Red(picColors(i).BackColor)
        gradtemp.G(i) = Green(picColors(i).BackColor)
        gradtemp.b(i) = Blue(picColors(i).BackColor)
    Next

Pic1.Cls
Call DrawMonth(Val(Text1), CLng(Combo1.Text))
End Sub

Private Sub Command2_Click()
SavePicture Pic1.Image, App.Path & "\images\pic.bmp"
End Sub

Private Sub Command3_Click()
Dim i As Long
    gradtemp.Count = 15
    For i = 0 To 15
        gradtemp.R(i) = Red(picColors(i).BackColor)
        gradtemp.G(i) = Green(picColors(i).BackColor)
        gradtemp.b(i) = Blue(picColors(i).BackColor)
    Next

Dim cc As Long
For cc = 1 To 12
    Pic1.Cls
    Call DrawMonth(Val(Text1), cc)
    DoEvents
    SavePicture Pic1.Image, App.Path & "\images\" & cc & ".bmp"
Next
End Sub

Private Sub Form_Activate()
If FontsAdded = False Then
    FontsAdded = True
    EnumFontFamilies Me.hdc, vbNullString, AddressOf EnumFontFamTypeProc, Combo2
    EnumFontFamilies Me.hdc, vbNullString, AddressOf EnumFontFamTypeProc, Combo3
    EnumFontFamilies Me.hdc, vbNullString, AddressOf EnumFontFamTypeProc, Combo4
    Combo2.SelStart = 0
    Combo2.SelLength = Len(Combo2)
    Combo2.SelText = "Arial"
    Combo3.SelStart = 0
    Combo3.SelLength = Len(Combo3)
    'Combo3.SelText = "Embossing Tape 1 BRK"
    Combo3.SelText = "Arial"
    Combo4.SelStart = 0
    Combo4.SelLength = Len(Combo4)
    Combo4.SelText = "Arial"

    cSun.Latitude = CDbl(53.9823)
    cSun.Longitude = CDbl(2.69367)
    cSun.TimeZone = CDbl(0)
    cSun.DateDay = #10/16/2007#
    cSun.TimeZone = 0
    'cSun.DaySavings = True
    'cSun.CalculateSun
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
    Set cSun = New clsSunrise

    For i = 1 To 12
    Combo1.AddItem i
    Next
    Combo1.ListIndex = 0
    For i = 2008 To 2050
        Text1.AddItem i
    Next
    Text1.ListIndex = 1
    For i = 1 To 100
        Combo5.AddItem i
    Next
    Combo5.ListIndex = 15
End Sub

Private Sub Form_Resize()
On Error Resume Next
Iox1.Width = ScaleWidth - Iox1.Left * 2
Iox1.Height = ScaleHeight - Iox1.Top
'Pic1.Left = 0
'Pic1.Top = 0
If HasLoaded = False Then
    HasLoaded = True
    Pic1.Height = Pic1.Height * 2
    Pic1.Width = Pic1.Width * 2
End If
End Sub

Private Sub picColors_Click(Index As Integer)
Dim tmpcol As Long
cd1.ShowColor
picColors(Index).BackColor = cd1.Color
End Sub

Private Sub Text1_Change()
'Call FindInfo(DateSerial(Val(Text1), 1, 1), DateSerial(Val(Text1) + 1, 1, 1) - DateSerial(Val(Text1), 1, 1))
End Sub

Private Sub Text1_Click()
ReDim hList(0)
Call FindInfo(DateSerial(Val(Text1), 1, 1), DateSerial(Val(Text1) + 1, 1, 1) - DateSerial(Val(Text1), 1, 1))
End Sub
