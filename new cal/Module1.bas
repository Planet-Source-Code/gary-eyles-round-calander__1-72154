Attribute VB_Name = "Module1"
Option Explicit

'Font enumeration types
Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64



Public Type NEWTEXTMETRIC
   tmHeight As Long
   tmAscent As Long
   tmDescent As Long
   tmInternalLeading As Long
   tmExternalLeading As Long
   tmAveCharWidth As Long
   tmMaxCharWidth As Long
   tmWeight As Long
   tmOverhang As Long
   tmDigitizedAspectX As Long
   tmDigitizedAspectY As Long
   tmFirstChar As Byte
   tmLastChar As Byte
   tmDefaultChar As Byte
   tmBreakChar As Byte
   tmItalic As Byte
   tmUnderlined As Byte
   tmStruckOut As Byte
   tmPitchAndFamily As Byte
   tmCharSet As Byte
   ntmFlags As Long
   ntmSizeEM As Long
   ntmCellHeight As Long
   ntmAveWidth As Long
End Type

'Public Const LF_FACESIZE = 32
'Public Const LF_FULLFACESIZE = 64

Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        'lfFaceName(1 To LF_FACESIZE) As Byte
        lfFacename As String * 33
End Type

'ntmFlags field flags
Public Const NTM_REGULAR = &H40&
Public Const NTM_BOLD = &H20&
Public Const NTM_ITALIC = &H1&

'tmPitchAndFamily flags
Public Const TMPF_FIXED_PITCH = &H1
Public Const TMPF_VECTOR = &H2
Public Const TMPF_DEVICE = &H8
Public Const TMPF_TRUETYPE = &H4

Public Const ELF_VERSION = 0
Public Const ELF_CULTURE_LATIN = 0

'EnumFonts Masks
Public Const RASTER_FONTTYPE = &H1
Public Const DEVICE_FONTTYPE = &H2
Public Const TRUETYPE_FONTTYPE = &H4

Public Declare Function EnumFontFamilies Lib "gdi32" _
   Alias "EnumFontFamiliesA" _
  (ByVal hdc As Long, _
   ByVal lpszFamily As String, _
   ByVal lpEnumFontFamProc As Long, _
   lParam As Any) As Long
   
Public Function DaylightSavingsTime2(CalDate As Date) As Boolean
    Dim Oct31 As String
    'If the date falls between November and
    '     March, then
    'it is standard time


    If Month(CalDate) > 10 Or Month(CalDate) < 4 Then
        DaylightSavingsTime2 = False
        Exit Function
    End If
    'If the date falls between May and Septe
    '     mber, then
    'it is daylight savings time


    If Month(CalDate) < 10 And Month(CalDate) > 4 Then
        DaylightSavingsTime2 = True
        Exit Function
    End If
    'If it is April...


    If Month(CalDate) = 4 Then
        'If the day of the month is less than th
        '     e
        'numbered day of the week (Sun=1, Sat=7)
        '     , then
        'it is standard time


        If Day(CalDate) < Weekday(CalDate) Then
            DaylightSavingsTime2 = False
            Exit Function
            'if not, then daylight time
        Else
            DaylightSavingsTime2 = True
            Exit Function
        End If
    End If
    'If it is october...


    If Month(CalDate) = 10 Then
        'If it is the last week in october, it i
        '     s
        'standard time
        Oct31 = "10/31/" & Year(CalDate)
        If (Weekday(CDate(Oct31)) >= _
        (Weekday(CalDate))) And (Day( _
        CalDate) > 24) Then
        DaylightSavingsTime2 = False
        Exit Function
        'any other week is daylight time
    Else
        DaylightSavingsTime2 = True
        Exit Function
    End If
End If
End Function
   

Public Function EnumFontFamTypeProc(lpNLF As LOGFONT, _
                                    lpNTM As NEWTEXTMETRIC, _
                                    ByVal FontType As Long, _
                                    lParam As ListBox) As Long

   Dim FaceName As String
   
   'If ShowFontType = FontType Then
        
     'convert the returned string from Unicode to ANSI
      FaceName = StrConv(lpNLF.lfFacename, vbUnicode)
        
     'add the font to the list
      lParam.AddItem Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
 
   'End If
   
  'return success to the call
   EnumFontFamTypeProc = 1

End Function

