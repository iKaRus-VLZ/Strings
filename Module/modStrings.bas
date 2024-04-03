Attribute VB_Name = "modStrings"
Option Base 0
Option Explicit
#Const APPTYPE = 0  ' 0=ACCESS, 1=EXCEL
'=========================
Private Const c_strModule As String = "modStrings"
'=========================
' ��������      : ������� ��� ������ �� ��������
' �����         : ������ �.�. (KashRus@gmail.com)
' ������        : 1.1.30.453854194
' ����          : 03.04.2024 10:03:56
' ����������    : ������� ��� Access x86, ������������ ��� x64, �� ������ �� �������������. _
'               : ��� ������ � Excel ������� APPTYPE=1
' v.1.1.30      : 12.03.2024 - ��������� � GroupsGet - ������ ������� ���������� ������ ��� ������� (����� �������� ����������� ��������� ����- � ����� -������� ��������� ����� If .. Then .. End If)
' v.1.1.29      : 21.12.2022 - ��������� � GroupsGet - ���������� �������������� ������. (�� ��� ������ �����������������)
' v.1.1.27      : 09.08.2022 - ��������� � DelimStringSet - �������� �������� SetUnique ��� �������� ������������ ����������� ��������
' v.1.1.26      : 11.02.2022 - ��������� � TaggedStringSet/Del - ���������� ������ ����������� ���� TagDelim ��������� ���������, ��������� �� ��������
' v.1.1.25      : 16.07.2020 - ���������� PlaceHoldersGet - ������� ������ ��������� � ������������� ���������� ���� ������ ���������� �������� ������������ �������������. ��������� ��������� ������������� ��������� (ReplaceExisting)
' v.1.1.24      : 24.03.2020 - ��������� ������� GroupGet - ��� ���������� �� ������ ����� (��������� ����������� � ������)
' v.1.1.23      : 04.02.2020 - ��������� � PlaceHoldersGet - �������� ��������� ������� �� ���� ��������� VBA.Like � VBS.RegExp
' v.1.1.22      : 03.02.2020 - ��������� PlaceHoldersGet - ��������� ��������� � ��������� �������� ���������� �� ������ �� �������; _
                               ������������� ������� ������ ������ � ��������������� ���������� ��� �������� ������������
' v.1.1.19      : 30.01.2020 - ��������� � PlaceHoldersSet - ��������� ����������� ������������� ������������� (������ ��. p_TermModify)
' v.1.1.17      : 29.11.2019 - ��������� ������� ��� ������ � ��������������� ����������� PlaceHoldersSetByIndex � PlaceHoldersSet
' v.1.1.16      : 31.10.2019 - ��������� � Tokenize - �������� �������������� �������� Positions - ������������ ������ ������� ��������� ������� � �������� ������ _
                               � ��������� ������� TokenString[Get","Set","Del] ����������/�������/�������� ������� �� ������, ����������� ��������� ��� DelimString � TaggedString _
                               ��� ���� ������� ������ � ����������� ��������� ��������� sBeg, sEnd ������������ ������� ��������� � ������
' v.1.1.12      : 24.09.2019 - ��������� ������� ������ �� �������� � ������������� DelimString[Get","Set","Del], _
                               ��������� ����������� ������ � �������������� ��������� (�� ����� ������) _
                               � ��������� ������� ������ �� �������� ������� ���������� TaggedString[Get","Set","Del]
' v.1.1.10      : 16.08.2019 - ��������� ������� ������������� ���������: SoundEx, PolyPhone _
                               � ������� ����������� ������������� ����������: ���������� ����� ���������������������, Levenshtein, Dice, ����������� Jaro etc. _
                               ��������� ��������� ������� ������������� �� ��������� ��������� - �������� ���������� �� ������ ���������.
' v.1.1.9       : 18.07.2019 - ��������� NumToWords - �������������� ����� � ����� � ��������� ���������� �� �������
' v.1.1.8       : 13.07.2019 - ������� ����������� � ������� ��������� ������ �� ����� �� ���������/������
' v.1.1.5       : 12.12.2018 - ��������� HyphenateWord - ��� ����������� ��������� � ������. ��������: http://www.cyberforum.ru/vba/thread792944.html
' v.1.1.4       : ��������� DeclineWords - ��������� ����������� �������� ���� ��� ��������� ��������������
' v.1.1.2       : ��������� DeclineWord - ������� ��������� �� �������. ��������� ����������� �������� �� ������.
'=========================
' ToDo: DeclineWord - ���������� ��� ������ � ��������� ���������� ��� ������� (��������� ��� ��������), ��� ������ ���������� � �������� ���������
' + �������� ����������� ������������� �������� � �������� � ������������� � ������� ������������ ������ �������
' + PlaceHoldersGet - ��������� ������������ ������������ ��������, �������� ���������� ����������, �������� �������������� ������ ���������� ��������
' - NumToWords      - ����������� �������� ����������� ����������� ������ >10^6
'=========================
'Private Const c_strEsc = "\" ' ������ ������ - ��� �������� ��� ��������� ���� ������� ������������� ��� ������� ������ (�� �����������, �� �����������)
' ������� �������
    ' ��������� �������� ������ ����: If iInStr (c_strSymbRusConsonDeaf,sChar) Then
Private Const c_strSymbRusConson = "���������������������"  ' ��������� �����
'Private Const c_strSymbRusConsonVoicPaired = "������", c_strSymbRusConsonVoicOnly = "����" ' ������� ������/��������
'Private Const c_strSymbRusConsonDeafPaired = "������", c_strSymbRusConsonDeafOnly = "����" ' ������ ������/��������
'Private Const c_strSymbRusConsonVoic = c_strSymbRusConsonVoicPaired & c_strSymbRusConsonVoicOnly ' ������� ���
'Private Const c_strSymbRusConsonDeaf = c_strSymbRusConsonDeafPaired & c_strSymbRusConsonDeafOnly ' ������ ���
'Private Const c_strSymbRusConsonHardSoft = "���������������" ' ������ ������/������ � ����������� �� �������
'Private Const c_strSymbRusConsonHardOnly = "���", c_strSymbRusConsonSoftOnly = "���" ' ������ ������/������
'Private Const c_strSymbRusConsonHissing = "����", c_strSymbRusConsonWhistling = "���"  ' �������/���������
'Private Const c_strSymbRusConsonSonar = "�����", c_strSymbRusConsonNoisy = "����������������"  ' ��������/������
Private Const c_strSymbRusVowel = "���������"   ' ������� �����
'Private Const c_strSymbRusVowelYotated = "���" ' ������������ (�������) ������� �����
'Private Const c_strSymbRusVowelSoft = "����"   ' ���������� ������� �����
'Private Const c_strSymbRusVowelHard = "�����"   ' ������ �� �� ���������� �����������
Private Const c_strSymbRusSign = "��"            ' �����
'Private Const c_strSymbRusSignSoft = "�"        ' ������ ����
'Private Const c_strSymbRusSignHard = "�"        ' ������ ����
' ���������� �������
Private Const c_strSymbRusAll = c_strSymbRusVowel & c_strSymbRusConson & c_strSymbRusSign
Private Const c_strSymbEngVowel = "aeiouy", c_strSymbEngConson = "bcdfghjklmnpqrstvwxz", c_strSymbEngSign = "" '"'`"
Private Const c_strSymbEngAll = c_strSymbEngVowel & c_strSymbEngConson & c_strSymbEngSign
' ����� � �������
Private Const c_strSymbDigits = "0123456789", c_strSymbMath = "+-*/\^|=", c_strSymbPunct = ".,?!:;-()" ' & "�"
Private Const c_strSymbCommas = "'""", c_strSymbParenth = "()[]{}<>", c_strSymbOthers = "_&@#$%~`"
Private Const c_strSymbSpaces = " " & vbCr & vbLf & vbNewLine & vbTab & vbVerticalTab
' ��� �������������� ����
Private Const c_strHexPref = "&H"
Private Const c_strOthers = " -~_"

Private Const c_idxPref = "i" ' ������� ���� ��������� ���������
' size convertion constants
Private Const PointsPerInch = 72
Private Const TwipsPerInch = 1440
Private Const CentimitersPerInch = 2.54                 '1 ���� = 127 / 50 ��
Private Const HimetricPerInch = 2540                    '1 ���� = 1000 * 127/50 himetrix
'
Private Const inch = TwipsPerInch                       '1 ���� = 1440 twips
Private Const pt = TwipsPerInch / PointsPerInch         '1 ����� = 20 twips
Private Const cm = TwipsPerInch / CentimitersPerInch    '1 �� = 566.929133858 twips
'--------------------------------------------------------------------------------
Public Enum DeclineCase         ' �����
    DeclineCaseUndef = 0
    DeclineCaseImen = 1         ' ��.�. (���/���)       Nominative
    DeclineCaseRod = 2          ' �.�.  (����/����)     Genitive
    DeclineCaseDat = 3          ' �.�.  (����/����)     Dative
    DeclineCaseVin = 4          ' �.�.  (����/���)      Accusative
    DeclineCaseTvor = 5         ' �.�.  (���/���)       Ablative
    DeclineCasePred = 6         ' �.�.  (� ���/� ���)   Prepositional
End Enum
Public Enum DeclineGend         ' ��� ("�|�|��")
    DeclineGendUndef = 0
    DeclineGendMale = 1         ' �.�.
    DeclineGendFem = 2          ' �.�.
    DeclineGendNeut = 3         ' �.�.
End Enum
Public Enum DeclineNumb         ' ����� ("��|��")
    DeclineNumbUndef = 0
    DeclineNumbSingle = 1       ' ��.�.
    DeclineNumbPlural = 2       ' ��.�.
End Enum
Public Enum SpeechPartType      ' ����� ����
    SpeechPartTypeUndef = 0
    SpeechPartTypeNoun = 1      ' ���������������
    SpeechPartTypeAdject = 2    ' ��������������
    SpeechPartTypeNumeral = 3   ' ������������
    SpeechPartTypeVerb = 4      ' ������
    SpeechPartTypeAdverb = 5    ' �������
    SpeechPartTypePronoun = 6   ' �����������
    SpeechPartTypePreposition = 7 ' �������
End Enum
Public Enum NumeralType         ' ��� ������������ ("��������������|����������")
    NumeralUndef = 0
    NumeralOrdinal = 1          ' ��������������
    NumeralCardinal = 2         ' ����������
End Enum
Public Enum SymbolType          ' ��� �������
    SymbolTypeUndef = 0
    SymbolTypeVowel = 1         ' �������
    SymbolTypeCons = 2          ' ���������
    SymbolTypeSign = 3          ' ����� ��������
    SymbolTypeNumb = 4          ' �����
End Enum
Public Enum AlphabetType        ' ��� ��������
    AlphabetTypeUndef = 0
    AlphabetTypeLatin = 1       ' ���������
    AlphabetTypeCyrilic = 2     ' �������������
End Enum
Public Type GroupExpr           ' ��� ��� �������� ��������� ����������� ����� � ��������� ���������� (��. GroupsGet)
    Text As String              ' ���������� ����� ��������� (��� ������)
    TextBeg As Long             ' ������� ������ ��������� � �������� ������ (������� ������)
    TextEnd As Long             ' ������� ����� ��������� � �������� ������ (������� ������)
    Bracket As Long             ' ��� ������/������� ������ (���������� ������� ������ ���� ������ � �������)
    Level As Long               ' ������� ����������� (0-��� ������, 1-������� ������, ... n-������ n-������)
End Type
'--------------------------------------------------------------------------------
' POINTER
'--------------------------------------------------------------------------------
'#If VBA7 = 0 Then       'LongPtr trick by @Greedo (https://github.com/Greedquest)
'Public Enum LongPtr
'    [_]
'End Enum
'#End If
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>  PtrSafe, LongPtr and LongLong
Private Const PTR_LENGTH As Long = 8
Private Const VARIANT_SIZE As Long = 24
#Else                   '<OFFICE97-2010>        Long
Private Const PTR_LENGTH As Long = 4
Private Const VARIANT_SIZE As Long = 16
#End If                 '<WIN32>
'--------------------------------------------------------------------------------
' SAFEARRAY
'--------------------------------------------------------------------------------
Private Const FADF_AUTO As Long = (&H1)
Private Const FADF_FIXEDSIZE As Long = (&H10)
Private Const FADF_HAVEVARTYPE As Long = (&H80)
Private Type SAFEARRAYBOUND         ' 8 bytes
    cElements As Long               ' +0 ���������� ��������� � �����������
    lLbound As Long                 ' +4 ������ ������� �����������
End Type
'Private Type SAFEARRAY
'    cDims           As Integer      ' +0 ����� ������������
'    fFeatures       As Integer      ' +2 ����, ������������ ��������� SafeArray
'    cbElements      As Long         ' +4 ������ ������ �������� � ������
'    cLocks          As Long         ' +8 C������ ������, ����������� ���������� ����������, ���������� �� ������.
'    dummyPadding    As Long         ' +8 (x64 only!)
'    pvData          As Long         ' +12(x86) ��������� �� ������
'                    As LongLong     ' +16(x64)
'    rgSAbound As SAFEARRAYBOUND     ' ����������� ��� ������ ����������� (������ = n*8 bytes, n- ���-�� ������������ �������)
'                                    ' +16(x86) rgSAbound.cElements (Long) - ���������� ��������� � �����������
'                                    ' +24(x64)
'                                    ' +20(x86) rgSAbound.lLbound (Long)   - ������ ������� �����������
'                                    ' +28(x64)
'End Type
Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
#If Win64 Then
    dummyPadding As Long
    pvData As LongLong
#Else
    pvData As Long
#End If
'    rgsabound0 As SAFEARRAYBOUND
   cElements    As Long
   lLbound      As Long
End Type
'--------------------------------------------------------------------------------
' MSVBA
'--------------------------------------------------------------------------------
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>
Private Declare PtrSafe Function VarPtrArray Lib "VBE7.dll" Alias "VarPtr" (ByRef Ptr() As Any) As LongPtr
#ElseIf VBA7 Then       '<WIN32 & OFFICE2010+>
Private Declare Function VarPtrArray Lib "VBE7.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
'#Else                   '<OFFICE2003-2010>
'Private Declare Function VarPtrArray Lib "VBE6.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
#Else                   '<OFFICE2000-2003>
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
'#Else                   '<OFFICE97-2000>
'Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
#End If                 '<WIN32>
'--------------------------------------------------------------------------------
' KERNEL32
'--------------------------------------------------------------------------------
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>
Private Declare PtrSafe Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoA" (ByVal locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare PtrSafe Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare PtrSafe Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
#Else                   '<WIN32>
Private Declare Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
#End If                 '<WIN32>
'--------------------------------------------------------------------------------
' USER32
'--------------------------------------------------------------------------------
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>
Private Declare PtrSafe Function GetDC Lib "user32.dll" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32.dll" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
#Else                   '<WIN32>
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
#End If                 '<WIN32>
'--------------------------------------------------------------------------------
' GDI32
'--------------------------------------------------------------------------------
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32A" (ByVal hdc As LongPtr, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32.dll" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32.dll" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function GetTextMetrics Lib "gdi32.dll" Alias "GetTextMetricsA" (ByVal hdc As LongPtr, ByRef lpMetrics As TEXTMETRIC) As Long
Private Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal e As Long, ByVal o As Long, ByVal w As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal q As Long, ByVal PAF As Long, ByVal f As String) As LongPtr
#Else                   '<WIN32>
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetTextMetrics Lib "gdi32.dll" Alias "GetTextMetricsA" (ByVal hdc As Long, ByRef lpMetrics As TEXTMETRIC) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal e As Long, ByVal o As Long, ByVal w As Long, ByVal i As Long, ByVal U As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal cp As Long, ByVal q As Long, ByVal PAF As Long, ByVal f As String) As Long
#End If                 '<WIN32>
'--------------------------------------------------------------------------------
' OLEAUT32
'--------------------------------------------------------------------------------
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>
Private Declare PtrSafe Function SysReAllocStringLen Lib "oleaut32.dll" (ByVal pBSTR As LongPtr, Optional ByVal pszStrPtr As LongPtr, Optional ByVal Length As Long) As Long
Private Declare PtrSafe Function SysAllocStringByteLen Lib "oleaut32.dll" (ByVal pBSTR As LongPtr, ByVal bLen As Long) As LongPtr
#Else                   '<WIN32>
Private Declare Function SysReAllocStringLen Lib "oleaut32.dll" (ByVal oleStr As Long, ByVal psz As Long, ByVal cch As Long) As Long
Private Declare Function SysAllocStringByteLen Lib "oleaut32.dll" (ByVal oleStr As Long, ByVal bLen As Long) As Long
#End If                 '<WIN32>
'Private Type TPROC
'    hMem As Long
'    vtPtr As Long
'End Type
'Private aProc() As TPROC
'
'Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
'Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private Const LOCALE_SLIST = &HC        ' ����������� ��������� ������
Private Const LOCALE_SDECIMAL = &HE     ' ���������� �����������
Private Const LOCALE_STHOUSAND = &HF    ' ����������� ��������

Const LOCALE_USER_DEFAULT = &H400
Const LOCALE_SYSTEM_DEFAULT = &H800

Private Type Size
    cX As Long
    cY As Long
End Type
Private Type TEXTMETRIC
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
End Type
Private Enum apiDeviceCapability
    HORZSIZE = 4
    VERTSIZE = 6
    HORZRES = 8
    VERTRES = 10
    LOGPIXELSX = 88        '  Logical pixels/inch in X
    LOGPIXELSY = 90        '  Logical pixels/inch in Y
End Enum

' ==================
Public Function RegEx() As Object
' ���������� ������ RegExp ��� ������ � ����������� �����������
' �� RegEx ��. https://regex.sorokin.engineer/ru/latest/regular_expressions.htmlStatic oRegEx As Object: If oRegEx Is Nothing Then Set oRegEx = CreateObject("VBScript.RegExp")
Static soRegEx As Object: If soRegEx Is Nothing Then Set soRegEx = CreateObject("VBScript.RegExp")
    Set RegEx = soRegEx
End Function
' ==================
' ������ ����������� ��������� ������� � ������ ������� ��������� ������� �
' http://www.xbeat.net/vbspeed/ ��. ����� �������� �� http://www.vbforums.com/showthread.php?540323-VB6-Faster-Split-amp-Join-(development)
' ==================
Public Sub xSplit(Expression As String, Result() As String, Optional Delimiter As String = " ") 'As Long
' Returns a zero-based, one-dimensional array containing a specified number of substrings.
'-------------------------
' Expression  - Required. String expression containing substrings and delimiters. If expression is a zero-length string, xSplit returns a single-element array containing a zero-length string.
' asToken()   - Required. One-dimensional string array that will hold the returned substrings. Does not have to be bound before calling xSplit, and is guaranteed to hold at least one element (zero-based) on return.
' Delimiter   - Optional. String character used to identify substring limits. If omitted, the space character (" ") is assumed to be the delimiter. If delimiter is a zero-length string, a single-element array containing the entire expression string is returned.
' returns number of elements
' ����������� Split �������� ������� �� �������� �������, xSplit - �� �������
'-------------------------
' v.1.0.0       : 08.12.2001 - original SplitB04 by Chris Lucas, cdl1051@earthlink.net from http://www.xbeat.net/vbspeed/c_Split.htm#SplitB04
'-------------------------
Dim c As Long, sLen As Long, DelLen As Long, tmp As Long, Results() As Long
'Dim lCount As Long
    sLen = LenB(Expression) \ 2: DelLen = LenB(Delimiter) \ 2
    If sLen = 0 Or DelLen = 0 Then ReDim Preserve Result(0 To 0): Result(0) = Expression: Exit Sub ': xSplit = 1: Exit Function     ' ������ ������
' ������� ����������� � ���������� �� �������
    ReDim Preserve Results(0 To sLen): tmp = InStr(Expression, Delimiter)
    Do While tmp
        Results(c) = tmp: c = c + 1
        tmp = InStr(Results(c - 1) + 1, Expression, Delimiter)
    Loop
' ��������� ������
    ReDim Preserve Result(0 To c)
    If c = 0 Then Result(0) = Expression: Exit Sub ': xSplit = 1: Exit Function      ' ��� ������������
    Result(0) = Left$(Expression, Results(0) - 1)
    For c = 0 To c - 2
        Result(c + 1) = Mid$(Expression, Results(c) + DelLen, Results(c + 1) - Results(c) - DelLen)
    Next c
    Result(c + 1) = Right$(Expression, sLen - Results(c) - DelLen + 1)
    'xSplit = c + 2
End Sub
Public Function xJoin(Arr() As String, _
    Optional Delimiter As String = " ", _
    Optional ByVal Count As Long = -1) As String
' ������ ������������ Join
'-------------------------
' v.1.0.0       : 01.10.2000 - original Join08 by by Matt Curland, mattcur@microsoft.com, www.PowerVB.com from http://www.xbeat.net/vbspeed/c_Join.htm#Join08
'-------------------------
' ��������� ��������� ������������, ������� �������� ��� ASM �������� �����
' Works with VB- or typelib-declared CopyMemory (pass strings with StrPtr)
'-------------------------
Dim Lower As Long
Dim Upper As Long
Dim cbDelim As Long
Dim cbTotal As Long
Dim i As Long
Dim cbCur As Long
Dim pCurDest As LongPtr, pDelim As LongPtr
    Lower = LBound(Arr)
    If Count = -1 Then
        Upper = UBound(Arr)
    Else
        Upper = Lower + Count - 1
    End If
    For i = Lower To Upper
        cbTotal = cbTotal + LenB(Arr(i))
    Next i
    cbDelim = LenB(Delimiter)
    If cbDelim Then cbTotal = cbTotal + cbDelim * (Upper - Lower)

    'Use API to avoid useless initialization
    CopyMemory ByVal VarPtr(xJoin), SysAllocStringByteLen(0, cbTotal), PTR_LENGTH
    'Use this instead if APIs are typelib-declared:
    ''xJoin = SysAllocStringByteLen(0, cbTotal)

    'Now, split into two different paths
    'a) No delimiter
    'b) Delimiter
    pCurDest = StrPtr(xJoin)
    If cbDelim = 0 Then
        For i = Lower To Upper
            cbCur = LenB(Arr(i))
            CopyMemory ByVal pCurDest, ByVal StrPtr(Arr(i)), cbCur
            pCurDest = pCurDest + cbCur
        Next i
    Else
        pDelim = StrPtr(Delimiter)
        For i = Lower To Upper - 1
            cbCur = LenB(Arr(i))
            CopyMemory ByVal pCurDest, ByVal StrPtr(Arr(i)), cbCur
            pCurDest = pCurDest + cbCur
            CopyMemory ByVal pCurDest, ByVal pDelim, cbDelim
            pCurDest = pCurDest + cbDelim
        Next i
        CopyMemory ByVal pCurDest, ByVal StrPtr(Arr(i)), LenB(Arr(i))
    End If
End Function
Public Function xReplace(ByRef Expression As String, _
    ByRef Find As String, ByRef ReplaceWith As String, _
    Optional ByVal Start As Long = 1, _
    Optional ByVal Count As Long = 2147483647, _
    Optional ByVal Compare As VbCompareMethod = vbBinaryCompare _
  ) As String
' Returns a string in which a specified substring has been replaced with another substring a specified number of times.
'-------------------------
' Expression  - Required. String expression containing substring to replace.
' Find    - Required. Substring being searched for. If Find is zero-length, xReplace returns a copy of Expression.
' ReplaceWith - Required. Replacement substring. If ReplaceWith is zero-length, xReplace returns a copy of Expression with all occurences of Find removed.
' Start   - Optional. Position within expression where substring search is to begin. If omitted, 1 is assumed. Must be used in conjunction with count.
' Count   - Optional. Number of substring substitutions to perform. If omitted, the default value is -1, which means make all possible substitutions. Must be used in conjunction with start.
' Compare - Optional. Numeric value indicating the kind of comparison to use when evaluating substrings. If omitted, the default value is 0, which means perform a binary comparison.
'-------------------------
' v.1.0.0       : 18.12.2000 - original Replace09 by Jost Schwider, jost@schwider.de from http://www.xbeat.net/vbspeed/c_Replace.htm#Replace09
'-------------------------
    If LenB(Find) Then
        If Compare = vbBinaryCompare Then
            p_ReplaceBin xReplace, Expression, Expression, Find, ReplaceWith, Start, Count
        Else
            p_ReplaceBin xReplace, Expression, LCase$(Expression), LCase$(Find), ReplaceWith, Start, Count
        End If
    Else
        xReplace = Expression
    End If
End Function
Private Static Sub p_ReplaceBin(ByRef Result As String, _
    ByRef Text As String, ByRef Search As String, _
    ByRef sOld As String, ByRef sNew As String, _
    ByVal Start As Long, ByVal Count As Long)
' by Jost Schwider, jost@schwider.de, 20001218
'-------------------------
Dim TextLen As Long, OldLen As Long, NewLen As Long
Dim ReadPos As Long, WritePos As Long, CopyLen As Long
Dim Buffer As String, BufferLen As Long
Dim BufferPosNew As Long, BufferPosNext As Long
  
    If Start < 2 Then
        Start = InStrB(Search, sOld)
    Else
        Start = InStrB(Start + Start - 1, Search, sOld)
    End If
    If Start Then
        OldLen = LenB(sOld): NewLen = LenB(sNew)
        Select Case NewLen
        Case OldLen
            Result = Text
            For Count = 1 To Count
                MidB$(Result, Start) = sNew
                Start = InStrB(Start + OldLen, Search, sOld)
                If Start = 0 Then Exit Sub
            Next Count
            Exit Sub
        Case Is < OldLen
            TextLen = LenB(Text)
            If TextLen > BufferLen Then Buffer = Text: BufferLen = TextLen
            ReadPos = 1: WritePos = 1
            If NewLen Then
                For Count = 1 To Count
                    CopyLen = Start - ReadPos
                    If CopyLen Then
                        BufferPosNew = WritePos + CopyLen
                        MidB$(Buffer, WritePos) = VBA.MidB$(Text, ReadPos, CopyLen)
                        MidB$(Buffer, BufferPosNew) = sNew
                        WritePos = BufferPosNew + NewLen
                    Else
                        MidB$(Buffer, WritePos) = sNew: WritePos = WritePos + NewLen
                    End If
                    ReadPos = Start + OldLen: Start = InStrB(ReadPos, Search, sOld)
                    If Start = 0 Then Exit For
                Next Count
            Else
                For Count = 1 To Count
                    CopyLen = Start - ReadPos
                    If CopyLen Then MidB$(Buffer, WritePos) = MidB$(Text, ReadPos, CopyLen): WritePos = WritePos + CopyLen
                    ReadPos = Start + OldLen
                    Start = InStrB(ReadPos, Search, sOld)
                    If Start = 0 Then Exit For
                Next Count
            End If
            If ReadPos > TextLen Then
                Result = LeftB$(Buffer, WritePos - 1)
            Else
                MidB$(Buffer, WritePos) = MidB$(Text, ReadPos)
                Result = LeftB$(Buffer, WritePos + LenB(Text) - ReadPos)
            End If
            Exit Sub
        Case Else
            TextLen = LenB(Text): BufferPosNew = TextLen + NewLen
            If BufferPosNew > BufferLen Then Buffer = Space$(BufferPosNew): BufferLen = LenB(Buffer)
            ReadPos = 1: WritePos = 1
            For Count = 1 To Count
                CopyLen = Start - ReadPos
                If CopyLen Then
                    BufferPosNew = WritePos + CopyLen: BufferPosNext = BufferPosNew + NewLen
                    If BufferPosNext > BufferLen Then Buffer = Buffer & Space$(BufferPosNext): BufferLen = LenB(Buffer)
                    MidB$(Buffer, WritePos) = MidB$(Text, ReadPos, CopyLen)
                    MidB$(Buffer, BufferPosNew) = sNew
                Else
                    BufferPosNext = WritePos + NewLen
                    If BufferPosNext > BufferLen Then Buffer = Buffer & Space$(BufferPosNext): BufferLen = LenB(Buffer)
                    MidB$(Buffer, WritePos) = sNew
                End If
                WritePos = BufferPosNext: ReadPos = Start + OldLen
                Start = InStrB(ReadPos, Search, sOld)
                If Start = 0 Then Exit For
            Next Count
            If ReadPos > TextLen Then
                Result = LeftB$(Buffer, WritePos - 1)
            Else
                BufferPosNext = WritePos + TextLen - ReadPos
                If BufferPosNext < BufferLen Then
                    MidB$(Buffer, WritePos) = MidB$(Text, ReadPos)
                    Result = LeftB$(Buffer, BufferPosNext)
                Else
                    Result = LeftB$(Buffer, WritePos - 1) & MidB$(Text, ReadPos)
                End If
            End If
            Exit Sub
        End Select
    Else
        Result = Text
    End If
End Sub
Public Function ReplaceMany(Source As String, _
    ParamArray ReplacePairs()) As String
' ���������� ������������� ������ � ������ ��������� �� ���������
'-------------------------
' Source - �������� ������
' ReplacePairs  - �������������� �������� � ���� "OldText=NewText"
'-------------------------
Const cDelim = ";", cTagDelim = "="
Dim Result As String
    On Error GoTo HandleError
    Result = Source: If Len(Result) = 0 Then GoTo HandleExit
Dim i As Long: i = 1 'LBound(Terms)             ' �������� � [%1%]
Dim sTerms As String:         sTerms = Join(ReplacePairs, cDelim)
Dim cTerms As New Collection, aKeys, sKey 'As String
    Call TaggedString2Collection(sTerms, cTerms, aKeys, cDelim, cTagDelim)
    For Each sKey In aKeys: Result = Replace(Result, sKey, cTerms(sKey)): Next 'sKey
HandleExit:     ReplaceMany = Result: Exit Function
HandleError:    Err.Clear: Resume HandleExit
End Function

Public Function Tokenize(Source As String, _
    Tokens() As String, _
    Optional Delims As String = c_strSymbSpaces & c_strSymbPunct & c_strSymbParenth & c_strSymbMath, _
    Optional Positions, _
    Optional IncEmpty As Boolean = False _
    ) As Long
' ��������� ������ �� ��������� �� ������ ������������
'-------------------------
' Source    - �������� ������
' Tokens()  - �� ������ �������� ������ ����� ���������� �� �������� ������ �� ������ ������������
' Delims    - ����� ��������� ������������
' Positions - (��������������) ������ ������� ������ ��������� � �������� ������ (���������� ��� ������/�������� ��������� ������)
' IncEmpty  = False - ������� ������ ��������� - ���������������� ����������� ����� ��������������� ��� ����
'           = True  - �������������� ������ ����� �������� ������ �������� ����� ����������������� �������������
'-------------------------
' v.1.1.0       : 31.10.2019 - �������� �������������� �������� Positions - ������������ ������ ������� ��������� ������� � �������� ������
' v.1.0.0       : 24.11.2000 - original Tokenize02 by Donald, donald@xbeat.net modified by G.Beckmann, G.Beckmann@NikoCity.de from http://www.xbeat.net/vbspeed/c_Tokenize.htm#Tokenize04
'-------------------------
Const ARR_CHUNK& = 1024
Dim cExp As Long, ubExpr As Long
Dim cDel As Long, ubDelim As Long
Dim aExpr() As Integer, aDelim() As Integer
Dim sa1 As SAFEARRAY1D, sa2 As SAFEARRAY1D
Dim cTokens As Long, iPos As Long
Dim bPos As Boolean
Dim Result As Long

    Result = -1
    On Error GoTo HandleError
    ubExpr = Len(Source): ubDelim = Len(Delims)
    ' ������� SAFEARRAY ��� �������� ������
    sa1.cbElements = 2:   sa1.cElements = ubExpr
    sa1.cDims = 1:        sa1.pvData = StrPtr(Source)
    ' ��������� ������ �������� �������� ������
    CopyMemory ByVal VarPtrArray(aExpr), VarPtr(sa1), PTR_LENGTH ' 4
    ' ������� SAFEARRAY ��� ������ ������������
    sa2.cbElements = 2:   sa2.cElements = ubDelim
    sa2.cDims = 1:        sa2.pvData = StrPtr(Delims)
    ' ��������� ������ �������� ������������
    CopyMemory ByVal VarPtrArray(aDelim), VarPtr(sa2), PTR_LENGTH ' 4
    
    ' �������������� �������������� �������
    If IncEmpty Then ReDim Preserve Tokens(ubExpr) Else ReDim Preserve Tokens(ubExpr \ 2)
    ' ��������� ������������� ���������� ������� ��������� ���������
    bPos = Not IsMissing(Positions): If bPos Then bPos = IsArray(Positions)
    If bPos Then If IncEmpty Then ReDim Preserve Positions(ubExpr) Else ReDim Preserve Positions(ubExpr \ 2)
    
    ubDelim = ubDelim - 1
    For cExp = 0 To ubExpr - 1
    ' ���������� ��� ������� �������� ������
        For cDel = 0 To ubDelim
    ' ���������� ��� ������� ������ ������������
            If aExpr(cExp) = aDelim(cDel) Then
                If cExp > iPos Then
        ' ���� ������� ������ �������� ������ ��������� � ������������
            ' � ����������� �� ��� ������������
            ' (���� ���������� ������ ����� ��� ������������ ���� �� cExp=iPos)
                ' ��������� �������� ������
                    Tokens(cTokens) = Mid$(Source, iPos + 1, cExp - iPos)
                    If bPos Then Positions(cTokens) = iPos + 1
                    cTokens = cTokens + 1
                ElseIf IncEmpty Then
            ' ��� ���� ������� ������ ������
                ' ��������� ������ ������
                    Tokens(cTokens) = vbNullString
                    If bPos Then Positions(cTokens) = iPos + 1
                    cTokens = cTokens + 1
                End If
        ' ��������� ������� ������ ���������� ������� ������
                iPos = cExp + 1: Exit For
            End If
        Next cDel
    Next cExp
    ' ���� ����� ���������� ����������� �������� ������� ��� ������� �������� ������ ������
    ' ��������� � �������
    If (cExp > iPos) Or IncEmpty Then
        Tokens(cTokens) = Mid$(Source, iPos + 1)
        If bPos Then Positions(cTokens) = iPos + 1
        cTokens = cTokens + 1
    End If
    ' �������� �������������� ������� �� ���������� ��������� ���������
    If cTokens = 0 Then Erase Tokens() Else ReDim Preserve Tokens(cTokens - 1)
    If bPos Then If cTokens = 0 Then Erase Positions() Else ReDim Preserve Positions(cTokens - 1)
    ' ���������� ���������� ��������� ���������
    Result = cTokens '- 1
    ' ������� ��������������� �������
    ZeroMemory ByVal VarPtrArray(aExpr), PTR_LENGTH '4
    ZeroMemory ByVal VarPtrArray(aDelim), PTR_LENGTH '4
HandleExit:  Tokenize = Result: Exit Function
HandleError: Result = -1: Err.Clear: Resume HandleExit
End Function
Public Function PlaceHoldersSetByIndex(Source As String, ParamArray Terms()) As String
' �������� �������������� ������� ���� [%n%] (��� n - ����� ���������), ���������� ������� ����������
'-------------------------
' Source - �������� ������
' Terms - �������������� ��������
Const LBr = "[%", RBr = "%]"                ' ��������� �����/������ ������ �������
Dim Result As String
    On Error GoTo HandleError
    Result = Source: If Len(Result) = 0 Then GoTo HandleExit
Dim i As Long: i = 1 'LBound(Terms)             ' �������� � [%1%]
Dim Term, sTemp As String
    If IsArray(Terms(0)) Then GoTo HandleArray
' ������� ������ ����������
    For Each Term In Terms
        sTemp = LBr & CStr(i) & RBr             ' ������� ������
        Result = Replace(Result, sTemp, Term)   ' ������
        i = i + 1
    Next
    GoTo HandleExit
HandleArray:
' �� ��� ������ ���� � ��������� �������� ������� ������
    For Each Term In Terms(0)
        sTemp = LBr & CStr(i) & RBr             ' ������� ������
        Result = Replace(Result, sTemp, Term)   ' ������
        i = i + 1
    Next
HandleExit:  PlaceHoldersSetByIndex = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function PlaceHoldersSetByNames(Source As String, _
    ParamArray NamedTerms()) As String
' �������� �������������� ������� ���� [%Param1%], ���������� ������� ����������
Const c_strProcedure = "PlaceHoldersSetByNames"
' Source - �������� ������
' NamedTerms  - �������������� �������� � ���� "Param1=Value1"
Const LBr = "[%", RBr = "%]"                ' ��������� �����/������ ������ �������
Const cDelim = ";"
Dim Result As String
    On Error GoTo HandleError
    Result = Source: If Len(Result) = 0 Then GoTo HandleExit
Dim i As Long: i = 1 'LBound(Terms)             ' �������� � [%1%]
Dim sTerms As String:         sTerms = Join(NamedTerms, cDelim)
Dim cTerms As New Collection: Call TaggedString2Collection(sTerms, cTerms, Delim:=cDelim)
    Result = PlaceHoldersSet(Result, cTerms, False, LBr, RBr)
HandleExit:     PlaceHoldersSetByNames = Result: Exit Function
HandleError:    Err.Clear: Resume HandleExit
End Function
Public Function PlaceHoldersSet(ByRef Source As String, _
    ByRef NamedTerms As Collection, _
    Optional AskMissing As Boolean = False, _
    Optional LBr As String = "[%", Optional RBr As String = "%]") As String
' �������� � ������ ������� �������������� ������� ���� [%Param1%] ���������� �� ���������
'-------------------------
' Source    - �������� ������
' NamedTerms - ��������� ������������� ��������. � �������� ����� �������������� ��������� ������ ���� �������� ���������
' AskMissing - ����������� �������� ������������� � ��������� ���������
' LBr/RBr   - �����/������ ������ ���������� ������� ����� ���������� (���� [%Param1%])
'-------------------------
' v.1.1.1       : 30.01.2020 - ��������� ����������� ������������� ������������� (������ ��. p_TermModify)
' v.1.1.0       : 20.01.2020 - ������ �������� ������ ����������. ������ ����� � �������� �������� � ��������� �������� ���������
'-------------------------
' ToDo: ������ ������ ����������� ��������� �������� ��������
'-------------------------
Dim Result As String
    On Error GoTo HandleError
    Result = Source: If Len(Result) = 0 Then GoTo HandleExit
Dim Term As String, Xpr As String, Key As String, Value As String
Dim i As Long: i = 1
' ��� ���.�������������. ������ ������������� ��. p_TermModify
Const c_ModLBr = "{", c_ModRBr = "}"
Dim Par As String, Pos As Long
    ' ���� ������� ���������� � ���������
    Do While p_FindNamedPlaceHolder(Result, Xpr, i, , LBr, RBr)
    ' ������� ������� ����������
        ' ����������� ������ �� ������� �������������� �������������
        Key = p_TermModify(Xpr, Par, Operation:=1)
        ' ���� �� ����� - ������ �������: Key = Xpr
    ' �������� �� �������� �� ��������� (��� ����������� ��� ����������)
        If p_IsExist(Key, NamedTerms, Term) Then
        ' ����� � ���������
        ElseIf AskMissing Then
        ' ����������� �������� ������������� � ��������� ����������
            Term = InputBox("������� �������� ���������� " & vbCrLf & Key & ":", "���������� �� �������!")
        ' ??? � ��������� � � ����� ��� ���������� �������������
            NamedTerms.Add Term, Key
        Else
        ' ���� �� ����� � �� ����������� - ��������� ��� ����
            ' ����� ������� ������ ������������� ����������
            ' ��� ������ �������� �������� ���������� ������������
            ' ����� �� ��� ���������� �������� ������ �������� �������������� ���������,
            ' �� ����� �� ��������� ������ ������������ ������� ���
            Term = LBr & Xpr & RBr: GoTo HandleNext
        End If
    ' ���������� ��������� (� ��������) ���������� �������� �� ������� � ��� ������� ����������
        Term = PlaceHoldersSet(Term, NamedTerms, AskMissing, LBr, RBr)
    ' ���� ���� ������������ - ��������� �� � Term
        If Len(Par) > 0 Then Term = p_TermModify(Term, Par, Operation:=0)
    '' ��������������, ��� ��� ����� �������� �������� ���������� ���������
        '    If EvalExpres Then
    '' ���� ������� ��������� ��������� � ������� �����������
        '    ' ��������� �������� ��������� �
        '    ' �������� ���������� ��������� � ��������� ��� ���������
        '        If p_IsEvalutable(Term, Value) Then
        '            Term = Value: With NamedTerms: .Remove (Key): .Add Key, Term: End With
        '        End If
        '    End If
    ' ���������� ������ ������� ���������� ���������� ��������� �� ����� ���������
        Result = Left$(Result, i - 1) & Replace(Result, LBr & Xpr & RBr, Term, i) ' ������
HandleNext:  i = i + Len(Term) 'If i > Len(Result) Then Exit Do ' ��������� ���� � ��������� ���� ������� ����������
    Loop
    'Erase Keys()
HandleExit:  PlaceHoldersSet = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Public Function PlaceHoldersGet(ByRef Source As String, ByRef Template As String, _
    Optional ByRef NamedTerms As Collection, Optional Keys, _
    Optional LBr As String = "[%", Optional RBr As String = "%]", _
    Optional ReplaceExisting As Integer = False, _
    Optional Method As Integer = 0, _
    Optional MultiSfx As String) As Boolean
' ��������� ������ �� ������������ ������� � ��������� �� ������ ��������� �������� ���������� ����������
'-------------------------
' Source    - �������� ������
' Template  - ������ ������, ����� ��������� �������������� ���������� ���� [%Param1%]
' NamedTerms - (������������ ��������) ��������� �������� �������������� ���������� ���� [%Param1%], ����������� �� �������� ������
' Keys      - (������������ ��������) ������ ��� ���������� ���� [%Param1%], ����������� �� �������� ������ (�������� ������� NamedTerms)
' LBr/RBr   - �����/������ ������ ���������� ������� ����� ���������� (���� [%Param1%])
' ReplaceExisting - ���������� ����������� � ��������� �������� � ������
' ���� ������ �������� ��������� ������ �� ���������� � ����� � ��� �� ������
'   0 - ����� ��������� ������ ��������
'  -1 - ����� ��������� ��������� ��������
'   1 - ����� ��������� ��� �������� � ���������� � ����������� ��������
' Method    - ������ ��������� �������� �������
'   0 - �� ������� ��������� (InStr)
'   1 - �� Like ��������� (InStrLike)
'   2 - �� RegEx ��������� (InStrRegEx)
' MultiSfx - ������� �������� ��� ������������� ���� (��� ReplaceExisting=1) �.�. ���-�� �������� ������������� � ������ ����������
'-------------------------
' v.1.0.2       : 16.07.2020 - �������� ������� ���� �� �������� - ������� ����������. ��������� ��������� ������������� ��������� (ReplaceExisting)
' v.1.0.1       : 04.02.2020 - �������� ��������� ������� �� ���� ��������� VBA.Like � VBS.RegExp
' v.1.0.0       : 03.02.2020 - �������� ������
'-------------------------
' ToDo: ������ ������ ����������� ���������: �������� ��������, ������ ���������� ��������, ����� ��� ����������� �������
' - ��� ReplaceExisting = -1 - ������ ��� ���������� � NamedTerms
'-------------------------
Const cSfx = "~&#" ' ������� �������� ��� ����� ���������� ��� ������������� ��������� (��-���������)
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    If Len(Source) = 0 Then GoTo HandleExit Else If Len(Template) = 0 Then GoTo HandleExit
Dim Xpr As String, Key As String, Item As String
Dim Part As String, Params As String, Found As String
Dim bLen As Integer: bLen = Len(LBr) + Len(RBr)
Dim bKeys As Boolean, aKeys() As String: bKeys = Not IsMissing(Keys)
Dim rBeg As Long, rEnd As Long, cBeg As New Collection
Dim tBeg As Long, tEnd As Long
Dim i As Long, j As Long, x As Long
Dim sKey As String
Dim bOK As Long
    i = 0
    rBeg = 1: tBeg = 1: tEnd = tBeg
    Set NamedTerms = New Collection
    If ReplaceExisting = 1 And Len(MultiSfx) = 0 Then MultiSfx = cSfx
    Do Until tBeg > Len(Template)
' ���� ������� ����������
        Found = vbNullString
    ' ���� ������� ������� ���������� �������� � ����������� ������
    ' ���� �� �������, �� �� ����� ������ ��� �������� ������� - ��������� ������� �� ����� ������
        ' (�������� �� ��������� ���������� ����������)
        If p_FindNamedPlaceHolder(Template, Xpr, tEnd, , LBr, RBr) Then Else tEnd = Len(Template) + 1
    ' ���� ����� �� ������� �� ����� ������� ������� ���������� �� ������ ���������
        Part = Mid$(Template, tBeg, tEnd - tBeg)
        If j = 0 Then ' ???
    ' ���� ��� ������ ����� ������� (�� ����������� ����������)- �������� ��� ��������� ��������� � ������
'            aBeg = InStrAll(rBeg, Source, Part, aFound, Method): jMax = UBound(aBeg)
            Do While rBeg <= Len(Source)
                j = j + 1
                Select Case Method
                Case 1:     rEnd = InStrLike(rBeg, Source, Part, Found)
                Case 2:     rEnd = InStrRegEx(rBeg, Source, Part, Found)
                Case Else:  rEnd = InStr(rBeg, Source, Part): Found = Part
                End Select
                If rEnd = 0 Then Exit Do
                rBeg = rEnd + Len(Found)
                'sKey = c_idxPref & j     ' ??? ��������� ��� ����� - ����������� ��� �������� ���������� ��� ������ ������������� �����������
                cBeg.Add rBeg ', sKey
            Loop
        End If
    ' ���� ���� � ����� - ����� ��������� ���������, ����� - �������� � �������
        If ReplaceExisting = -1 Then j = cBeg.Count Else j = 1
    ' ���� ��� ������ ������ ������� ������ ����� ������ (����� ����������) - ���� ������
        If Len(Key) = 0 Then GoTo HandleNext
        i = i + 1   ' ����������� ������ ���������� � �������
        Do Until j > cBeg.Count
' ���������� ��������� ����������
            If Len(Part) > 0 Then
' ���������� ����������� ������ � ������� ������ �������
    ' ���� ����� ����������� ���������� ���� �������� ������� ��� �������������
' !!! ��������� ��� ��������� ��� � ����� ������ ������ ����������� ����������
    ' �������� ���� ��������� ����� � ������ ����� �� ���� ������� Params �������� �����
    ' ���������� ����������� ��� �� ����������.
' ���� ����� ����� ��� ��� ����������� ���������� ��������
            Select Case Method
            Case 1:     rEnd = InStrLike(cBeg(j), Source, Part, Found)
            Case 2:     rEnd = InStrRegEx(cBeg(j), Source, Part, Found)
            Case Else:  rEnd = InStr(cBeg(j), Source, Part): If rEnd > 0 Then Found = Part Else Found = vbNullString
            End Select
            End If
        ' ���� ����� �� ������ - ����������� �������� �� ���������� ��������� � ��������� � ����������
            bOK = rEnd > 0
            If Not bOK Then GoTo HandleNotOk
        ' ������� ������� ����������� - ��������� �������� ����������� ����������
            Item = Mid$(Source, cBeg(j), rEnd - cBeg(j))
'' <<< ����� ����� ��������� ������������ Item ��������� � Params
'            If Len(Params) > 0 Then
'    '         ' ���� Item �� ����� Params - ������� �������� �� �������� ����������,
'    '         ' � ��������� �������� ������� - ??? �������� ��� �����������
'Stop
'                Item = p_TermModify(Item, Params, Operation:=0)
'                bOk = ??? '
'            End If
'            If Not bOk then GoTo HandleNotOk
    ' ��������� ������� ��������� ��������� ��� ����������� �������������
            rBeg = rEnd + Len(Found)
            cBeg.Remove (j): If j <= cBeg.Count Then cBeg.Add rBeg, Before:=j Else cBeg.Add rBeg, After:=cBeg.Count
    ' ��������� ��� � ��������� ��� ������������ ��������
        ' ���� ��������� ������������ ��������� - ���� ����������� ��� ���������� �� �������
        ' ���� ��������� ��� ���������� - ��������� �� ������ ����� ���������� �� ������� � ������� �������� � ���������
            sKey = Key
            If ReplaceExisting = 1 And j > 1 Then sKey = sKey & MultiSfx & (j - 1)
            ' ����� ����� ��������� ���������� � NamedTerms ���,
            ' ����� ��� ������� ��������� � ������ ������� ��������� ������
            ' ��� �������� �������� ���� ����� ������������ �������� �������������
'Stop
            If NamedTerms.Count = 0 Then NamedTerms.Add Item, sKey Else NamedTerms.Add Item, sKey, After:=j * i - 1    ', Before:=
            ' ����������� ������ ���������� � ����������
            If ReplaceExisting = 1 Then j = j + 1 Else Exit Do
            GoTo HandleNextVar
HandleNotOk:
    ' ������� ������� �� �����������
        ' ������� ��� �� ��������� ���������� ��� �������������� ����������� ������� ���� �����
            cBeg.Remove (j)
        ' ����� ���� ��������� �������� ����������� ���������� ����������������� ��������
            If ReplaceExisting <> 1 Then
            ' ��� ������ ��������� ���������� - ������ �������� NamedTerms � aKeys
                Set NamedTerms = New Collection
                If ReplaceExisting = -1 Then j = cBeg.Count
            Else
            ' ��� ������������� - ���� ������ ���������� ����������
            ' �� ��������� (�������) ����� �� ������� j � NamedTerms
                For x = 1 To i - 1: NamedTerms.Remove (j * i - x): Next x
            End If
HandleNextVar:
        Loop
HandleNext:
        If tEnd > Len(Template) Then Exit Do
' ��������� �� ������� ���������� ���������� ��� ��� �������������� �������������
        Key = p_TermModify(Xpr, Params, Operation:=1)
' ��������� ��������� � ������ ������� �� ������ ����� ��������� ������� ����������
        tBeg = tEnd + Len(Xpr) + bLen
        tEnd = tBeg
    Loop
'HandleExitDo:
    If bKeys Then Keys = p_GetCollKeys(NamedTerms)
HandleExit:  PlaceHoldersGet = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_FindNamedPlaceHolder(ByRef Source As String, _
    Optional ByRef Name As String, _
    Optional ByRef sBeg As Long = 1, Optional ByRef sEnd As Long = 0, _
    Optional LBr As String = "[%", Optional RBr As String = "%]") As Boolean
' ���� � ������ ������� ����������, ������������ �������������, ���������� � ��� � �������
'-------------------------
' Source - ������ � ������� ������������ �����
' Name - ��� ��������� ���������� (��� ������)
' sBeg - ������� ������ ��������� ���������� � ������ (������� ������)
' sEnd - ������� ����� ��������� ���������� � ������ (������� ������)
' LBr/RBr - �����/������ ������ ���������� ������� ����� ���������� (���� %Param1%)
'-------------------------
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    If Len(Source) = 0 Then GoTo HandleExit
Dim pBeg As Long, pEnd As Long
    ' ���� � ��������� ����� ������
    pBeg = InStr(sBeg, Source, LBr): If pBeg = 0 Then GoTo HandleExit Else sBeg = pBeg: pBeg = pBeg + Len(LBr)
    ' ���� � ��������� ������ ������
    pEnd = InStr(pBeg, Source, RBr): If pEnd = 0 Then GoTo HandleExit Else sEnd = pEnd + Len(RBr)
    ' �������� ������ ����� ��������
    Name = Mid$(Source, pBeg, pEnd - pBeg)
    Result = True 'Len(Name) > 0
HandleExit:  p_FindNamedPlaceHolder = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Private Function p_TermModify(ByVal Term As String, ByRef Params As String, _
    Optional Operation = 0, Optional ReplaceExisting As Integer = False) As String
' ������������ ������ � ��������������
'-------------------------
' Term      - �������������� ��������
' Params    - ������ ����������
' Operation - �������� ������������ ��� ���������
'   0 - ��������� ��������� ������ ������������� Params � �������� Term � ���������� � Result
'   1 - ���������� ������� � ������ �������������, ����������:
'       � Params - ��������� ������ �������������, � Result - ��� ���������
' ������� � ����� ������� ����� ��� �������� �������� ������� ������������� ��������� � ����� �����
' ReplaceExisting - ���� ������������ ��������� ��� ����������� ���������� � ���������� ������ ��� ������� ����������
'   0 - ��������� ������, ����������� ����� ��������������
'  -1 - ��������� ������ ���� ����� ���������� - ��������� ���������
'-------------------------
' v.1.0.3       : 05.02.2020 - ��� �������� ��������� �������� � ��������� ������� ������������� ���������� ���� � ���������� �������������
'-------------------------
' ������ ������ �������������: {�����������1:��������1-1,...,��������1-X1;...;�����������N:��������N-1,...,��������N-XN}
Const c_ModLBr = "{", c_ModRBr = "}" ' ������ ���������� ������ ������������� � ���������
Const cXprDelim = ";" ' ����������� ��������� ������������� � ������
Const cNamDelim = ":" ' ����������� �����/���������� ������������
Const cParDelim = "," ' ����������� ���������� ������������
Dim Pos As Long
Dim Result As String: Result = Term
    On Error GoTo HandleError
    If Len(Term) = 0 Then GoTo HandleExit
    Select Case Operation
    Case 0
' ��������� Term � ���������� ���������� ������������
        If Len(Params) = 0 Then GoTo HandleExit
        Dim Xpr, Par As String
    ' �������� ������ ������������ � �����������
        For Each Xpr In Split(Params, cXprDelim)
    ' ���������� ������ �������������
        ' �������� ��� ������������ � ������ ��� ����������
            Pos = InStr(1, Xpr, cNamDelim)
            If Pos > 0 Then Par = Mid$(Xpr, Pos + Len(cNamDelim)): Xpr = Left$(Xpr, Pos - 1)
    ' ��������� �����������
            Result = p_TermModifyXprGet(Result, Xpr, Par, cParDelim, ReplaceExisting)
        Next Xpr
    Case Else
' ���������� �� Term ����� ���������, ����������� ������� � ��� �������������
    ' ���������� ������ � ��������� ������ �������������
        Result = Term: Params = vbNullString
        Pos = InStr(1, Result, c_ModLBr): If Pos = 0 Then GoTo HandleExit
        If Right$(Result, Len(c_ModRBr)) <> c_ModRBr Then GoTo HandleExit
        Term = Left$(Result, Pos - 1): Pos = Pos + Len(c_ModLBr)
        Params = Mid$(Result, Pos, Len(Result) - Pos - Len(c_ModRBr) + 1)
        Result = Term
    End Select
HandleExit:  p_TermModify = Result: Exit Function
HandleError: Result = Term: Err.Clear: Resume HandleExit
End Function
Private Function p_TermModifyXprGet(Term As String, _
    Modificator, Optional Params As String, _
    Optional ParDelim As String, _
    Optional ReplaceExisting As Integer = False) As String
' ��������� ����������� � �������� � ���������� ���������
'-------------------------
' Term - �������������� ��������
' Modificator - ��� ������������ ������������ � ��������
' Params - ����� ���������� ������������
' ParDelim - ����������� ���������� � ������
' ReplaceExisting - ���� ������������ ������ ������� �� ���������� ���������
'-------------------------
' v.1.0.2       : 31.01.2020 - ������ ������ �������� ���������� �������� ������������� �� ����� �������
'-------------------------
Dim Result As String: Result = Term
    On Error GoTo HandleError
    If Len(Modificator) = 0 Then GoTo HandleExit
Dim sFun As String, sKey As String, sVal
Dim cPar As Collection
    Select Case LCase(Modificator)
' <<< ����� ����� ������� ���������� ����� ������������� � ������ ���������� ��� �������
' !!! ������� �� �������� ���������� � ���������������� �������� !!! - �� ����� ���������� �� ����������, ���������� ��������� ���� ������
    Case "�������", "ucase":    Result = UCase(Result)
    Case "������", "lcase":     Result = LCase(Result)
    Case "�����������", "pcase": Result = StrConv(Result, vbProperCase)
    Case "��������", "decline":   sFun = "DeclineWords('" & Result & "'"
                ' �������� ��������� ������������
                    Set cPar = p_TermModifyParGet(Params, ParDelim, ReplaceExisting)
                ' ��������� ��������� �������
                    sKey = "NewCase": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    sKey = "NewNumb": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    sKey = "NewGend": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    sKey = "Animate": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    sKey = "IsFio": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    'sKey = "SkipWords": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, vbNullString)
                    sFun = sFun & ")"
                    If p_IsEvalutable(sFun, Result) Then Else Err.Raise vbObjectError + 512
    Case "�����������", "numtowords": sFun = "NumToWords(" & Result
                ' �������� ��������� ������������
                    Set cPar = p_TermModifyParGet(Params, ParDelim, ReplaceExisting)
                ' ��������� ��������� �������
                    sKey = "NewCase": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    sKey = "NewNumb": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    sKey = "NewGend": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    sKey = "Animate": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    sKey = "NewType": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    'sKey = "Unit": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, vbNullString)
                    'sKey = "SubUnit": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, vbNullString)
                    'sKey = "DecimalPlaces": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    'sKey = "TranslateFrac": sFun = sFun & "," & IIf(p_IsExist(sKey, cPar, sVal), sVal, 0)
                    sFun = sFun & ")"
                    If p_IsEvalutable(sFun, Result) Then Else Err.Raise vbObjectError + 512
    Case "�������", "in" ' �������� �������������� �������� ������ ���������� ��������, ��������� � Params
        ' �� PlaceHolderGet - ���������� �������� �� ������ - ��������� ��������� �������� �� ������ - ��� ������������ - ����������, ��� �������������� ���������� ����� � ������
        ' �� PlaceHolderSet - ��������� �������� � ������ - ��������� ������������� �������� �� ������������ ������ - ���� ������������� - �����������, � ���� ��� - ???
                ' �������� ��������� ������������ (������ ���������� ��������)
                    Result = vbNullString
                    Set cPar = p_TermModifyParGet(Params, ParDelim, ReplaceExisting)
                    For Each sVal In cPar
                        If Left$(Term, Len(sVal)) = sVal Then Result = sVal: Exit For
                    Next sVal
    Case "���", "is"     ' �������� ������������ ����, ��������� � Params
                    Result = vbNullString
                    Set cPar = p_TermModifyParGet(Params, ParDelim, ReplaceExisting)
                    Dim l As Long
                    For Each sVal In cPar
                        For l = Len(Term) To 1 Step -1
                            sKey = Left$(Term, l)
                            Select Case sVal
                            Case "�����", "num":    If IsNumeric(sKey) Then Result = sKey
                            Case "����", "date":    If IsDate(sKey) Then Result = sKey
                            Case "�����", "word":   sFun = "^[a-zA-Z�-��-߸�]*$": RegEx.Pattern = sFun: If RegEx.Test(sKey) Then Result = sKey
                            Case "�����", "text":   Result = sKey ' ��� ����� �������� ��
                            Case "��������", "rus": sFun = "^[�-��-߸�]*$": RegEx.Pattern = sFun: If RegEx.Test(sKey) Then Result = sKey
                            Case "��������", "eng": sFun = "^[a-zA-Z]*$": RegEx.Pattern = sFun: If RegEx.Test(sKey) Then Result = sKey
                            'Case "���", "var":   sFun = "^[a-zA-Z�-��-߸�][_a-zA-Z�-��-߸�0-9]*$": RegEx.Pattern = sFun: If RegEx.Test(sKey) Then Result = sKey
                            Case Else: GoTo HandleExit ' ����� ������� �������� �� �����-�� ����������� ����� ��������
                            End Select
                            If Len(Result) > 0 Then Exit For
                        Next l
                        If Len(Result) > 0 Then Exit For
                    Next sVal
'    Case "����","if"     ' �������� ������� ���������� ������� ��������������� �������
'    Case "�����","choose"    ' �������� ���������� ������� ���������� ������� ��������������� ������� ��������� �������
    Case Else
    End Select
HandleExit:  p_TermModifyXprGet = Result: Exit Function
HandleError: Result = Term: Err.Clear: Resume HandleExit
End Function
Private Function p_TermModifyParGet(Params As String, _
    Optional ParDelim As String, _
    Optional ReplaceExisting As Integer = False) As Collection
' ������������ ����� ���������� ������������� ������ � ��������� ���������� ������� � ���������� � ���� ���������
'-------------------------
' Params - ����� ���������� ������������
' ParDelim - ����������� ���������� � ������
' ReplaceExisting - ���� ������������ ������ ������� �� ���������� ���������
Dim sKey As String, sVal 'As String
Dim cPar As New Collection, Par
    For Each Par In Split(Params, ParDelim)
        Select Case LCase(Par)
' <<< ����� ����� ������� ���������� ����� ���������� �������, ���������� �������������� � �� ��������
        Case "�����":   sKey = "NewType": sVal = NumeralOrdinal
        Case "�����":   sKey = "NewType": sVal = NumeralCardinal
        Case "��":      sKey = "NewCase": sVal = DeclineCaseImen
        Case "���":     sKey = "NewCase": sVal = DeclineCaseRod
        Case "���":     sKey = "NewCase": sVal = DeclineCaseDat
        Case "���":     sKey = "NewCase": sVal = DeclineCaseVin
        Case "��":      sKey = "NewCase": sVal = DeclineCaseTvor
        Case "����":    sKey = "NewCase": sVal = DeclineCasePred
        Case "��":      sKey = "NewNumb": sVal = DeclineNumbSingle
        Case "��":      sKey = "NewNumb": sVal = DeclineNumbPlural
        Case "���":     sKey = "NewGend": sVal = DeclineGendMale
        Case "���":     sKey = "NewGend": sVal = DeclineGendFem
        Case "c�":      sKey = "NewGend": sVal = DeclineGendNeut
        Case "����":    sKey = "Animate": sVal = True
        Case "���":     sKey = "IsFio":   sVal = True
        Case Else:      sKey = c_idxPref & Par: sVal = Par ' ������ ������ ��������� ��� ���� ����� ��� �����-�� �����
        'Case Else:      GoTo HandleNext                    ' ����������� ����������
        End Select
    ' ��������� �������� � ���������
        If p_IsExist(sKey, cPar) Then If Not ReplaceExisting Then GoTo HandleNext Else cPar.Remove sKey
        cPar.Add sVal, sKey
HandleNext: Next Par: Set p_TermModifyParGet = cPar
End Function

Public Function GroupsGet(Source As String, _
    ByRef cGroups As Collection, _
    Optional UsePlaceHolders As Boolean = False, _
    Optional Templates, Optional TermDelim = "@", Optional TempDelim = ";", _
    Optional aGroups) As Boolean
' ���������� ��������� ����� ������������ � ������ (��������� ����������� � ������)
'-------------------------
' Source    - ��������� ���������� ������
' cGroups   - (������������) ��������� ����������� ������ ������ �������� ������������� ������ ������ � ������� �������
'             ����������� ������� ������������� ����������� ������ � ��������� Br
'             ��������� ����� ��� ����������� ������������� ���������� �������� � ��������� PlaceHoldersGet/Set
' UsePlaceHolders - ���� True � Text, ����� ���������� ��������� ���������� �������������� ������ �� ����� �������� ������� ����������
'             ����: ([%1%])+([%2%]), ��� 1,2.. - ������� ��������� ��������� cGroups �������� ���������� ������
'             ����� - ������ ��������� ��������� ������������ � �������.
' Templates - ������ ��� ������ ����� ���������� ������� ���������� �����
'             �.�. ������ ����������� ������ ��������� �����-�������, ��� �����-����� ���������� ������
'             ���� ����� ������� ���� ����������� �� ���� ���������� � ������� ���������� ������������
' TermDelim - ����������� ��������� (���������� ������ ��� ����������� ������������ �������� ������) � ������ �������
' TempDelim - ����������� �������� � ������
' aGroups   - (������������) ������ ������� ��������� ������ (������� �����/������ �����/����������� �����) ����� ������ ���� ������ ����������� ������� � ������
' ����������: True  - ���� ��������� ������� ���������,
'             False - ���� ��������� �������� ���������� ������ ��� �� ���������
'-------------------------
' v.1.0.2       : 12.03.2024 - ������ ������� ���������� ������ ��� �������
' v.1.0.1       : 21.12.2022 - ���������� �������������� ������. (�� ��� ������ �����������������)
' v.1.0.0       : 24.03.2020 - �������� (����� ������ � �������) ������
'-------------------------
' �������:
' 1) strText = "Do: If True Then 1 Else 0 End If: Loop"
'    strTemp = "If @ Then @ Else @ End If;Do: @: Loop"
'    Call GroupsGet(strText, cGroup, True, strTemp)
' 2) strText = "((5+2)+3*(4+5)^4)-97"
'    Call GroupsGet(strText, cGroup, True)
'-------------------------
' ToDo: - �������� ����������� �� ������ ���������� ������, �� �������� ����� ����������� � ������ �������� � ���������� ����� ��������� �� ������, - ���� �������������
'         ��������� �������:
'           1) "������-�������" - ��� ������� �� ������������ ������� ����� ������ ������������ � ����� � ������ ��������� � ������ � ���������� ������ �� ���������� �������
'               ����� - ����������� �������� ����� � ��� �� ����������
'           2) "����������" - ��� ������� ��������� ��������� ���� ����������� �� �������� �������� � ����������� �� �� ���� ����������� �� �������
'               ����� - �������� � ���� �������� ��������� �� ��� ����� �������/��������
'           3) "��������������" - ����� ������� ������ ����������� ������� �� ���� �� ���������� � ������� ���������� ������������
'               ����� - ��������� �� ����� ��������� (�������� ������������ ���������, ����������� �� ����������� ����� ��������, .. ���???)
'               � �� ���� ��� ��� �������� ��� ����������� ������ ��������, �������� ��������������� ��-�� �������� � �������������� ���������
'           ����� ���������� ����� ������ �� 3 � 1.
'           ������� ������������ �������������� ������ (�� �������), ����� ��������� ���������� ���������,
'           � �� ��� �� ��������� ����������� ������� ��������� ��� �������
'       - ������������� ������������� � �������������� ��������� � �������� �������� (@[,@]) �.�. (@);(@,@),(@,@,@) � �.�.
'-------------------------
#Const TestErr = False          ' ��������� ������ ��������������� ������
Const cPref = "Br"              ' ������� ������������ �������� ���������
Const errUnclosedExp = vbObjectError + 511
Const errIncompleteExp = vbObjectError + 512
Dim Result As Boolean ': Result = False
On Error GoTo HandleError
    If Len(Source) = 0 Then Result = True: GoTo HandleExit
' ������ ���������� ������� ����� (������)
Dim sTerm, sName As String
Dim aTemp                       ' ������ �������� ��������� �������
Dim t As Long                   ' ������ ������� � �������
    If IsMissing(Templates) Then
' �� ������ - ���� ����� ������ ��-���������
        aTemp = Array(Array("(", ")"), Array("[", "]"), Array("{", "}"), Array("<", ">"), Array("%", "%"), Array("'", "'"), Array("""", """"))
    ElseIf IsArray(Templates) Then
' ������ ���������� �������� (�� �����������)
        ReDim aTemp(LBound(aTemp), UBound(aTemp)): For t = LBound(aTemp) To UBound(aTemp): aTemp(t) = Split(Templates(t), TermDelim): Next t
    Else
' ������ �������
        ReDim aTemp(0 To 0): For Each sTerm In Split(Templates, TempDelim): ReDim Preserve aTemp(0 To t): aTemp(t) = Split(sTerm, TermDelim): t = t + 1: Next sTerm
    End If
Dim l As Long           ' ������ ������������ �������� ������� = LBound(aTemp(t)) - ����������� ������, =UBound(aTemp(t)) - ����������� ������, ��������� - ������������� ������
Dim i As Long           ' ������� ������� � ����������� ������
Dim j As Long           ' ������ �����
Dim g As Long           ' ������ �������� ������� ��� �������� ���������� �������
Dim aStack() As Long    ' ��������� ���� ������
Const sStep = 3         ' ��� ��������� ������� �����
                ' +0    '(.TempNum) ����� ������� �� �������
                ' +1    '(.TempItm) ����� �������� �������
                ' +2    '(.TempBeg) ������� ������ ��������� � �������� ������ (������� ������)
'Dim aGroups() As Long ' ������ ��� �������� ���������� �������
Const gStep = 7         ' ��� ��������� ������� ������� ���������
                ' +0    '(.TextLev) ������� ����������� (0-��� ������, 1-������� ������, ... n-������ n-������)
                ' +1    '(.TextBeg) ������� ������ ������ ��������� � �������� ������ (����� ����������� ������)
                ' +2    '(.TextEnd) ������� ����� ������ ��������� � �������� ������ (�� ����������� ������)
                ' +3    '(.TempBeg) ������� ������ ��������� � �������� ������ (������� ������)
                ' +4    '(.TempEnd) ������� ����� ��������� � �������� ������ (������� ������)
                ' +5    '(.TempNum) ����� ������� �� �������
                ' +6    '(.TempItm) ����� �������� � �������
Dim iBeg As Long, iEnd As Long, iLen As Long
    ReDim aGroups(1 To 1) As Long
    Set cGroups = New Collection
' ���� ������ ������ � ������ ��������� ����
    i = 1
    Do Until i > Len(Source)
' ��������� ������� � ������� �������
        iLen = 1        ' ������������� ������ �����������
        If j > 0 Then
    ' ���� � ����� ���� ���������� ������
            ' ��������� ��������� ������� ��� ������� � ������� �����
            t = aStack(j - 2)                       '(.TempNum) ����� ������� �� �������
            l = aStack(j - 1)                       '(.TempItm) ����� �������� �������
            sTerm = aTemp(t)(l + 1)
            If sTerm = Mid$(Source, i, Len(sTerm)) Then
            ' ���� ��������� - ��������� �������� ������ � ���������
            ' ������� �������� � ���������
                g = g + gStep: ReDim Preserve aGroups(1 To g) 'As Long
                iLen = Len(aTemp(t)(l))             ' ����� ��������� ����������� �������� �������
                iBeg = aStack(j - 0)                ' ������� ������ ��������� ����������� �������� ������� � �������� ������
                iEnd = iBeg + iLen                  ' ������� ����� ��������� ����������� �������� ������� � �������� ������
                iLen = Len(sTerm)                   ' ����� ��������� �������� �������� �������
                l = l + 1                           ' ��������� � ���������� �������� ������� �������
                
                aGroups(g - 6) = j \ sStep        '(.TextLev) ������� ����������� (0-��� ������, 1-������� ������, ... n-������ n-������)
                aGroups(g - 5) = iEnd             '(.TextBeg) ������� ������ ������ ��������� � �������� ������ (����� ����������� ������)
                aGroups(g - 4) = i                '(.TextEnd) ������� ����� ������ ��������� � �������� ������ (�� ����������� ������)
                aGroups(g - 3) = iBeg             '(.TempBeg) ������� ������ ��������� � �������� ������ (������� ������)
                aGroups(g - 2) = i + iLen         '(.TempEnd) ������� ����� ��������� � �������� ������ (������� ������)
                aGroups(g - 1) = t                '(.TempNum) ����� ������� �� �������
                aGroups(g - 0) = l                '(.TempItm) ����� �������� � �������
            ' ������� ���������� ������ � �������������� ���������
                sTerm = Mid$(Source, iEnd, i - iEnd)                    ' ��������� �� ������� ������ ��������
                sName = cPref & (g \ gStep): cGroups.Add sTerm, sName   ' ��������� � ���������
                If l = UBound(aTemp(t)) Then
            ' ���� ������� ������� ����������� - ��������� ������� �����
                    j = j - sStep: If j > 0 Then ReDim Preserve aStack(1 To j) Else Erase aStack   ' ��������� ����
                Else
            ' ����� ����������� � ����� ������� �������� ������� � ��� �������
                    aStack(j - 1) = l               '(.TempLev) ����� �������� �������
                    aStack(j - 0) = i               '(.TempBeg) ������� ������ ��������� � �������� ������ (������� ������)
                End If
                GoTo HandleNextSym                  ' �������� ������ � �������� - ������� � ���������� �������
            End If
        End If
' ��������� ������ (�����������) ������� ���� ��������
        For t = LBound(aTemp) To UBound(aTemp)
            l = LBound(aTemp(t)): sTerm = aTemp(t)(l)
            If sTerm = Mid(Source, i, Len(sTerm)) Then
            ' ���� ����������� ������� ��������� � ������� ���������� ������
                If l < UBound(aTemp(t)) Then
                ' ���� ������� ������� �� ����������� (������� ����������� ����� � ����������� � �����������) - ������� ��� � ����
                    j = j + sStep: ReDim Preserve aStack(1 To j) ' ����������� ����
                    aStack(j - 2) = t               '(.TempNum) ����� ������� �� �������
                    aStack(j - 1) = l               '(.TempItm) ����� �������� �������
                    aStack(j - 0) = i               '(.TempBeg) ������� ������ ��������� � �������� ������ (������� ������)
                End If
                iLen = Len(sTerm)                   ' ������� ������� � ������ �� ����� ���������� ���������
                GoTo HandleNextSym ': Exit For      ' �������� ������ � �������� - ������� � ���������� �������
            End If
        Next t
#If TestErr Then
' ����� ������������� ��������� �������� ���������� ������ - ��������� ������������ ��� � ��������� ������� �� ���������� ������������ � ������� ������� ������
        For t = LBound(aTemp) To UBound(aTemp)
            For l = LBound(aTemp(t)) + 1 To UBound(aTemp(t))
            sTerm = aTemp(t)(l): If sTerm = Mid$(Source, i, Len(sTerm)) Then sName = aTemp(t)(l - 1): Err.Raise errIncompleteExp
        Next l: Next t
#End If
' ������� �������� ������ - ���������� ������ - ������ ��������� � ���������� �������
HandleNextSym: i = i + iLen    ' ������� ��������� � ������ �� ��������� ����� ������������������� ������
    Loop
'#If TestErr Then
'' ��������� ������� � ����� ���������� ������
    If j <> 0 Then sTerm = Join(aTemp(aStack(j - 2)), "..."): i = aStack(j): Err.Raise errUnclosedExp
'#End If
    Erase aStack
' ���� UsePlaceHolders=False ������������� ��������� ������� ��������� � ��������� ���,
' �.�. ��� ��������� � Source, �� ��� ������������ - �������.
    ' ��������� ������� ������� ������� ���������� ��� �������� ������
' ������� �������� � ���������
    g = g + gStep: ReDim Preserve aGroups(1 To g) 'As tTerm
    'aGroups(g - 6) = 0                '(.TextLev) ������� ����������� (0-��� ������)
    aGroups(g - 5) = 1                '(.TextBeg) ������� ������ ������ ��������� � �������� ������ (����� ����������� ������)
    aGroups(g - 4) = Len(Source) + 1  '(.TextEnd) ������� ����� ������ ��������� � �������� ������ (�� ����������� ������)
    aGroups(g - 3) = 1                '(.TempBeg) ������� ������ ��������� � �������� ������ (������� ������)
    aGroups(g - 2) = Len(Source) + 1  '(.TempEnd) ������� ����� ��������� � �������� ������ (������� ������)
    aGroups(g - 1) = -1               '(.TempNum) ����� ������� �� �������
    aGroups(g - 0) = -1               '(.TempItm) ����� �������� � �������
    sTerm = Source                      '
    sName = cPref & (g \ gStep): cGroups.Add sTerm, sName     ' ��������� � ���������
    Result = True:
' ���� ���� ��������� ������� ������� ��������� � ������� - ������ ���
    If UsePlaceHolders Then Call p_GroupsPlaceHoldersSet(cGroups, aGroups)
HandleExit:     GroupsGet = Result: Exit Function
HandleError:    Select Case Err
    Case errUnclosedExp:    Debug.Print "������! ������������� ��������� """ & sTerm & """ � ������� " & i & " � ������: """ & Source & """"
    Case errIncompleteExp:  Debug.Print "������! """ & sTerm & """ ��� """ & sName & """ � ������� " & i & " � ������: """ & Source & """"
    Case Else: Stop: Resume 0
    End Select
    Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_GroupsPlaceHoldersSet(ByRef cGroups As Collection, aGroups) As Boolean
' ������������� ��������� ��������� ����������� ����� ������� �������� ��������� �����������
Dim Result As Boolean ': Result = False
' �������� � ��������� ������� ��� ��������� ����������
On Error GoTo HandleError
Const cLBr = "[%", cRBr = "%]"  ' ������ ��� ������ �� �������� ������� ����������.
Const cPref = "Br"              ' ������� ������������ �������� ���������
Const gStep = 7                 ' ��� ��������� ������� ������� ���������
Dim g As Long                   ' ������ �������� ������� ��� �������� ���������� �������
Dim i As Long, j As Long, iMin As Long
Dim iBeg As Long, iEnd As Long
Dim jBeg As Long, jEnd As Long
Dim iLvl As Long, jLvl As Long
Dim sTerm, sName As String
    i = UBound(aGroups) \ gStep: iMin = 1
    Do While i > iMin 'For i = i To 2 Step -1
' ��������� ��� ����������� �������� ������� � �������� �������
    ' �������� � ������� ������������� ��� ��� �������� (�� � ������� Level ������) ����� ����
    ' ��������� - ���� ������� ������������ (j) �������� ����� ������ ������������ (i)
    ' �������� ���������� ������������ �������� � ����������� �� ���������� ���������
    ' ������� ������� ������� �������� � ����������� �������� �� ������ �������������� ���������
            j = iMin
        ' ����� ��������� �� �������
            sName = cPref & i
            sTerm = cGroups(sName)
        ' ������� � ������� ������������ ��������� � �������� ������
            g = (i - 1) * gStep + 1
            iLvl = aGroups(g + 0)
            iBeg = aGroups(g + 1)
            iEnd = aGroups(g + 2)
        Do While j < i 'For j = 1 To i - 1
    ' ��������� ��� ����������� �������� ������� � ������ ������� �� ������������
        ' ������� � ������� ������������ ��������� � �������� ������
            g = (j - 1) * gStep + 1
        ' ����������� ������� ������ ������������ ����������� ������ ����������� ������������ ������������
            jLvl = aGroups(g + 0): If iLvl <> (jLvl - 1) Then GoTo HandleNextJ    'iLvl > jLvl -> Next j
        ' ������� ������������ �������� ������� ������ � �������� ������ ������������
            jBeg = aGroups(g + 1): If iBeg > jBeg Then GoTo HandleNextJ           'iBeg > jBeg -> Next j
            jEnd = aGroups(g + 2): If iEnd < jEnd Then GoTo HandleNextJ           'iEnd < jEnd -> Next j
        ' ���� ������ ������� j
            ' ����� ������� �� ����� ������ (������ ��������� ��������)
            sTerm = Left$(sTerm, Len(sTerm) - (iEnd - jBeg)) & cLBr & j & cRBr & Right$(sTerm, iEnd - jEnd)
        ' � �������� ������ ������� ���������������� ���������
            iBeg = aGroups(g + 3) + 1     ' ������� ������� �� ������ �������������� ��������� �� ������� ����� ���������
            If j = iMin Then iMin = j + 1   ' ������� ��������� ������ ������� ����������� ��������� �������
                                            ' (��������� ������ ������� ��� �� ���������� ��� ������ ��� ��������� �����)
        '' ���� �������� ������� j
        '    ' ����� ������� �� ������ ������ (����� ��������� ��������)
        '    ??? 'sTerm = Left$(sTerm, (iEnd - jBeg + 1)) & cLBr & j & cRBr & Right$(sTerm, Len(sTerm) - iBeg - jEnd)
        '    ' � �������� ������� ������� ���������������� ���������
        '    iEnd = aGroups(g + 4) - 1      ' ������� ������� �� ����� �������������� ��������� �� ������� �� ���������
        '    ??? 'If j = iMin Then iMin = j   '
        ' ��������� ��������� �� �������� ��������
            If iEnd <= iBeg Then Exit Do   ' ������ �������� ��������� �������
            If i = iMin Then iMin = i + 1   ' ??? ������� ��������� ������ ������� ����������� ��������� �������
                                            ' (��������� ����������� ������ ������� ��� �� ���������� ��� ������ ��� ��������� �����)
HandleNextJ: j = j + 1
    Loop 'Next j
    ' ���������� ��������� ������� � ���������
        With cGroups: .Remove sName: .Add sTerm, sName, After:=i - 1: End With
HandleNextI: i = i - 1
    Loop 'Next i
HandleExit:  p_GroupsPlaceHoldersSet = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function GroupText(Source As String, idx As Long, _
    Optional Templates, Optional TermDelim = "@", Optional TempDelim = ";" _
    ) As String
' ���������� ���������� ������ �������� �� ������ (��������� ����������� � ������)
'-------------------------
' Source    - ��������� ���������� ������
' Idx       - ������ ������ ���������� ������� ����������
' Templates - ������ ��� ������ ����� ���������� ������� ���������� �����
'             �.�. ������ ����������� ������ ��������� ��������� ������ ���� ������� �������
' TermDelim  - ����������� ��������� (���������� ������ ��� ����������� ������������ �������� ������) � ������ �������
' TempDelim  - ����������� �������� � ������
'-------------------------
On Error GoTo HandleError
Dim cGroups As Collection: Call GroupsGet(Source, cGroups, , Templates, TermDelim, TempDelim)    'If Not .. Then Err.Raise vbObjectError + 512 'Exit Function
    GroupText = cGroups(idx) ' ����
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function InStrLike(Start As Long, _
    String1 As String, String2 As String, _
    Optional Found As String, _
    Optional Compare As VbCompareMethod = vbTextCompare) As Long
' InStr ����������� ������ ���������� �� ����� ��������� Like
'-------------------------
' Start     - ��������� �������
' String1   - ������ � ������� ���������� �����
' String2   - ������ ���������� ������ ����� ������
' Found     - (������������) �������� �� ����� ���������
' Compare   - ������ ���������
' ���������� ������� ������� ��������� String2 � String1 ������� � ������� Start
'-------------------------
' v.1.0.1       : 05.02.2020 - ��������� ��� ������� ��������� � ������������ ����
' v.1.0.0       : 23.02.2003 - original by VictorB212 from http://www.vbforums.com/showthread.php?232259-InStrLike-(debugging-help-required)
'-------------------------
Const cSymBeg = "^" ' ������ �������� � ������ (��� � RegEx) - ������������ �.� ��������� ������ ����� Start ��� 0
Const cSymEnd = "$" ' ������ �������� � �����
Dim Result As Long: Result = False
    On Error GoTo HandleError
    If Start <= 0 Then Start = 1
    Found = vbNullString
Dim S1 As String, S2 As String
    S1 = Mid$(String1, Start): S2 = String2
' ��������� ������ � ������ ������
Dim bBeg As Boolean: bBeg = Left$(String2, Len(cSymBeg)) = cSymBeg:  If bBeg Then S2 = Mid$(S2, Len(cSymBeg) + 1) Else S2 = "*" & S2
Dim bEnd As Boolean: bEnd = Right$(String2, Len(cSymEnd)) = cSymEnd: If bEnd Then S2 = Left$(S2, Len(S2) - Len(cSymBeg)) Else S2 = S2 & "*"
    If Compare = vbTextCompare Then S1 = UCase$(S1): S2 = UCase$(S2)
' ��������������� �������� ������������ ������ �������
Dim iSgn As Integer: iSgn = S1 Like S2: If Not iSgn Then GoTo HandleExit
Dim lLen As Long, lPos As Long
' ���� ������ �������
HandleRightBound:
    lLen = Len(S1): lPos = lLen
    If bEnd Then GoTo HandleLeftBound           ' ���� �������� � ������� ���� - ������ ������� �������� - ���� �����
    Do
        If Not iSgn Then iSgn = 1 Else If Not (Left$(S1, lPos - 1) Like S2) Then Exit Do
        lLen = lLen \ 2: If lLen < 1 Then lLen = 1  ' ����� �������� �������
        lPos = lPos + iSgn * lLen                   ' ��������� ������� ���� ���������, ����� - �����������
        iSgn = Left$(S1, lPos) Like S2              ' ��������� ����������
    Loop
' ���� ����� �������
HandleLeftBound:
    S1 = Left$(S1, lPos)                            ' �������� ������ �� �������� �������
    lLen = Len(S1)
    If bBeg Then lPos = 0: GoTo HandleResult        ' ���� �������� � ������ ���� - ����� ������� �������� - �������� ���������
    lPos = lLen                     '
    Do
        If Not iSgn Then iSgn = 1 Else If Not (Right$(S1, lPos - 1) Like S2) Then Exit Do
        lLen = lLen \ 2: If lLen < 1 Then lLen = 1  ' ����� �������� �������
        lPos = lPos + iSgn * lLen                   ' ��������� ������� ���� ���������, ����� - �����������
        iSgn = Right$(S1, lPos) Like S2             ' ��������� ����������
    Loop
    lLen = lPos: lPos = Len(S1) - lPos
' ���������� ���������
HandleResult:
    Result = Start + lPos: If Result Then Found = Mid$(String1, Result, lLen)
HandleExit:  InStrLike = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function InStrRegEx( _
    Start As Long, _
    String1 As String, String2 As String, _
    Optional Found As String, _
    Optional Compare As VbCompareMethod = vbTextCompare) As Long
' InStr ����������� ������ ���������� �� ����� ��������� RegEx
'-------------------------
' Start     - ��������� �������
' String1   - ������ � ������� ���������� �����
' String2   - ������ ���������� ������ ����� ������
' Found     - (������������) �������� �� ����� ���������
' ���������� ������� ������� ��������� String2 � String1 ������� � ������� Start
'-------------------------
' v.1.0.0       : 25.08.2010 - original by BC_Programmer https://www.computerhope.com/forum/index.php?topic=109171.msg736986#msg736986
'-------------------------
Dim Result As Long: Result = False
    On Error GoTo HandleError
    If Start <= 0 Then Start = 1
    Found = vbNullString
Dim S1 As String: S1 = Mid$(String1, Start)   'shortened version of String1
Dim oMatches As Object, oMatch As Object
    ' ����� RegExp � �������� ��� �����
    With RegEx: .IgnoreCase = (Compare = vbTextCompare): .Pattern = String2: Set oMatches = .Execute(S1): End With
    If oMatches.Count = 0 Then GoTo HandleError
    For Each oMatch In oMatches
        Result = oMatch.FirstIndex + Start: Found = oMatch
        Exit For
    Next
HandleExit:  InStrRegEx = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function InStrAll( _
    Start As Long, _
    String1 As String, String2 As String, _
    Optional Found, _
    Optional Method As Integer = 0, _
    Optional Compare As VbCompareMethod = vbTextCompare)
' InStr ����������� ������ ��� ���������� �� ���������
Const c_strProcedure = "InStrAll"
' Start     - ��������� �������
' String1   - ������ � ������� ���������� �����
' String2   - ������ ���������� ������ ����� ������
' Found     - (������������) ������ ��������� �� ����� ��������
' Method    - ������ ��������� �������� �������
'   0 - �� ������� ��������� (InStr)
'   1 - �� Like ��������� (InStrLike)
'   2 - �� RegEx ��������� (InStrRegEx)
' ���������� ������ ������� ��������� String2 � String1 ������� � ������� Start
'-------------------------
' v.1.0.0       : 16.07.2020 -
'-------------------------
    On Error GoTo HandleError
    If Start <= 0 Then Start = 1
Dim sFound As String
Dim bFound As Boolean: bFound = Not IsMissing(Found)
Dim S1 As String: S1 = Mid$(String1, Start)   'shortened version of String1
Dim lPos As Long, lLen As Long: lLen = Len(S1): lPos = 1
Dim aResult() As Long, aFound() As String
'Dim cResult As New Collection, cFound As New Collection
Dim i As Long: i = 0
    Do While lPos <= lLen
        Select Case Method
        Case 1:     lPos = InStrLike(lPos, S1, String2, sFound)
        Case 2:     lPos = InStrRegEx(lPos, S1, String2, sFound)
        Case Else:  lPos = InStr(lPos, S1, String2): sFound = String2
        End Select
        If lPos = 0 Then Exit Do
        i = i + 1
        ReDim Preserve aResult(1 To i): aResult(i) = lPos: If bFound Then ReDim Preserve aFound(1 To i): aFound(i) = sFound
        'cResult.Add lPos: If bFound Then cFound.Add sFound
        lPos = lPos + Len(sFound)
    Loop
'HandleExit:  Set InStrAll = cResult: If bFound Then Set Found = cFound: Exit Function
HandleExit:  InStrAll = aResult: If bFound Then Found = aFound: Exit Function
HandleError: Erase aResult: Err.Clear: Resume HandleExit
End Function
Public Static Function InStrCount( _
    ByRef Text As String, _
    ByRef Find As String, _
    Optional ByVal Start As Long = 1, _
    Optional ByVal Compare As VbCompareMethod = vbBinaryCompare _
    ) As Long
' ���������� ���������� �������� � ������
Const c_strProcedure = "InStrCount"
' Text - ����� � ������� ������������ �����
' Find - ������� ��������� ������� ���������� ��������� ������� ������������
' Start - ��������� ������� ������
' Compare - ��� ���������
'-------------------------
' v.1.0.1       : 21.11.2001 - original InStrCount04 by Jost Schwider, jost@schwider.de from http://www.xbeat.net/vbspeed/c_InStrCount.htm#InStrCount04
'-------------------------
Const MODEMARGIN = 8
Dim TextAsc() As Integer
Dim TextData As LongPtr, TextPtr As LongPtr
Dim FindAsc(0 To MODEMARGIN) As Integer
Dim FindLen As Long
Dim FindChar1 As Integer
Dim FindChar2 As Integer
Dim i As Long

    If Compare = vbBinaryCompare Then
' �������� ���������
        FindLen = Len(Find)
        If FindLen Then
    ' �������� ������� ����������
            If Start < 2 Then Start = InStrB(Text, Find) Else Start = InStrB(Start + Start - 1, Text, Find)
            If Start Then
        ' ���� ����������� ���������
                InStrCount = 1
                If FindLen <= MODEMARGIN Then
            ' ��� ���� ������� ��������� �� MODEMARGIN - ������� ������
                ' ���������� ���������� �������
                    If TextPtr = 0 Then ReDim TextAsc(1 To 1): TextData = VarPtr(TextAsc(1)):
                    CopyMemory TextPtr, ByVal VarPtrArray(TextAsc), PTR_LENGTH: TextPtr = TextPtr + 8 + PTR_LENGTH
                ' ������������� �������
                    CopyMemory ByVal TextPtr, ByVal VarPtr(Text), PTR_LENGTH            'pvData
                    CopyMemory ByVal TextPtr + PTR_LENGTH, Len(Text), 4 ' PTR_LENGTH    'nElements
                    Select Case FindLen
                    Case 1 ' � ������ ���� ����
                        FindChar1 = AscW(Find)
                        For Start = Start \ 2 + 2 To Len(Text)
                            If TextAsc(Start) = FindChar1 Then InStrCount = InStrCount + 1
                        Next Start
                    Case 2 ' � ������ ��� �����
                        FindChar1 = AscW(Find): FindChar2 = AscW(Right$(Find, 1))
                        For Start = Start \ 2 + 3 To Len(Text) - 1
                            If TextAsc(Start) = FindChar1 Then
                                If TextAsc(Start + 1) = FindChar2 Then
                                    InStrCount = InStrCount + 1: Start = Start + 1
                                End If
                            End If
                        Next Start
                    Case Else ' � ������ ������ ���� ������
                        CopyMemory ByVal VarPtr(FindAsc(0)), ByVal StrPtr(Find), FindLen + FindLen
                        FindLen = FindLen - 1
                        ' ������ ��� �����
                        FindChar1 = FindAsc(0): FindChar2 = FindAsc(1)
                        For Start = Start \ 2 + 2 + FindLen To Len(Text) - FindLen
                            If TextAsc(Start) = FindChar1 Then
                                If TextAsc(Start + 1) = FindChar2 Then
                                    For i = 2 To FindLen
                                        If TextAsc(Start + i) <> FindAsc(i) Then Exit For
                                    Next i
                                    If i > FindLen Then
                                        InStrCount = InStrCount + 1: Start = Start + FindLen
                                    End If
                                End If
                            End If
                        Next Start
                    End Select
                ' ��������������� �������� �� �������
                    CopyMemory ByVal TextPtr, TextData, PTR_LENGTH 'pvData
                    CopyMemory ByVal TextPtr + PTR_LENGTH, 1&, 4 'PTR_LENGTH  'nElements
                Else
            ' ��� ������� ���� - ������� ������
                    FindLen = FindLen + FindLen
                    Start = InStrB(Start + FindLen, Text, Find)
                    Do While Start
                        InStrCount = InStrCount + 1
                        Start = InStrB(Start + FindLen, Text, Find)
                    Loop
                End If 'FindLen <= MODEMARGIN
            End If 'Start
        End If 'FindLen
    Else
' ��������� ���������
    ' ���������� ������� �������
        InStrCount = InStrCount(LCase$(Text), LCase$(Find), Start)
    End If
End Function
Public Function Replicate(ByVal Number As Long, Pattern As String) As String
' Returns a pattern replicated in a string a specified number of times.
'-------------------------
' Number  - Required. Number of replications desired.
' Pattern - Required. Character pattern to replicate.
'-------------------------
' v.1.0.2       : 06.12.2000 - original Replicate05 by Donald, donald@xbeat.net from http://www.xbeat.net/vbspeed/c_Replicate.htm#Replicate05
'-------------------------
Dim lp As Long
    If Number > 0 Then
        lp = Len(Pattern)
        Select Case lp
        Case Is > 1:    Replicate = Space$(Number * lp): Mid$(Replicate, 1, lp) = Pattern: If Number > 1 Then Mid$(Replicate, lp + 1) = Replicate
        Case 1:         Replicate = String$(Number, Pattern)
        End Select
    End If
End Function
Public Function WordWrap( _
    ByRef Text As String, _
    ByVal Width As Long, _
    Optional ByRef CountLines As Long) As String
' ��������� ������ �� �������� ���������� �������� � �������������� ������� ������� ������.
'-------------------------
' Text  - ����������� ������
' Width - ����� ������ � ��������
'-------------------------
' v.1.0.0       : 13.09.2004 - original WordWrap01 by Donald, donald@xbeat.net from http://www.xbeat.net/vbspeed/c_WordWrap.htm#WordWrap01
'-------------------------
Dim i As Long, lenLine As Long
Dim posBreak As Long, cntBreakChars As Long
Dim abText() As Byte, abTextOut() As Byte
Dim ubText As Long

    If Width <= 0 Then CountLines = 0: Exit Function
    If Len(Text) <= Width Then CountLines = 1: WordWrap = Text: Exit Function
    abText = StrConv(Text, vbFromUnicode): ubText = UBound(abText)
    ReDim abTextOut(ubText * 3)     'dim to potential max
    For i = 0 To ubText
        Select Case abText(i)
        Case 32, 45: posBreak = i   ' ������ � �������
        Case Else
        End Select
        abTextOut(i + cntBreakChars) = abText(i)
        lenLine = lenLine + 1
        If lenLine > Width Then
            If posBreak > 0 Then
                If posBreak = ubText Then Exit For ' don't break at the very end
                ' ������ ����� ������� ��� ��������
                abTextOut(posBreak + cntBreakChars + 1) = 13  'CR
                abTextOut(posBreak + cntBreakChars + 2) = 10  'LF
                i = posBreak: posBreak = 0
            Else ' cut word
                abTextOut(i + cntBreakChars) = 13     'CR
                abTextOut(i + cntBreakChars + 1) = 10 'LF
                i = i - 1
            End If
            cntBreakChars = cntBreakChars + 2: lenLine = 0
        End If
    Next
    CountLines = cntBreakChars \ 2 + 1
    ReDim Preserve abTextOut(ubText + cntBreakChars)
    WordWrap = StrConv(abTextOut, vbUnicode)
End Function
Public Function Compress( _
    sExpression As String, _
    sCompress As String, _
    Optional Compare As VbCompareMethod = vbBinaryCompare) As String
' Returns a string where multiple adjacent occurrences of a specified substring are compressed to just one occurrence. For example, the function will compress multiple spaces within a string down to single spaces.
'-------------------------
' sExpression - Required. String expression containing substring sequences to be compressed. If sExpression is a zero-length string, Compress returns a zero-length string as well.
' sCompress - Required. The single string whereof sequences are to be compressed. Read sCompress as "compress multiples of". Can be longer than one char. If sCompress is a zero-length string, the function returns sExpression unmodified.
' Compare - Optional. Numeric value indicating the kind of comparison to use when evaluating substrings. If omitted, the default value is 0, which means a binary comparison is performed.
'-------------------------
' v.1.0.0       : 04.11.2001 - original Compress06 by Tom Winters, tom@interplanetary.freeserve.co.uk from http://www.xbeat.net/vbspeed/c_Compress.htm#Compress06
'-------------------------
Dim sExp$, sFind$, lLenCompress&, lLenExpression&
Dim lChrPosition&

    lLenExpression = Len(sExpression)
    If lLenExpression = 0 Then Exit Function
    lLenCompress = Len(sCompress)
    If lLenCompress <> 0 Then
        If lLenCompress = 1 Then
    ' ������� ������ �������� �� ������ �������
            If lLenExpression < 10 Then
        ' ��������� ������ ��������
                sFind = sCompress + sCompress ' ���� ��� ���������� ������� ������
                Compress = sExpression
                ' ����������� ���� ���� ����������
                ' ���� ����� ������ �������� �� �������� �������
                Do
                    lChrPosition = InStr(1, Compress, sFind, Compare)
                    If lChrPosition = 0 Then Exit Function
                    sExp = Left$(Compress, lChrPosition)
                    Compress = Right$(Compress, Len(Compress) - Len(sExp) - lLenCompress)
                    Compress = sExp + Compress
                Loop
            Else
    ' ��� ������� ����� ������
        ' ��������� ���������� ������ ������
            ' Ideally we'd check the the entire string for segment matches,
            ' but if we do that we'll be here for ever
            ' So, we'll use a reasonable compromise..
        ' �������� ������ 12 �������� ������  ���� 2/3 ����������
            Dim sNewSearchString As String
                sExp = Left$(sExpression, 12)
                sNewSearchString = String$(8, sCompress)
                lChrPosition = InStr(1, sExp, sNewSearchString, Compare)
                ' ���� � ��� ������� ������������� ������ ������
                If lChrPosition > 0 Then
        ' ������� �������������� ������������� ������ � ������� ���������
                Dim lLenNewSearchString As Long, lLenFind2 As Long, lStringSizeCounter As Long
                    lLenFind2 = lLenCompress + lLenCompress
                    lStringSizeCounter = (lLenExpression - lLenFind2)
                    ' Make new search string divisible by 2
                    lStringSizeCounter = lStringSizeCounter + (lStringSizeCounter And 1)
                    ' ������� ����� ������ ������
                    sNewSearchString = String$(lStringSizeCounter, sCompress)
                    lLenNewSearchString = Len(sNewSearchString)
                    lStringSizeCounter = 0
                    Compress = sExpression
                    sFind = sCompress + sCompress
                ' ���� �� ���� ������� ������ ������� ������ ������� ������� ������������������, ����� - �������
                ' ������� ��� ������� ����� ����� ����������� ����� ���������� �������� ����� ������������������ ���� ���������� �� ����� �������
                ' ���� ����� ���������� ������ ������������������ � �������� ��������������������
                    Do
                        Do
                            lChrPosition = InStr(1, Compress, sNewSearchString, Compare)
                            If lChrPosition = 0 Then Exit Do
                            sExp = Left$(Compress, lChrPosition)
                            Compress = Right$(Compress, Len(Compress) - Len(sExp) - lLenNewSearchString + lLenCompress)
                            Compress = sExp + Compress
                            ' Reset quantity removed from search string
                            lStringSizeCounter = 0
                        Loop
                        lChrPosition = InStr(1, Compress, sFind, Compare)
                        If lChrPosition = 0 Then Exit Function
                        ' Increase size of search string counter
                        lStringSizeCounter = lStringSizeCounter + lLenCompress
                        ' Make search string smaller and try again
                        sNewSearchString = Right$(sNewSearchString, Len(Compress) - lStringSizeCounter)
                        lLenNewSearchString = Len(sNewSearchString)
                    Loop
                End If
            End If
        End If
' ��������� Unicode
    ' We can save a lot of work by not passing Unicode strings to our byte arrays
    ' so we test for Unicode first and if true, use another Instring method
    Dim lCharacter As Long, lAsciiValue As Long
        For lCharacter = 1 To lLenCompress
            lAsciiValue = Asc(Mid$(sCompress, lCharacter, 1))
            If lAsciiValue > 127 Then
                ' If we have a Unicode multiple search character
                Dim bGo As Boolean, lPosition&
                sExp = sExpression
                ' This method is fast with more complex expressions
                Do While Len(sExp) > 0
                    bGo = False
                    lPosition = InStr(1, sExp, sCompress, Compare)
                    If Mid$(sExp, lPosition + lLenCompress, lLenCompress) = sCompress Then
                        If lPosition = 1 Then
                            bGo = True
                        End If
                    End If
                    If bGo Then
                        sExp = Right$(sExp, Len(sExp) - lLenCompress)
                    Else
                        Compress = Compress + Left$(sExp, 1)
                        sExp = Right$(sExp, Len(sExp) - 1)
                    End If
                Loop
                Exit Function
            End If
        Next
' ������ � �������������� ��������� �������
    Dim bMatch As Boolean, bMatchResult1 As Boolean, bMatchResult2 As Boolean
    Dim lLenExpressionArray&, lLenCompressArray&, lbytePosition&, lNewCounter&
    Dim byExpressionArray() As Byte, byNewArray() As Byte, byCompressArray() As Byte
    Dim lNearEndofExpression&, lExpCounter&, lLenCompressArrayplus1&
    ' Set case according to status of comparison
        If Compare = vbTextCompare Then sExpression = LCase$(sExpression): sCompress = LCase$(sCompress)
        ' ����������� ������ � �������� ������
        byExpressionArray = sExpression: byCompressArray = sCompress
        ' �������� ������ ��������� ������� �� ����� ��������� ����� ������ (������� ������� ��� UBound)
        lLenExpressionArray = lLenExpression + lLenExpression - 1: lLenCompressArray = lLenCompress + lLenCompress - 1
        ReDim byNewArray(lLenExpressionArray): lNewCounter = 0
        bMatch = Left$(sExpression, 1) = sCompress
' C������ �������������� ��������� ������ ��� ������ ��������� �������
        ' Equate character match in boolean terms, applying logic "And"
        If Not bMatch And (lLenCompressArray = 1) Then
            For lbytePosition = 1 To lLenExpressionArray
                ' For/Next loops work faster advancing from 1
                ' But we want lbytePosition to start at byte zero so;
                lbytePosition = lbytePosition - 1
                If byExpressionArray(lbytePosition) <> byCompressArray(0) Then
                    ' Set this byte of byNewArray() equal to byExpressionArray()
                    byNewArray(lNewCounter) = byExpressionArray(lbytePosition): lNewCounter = lNewCounter + 2
                Else
                    ' Set this byte of byNewArray() equal to byCompressArray()
                    If byExpressionArray(lbytePosition - 2) <> byCompressArray(0) Then
                        byNewArray(lNewCounter) = byCompressArray(0): lNewCounter = lNewCounter + 2
                    End If
                End If
                ' As we want to advance 2 Hex places on each loop cycle
                ' (the range of 1 byte being 0-255 (&H0 to &HFF, 2^8 or 8-bits))
                ' we add 2 to lbytePosition, (a touch faster than using "Step 2")
                lbytePosition = lbytePosition + 2
            Next
        Else
' ������� ��������������� ��������� ������ ��� ������ ��������� �������
    ' For a 3 character compress string (could be any length here)
    ' if the expression contains "xbz" and the findstring is "abc"
    ' bMatchResult1 of ("x" And "a") = False
    ' bMatchResult2 of ("b" And "b") = True
    ' We "And" these two results so that (in this example);
    ' (bMatchResult1 And bMatchResult2) = (False And True) = False
    ' Pass this result to bMatchResult1 and pass a new test to bMatchResult2
    ' bMatchResult2 result: ("z" And "c") = False
    ' -> (bMatchResult1 And bMatchResult2) = (False And False) = False!
    
    ' We only get a result of True, if all characters in the stream match
            lNewCounter = 0
            lLenCompressArrayplus1 = lLenCompressArray + 1
            lNearEndofExpression = lLenExpressionArray - (lLenCompressArray - 1)
            bMatchResult1 = True
            ' lbytePosition is at start of expression, so we init subsequent loop by
            ' obtaining a match result for bFindArray() in bExpressionArray()
            For lbytePosition = 1 To lLenCompressArrayplus1
                lbytePosition = lbytePosition - 1
                ' Equate character match in boolean terms
                bMatchResult2 = byExpressionArray(lbytePosition) = byCompressArray(lbytePosition)
                ' Colate results to date and save
                bMatchResult1 = bMatchResult1 And bMatchResult2
                ' Provided no match was found in stream
                If Not bMatchResult1 Then
                    ' Reset byNewArray() counter to zero and exit loop
                    lNewCounter = 0
                    Exit For
                End If
                ' ..or advance both array counters and check next character
                byNewArray(lNewCounter) = byCompressArray(lbytePosition)
                lNewCounter = lNewCounter + 2
                lbytePosition = lbytePosition + 2
            Next
            ' With this result for first term done, we loop to build bNewArray()
            For lExpCounter = 1 To lLenExpressionArray
                lExpCounter = lExpCounter - 1
                ' Reset current result to False
                bMatch = False
                ' If we're not at end of the expression
                If lExpCounter < lNearEndofExpression Then
                    bMatch = True
                    ' For each element in byCompressArray(), check for match with expression
                    For lbytePosition = 1 To lLenCompressArray
                        lbytePosition = lbytePosition - 1
                        bMatchResult2 = byExpressionArray(lExpCounter + lbytePosition) = byCompressArray(lbytePosition)
                        bMatch = bMatch And bMatchResult2
                        lbytePosition = lbytePosition + 2
                    Next
                End If
                ' If no match found, set just one byte in byNewArray()
                ' to current byExpressionArray() byte value..
                If Not bMatch Then
                    byNewArray(lNewCounter) = byExpressionArray(lExpCounter)
                    ' ..and advance both counters to next byte
                    lNewCounter = lNewCounter + 2
                    lExpCounter = lExpCounter + 2
                ' Provided the previous pass through loop did not achieve a match
                ElseIf Not bMatchResult1 Then
                    ' Work through find string stream [ bFindArray() ]
                    For lbytePosition = 1 To lLenCompressArray
                        lbytePosition = lbytePosition - 1
                        ' Set current byte in byNewArray() = byCompressArray()
                        ' and advance all counters
                        byNewArray(lNewCounter) = byCompressArray(lbytePosition)
                        lNewCounter = lNewCounter + 2
                        lExpCounter = lExpCounter + 2
                        lbytePosition = lbytePosition + 2
                    Next
                Else
                    ' If no match was found
                    ' just advance lExpressionCounter beyond length of search string
                    lExpCounter = lExpCounter + lLenCompressArrayplus1
                End If
                ' ..accumulating results thus far
                bMatchResult1 = bMatch
            ' ..and try again
            Next
        End If
' ����������� byNewArray() � ������ � �������� ������ ������ �� ���������� (������� ������� ��� ReDim Preserve � �������� ��������)
        Compress = byNewArray: Compress = Left$(Compress, lNewCounter * 0.5)
        Exit Function
    Else
' Handle Error
    ' ���� ������� ������������������ ���� ������� ����� ���������� �������� ���������
        Compress = sExpression
    End If
End Function
Public Function IsAscII(txt As String) As Boolean
' ��������� �������� �� ������ ASCII �������
'-------------------------
    If Len(txt) = LenB(txt) Then IsAscII = True: Exit Function
Dim i As Long
    For i = 1 To Len(txt)
        If Asc(MidB$(txt, 2 * i, 1) & vbNullChar) <> 0 Then Exit Function  ' False
    Next i
    IsAscII = True
End Function
' ==================
' ������� ��� ���������/�����������/������������ �����
' ==================
Public Function DelimStringGet(ByRef Source As String, _
    ByVal Pos As Long, _
    Optional Delim As String = " ", _
    Optional sBeg As Long, Optional sEnd As Long _
    ) As String
' ���������� �������� ������ � ������������� � ��������� ��������
'-------------------------
' Source    - �������� ������
' Pos       - ������� ����������� ���������
' Delim     - �����������
' sBeg,sEnd - ���������� ������� ������ � ��������� ����������� ��������� � ��������
'-------------------------
Dim Result As String: Result = vbNullString
    If Len(Source) = 0 Then GoTo HandleExit
'    If Pos < 1 then Goto HandleExit
'' ��� ������ Split - ��������, �� ��������� �������
    'Result = Split(Source, Delim)(Pos - 1)
'' ��� ������ InStr - ���� �������, �� ������ �������
    'Dim i As Long: i = 1: sBeg = 1
    'Do
    '    sEnd = InStr(sBeg, Source, Delim)
    '    If sEnd = 0 Then sEnd = Len(Source) + 1: Exit Do
    '    i = i + 1: If i > Pos Then Exit Do
    '    sBeg = sEnd + Len(Delim)
    'Loop
    'Result = Mid$(Source, sBeg, sEnd - sBeg)
' ������� � ������ �������� ������ ��������� - ������������������ ������ � ������������� Split, �� ������� ������� ������������� ��� �����������
    ' ��������� ������������ ������������� ������� (� ����� ������)
    Call p_GetSubstrBounds(Source, Pos, sBeg, sEnd, Delim)
    Result = Mid$(Source, sBeg, sEnd - sBeg)
HandleExit:  DelimStringGet = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function DelimStringDel(Source As String, _
    ByVal Pos As Long, _
    Optional Delim As String = " ", _
    Optional sBeg As Long, Optional sEnd As Long _
    ) As String
' ���������� �������� ������ � ������������� ��� �������� � ��������� ��������
'-------------------------
' Source    - �������� ������
' Pos       - ������� ��������� ���������
' Delim     - �����������
' sBeg,sEnd - ���������� ������� ��������� ���������
'-------------------------
Dim Result As String: Result = Source
    On Error GoTo HandleError
    If Len(Source) = 0 Then GoTo HandleExit
'    If Pos < 1 then Goto HandleExit
'' ��� ������ Split
'Dim arr() As String
'    arr = Split(Result, Delim): arr(Pos - 1) = vbNullString
'    Result = Replace(Join(arr, Delim), Delim & Delim, Delim): Erase arr()
'' ��� ������ InStr ���������� DelimStringGet
' ������� � ������ �������� ������ ��������� - ����� ������������ ������������� ������� (� ����� ������)
    Call p_GetSubstrBounds(Source, Pos, sBeg, sEnd, Delim)
    If sBeg = 1 Then sEnd = sEnd + Len(Delim) Else sBeg = sBeg - Len(Delim)
    Result = Left$(Source, sBeg - 1) & Mid$(Source, sEnd)
HandleExit:  sEnd = sBeg: DelimStringDel = Result: Exit Function
HandleError: Result = Source: Err.Clear: Resume HandleExit
End Function
Public Function DelimStringSet(Source As String, _
    ByVal Pos As Long, ByVal Data As String, _
    Optional Delim As String = " ", _
    Optional SetUnique As Integer = False, _
    Optional Overwrite As Boolean = False, _
    Optional sBeg As Long, Optional sEnd As Long _
    ) As String
' ��������� � ������ � ������������� ������� � ������� � ��������� ��������
'-------------------------
' Source    - �������� ������
' Pos       - ������� ������� ��������
' Data      - ����������� ������
' Delim     - �����������
' SetUnique = False - ������� ��������� ���������� �� � ������� � ��������
'           = True  - ������� ����� ����������� ������ ���� ��������� ����������� � ��������, ����� - �������� ������ ��������� ��� ���������
'           = 1     - ������� ����� ����������� ������ ���� ��������� ����������� � ��������, ���� ��������� ��� ������������ � ��������, - ��������� ����� ������� �� ��������, � ����� ��������� � ��������� �������
' Overwrite = False - ������� �� ������� (��� Pos>0 ������� ����� ��������� ��������, ��� Pos<0 - �����. �.�. Pos=1 - ������� �������, � Pos=-1 - � ����� ������)
'           = True  - ������� � �������  �������� ������ � ��������� �������,
'             (!)     � ������������ � SetUnique<>0 ����� ��������� � ����������� �����������
' sBeg,sEnd - ���������� ������� ������ � ��������� ����������� ��������� � ��������
'-------------------------
' v.1.0.1       : 09.08.2022 - �������� �������� SetUnique ��� �������� ������������ ����������� ��������
'-------------------------
Dim Result As String: Result = Source
    On Error GoTo HandleError
    If Len(Result) = 0 Then Result = Data: GoTo HandleExit
    ' ��������� �������� ������ �� ������� ��������� ���������,
    If SetUnique Then
        Select Case SetUnique
        Case 1      ' ���� ����������� ������� ��� ���� � ������ - ������� ��������� ��������� � ����������
            If Result = Data Then GoTo HandleExit
            Result = Replace(Result, Delim & Data & Delim, Delim)
            If Left$(Result, Len(Data & Delim)) = Data & Delim Then Result = Mid$(Result, Len(Data & Delim) + 1)
            If Right$(Result, Len(Delim & Data)) = Delim & Data Then Result = Left$(Result, Len(Result) - Len(Data & Delim))
        Case True   ' ���� ����������� ������� ��� ���� � ������ - �����
            If Result = Data Then GoTo HandleExit
            sBeg = InStr(1, Result, Delim & Data & Delim): If sBeg Then sBeg = sBeg + Len(Delim): GoTo HandleExit
            If Left$(Result, Len(Data & Delim)) = Data & Delim Then sBeg = 1: GoTo HandleExit
            If Right$(Result, Len(Delim & Data)) = Delim & Data Then sBeg = Len(Result) - Len(Data) + 1: GoTo HandleExit
        End Select
    End If
    ' ��������� ������� �������
    Select Case Pos
    Case 1:     If Not Overwrite Then sBeg = 1:           Result = Data & Delim & Result: GoTo HandleExit
    Case -1:    If Not Overwrite Then sBeg = Len(Result): Result = Result & Delim & Data: GoTo HandleExit
    End Select
'    If Pos < 1 then Goto HandleExit
'' ��� ������ Split
'    Dim arr() As String:arr = Split(Result, Delim)
'    If Overwrite Then arr(Pos - 1) = Data Else arr(Pos - 1) = Data & Delim & arr(Pos - 1)
'    Result = Join(arr, Delim): Erase arr()
'' ��� ������ InStr ���������� DelimStringGet
' ������� � ������ �������� ������ ��������� - ����� ������������ ������������� ������� (� ����� ������)
    If Overwrite Then
    ' ������� � �������
        ' �������� ������� �� ������, �������
        Call p_GetSubstrBounds(Result, Pos, sBeg, sEnd, Delim)
    Else
    ' ������� �� �������
        ' �������� ������� �� ������, �������, ��������� ����� �� ������� ������
        ' � ���������� ��� � ������������ ��������� ������ (Sgn(Pos) = -1) �.�. �� ������ �-��� (�������������� Pos)
        If (Sgn(Pos) = -1) = p_GetSubstrBounds(Result, Pos, sBeg, sEnd, Delim) Then
        ' ���� ����� �� ������� � ����������� �� ������    (bRes = False; bDir = False)
        ' ��� ��� ������ �� ������� � ����������� �� ����� (bRes = True;  bDir = True)
        ' ������� �����:  "PrevVal & Delim & NewVal"
            Data = Mid$(Result, sBeg, sEnd - sBeg) & Delim & Data: Pos = Pos + 1
        Else
        ' ��� ������ �� ������� � ����������� �� ������    (bRes = True;  bDir = False)
        ' ��� ���� ����� �� ������� � ����������� �� ����� (bRes = False; bDir = True)
        ' ������� �����: "NewVal & Delim & PrevVal"
            Data = Data & Delim & Mid$(Result, sBeg, sEnd - sBeg)
        End If
    End If
    Result = Left$(Result, sBeg - 1) & Data & Mid$(Result, sEnd)
HandleExit:  sEnd = sBeg + Len(Data): DelimStringSet = Result: Exit Function
HandleError: Result = Source: Err.Clear: Resume HandleExit
End Function
Public Function DelimStringShrink(ByVal Source As String, _
    Optional Delim As String = " " _
    ) As String
' ������� �� ������ � ������������� ������������� �������� �������� ������ ������ ���������
'-------------------------
' Source - �������� ������
' Delim - �����������
'-------------------------
Dim Arr() As String, Col As Collection
Dim Result As String

    On Error GoTo HandleError
    Result = Trim$(Source)
    If Len(Result) = 0 Then GoTo HandleExit
' ��������� ������
    'Call xSplit(Result, Arr, Delim)
    Arr = Split(Result, Delim)
    Set Col = New Collection
Dim i As Long, iMax As Long: i = LBound(Arr): iMax = UBound(Arr)
    On Error Resume Next
' ��������� � ���������
Dim Itm
    Do Until i > iMax
        Itm = Trim$(Arr(i))
        Col.Add Itm, c_idxPref & Itm: Err.Clear: i = i + 1
    Loop
    On Error GoTo HandleError
' �������� ������
    Result = vbNullString
    For Each Itm In Col
        Result = Result & Delim & Itm
    Next Itm
    Result = Mid$(Result, Len(Delim) + 1)
HandleExit:  Erase Arr: Set Col = Nothing
             DelimStringShrink = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function DelimStringSimile(ByVal String1 As String, ByVal String2 As String, _
    Optional Delim As String = " ", _
    Optional Compare As VbCompareMethod = vbTextCompare _
    ) As Boolean
' ���� � ������� � ������������� ����������� ��������, ���������� True ��� ���������� ������� ����������
'-------------------------
' String1, String2  - ������ � ������������� �������� ������� ����� ������������ ����� �����
' Delim             - �����������
' Compare           - ����� ���������
'-------------------------
Dim Result As Boolean ':Result = False
    On Error GoTo HandleError
    String1 = Trim$(String1): If Len(String1) = 0 Then Err.Raise vbObjectError + 512
    String2 = Trim$(String2): If Len(String2) = 0 Then Err.Raise vbObjectError + 512
    If Compare = vbTextCompare Then String1 = UCase(String1): String2 = UCase(String2)
Dim a1() As String: a1 = Split(String1, Delim) 'Call xSplit(String1, a1, Delim)
Dim a2() As String: a2 = Split(String2, Delim) 'Call xSplit(String2, a2, Delim)
Dim i As Long, j As Long: i = LBound(a1): j = LBound(a2)
    Do
        If i > UBound(a1) Then i = LBound(a1): j = j + 1: If j > UBound(a2) Then Exit Do
        Result = Trim$(a1(i)) = Trim$(a2(j)): i = i + 1
    Loop Until Result
HandleExit:  DelimStringSimile = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function TokenStringGet(Source As String, _
    ByVal Pos As Long, _
    Optional Delims As String = c_strSymbSpaces & c_strSymbPunct & c_strSymbParenth & c_strSymbMath, _
    Optional IncEmpty As Boolean = False, _
    Optional DelimsLeft As String, Optional DelimsRight As String, _
    Optional sBeg As Long, Optional sEnd As Long _
    ) As String
' ���������� �������� ������ �� �������������� ������������� � ��������� ��������
'-------------------------
' Source    - �������� ������
' Pos       - ������� ������������ ��������
' Delims    - ����� ������������ ��� ��������� �������� ������
' IncEmpty  = False - ������� ������ ��������� - ���������������� ����������� ����� ��������������� ��� ����
'           = True  - �������������� ������ ����� �������� ������ �������� ����� ����������������� �������������
' DelimsLeft/DelimsRight    - ���������� ����������� ������������� ����� � ������ �� ������������ ������
' sBeg,sEnd - ���������� ������� ������ � ��������� ����������� ��������� (������) � ��������
'-------------------------
Dim Result As String: Result = Source
    
    On Error GoTo HandleError
    If Len(Source) = 0 Then Result = vbNullString: GoTo HandleExit
Dim aData() As String, aPos() As Long
Dim aMin As Long: aMin = 1
Dim aMax As Long: aMax = Tokenize(Source, aData(), Delims, aPos(), IncEmpty)
    If Pos < 0 Then Pos = aMax + Pos + 1
    If Pos < aMin Then Pos = aMin Else If Pos > aMax Then Pos = aMax
    ' ���� ������������ �� ���������� ������ (�����)
    sEnd = aPos(Pos - 1): If Pos = 1 Then sBeg = 1 Else sBeg = aPos(Pos - 2) + Len(aData(Pos - 2))
    If sEnd > sBeg Then DelimsLeft = Mid$(Source, sBeg, sEnd - sBeg) Else DelimsLeft = vbNullString
    ' ���� ������������ ����� ���������� ������ (������)
    sBeg = aPos(Pos - 1) + Len(aData(Pos - 1)): If Pos = aMax Then sEnd = Len(Source) Else sEnd = aPos(Pos)
    If sEnd > sBeg Then DelimsRight = Mid$(Source, sBeg, sEnd - sBeg) Else DelimsRight = vbNullString
    ' �������� ������
    Result = aData(Pos - 1): sBeg = aPos(Pos - 1): sEnd = sBeg + Len(Result)
HandleExit:  TokenStringGet = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function TokenStringSet(Source As String, _
    ByVal Pos As Long, ByVal Data As String, _
    Optional Delims As String = c_strSymbSpaces & c_strSymbPunct & c_strSymbParenth & c_strSymbMath, _
    Optional IncEmpty As Boolean = False, _
    Optional Overwrite As Boolean = False, _
    Optional NewDelim As String = " ", _
    Optional sBeg As Long, Optional sEnd As Long _
    ) As String
' ��������� � ������ �� �������������� ������������� ������� � ������� � ��������� ��������
'-------------------------
' Source    - �������� ������
' Pos       - ������� ������� ��������
' Data      - ����������� ������
' Delims    - ����� ������������ ��� ��������� �������� ������
' IncEmpty  = False - ������� ������ ��������� - ���������������� ����������� ����� ��������������� ��� ����
'           = True  - �������������� ������ ����� �������� ������ �������� ����� ����������������� �������������
' Overwrite = False - ������� �� ������� (��� Pos>0 ������� ����� ��������� ��������, ��� Pos<0 - �����. �.�. Pos=1 - ������� �������, � Pos=-1 - � ����� ������)
'           = True  - ������� � �������  �������� ������ � ��������� �������
' NewDelim - ����������� ����������� (����������� ��� Overwrite=False) ����������� ����� Data ����� ��������� �������
' sBeg,sEnd - ���������� ������� ������ � ��������� ����������� ��������� (������) � ��������
'-------------------------
Dim Result As String: Result = Source
    
    On Error GoTo HandleError
    If Len(Source) = 0 Then Result = Data: Pos = 1: GoTo HandleExit
    Select Case Pos
    Case 1:     If Not Overwrite Then sBeg = 1:           Result = Data & NewDelim & Source: GoTo HandleExit
    Case -1:    If Not Overwrite Then sBeg = Len(Source): Result = Source & NewDelim & Data: GoTo HandleExit
    End Select
Dim aData() As String, aPos() As Long
Dim aMin As Long: aMin = 1
Dim aMax As Long: aMax = Tokenize(Source, aData(), Delims, aPos(), IncEmpty)
Dim sTemp As String: sTemp = Data
    If Pos < 0 Then Pos = aMax + Pos + 2
    If Pos < aMin Then Pos = aMin Else If Pos > aMax Then Pos = aMax
    sBeg = aPos(Pos - 1)
    If Overwrite Then
    ' ������ �������� ��������� ������ ��������. ��������� ����������� �����������
        sEnd = sBeg + Len(aData(Pos - 1))
    Else
    ' ������� ������ ����� ���������. �������� ����������� (NewDelim)
        sEnd = sBeg: sTemp = sTemp & NewDelim
    End If
    Result = Left$(Result, sBeg - 1) & sTemp & Mid$(Result, sEnd)
HandleExit:  sEnd = sBeg + Len(Data): TokenStringSet = Result: Exit Function
HandleError: Result = Source: Err.Clear: Resume HandleExit
End Function
Public Function TokenStringDel(Source As String, _
    ByVal Pos As Long, _
    Optional Delims As String = c_strSymbSpaces & c_strSymbPunct & c_strSymbParenth & c_strSymbMath, _
    Optional IncEmpty As Boolean = False, _
    Optional SubDelims As Boolean = True, Optional NewDelim As String = " ", _
    Optional sBeg As Long, Optional sEnd As Long _
    ) As String
' ������� ����� � ��������� ������� �� ������ �� �������������� �������������
'-------------------------
' Source    - �������� ������
' Pos       - ������� ��������� ���������
' Delims    - ����� ������������ ��� ��������� �������� ������
' IncEmpty  = False - ������� ������ ��������� - ���������������� ����������� ����� ��������������� ��� ����
'           = True  - �������������� ������ ����� �������� ������ �������� ����� ����������������� �������������
' SubDelims = False - ����������� ������ ���������� �������� �������� � �������� ������,
'           = True  - ����������� ���������� ����� �������� �������� ����� �������� �� EndDelims
' NewDelim - �������� ����������� (����������� ��� SubDelims = True) ����������� ����� Data ����� ��������� �������
' sBeg,sEnd - ���������� ������� ��������� ��������� (������)
'-------------------------
Dim Result As String: Result = Source
    
    On Error GoTo HandleError
    If Len(Source) = 0 Then Pos = 1: GoTo HandleExit
Dim aData() As String, aPos() As Long
Dim aMin As Long: aMin = 1
Dim aMax As Long: aMax = Tokenize(Source, aData(), Delims, aPos(), IncEmpty)
Dim sTemp As String ':  sTemp = vbNullString
    If Pos < 0 Then Pos = aMax + Pos + 1
    If Pos < aMin Then Pos = aMin Else If Pos > aMax Then Pos = aMax
    If SubDelims Then
    ' ����������� ����� � ������ ��������� �, ���� ������� � �������� ������,- ���������� �� �����
        Select Case Pos
        Case aMin: sBeg = 1: sEnd = aPos(Pos)
        Case aMax: sBeg = aPos(Pos - 2) + Len(aData(Pos - 2)): sEnd = Len(Result) + 1
        Case Else: sBeg = aPos(Pos - 2) + Len(aData(Pos - 2)): sEnd = aPos(Pos): sTemp = NewDelim
        End Select
    Else
    ' ��������� ����������� �����������
        sBeg = aPos(Pos - 1)
        sEnd = sBeg + Len(aData(Pos - 1))
    End If
    Result = Left$(Result, sBeg - 1) & sTemp & Mid$(Result, sEnd)
HandleExit:  sEnd = sBeg: TokenStringDel = Result: Exit Function
HandleError: Result = Source: Err.Clear: Resume HandleExit
End Function
Public Function TaggedStringGet(Source As String, _
    ByRef Tag As String, _
    Optional Delim As String = ";", Optional TagDelim As String = "=", _
    Optional Compare As VbCompareMethod = vbTextCompare, _
    Optional sBeg As Long, Optional sEnd As Long _
    ) As String
' ���������� �������� �������� ������ � ������������� � ��������� ������
'-------------------------
' ����� ��� ���� ����� ������������ ����� ����� TaggedValues �� �����: https://www.sql.ru/forum/661816/vdrug-u-kogo-est-dlya-obrabotki-v-vba-strok-svoystv
' �� ������ � ����� ������� ������ �������
' ������� ����� �� ��������� ������������ ����, - ����� ���������� ������ ���������
' Source    - ������ ��������� ���� "Tag1=Val1;...TagN=ValN"
' Tag       - ��� (Tag) ��������. ���� Tag �� ����� - ������� ����� ������� �� Pos
' Delim     - ����������� ��� ��� (Tag) / �������� (Val)
' TagDelim  - ����������� ����� (Tag) � �������� (Val) � ����
' sBeg,sEnd - ���������� ������� ������ � ��������� ����������� ��������� (�������� ����) � ��������
' Compare   - ��� ��������� (vbBinaryCompare/vbTextCompare)
' ���������� �������� (Val) �������� � ��������� ������ (Tag)
'-------------------------
'' ! ���� ����� ������ �� ������� ��������� - ���� ����������������� ����� ������
'    � �������� Optional ByRef Pos As Long = 0, _
'' Pos -     �� ����� ������� ������������� ��������. (���� ������ Tag �� ������������)
''           >0 - ������� ������������ ������ ������
''           <0 - ������� ������������ ����� ������
''           �� ������ ������� ����������� �������� ������������ ������ ������
'-------------------------

'RegExp: Mask = Delim & Tag & TagDelim & "(.+?)" & Delim
'        With Regex.Execute(Delim & Source & Delim)(0)
'           Val = .SubMatches(0)
'           Pos = .Length + Len(Tag & TagDelim) - Len(Delim)
'        End With
Dim Result As String: Result = vbNullString
    On Error GoTo HandleError
    If Len(Source) = 0 Then GoTo HandleExit
    sBeg = 1
    If Len(Tag) > 0 Then
' ���� �� Tag
Dim pLen As Long: pLen = Len(Tag & TagDelim): If Len(Source) < pLen Then GoTo HandleExit 'Pos = 0: GoTo HandleExit
        sBeg = sBeg + pLen
'        If StrComp(Mid$(Source, 1, pLen), Tag & TagDelim, Compare) = 0 ' If Mid$(tmpSrc, 1, pLen) = tmpTag & TagDelim Then
'            Pos = 1
'        Else
        If StrComp(Mid$(Source, 1, pLen), Tag & TagDelim, Compare) <> 0 Then ' If Mid$(tmpSrc, 1, pLen) <> tmpTag & TagDelim Then
'        ' ���� ��� � �������� ������ �� � ������ ������
'        ' ���� ��������� � �������� ������ ������������ � Delim � ��������������� TagDelim
            sBeg = InStr(1, Source, Delim & Tag & TagDelim, Compare)
            If sBeg = 0 Then GoTo HandleExit 'Pos = 0: GoTo HandleExit
            sBeg = sBeg + Len(Delim) + pLen
'            Pos = InStrCount(Left$(Source, sBeg), Delim) + 1
        End If
'        ' ������ � sBeg ������� ������ �������� ����
'        ' ���� ������� ����� �������� ���� � ������� ������ �� ���������� Delim
        sEnd = InStr(sBeg, Source, Delim, Compare): If sEnd = 0 Then sEnd = Len(Source) + 1
        Result = Mid$(Source, sBeg, sEnd - sBeg)
'    Else
'' ���� �� Pos
'        Call p_GetSubstrBounds(Source, Pos, sBeg, sEnd, Delim)
'        Result = Mid$(Source, sBeg, sEnd - sBeg)
'        sEnd = InStr(Result, TagDelim): sBeg = sEnd + Len(TagDelim)
'        Tag = Left$(Result, sEnd - 1): Result = Mid$(Result, sBeg)
    End If
HandleExit:  TaggedStringGet = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function TaggedStringSet(Source As String, _
    ByRef Tag As String, Optional ByRef Data As String, _
    Optional Delim As String = ";", Optional TagDelim As String = "=", _
    Optional sBeg As Long, Optional sEnd As Long, _
    Optional Compare As VbCompareMethod = vbTextCompare _
    ) As String
' ������������� �������� �������� ������ � ������������� � ��������� ������
'-------------------------
' Source -  ������ ��������� ���� "Tag1=Val1;...TagN=ValN"
' Tag -     ��� (Tag) ���������������� ��������, ���� ������� ����������� - ����� ��������
' Data -    �������� (Val) ���������������� ��������
' Delim -   ����������� ��� ��� (Tag) / �������� (Val)
' TagDelim - ����������� ����� (Tag) � �������� (Val) � ����
' sBeg,sEnd - ���������� ������� ������ � ��������� ����������� ��������� (�������� ����) � ��������
' Compare - ��� ��������� (vbBinaryCompare/vbTextCompare)
' ���������� ������ ��������� � ������ ������������ ��������
'-------------------------
'' ! ���� ����� ������ �� ������� ��������� - ���� ����������������� ����� ������
'    � �������� Optional ByRef Pos As Long = 0, _
'' Pos -     �� ����� ������� ������� ��������. (���������� ����� ������� � �������������� ������)
''           0  - ����� �������� � ������� ���������� �������� ��� ����� ��� ��� ����������
''           >0 - ������� ������������ ������ ������, ������� ����� ��������� ��������
''           <0 - ������� ������������ ����� ������, ������� ����� ��������� �������
''           �� ������ �������� ������� ������������ �������� ������������ ������ ������
'-------------------------
' ! ������� ����� �� ��������� ������������ ����, - ����� �������� ������ ���������
'-------------------------
' v.1.0.3       : 11.02.2022 - ���������� ������ ����������� ���� TagDelim ��������� ���������, ��������� �� ��������
'-------------------------
Dim Result As String: Result = Source
    On Error GoTo HandleError
' ������ ��� ����
    If Len(Tag) = 0 Then
        Result = Source
'' ���� �� Pos � �������� ��� Tag
'   ' ��� ��������� �� ������� - �� ���� �������� �� ������� ������, - ������ ����� ��������� ���������
'        Call p_GetSubstrBounds(Result, Pos, sBeg, sEnd, Delim)
'        Tag = Split(Mid$(Result, sBeg, sEnd - sBeg), TagDelim)(0)
'        sTemp = Tag & TagDelim & Data   ' "Tag=Val" - ��� ������ ��� �������
        GoTo HandleExit
    End If
' ������ ��������
    If Len(Data) = 0 Then Result = TaggedStringDel(Source, Tag): GoTo HandleExit
' ������ ������
    If Len(Source) = 0 Then If Len(Tag) > 0 Then Result = Tag & TagDelim & Data: GoTo HandleExit ': Pos = 1
Dim sTemp As String
Dim pLen As Long
'' ���� �� Tag
'    ' ����� ������� ���� � ����� ������
    sBeg = 0: sEnd = 0 'Len(Result)
    pLen = 1 'Len(Tag) + Len(TagDelim)
    sTemp = Tag & TagDelim & Data ' "Tag=Val"
    If StrComp(Left$(Source, Len(Tag) + Len(TagDelim)), Tag & TagDelim, Compare) = 0 Then  ' Left$(Source, Len(Tag) + Len(TagDelim)) = Tag & TagDelim Then
    ' ��� ���� ������� � ������ ������ ("Tag=...")
        sBeg = 1
    Else
    ' ���� � �������� ������ ��� ���� � ������������ ���/�������� ("...;Tag=...")
        sBeg = InStr(pLen + 1, Result, Delim & Tag & TagDelim, Compare)
        If sBeg > 0 Then
    ' ��� ���� � ������������ ���/�������� ������� � �������� ������ ("...;Tag=...")
            pLen = pLen + Len(Delim)
        Else
    ' ���� �� ������� - ��������� ��� ��� �������� � ����������� ���/��������
            pLen = Len(Tag) + Len(Delim)
            If StrComp(Left$(Source, pLen), Tag & Delim, Compare) = 0 Then 'Left$(Source, pLen) = Tag & Delim Then
        ' ��������� � ������ ������ ("Tag;...")
                sBeg = 1: sEnd = Len(Tag) + 1
            ElseIf StrComp(Source, Tag, Compare) = 0 Then  'tmpSource = tmpTag Then
        ' ��������� ��� ������ ("Tag")
                sBeg = 1: sEnd = Len(Result) + 1
            ElseIf StrComp(Right$(Source, pLen), Delim & Tag, Compare) = 0 Then  'Right$(tmpSource, pLen) = tmpDelim & tmpTag Then
        ' ��������� � ����� ������ ("...;Tag")
                sBeg = Len(Result) - Len(Tag): sEnd = Len(Result) + 1
            Else
        ' ��������� � �������� ������ ("...;Tag;...")
                sBeg = InStr(1, Result, Delim & Tag & Delim, Compare): sEnd = sBeg + pLen
            End If
        End If
    End If
' ���� ������ ������� � ����� �� ��������� ���� �������� �����������
    If sBeg > 0 And sEnd = 0 Then sEnd = InStr(sBeg + pLen + 1, Result, Delim, Compare): If sEnd = 0 Then sEnd = Len(Result) + 1
'        bFound = sBeg > 0

'        If Pos = 0 Then
'    ' ������� � ������� ���������� �������� ��� � ����� ������
'        ' ������������ ������� ���������
'            ' ������ �� � ������ ������� ����� �����������
'            ' �� ������ - ������� � �����, ��������� ����� �������� �����������
        If sBeg = 0 Then sBeg = Len(Result) + 1: sEnd = sBeg
'        ' �������� ������� ��������� �������
        If sBeg > 1 Then
            sTemp = Delim & sTemp
'                Pos = InStrCount(Left$(Result, sBeg), Delim) + 2
'            Else
'                Pos = 1
        End If
'        Else
'    ' �������� ��������� � ������� ������� � ������� � ��������� �������
'        ' �������� � ��������� �������
'            If sBeg > 0 Then
'                If sBeg = 1 Then sEnd = sEnd + Len(Delim)
'                Result = Left$(Source, sBeg - 1) & Mid$(Source, sEnd)
'            End If
'        ' �������� ������� ��������� ������� � ��������� ����������� ������
'            'Result = DelimStringSet(Result, Pos, sTemp, Delim, Overwrite:=False): GoTo HandleExit
'            Select Case Pos
'            Case 1:     Result = sTemp & Delim & Result: GoTo HandleExit
'            Case -1:    Result = Result & Delim & sTemp: Pos = InStrCount(Result, Delim) + 2: GoTo HandleExit
'            End Select
'            If (Sgn(Pos) = -1) = p_GetSubstrBounds(Source, Pos, sBeg, sEnd, Delim) Then
'                sTemp = Mid$(Result, sBeg, sEnd - sBeg) & Delim & sTemp: If Not bFound Then Pos = Pos + 1
'            Else
'                sTemp = sTemp & Delim & Mid$(Result, sBeg, sEnd - sBeg)
'            End If
'        End If

' ������� � ��������� �������
    Result = Left$(Result, sBeg - 1) & sTemp & Mid$(Result, sEnd)
HandleExit:  sBeg = sEnd - Len(Data): TaggedStringSet = Result: Exit Function
HandleError: Result = Source: Err.Clear: Resume HandleExit
End Function
Public Function TaggedStringDel(Source As String, _
    ByRef Tag As String, _
    Optional Delim As String = ";", Optional TagDelim As String = "=", _
    Optional sBeg As Long, Optional sEnd As Long, _
    Optional Compare As VbCompareMethod = vbTextCompare _
    ) As String
' ������� ������� ������ � ������������� � ��������� ������
'-------------------------
' Source - ������ ��������� ���� "Tag1=Val1;...TagN=ValN"
' Tag -     ��� (Tag) ��������. ���� Tag �� ����� - ������� ����� ������� �� Pos
' Delim -   ����������� ��� ��� (Tag) / �������� (Val)
' TagDelim - ����������� ����� (Tag) � �������� (Val) � ����
' sBeg,sEnd - ���������� ������� ������ � ��������� ����������� ��������� � ��������
' Compare - ��� ��������� (vbBinaryCompare/vbTextCompare)
' ���������� ������ ��������� ��� ���������� ��������
'-------------------------
'' ! ���� ����� ������ �� ������� ��������� - ���� ����������������� ����� ������
'    � �������� Optional ByRef Pos As Long = 0, _
'' Pos -     �� ����� ������� ���������� ��������. (���� ������ Tag �� ������������)
''           >0 - ������� ������������ ������ ������
''           <0 - ������� ������������ ����� ������
''           �� ������ ������� �������� ��������������� ��������� ������������ ������ ������ ��� 1
'-------------------------
' ! ������� ����� �� ��������� ������������ ����, - ����� ������� ������ ���������
'-------------------------
Dim Result As String: Result = Source
    On Error GoTo HandleError
    If Len(Source) = 0 Then GoTo HandleExit
    If Len(Tag) > 0 Then
'' ���� �� Tag
'    ' ����� ������� ���� � ����� ������
'' ���� �� Tag
'    ' ����� ������� ���� � ����� ������
        sBeg = 0: sEnd = 0 'Len(Result)
Dim pLen As Long
        pLen = 1 'Len(Tag) + Len(TagDelim)
        'sTemp = Tag & TagDelim & Data   ' "Tag=Val"
        If StrComp(Left$(Source, Len(Tag) + Len(TagDelim)), Tag & TagDelim, Compare) = 0 Then ' Left$(tmpSource, Len(Tag) + Len(TagDelim)) = tmpTag & tmpTagDelim Then
        ' ��� ���� ������� � ������ ������ ("Tag=...")
            sBeg = 1
        Else
        ' ���� � �������� ������ ��� ���� � ������������ ���/�������� ("...;Tag=...")
            sBeg = InStr(pLen + 1, Result, Delim & Tag & TagDelim, Compare)
            If sBeg > 0 Then
        ' ��� ���� � ������������ ���/�������� ������� � �������� ������ ("...;Tag=...")
                pLen = pLen + Len(Delim)
            Else
        ' ���� �� ������� - ��������� ��� ��� �������� � ����������� ���/��������
                pLen = Len(Tag) + Len(Delim)
                If StrComp(Left$(Source, pLen), Tag & Delim, Compare) = 0 Then  ' Left$(tmpSource, pLen) = tmpTag & tmpDelim Then
            ' ��������� � ������ ������ ("Tag;...")
                    sBeg = 1: sEnd = Len(Tag) + 1
                ElseIf StrComp(Source, Tag, Compare) = 0 Then   ' tmpSource = tmpTag Then
            ' ��������� ��� ������ ("Tag")
                    sBeg = 1: sEnd = Len(Result) + 1
                ElseIf StrComp(Right$(Source, pLen) = Delim & Tag, Compare) = 0 Then   'Right$(tmpSource, pLen) = tmpDelim & tmpTag Then
            ' ��������� � ����� ������ ("...;Tag")
                    sBeg = Len(Result) - Len(Tag): sEnd = Len(Result) + 1
                Else
            ' ��������� � �������� ������ ("...;Tag;...")
                    sBeg = InStr(1, Result, Delim & Tag & Delim, Compare): sEnd = sBeg + pLen
                End If
            End If
        End If
        If sBeg = 0 Then GoTo HandleExit ' ��� �� ������
'        ' ���� ������ ��������� ������� ���� ��� �����
        If sEnd = 0 Then sEnd = InStr(sBeg + pLen + 1, Result, Delim, Compare): If sEnd = 0 Then sEnd = Len(Result) + 1
'        ' �������� ������� ��������� ��������
        If sBeg <= 1 Then
            sEnd = sEnd + Len(Delim) ': Pos = 1
'        Else
'            Pos = InStrCount(Left$(Result, sBeg), Delim) + 1
        End If
'    Else
'' ���� �� Pos
'    ' ��� ��������� �� ������� - �� ���� �������� �� ������� ������, - ������ ����� ��������� ���������
'        Call p_GetSubstrBounds(Result, Pos, sBeg, sEnd, Delim)
'        'Tag = Split(Mid$(Result, sBeg, sEnd - sBeg), TagDelim)(0)
'        If sBeg > 1 Then sBeg = sBeg - Len(Delim) Else sEnd = sEnd + Len(Delim)
'        If Pos > 1 Then Pos = Pos - 1
    End If
' �������� ��������� �������
    Result = Left$(Result, sBeg - 1) & Mid$(Result, sEnd)
HandleExit:  sEnd = sBeg: TaggedStringDel = Result: Exit Function
HandleError: Result = Source: Err.Clear: Resume HandleExit
End Function
Public Function TaggedString2Collection(Source As String, _
    Optional Tags As Collection, Optional Keys, _
    Optional Delim As String = ";", Optional TagDelim As String = "=", _
    Optional ReplaceExisting As Integer = True _
    ) As Boolean
' ����������� ������ ����������� ���������� � ��������� �������� � ������� ����� ����� ����
'-------------------------
' Source - ������ ��������� ���� "Tag1=Val1;...TagN=ValN"
' Tags - ������������ ��������� �������� �����. ���� �� ����� �������� �������� ��������� ����� �������� ����� ��������� � ���
' Keys - (���� ������) ���������� ������ ������ ��������� (Tag)
' Delim -   ����������� ��� ��� (Tag) / �������� (Val)
' TagDelim - ����������� ����� (Tag) � �������� (Val) � ����
' ReplaceExisting - ���������� ��������� ��� ����������� ���������� � ���������� ������
'   0 - � ��������� ����� ��������� ������ ���������, ����������� ����� ��������������
'  -1 - ���������� � ���������� ������ ����� ���������� - � ��������� ��������� ��������� ���������
'-------------------------
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    Do While Right(Source, Len(Delim)) = Delim: Source = Trim(Left$(Source, Len(Source) - Len(Delim))): Loop
    If Len(Source) = 0 Then GoTo HandleExit
    If Tags Is Nothing Then Set Tags = New Collection
Dim Arr() As String, Term 'As String
Dim sKey As String, vVal 'As String
    Arr() = Split(Source, Delim)
    For Each Term In Arr
        sKey = Split(Term, TagDelim)(0) ' �������� ���
        If p_IsExist(sKey, Tags) Then If ReplaceExisting Then Tags.Remove sKey Else GoTo HandleNext
        vVal = Split(Term, TagDelim)(1) ' �������� ��������
        Tags.Add vVal, sKey
HandleNext:
    Next
    Erase Arr
    If Not IsMissing(Keys) Then Keys = p_GetCollKeys(Tags)
    Result = True
HandleExit:  TaggedString2Collection = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function TaggedCollection2String(Tags As Collection, _
    Optional Delim As String = ";", Optional TagDelim As String = "=" _
    ) As String
' ����������� ��������� �������� � ������� ����� ����� ���� � ������ ����������� ����������
'-------------------------
Dim Result As String: Result = vbNullString
    On Error GoTo HandleError
    If Tags Is Nothing Then GoTo HandleExit
    If Tags.Count = 0 Then GoTo HandleExit
Dim Keys() As String: Keys() = p_GetCollKeys(Tags)
Dim i As Long
    For i = 1 To Tags.Count
        Result = Result & Delim & Keys(i) & TagDelim & Tags.Item(i)
    Next i
    If Left$(Result, Len(Delim)) = Delim Then Result = Mid$(Result, Len(Delim) + 1)
HandleExit:  TaggedCollection2String = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Private Function p_GetSubstrBounds(ByRef Source As String, _
    ByRef Pos As Long, ByRef sBeg As Long, ByRef sEnd As Long, _
    Optional Delim As String = " ") As Boolean
' ���������� ����� � ������� ������ � ����� ��������� � ������ � �������������
'-------------------------
' Source -  �������� ������
' Pos -     �� ����� ������� �������� ��������.
'           >0 - ������� ������������ ������ ������
'           <0 - ������� ������������ ����� ������
'           �� ������ ������� �������� ������������ ������ ������
' sBeg, sEnd - ���������� ������� ��������� � ������
' Delim -   �����������
' ���������� True ���� �������� ������� � �������� ������, ����� False
'-------------------------
Dim Result As Boolean
    On Error GoTo HandleError
Dim i As Long
    i = 1: sBeg = 1
    If Pos >= 0 Then
' ������� �� ������
    ' ��������� ��� ������ � ������ �� ������������, �������� ����� ���������
        If Pos = 0 Then Pos = 1
        Do
            sEnd = InStr(sBeg, Source, Delim)
            If sEnd = 0 Then sEnd = Len(Source) + 1:  Exit Do
            i = i + 1: If i > Pos Then Exit Do
            sBeg = sEnd + Len(Delim)
        Loop
        Result = (Pos <= i): If Not Result Then Pos = i  ' ������� ���� ������� �������
    Else
' ������� �� �����
    ' ������� 1: ��������� ��� ������ � ����� �� ������������, �������� ����� ���������
    ' ������� 2: ��������� ��� ������ � ������ ����������� ���������� �������� � �������� ������ ������� ������������
    Dim aPos() As Long
    ' ������������ ���������� ���������� � ������, �������� ������ ������� ������������
        Do
            ReDim Preserve aPos(1 To i): aPos(i) = sBeg
            sBeg = InStr(sBeg, Source, Delim) '
            i = i + 1
            If sBeg > 0 Then sBeg = sBeg + Len(Delim) Else Exit Do
        Loop
    ' ��������� ������� ������������ ����� ������ � ������� �� ������
        Pos = i + Pos
        Select Case Pos
        Case 1 To i: Result = True              ' ������� � �������� ������
        Case Is < 1: Result = False: Pos = 1    ' ������� ���� ������ �������
        'Case Is > i: Result = False: Pos = i    ' ������� ���� ������� �������
        End Select
    ' ����� ������� ��������� � ��������� �������� �� �������
        sBeg = aPos(Pos): If Pos < (i - 1) Then sEnd = aPos(Pos + 1) - Len(Delim) Else sEnd = Len(Source) + 1
        Erase aPos()
    End If
HandleExit:  p_GetSubstrBounds = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
' ==================
' ������� ��� ���������/������������� ��������� ������ � �������������
' ==================
#If APPTYPE = 0 Then ' ������ ��� Access
Public Sub TextToArrayByControl(TextString As String, _
    vControls As Variant, _
    Optional Separators As String = " �.,;:!?()[]{}�+-*/\|" & vbTab & vbCrLf)
' ��������� ����� � ������ ���������� ������������ �� ������, ��������������� ������ �����, � ������������ � �� �����
'-------------------------
' TextString - ������ ������ ������� ���������� �������
' vControls  - ��������� ��� ������ ����� � ������� ���� ������������ �����
' Separators - ������ ������������ �� ������� ����� ���� ����� ���� ������� ������ ������� �������� - � �������� ������ ����� ������
'-------------------------
Dim hFont As LongPtr, hOldFont As LongPtr
Dim WidthInPix As Long
Dim sz As Size, tm As TEXTMETRIC
Dim strText As String, strRest As String, strTemp As String
Dim aWords() As String
Dim spLen As Long, spPos As Integer, spPosNext As Integer
Dim tWidth As Long, tHeight As Long
Dim i As Long, iMax As Long, w As Long
Dim aCtl As Variant, ctl  As Variant

Dim strMessage As String, strComment As String
Dim Result As Boolean

    Result = False
    On Error GoTo HandleError
    If Len(TextString) = 0 Then GoTo HandleExit
    If IsArray(vControls) Then
        aCtl = vControls
    ElseIf TypeOf vControls Is Collection Then
        Set aCtl = vControls
    ElseIf TypeOf vControls Is Access.Control Then
        aCtl = Array(vControls)
    Else
        Err.Raise vbObjectError + 512
    End If
Dim tDC As LongPtr:           tDC = GetDC(0)
Dim PIXEL_PER_INCH_X As Long: PIXEL_PER_INCH_X = GetDeviceCaps(tDC, LOGPIXELSX)
' ��������� ������
    Call Tokenize(TextString, aWords, Separators)
    i = LBound(aWords): iMax = UBound(aWords)
    spLen = 1
    'strRest = Text
    ' �������: vbCrLf ������ �� vbCr ����� ������ ������� ������ ������
    strRest = Replace(TextString, vbCrLf, vbCr)
    For Each ctl In aCtl
        w = 0
        strText = vbNullString
        Do
        ' ���������� ����� ������
            If i < iMax Then
                strTemp = aWords(i)
                spPos = Len(strTemp) + 1
                spPosNext = InStr(spPos, strRest, aWords(i + 1))
                spLen = spPosNext - spPos
            Else
                strTemp = strRest
                spPos = Len(strTemp) + 1
                spPosNext = spPos
                spLen = 0
            End If
            ' ������� ������ �������� Chr(&HAD)
            strTemp = Replace(strText & strTemp, Chr(&HAD), vbNullString)
            ' ������ ����� ���������� ������ + ������� �������� + ������� �����������
            strTemp = strTemp & Mid$(strRest, spPos, spLen)
            spLen = Len(Trim$(strTemp)): If spLen = 0 Then spLen = 1
        ' �������� ������ ������
            hFont = p_HFontByControl(ctl)
            hOldFont = SelectObject(tDC, hFont)
            GetTextExtentPoint32 tDC, strTemp, spLen, sz
            SelectObject tDC, hOldFont
            DeleteObject hFont
        ' �������: w=0, - ��� ���� �������
            WidthInPix = ctl.Width * (PIXEL_PER_INCH_X / TwipsPerInch)
            If sz.cX <= WidthInPix Or w = 0 Then
            ' ���� ������ ����� � ������ ������ ������� ������ - �� ����� ����,
            ' ����� �������� � ������� �����
                If sz.cX > tWidth Then tWidth = sz.cX
                strRest = Mid$(strRest, spPosNext)
                strText = strTemp
                i = i + 1
                w = w + 1
            End If
        ' ���������� ������ ������ � �������� ��������
        Loop Until (i > iMax) Or (WidthInPix < sz.cX) '(WidthInPix < (sz.cx * (1 + spLen) / spLen))
        ctl.Value = strText
        'tHeight = tHeight + sz.cy
    Next ctl
' �������� �������� ������ � �� ������ � ��������
'    WidthInPix = tWidth: HeightInPix = tHeight
'    GetTextMetrics tDC, tm
'    Overhang = tm.tmOverhang ' ������� ��� ��������� � ������� �������
HandleExit:  SelectObject tDC, hOldFont
             DeleteObject hFont: ReleaseDC 0, tDC
             Exit Sub
HandleError: Err.Clear: Resume HandleExit
End Sub
Public Function TextToArrayByWidth(TextString As String, WidthInPix As Long, Optional HeightInPix, _
    Optional Separators As String = " �.,;:!?()[]{}�+-*/\|" & vbTab & vbCrLf, _
    Optional OutLines, Optional Overhang As Long, Optional OutDelimiter = vbCrLf, _
    Optional hFont As LongPtr, Optional hdc As LongPtr = 0) As String
' , Optional OutLineWidth, Optional OutLineHeight
' ��������� ����� � ������ ���������� ������������ �� ������, ��������������� �������� �������� ������� � ���������� ������
'-------------------------
' TextString - ������ ������ ������� ���������� �������
' WidthInPix - �� ����� - ������������ ������ ��������� ������,
'              �� ������ - �������� ������ ��������� ������
' HeightInPix - �� ������ - �������� ������ ��������� ������
' Separators - ������ ������������ �� ������� ����� ���� �����.
'       ���� ������� ������ ������� �������� - � �������� ������ ����� ������
' OutLines - ������ ����� ��������� ������
' Overhang - �������� ��� ������������� ������� ��� ���������, ������ � ��. �������
' OutDelimiter - ����������� ����� � �������� ������
' hFont - hFont ������ ��� �������� ������������ ���������
' hDC - hDC -������� ���� ����� ���������� �����
'' OutLineWidth, OutLineHeight - ������� �������� ����� ��������� ������
'-------------------------
Dim Result As String
Dim sz As Size, tm As TEXTMETRIC
Dim strText As String, strRest As String, strTemp As String
Dim aWords() As String, aText() As String ', aWidth() As Long, aHeight() As Long
Dim spLen As Long, spPos As Integer, spPosNext As Integer
Dim tWidth As Long, tHeight As Long
Dim i As Long, iMax As Long, ii As Long, w As Long

    On Error GoTo HandleError
    Result = vbNullString
    tWidth = 0: tHeight = 0
    If hFont = 0 Then hFont = p_HFontByControl() 'GoTo HandleExit
    If WidthInPix < 0 Then GoTo HandleExit
Dim WidthMax As Long: WidthMax = WidthInPix - 2 ' �������� )
Dim tDC As LongPtr, hOldFont As LongPtr
    If hdc = 0 Then tDC = GetDC(0) Else tDC = hdc

Dim PIXEL_PER_INCH_X As Long: PIXEL_PER_INCH_X = GetDeviceCaps(tDC, LOGPIXELSX)
    hOldFont = SelectObject(tDC, hFont) 'select font into the DC
    
    Call Tokenize(TextString, aWords, Separators)
    i = LBound(aWords): iMax = UBound(aWords)
    ii = 0: spLen = 1
    'strRest = Text
    ' �������: vbCrLf ������ �� vbCr ����� ������ ������� ������ ������
    strRest = Replace$(TextString, vbCrLf, vbCr)
    Do
        w = 0
        strText = vbNullString
        Do
        ' ���������� ����� ������
            If i < iMax Then
                strTemp = aWords(i)
                spPos = Len(strTemp) + 1
                spPosNext = InStr(spPos, strRest, aWords(i + 1))
                spLen = spPosNext - spPos
            Else
                strTemp = strRest
                spPos = Len(strTemp) + 1
                spPosNext = spPos
                spLen = 0
            End If
            ' ������� ������ �������� Chr(&HAD)
            strTemp = Replace$(strText & strTemp, Chr(&HAD), vbNullString)
            ' ������ ����� ���������� ������ + ������� �������� + ������� �����������
            strTemp = strTemp & Mid$(strRest, spPos, spLen)
            spLen = Len(Trim$(strTemp)): If spLen = 0 Then spLen = 1
        ' �������� ������ ������
            GetTextExtentPoint32 tDC, strTemp, spLen, sz
        ' �������: w=0, - ��� ���� �������
            If sz.cX <= WidthMax Or w = 0 Then
            ' ���� ������ ����� � ������ ������ ������� ������ - �� ����� ����,
            ' ����� �������� � ������� �����
                If sz.cX > tWidth Then tWidth = sz.cX
                strRest = Mid$(strRest, spPosNext)
                strText = strTemp
                i = i + 1
                w = w + 1
            End If
        Loop Until (i > iMax) Or (WidthMax < sz.cX) '(WidthMax < (sz.cx * ((1 + spLen) / spLen))
        ReDim Preserve aText(ii): aText(ii) = Trim$(strText)
'        ReDim Preserve aWidth(ii): aWidth(ii) = sz.CX
'        ReDim Preserve aHeight(ii): aHeight(ii) = sz.CY
        tHeight = tHeight + sz.cY
'        Result = Result & OutDelimiter & strText
    ' ���� �������� ����� - �������
        If Len(strRest) = 0 Then Exit Do
        ii = ii + 1
    Loop
' �������� �������� ������ � �� ������ � ��������
    WidthInPix = tWidth: HeightInPix = tHeight
    Result = Join(aText, OutDelimiter)
    OutLines = aText:           Erase aText
'    OutLineWidth = aWidth:      Erase aWidth
'    OutLineHeight = aHeight:    Erase aHeight
    GetTextMetrics tDC, tm
    Overhang = tm.tmOverhang ' ������� ��� ��������� � ������� �������
    
HandleExit:  SelectObject tDC, hOldFont: If hdc = 0 Then ReleaseDC 0, tDC
             TextToArrayByWidth = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
#End If
' ==================
' ������� ��������/�������������� �����/��������
' ==================
Public Function TextTranslit(ByVal Source As String, Optional Direction As Byte = 0) As String
' �������������� ��������� ��� �������� �������� �� ���� � 52535.1-2006
'-------------------------
' Direction = 0 - �������������� (��� > ���)
'             1 - �������� �������������� (��� > ���)
'-------------------------
Dim Result As String: Result = vbNullString
    On Error GoTo HandleError
Dim i As Integer, j As Integer, c As String: i = 1
Const cSymbRus = "��������������������������������"
Dim TransLat(): TransLat = Array("shch", "zh", "kh", "tc", "ch", "sh", "iu", "ia", "a", "e", "e", "i", "o", "u", "y", "e", "i", "b", "v", "g", "d", "z", "k", "l", "m", "n", "p", "r", "s", "t", "f", "", "")
Dim cLen As Integer: cLen = 1
    If Direction = 0 Then   ' ��� >> ���
        Do Until i > Len(Source)
            c = Mid$(Source, i, cLen)           ' ������� ��� ������
            j = InStr(1, cSymbRus, LCase$(c))    ' ����� �������� �������
            If j > 0 Then j = j - 1: If c = LCase$(c) Then c = TransLat(j) Else c = StrConv(TransLat(j), vbProperCase)
            Result = Result & c
            i = i + cLen
        Loop
    Else                    ' ��� >> ���
Const cSymbLat = "chjqwx" ''"
Dim TransRus(): TransRus = Array("�", "�", "��", "�", "�", "��") ', "�")
        Do Until i > Len(Source)
        ' ��������� �� ������� �������������� �� ����
            For j = 0 To UBound(TransLat)       ' ����� �������� �������
                c = TransLat(j): cLen = Len(c): If cLen = 0 Then Exit For
                If LCase(Mid$(Source, i, cLen)) = c Then c = Mid$(cSymbRus, j + 1, 1): Exit For
            Next j
            If Len(c) = 0 Then
        ' ���� �� ������� ����� ��� ������� ���� ��������� �� �������� �������
                cLen = 1: c = Mid$(Source, i, cLen) ' ������� ��� ������
                j = InStr(1, cSymbLat, LCase$(c))    ' ����� �������� �������
                If j > 0 Then c = TransRus(j - 1)   ' ���� ������ - ����
            End If
            If UCase$(Mid$(Source, i, 1)) = Mid$(Source, i, 1) Then c = StrConv(c, vbProperCase)
            Result = Result & c
            i = i + cLen
        Loop
    End If
HandleExit:  TextTranslit = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function TextCorrectLatRus(Source As String, Optional Direction As Byte = 0) As String
' ���������� �������� Lat<=>Rus
'-------------------------
' Direction = 0 - �������� ��������� ������ ��������������� �������
' Direction = 1 - �������� ������� ������ ��������������� ���������
'-------------------------
Dim Result As String
    Result = Source
    If Len(Source) = 0 Then GoTo HandleExit
Const cSymbLat As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const cSymbRus As String = "��������������������������"
Dim strSearch As String, strReplace As String
    If Direction = 0 Then   'Lat=>Rus
        strSearch = cSymbLat: strReplace = cSymbRus
    Else                    'Rus=>Lat
        strSearch = cSymbRus: strReplace = cSymbLat
    End If
Dim Char As String * 1, bUpper As Boolean
Dim c As Long, cMax As Long, i As Byte: c = 1: cMax = Len(Source)
    Do Until c > cMax
        Char = VBA.Mid$(Result, c, 1): bUpper = Char = VBA.UCase$(Char)
        i = InStr(1, strSearch, VBA.UCase$(Char)): If i = 0 Then GoTo HandleNext
        Char = VBA.Mid$(strReplace, i, 1): If bUpper Then Else Char = VBA.LCase$(Char)
        Mid$(Result, c, 1) = Char
HandleNext: c = c + 1
    Loop
HandleExit:  TextCorrectLatRus = Result
HandleError: Result = Source: Err.Clear: Resume HandleExit
End Function
Public Function TextContainsAlpha(Source As String) As Boolean
' ��������� ������� � ������ �������� ���������� � ������ ����������
'-------------------------
Dim c As Long, cMax As Long
Dim Char As String * 1
Dim Result As Boolean
    Result = False
    On Error GoTo HandleError
' ������ ����������� �������
Dim PermissedSymb As String: PermissedSymb = VBA.UCase$(c_strSymbRusAll & c_strOthers)  '(c_strOthers)
' ��������� ��� ������� ���� �� ������ ������ �� �� ������
    c = 1: cMax = Len(Source)
    Do Until c > cMax
        Char = VBA.UCase$(VBA.Mid$(Source, c, 1))
        Select Case Char
        Case "A" To "Z", "a" To "z", "0" To "9"
        Case Else: If InStr(1, PermissedSymb, Char) = 0 Then Result = True: Exit Do
        End Select
        c = c + 1
    Loop
HandleExit:  TextContainsAlpha = Result
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function TextAlpha2Code(Source As String, _
    Optional Encoding As Byte = 0, _
    Optional Prefix As String = "%") As String
' �������� ������� �� �������� � ������ ���������� �� ����������������� ����� ���� %XX
'-------------------------
' Source   - ���������� ������
' Encoding - ��� ��������� 0-cp1251, 1-UTF-8, 2-URL ��� (��� � ��������� ��������)
' Prefix   - ������� ���� �������: "%","\u","=" ��� ��
'-------------------------
Dim Result As String
    Result = vbNullString
    On Error GoTo HandleError
Dim c As Long, cMax As Long, cLen As Byte: c = 1: cMax = Len(Source): If cMax = 0 Then GoTo HandleExit
' ���������� ��������� �����������
    Select Case Encoding
    Case 0: cLen = 2  ' cp1251   Prefix = "%"
    Case 1: cLen = 4  ' UTF-8    Prefix = "\u"
    Case 2: cLen = 2  ' URL ���  Prefix = "%" ��� "="
    Case Else: Err.Raise vbObjectError + 512
    End Select
' ������ �������������� ����������� ������� ������ a-z,A-Z � 0-9. ����� ������� � �������� �������
Dim PermissedSymb As String: PermissedSymb = Replace(VBA.UCase$(c_strOthers), " ", "") '(c_strSymbRusAll & c_strOthers)  '(c_strOthers)
' ��������� ��� ������� ������
Dim Char As String, Code As String
    Do Until c > cMax
        Char = VBA.Mid$(Source, c, 1)
        Select Case Char
        Case "A" To "Z", "a" To "z", "0" To "9"
        Case Else: If InStr(1, PermissedSymb, VBA.UCase$(Char)) > 0 Then GoTo HandleNext
            Select Case Encoding
            Case 0: Code = VBA.Hex$(Asc(Char)):  Char = Prefix & String(cLen - Len(Code), "0") & Code
            Case 1: Code = VBA.Hex$(AscW(Char)): Char = Prefix & String(cLen - Len(Code), "0") & Code
            Case 2: Code = VBA.Hex$(AscW(Char)): Code = String(2 * cLen - Len(Code), "0") & Code
                    Select Case VBA.Left$(Code, cLen)
                    Case "00": Char = Prefix & VBA.Mid$(Code, cLen + 1)
                    Case "04": Char = Prefix & Hex$(&HD0 + (CLng(c_strHexPref & VBA.Mid$(Code, cLen + 1)) \ &H40)) & Prefix & Hex$(&H80 + (CLng(c_strHexPref & VBA.Mid$(Code, cLen + 1)) Mod &H40))
                    Case Else: Char = Prefix & VBA.Left$(Code, cLen) & Prefix & VBA.Mid$(Code, cLen + 1)
                    End Select
            End Select
        End Select
HandleNext:  Result = Result & Char: c = c + 1
    Loop
HandleExit:  TextAlpha2Code = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function TextCode2Alpha(Source As String, _
    Optional Encoding As Byte = 0, _
    Optional Prefix As String = "%") As String
' �������� ��� ���� %XX, �������� �� �������� � ������ ����������, �� ���������
'-------------------------
' Source   - ������������ ������
' Encoding - ��� ��������� 0-cp1251, 1-UTF-8, 2-URL ��� (��� � ��������� ��������)
' Prefix   - ������� ���� �������: "%","\u","=" ��� ��
'-------------------------
Dim Result As String
    On Error GoTo HandleError
    Result = vbNullString
    On Error GoTo HandleError
Dim c As Long, cMax As Long, cLen As Byte: c = 1: cMax = Len(Source): If cMax = 0 Then GoTo HandleExit
' ���������� ��������� �����������
    Select Case Encoding
    Case 0: cLen = 2  ' cp1251   Prefix = "%"
    Case 1: cLen = 4  ' UTF-8    Prefix = "\u"
    Case 2: cLen = 2  ' URL ���  Prefix = "%" ��� "="
    Case Else: Err.Raise vbObjectError + 512
    End Select
' ��������� ��� ������� ������
Dim Char As String, Code As String, Cod2 As String
    Do Until c > cMax
        If VBA.Mid$(Source, c, Len(Prefix)) <> Prefix Then
' ����������� (����������������) ������
            Char = VBA.Mid$(Source, c, 1)
        Else
' ���� ������� ����������� ������ - �������������� ���
            Code = VBA.UCase$(VBA.Mid$(Source, c + Len(Prefix), cLen))
            Select Case Encoding
            Case 0: Code = c_strHexPref & Code: If IsNumeric(Code) Then Char = VBA.Chr$(Val(Code)):  c = c + cLen + Len(Prefix) - 1
            Case 1: Code = c_strHexPref & Code: If IsNumeric(Code) Then Char = VBA.ChrW$(Val(Code)): c = c + cLen + Len(Prefix) - 1
            Case 2: Code = c_strHexPref & Code: c = c + cLen + Len(Prefix) - 1
            ' ������� U+0000..U+00FF >> %00..%FF
            ' ������� U+0400..U+04FF >> %D0%80..%D0%BF;%D1%80..%D1%BF;%D2%80..%D2%BF;%D3%80..%D3%BF
                    If IsNumeric(Code) Then
                        Select Case CLng(Code)
                        Case &HD0 To &HD3:
                            If VBA.Mid$(Source, c + 1, Len(Prefix)) = Prefix Then
                                Cod2 = c_strHexPref & VBA.UCase$(VBA.Mid$(Source, c + Len(Prefix) + 1, cLen))
                                If IsNumeric(Cod2) Then
                                    Select Case CLng(Cod2)
                                    Case &H80 To &HBF: Code = c_strHexPref & Hex(Val(&H400 + &H40 * (Val(Code) Mod &HD0) + Val(Cod2) - &H80)): c = c + cLen + Len(Prefix) '- 1
                                    Case Else:         Code = Code & Mid$(Cod2, Len(Prefix) + 1)
                                    End Select
                                End If
                            End If
                        Case Else:
                        End Select
                    End If
                    Char = VBA.ChrW$(Val(Code))  ' ������ �������
            End Select
        End If
HandleNext:  Result = Result & Char: c = c + 1
    Loop
HandleExit:  TextCode2Alpha = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function TextClearAlpha( _
    Source, _
    Optional ReplaceWith As String = "_", _
    Optional AllowedSymbols As String)
' ������� ������, ������� ������� �� �������� � ������ ���������� ��������� ��������
'-------------------------
' �������� ����� � ��������� �����.
'-------------------------
Dim c As Long, cMax As Long
Dim i As Long, iMax As Long
Dim strSource As String, strResult As String
Dim Char As String, PrevChar As String
Dim Result

    On Error GoTo HandleError
    Result = vbNullString
    If Len(AllowedSymbols) = 0 Then AllowedSymbols = c_strSymbDigits & c_strSymbRusAll & c_strSymbEngAll
    AllowedSymbols = UCase$(AllowedSymbols)
    If IsArray(Source) Then
        i = LBound(Source): iMax = UBound(Source)
        If iMax < i Then GoTo HandleExit
        ReDim Result(i To iMax)
        strSource = Source(i)
    Else
        i = 1: iMax = 1
        strSource = CStr(Source)
    End If
    Do
        cMax = Len(strSource)
        strResult = vbNullString
        For c = 1 To cMax
            Char = Mid$(strSource, c, 1)
            If InStr(1, AllowedSymbols, UCase$(Char)) = 0 Then
                If PrevChar <> ReplaceWith Then Char = ReplaceWith Else Char = vbNullString
                PrevChar = ReplaceWith
             Else
                PrevChar = Char
            End If
            strResult = strResult & Char
        Next c
        If Right$(strResult, Len(ReplaceWith)) = ReplaceWith Then strResult = Left$(strResult, Len(strResult) - Len(ReplaceWith))
        If Left$(strResult, Len(ReplaceWith)) = ReplaceWith Then strResult = Right$(strResult, Len(strResult) - Len(ReplaceWith))
        strResult = Trim$(strResult)
        If Len(strResult) < 1 Then strResult = vbNullString 'c_strEmptyString
        If IsArray(Source) Then
            Result(i) = strResult
            If i >= iMax Then Exit Do
            i = i + 1: strSource = Source(i)
         Else
            Result = strResult
            Exit Do
        End If
    Loop
HandleExit:  TextClearAlpha = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function GetCharType(Letter As String, Optional AlphaType As Byte = 0) As Byte
' ��������� ����� � ���������� ���������
'-------------------------
' �� ������: 1-�������,2-���������,3-����(��),� �.�.,0-�� ����������
' AlphaType = 1-����� ���������� ��������, 2-����� �������� ��������, 3-�����, 0-���� ������
'-------------------------
Dim Result As Byte: Result = 0 ': SymbType = 0
Dim sChar As String * 1: sChar = LCase$(Left$(Trim$(Letter), 1))

    If InStr(c_strSymbDigits, sChar) > 0 Then Result = SymbolTypeNumb:      AlphaType = AlphabetTypeUndef:      GoTo HandleExit
    If InStr(c_strSymbRusVowel, sChar) > 0 Then Result = SymbolTypeVowel:   AlphaType = AlphabetTypeCyrilic:    GoTo HandleExit
    If InStr(c_strSymbRusConson, sChar) > 0 Then Result = SymbolTypeCons:   AlphaType = AlphabetTypeCyrilic:    GoTo HandleExit
    If InStr(c_strSymbRusSign, sChar) > 0 Then Result = SymbolTypeSign:     AlphaType = AlphabetTypeCyrilic:    GoTo HandleExit
    If InStr(c_strSymbEngVowel, sChar) > 0 Then Result = SymbolTypeVowel:   AlphaType = AlphabetTypeLatin:      GoTo HandleExit
    If InStr(c_strSymbEngConson, sChar) > 0 Then Result = SymbolTypeCons:   AlphaType = AlphabetTypeLatin:      GoTo HandleExit
    If InStr(c_strSymbEngSign, sChar) > 0 Then Result = SymbolTypeSign:     AlphaType = AlphabetTypeLatin:      GoTo HandleExit
'    If InStr(c_strSymbMath, sChar)>0 Then Result = : AlphaType = AlphabetTypeUndef:GoTo HandleExit
'    If InStr(c_strSymbPunct, sChar)>0 Then Result = : AlphaType = AlphabetTypeUndef:GoTo HandleExit
'    If InStr(c_strSymbCommas, sChar)>0 Then Result = : AlphaType = AlphabetTypeUndef:GoTo HandleExit
'    If InStr(c_strSymbParenth, sChar)>0 Then Result = : AlphaType = AlphabetTypeUndef:GoTo HandleExit
'    If InStr(c_strSymbOthers, sChar)>0 Then Result = SymbolTypeUndef:       AlphaType = AlphabetTypeUndef:      GoTo HandleExit
HandleExit: GetCharType = Result
End Function
' ==================
' ������� ��� ��������� ����
' ==================
Public Function PolyPhone(ByVal Word As String, Optional FuzzyIdx As Boolean = False) 'As String
'Polyphon: An Algorithm for Phonetic String Matching in Russian Language (Paramonov V.V., Shigarov A O., Ruzhnikov G.M. )
'-------------------------
' FuzzyIdx - ���� True  - ���������� �������� ��� ��� ��������� ���������
'            ���� False - ���������� ������������ ���
'-------------------------
'http://td.icc.ru/files/papers/Paramonov_ICIST2016.pdf
'https://cyberleninka.ru/article/n/obzor-algoritmov-foneticheskogo-kodirovaniya
'-------------------------
Dim Result As String, x As Long: Result = vbNullString
    On Error GoTo HandleError
Dim i As Long, sChar As String
    For i = 1 To Len(Word)
        sChar = Mid$(Word, i, 1)
        Select Case sChar
        Case "B":                     sChar = "�" ' ����������� ������ ��������� ���� ������ ���� �������� ��������:
        Case "M":                     sChar = "�"
        Case "H":                     sChar = "�"
        Case "A", "a":                sChar = "�"
        Case "E", "e":                sChar = "�"
        Case "O", "o":                sChar = "�"
        Case "C", "c":                sChar = "�"
        Case "X", "x":                sChar = "�"
        Case "�", "�", "�", "�":      sChar = vbNullString  ' �������� ���� �, �.
        Case "�" To "�", "�" To "�":  sChar = UCase$(sChar)
        Case Else:                    sChar = vbNullString  ' �������� ���� ����, �� ������������� �������� �������� �����.
        End Select
        ' ������ ���� ���������� ���� �����.
        If sChar = UCase$(Mid$(Word, i + 1, 1)) Then i = i + 1: GoTo HandleNext
        ' ������ ��������� ����
        Select Case sChar
        Case "�", "�", "�", _
             "�", "�", "�", "�", "�": sChar = "�"
        Case "�":                     sChar = "�"
        Case "�":                     sChar = "�"
        Case "�":                     sChar = "�"
        Case "�":                     sChar = "�"
        Case "�":                     sChar = "�"
        Case "�":                     sChar = "�"
        Case "�":                     sChar = "�"
        Case "�":                     sChar = "�"
        Case "�":                     sChar = "�"
        End Select
        ' ���������� �����������:
        If Len(Result) > 3 Then
            Select Case Right$(Result, 4) & sChar
            Case "�����": Mid$(Result, Len(Result) - 3, 4) = "����": sChar = vbNullString: GoTo HandleNext
            End Select
        End If
        If Len(Result) > 2 Then
            Select Case Right$(Result, 3) & sChar
            Case "����": Mid$(Result, Len(Result) - 2, 3) = "�C�": sChar = vbNullString: GoTo HandleNext
            Case "����": Mid$(Result, Len(Result) - 2, 3) = "C��": sChar = vbNullString: GoTo HandleNext
            End Select
        End If
        If Len(Result) > 1 Then
            Select Case Right$(Result, 2) & sChar
            Case "���": Mid$(Result, Len(Result) - 1, 2) = "� ": Result = Trim$(Result): sChar = vbNullString: GoTo HandleNext
            Case "���": Mid$(Result, Len(Result) - 1, 2) = "� ": Result = Trim$(Result): sChar = vbNullString: GoTo HandleNext
            Case "���": Mid$(Result, Len(Result) - 1, 2) = "� ": Result = Trim$(Result): sChar = vbNullString: GoTo HandleNext
            Case "���": Mid$(Result, Len(Result) - 1, 2) = "� ": Result = Trim$(Result): sChar = vbNullString: GoTo HandleNext
            Case "���": Mid$(Result, Len(Result) - 1, 2) = "��": sChar = vbNullString: GoTo HandleNext
            Case "���": Mid$(Result, Len(Result) - 1, 2) = "��": sChar = vbNullString: GoTo HandleNext
            Case "���": Mid$(Result, Len(Result) - 1, 2) = "C�": sChar = vbNullString: GoTo HandleNext
            Case "���": Mid$(Result, Len(Result) - 1, 2) = "C�": sChar = vbNullString: GoTo HandleNext
            Case "���": Mid$(Result, Len(Result) - 1, 2) = "�A": sChar = vbNullString: GoTo HandleNext
            Case "���": Mid$(Result, Len(Result) - 1, 2) = "��": sChar = vbNullString: GoTo HandleNext
            Case "���": Mid$(Result, Len(Result) - 1, 2) = "�C": sChar = vbNullString: GoTo HandleNext
            Case "���": Mid$(Result, Len(Result) - 1, 2) = "��": sChar = vbNullString: GoTo HandleNext
            Case "���": Mid$(Result, Len(Result) - 1, 2) = "��": sChar = vbNullString: GoTo HandleNext
            Case "���": Mid$(Result, Len(Result) - 1, 2) = "��": sChar = vbNullString: GoTo HandleNext
            Case "���": Mid$(Result, Len(Result) - 1, 2) = "��": sChar = vbNullString: GoTo HandleNext
            Case "���": Mid$(Result, Len(Result) - 1, 2) = "P�": sChar = vbNullString: GoTo HandleNext
            Case "���": Mid$(Result, Len(Result) - 1, 2) = "P�": sChar = vbNullString: GoTo HandleNext
            Case "��A": Mid$(Result, Len(Result) - 1, 2) = "A�": sChar = "A": GoTo HandleNext
            Case "���": Mid$(Result, Len(Result) - 1, 2) = "��": sChar = "�": GoTo HandleNext
            End Select
        End If
        If Len(Result) > 0 Then
            Select Case Right$(Result, 1) & sChar
            Case "��": Mid$(Result, Len(Result), 1) = "�": sChar = vbNullString: GoTo HandleNext
            Case "��": Mid$(Result, Len(Result), 1) = "�": sChar = vbNullString: GoTo HandleNext
            Case "��": Mid$(Result, Len(Result), 1) = "�": sChar = vbNullString: GoTo HandleNext
            Case "��": Mid$(Result, Len(Result), 1) = "�": sChar = vbNullString: GoTo HandleNext
            Case "��": Mid$(Result, Len(Result), 1) = "�": sChar = vbNullString: GoTo HandleNext
            Case "��": Mid$(Result, Len(Result), 1) = "�": sChar = vbNullString: GoTo HandleNext
            Case "��": Mid$(Result, Len(Result), 1) = "�": sChar = vbNullString: GoTo HandleNext
            Case "��": Mid$(Result, Len(Result), 1) = "�": sChar = vbNullString: GoTo HandleNext
            Case "��": Mid$(Result, Len(Result), 1) = "C": sChar = "�": GoTo HandleNext
            Case "��": Mid$(Result, Len(Result), 1) = "�": sChar = "�": GoTo HandleNext
            End Select
        End If
HandleNext:
        If FuzzyIdx Then
        ' ��������� �������� �������� �������� ��������
            Select Case sChar
            Case "�": x = x + 2
            Case "�": x = x + 3
            Case "�": x = x + 5
            Case "�": x = x + 7
            Case "�": x = x + 11
            Case "�": x = x + 13
            Case "�": x = x + 17
            Case "�": x = x + 19
            Case "�": x = x + 23
            Case "�": x = x + 29
            Case "�": x = x + 31
            Case "�": x = x + 37
            Case "�": x = x + 41
            Case "�": x = x + 43
            Case "�": x = x + 47
            Case "�": x = x + 53
            Case "�": x = x + 59
            End Select
        End If
        Result = Result & sChar
    Next i
HandleExit:  PolyPhone = IIf(FuzzyIdx, x, Result): Exit Function
HandleError: Result = vbNullString: x = 0: Err.Clear: Resume HandleExit
End Function
Public Function MetaPhoneRu1(ByVal Word As String) As String
'�������������� ������� �������, �� �� �����������.
'-------------------------
'��������: http://forum.aeroion.ru/topic461.html
Const alf$ = "�����������������������������ۨ", _
      cns1$ = "�����", _
      cns2$ = "�����", _
      cns3$ = "����������������", _
      ch$ = "����ߨ�", _
      ct$ = "�������"
'alf - ������� ����� ����������� ����, cns1 � cns2 - ������� � ������
'���������, cns3 - ���������, ����� �������� ������� ����������,
'ch, ct - ������� � ������ �������
Dim s$, v$, i&, b&, c$
'S, V - ������������� ������, i - ������� �����, B - �������
'���������� ��������, c$ - ������� ������

'��������� � ������� �������, ��������� ������ ������� �� alf
'����������� ������ � ������, �������� � S:
    Word = UCase$(Word): s = " "
    For i = 1 To Len(Word)
        c = Mid$(Word, i, 1)
        If InStr(alf, c) Then s = s & c
    Next i
    If Len(s) = 1 Then Exit Function
    '�������� ���������:
    Select Case Right$(s, 6)
    Case "������":      s = Left$(s, Len(s) - 6) & "@"
    Case "������":      s = Left$(s, Len(s) - 6) & "#"
    Case "������":      s = Left$(s, Len(s) - 6) & "$"
    Case "������":      s = Left$(s, Len(s) - 6) & "%"
    End Select
    
    Select Case Right$(s, 3)
    Case "���", "���":  s = Left$(s, Len(s) - 3) & "9"
    Case "���":         s = Left$(s, Len(s) - 3) & "1"
    Case "���":         s = Left$(s, Len(s) - 3) & "3"
    End Select
    
    Select Case Right$(s, 2)
    Case "��", "��":    s = Left$(s, Len(s) - 2) & "4"
    Case "��":          s = Left$(s, Len(s) - 2) & "6"
    Case "��", "��":    s = Left$(s, Len(s) - 2) & "7"
    Case "��", "��":    s = Left$(s, Len(s) - 2) & "5"
    Case "��":          s = Left$(s, Len(s) - 2) & "8"
    Case "��", "��":    s = Left$(s, Len(s) - 2) & "2"
    Case "��", "��":    s = Left$(s, Len(s) - 2) & "0"
    End Select
    '�������� ��������� ������, ���� �� - ������� ���������:
    b = InStr(cns1, Right$(s, 1))
    If b Then Mid$(s, Len(s), 1) = Mid$(cns2, b, 1)
    '�������� ����:
    For i = 2 To Len(s)
        c = Mid$(s, i, 1)
        b = InStr(ch, c)
        If b Then Mid$(s, i, 1) = Mid$(ct, b, 1) '������ �������
        If InStr(cns3, c) Then '��������� ���������
            b = InStr(cns1, Mid$(s, i - 1, 1))
            If b Then Mid$(s, i - 1, 1) = Mid$(cns2, b, 1)
        End If
    Next i
    '��������� �������, ������� ������ ������:
    For i = 2 To Len(s)
        c = Mid$(s, i, 1)
        If c <> Mid$(s, i - 1, 1) Then v = v & c
    Next i
    MetaPhoneRu1 = v
End Function
Public Function MetaPhoneRu2(ByVal Word As String) As String
'������ ������� �������, ������.
'-------------------------
'��������: http://forum.aeroion.ru/topic461.html
'�������� ��, �� ���.; ������� �������������.
Const alf$ = "����������������������������٨�", _
      cns1$ = "�����", _
      cns2$ = "�����", _
      cns3$ = "����������������", _
      ch$ = "����ߨ�", _
      ct$ = "�������"
'alf - ������� ����� ����������� ����, cns1 � cns2 - ������� � ������
'���������, cns3 - ���������, ����� �������� ������� ����������,
'ch, ct - ������� � ������ �������
Dim s$, v$, i&, b&, c$, old_c$
'S, V - ������������� ������, i - ������� �����, B - �������
'���������� ��������, c$ - ������� ������, c_old$ - ����������
'������

'��������� � ������� �������, ��������� ������
'������� �� alf � �������� � S:
    Word = UCase$(Word)
    For i = 1 To Len(Word)
        c = Mid$(Word, i, 1)
        If InStr(alf, c) Then s = s & c
    Next i
    If Len(s) = 0 Then Exit Function
    '������� ���������:
    Select Case Right$(s, 6)
    Case "������":                  s = Left$(s, Len(s) - 6) & "@"
    Case "������":                  s = Left$(s, Len(s) - 6) & "#"
    Case "������":                  s = Left$(s, Len(s) - 6) & "$"
    Case "������":                  s = Left$(s, Len(s) - 6) & "%"
    Case Else
        If Right$(s, 4) = "����" Or Right$(s, 4) = "����" Then
            s = Left$(s, Len(s) - 4) & "9"
        Else
            Select Case Right$(s, 3)
            Case "���", "���":      s = Left$(s, Len(s) - 3) & "9"
            Case "���":             s = Left$(s, Len(s) - 3) & "1"
            Case "���", "���":      s = Left$(s, Len(s) - 3) & "4"
            Case "���":             s = Left$(s, Len(s) - 3) & "3"
            Case Else
                Select Case Right$(s, 2)
                Case "��", "��":    s = Left$(s, Len(s) - 2) & "4"
                Case "��":          s = Left$(s, Len(s) - 2) & "6"
                Case "��", "��":    s = Left$(s, Len(s) - 2) & "7"
                Case "��", "��":    s = Left$(s, Len(s) - 2) & "5"
                Case "��":          s = Left$(s, Len(s) - 2) & "8"
                Case "��", "��":    s = Left$(s, Len(s) - 2) & "2"
                Case "��", "��":    s = Left$(s, Len(s) - 2) & "0"
                End Select
            End Select
        End If
    End Select
    '�������� ��������� ������, ���� �� - ������� ���������:
    b = InStr(cns1, Right$(s, 1))
    If b Then Mid$(s, Len(s), 1) = Mid$(cns2, b, 1)
    old_c = " "
    '�������� ����:
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        b = InStr(ch, c)
        If b Then   '���� �������
            If old_c = "�" Or old_c = "�" Then
                If c = "�" Or c = "�" Then '�������������� � �������
                    old_c = "�": Mid$(v, Len(v), 1) = old_c
                Else '���� �� �������������� � �������, � ������ �������
                    If c <> old_c Then v = v & Mid$(ct, b, 1)
                End If
            Else    '���� �� �������������� � �������, � ������ �������
                If c <> old_c Then v = v & Mid$(ct, b, 1)
            End If
        Else        '���� ���������
            If c <> old_c Then '��� ����������
                If InStr(cns3, c) Then '��������� ���������
                    b = InStr(cns1, old_c)
                    If b Then old_c = Mid$(cns2, b, 1): Mid$(v, Len(v), 1) = old_c
                End If
                If c <> old_c Then v = v & c '��� ������
            End If
        End If
        old_c = c
    Next i
    MetaPhoneRu2 = v
End Function

Public Function MetaPhoneRu3(ByVal Word As String) As Long
'������ ������� ���������� ����� �24-������ �����.
'-------------------------
'��������: http://forum.aeroion.ru/topic461.html
'������������ ���������� ������ ���� �������� ������������ �����, ����������� ��� ��������
'������-�����.  ��������� ����������� SoundEx ��� �������� ������, �������� ���� ������.

'������ ������� MetaPhoneRu �� ����� �������� �������� ������ ���� ������ ������� 7�����
'������������ ���� (��� ������ Unicode�� 14 ����). Ÿ ����� ��������� ������ Long,
'���������� ������ 4������.

'��������, ����� 24 ��������� ����� ������� ������� �����. �������� �� �24-������
'������� ��������� ���, ��������, ���:
'��� ���� �1�, �ʻ�� �8�, �ͻ�� �11�. �����
'���� ����������� �24 * 8 + 1 = 193;
'���ͻ�� �24 * (24 * 8 + 1) +  11 = 4643.

'���������� ������� �MetaPhoneRu ������� ������ ����� ��������� �������� ������� 24.
' ����� ���� Long ������� ��������� ������ 6�����, � ������� ������� �������: �����������
'� ��������������� ����������� �� ����� ��� ��������� �� ������λ.

'�������� ���� ���������� ���������� ������� ���������: ��������� � ����������� ����� �����
'������ �����, ���� ������ ����� ���� ����� ������� �� ��. ��� ���������������� ��������
'������ ����� � ���������� ����������� �����, ������������ ��-�� ����, ��� 24 �� ��������
'�������� ������. ����� 24^6 (����� ���������� �� 24-� �������� �� ����� ��������) � 2^32
'(����� ��������� ��������� 4-������� ���������� Long) ������� ��������� �������, ��� ��
'���������� 24-������ �����. ��� ���������� � ������������ ��� �������� �����������
'��� ����� ��������� 2^32 / 24^6 = 22 ��������� ���������, �� ������� ���������� ������������
'������������.

'������� ��������� �� �������� ���� ������� MetaPhoneRu ���� �� ����� ������� �������� ����
'������� ����� �log2(256) / log2(24) = �1,7 ���, �� ��� ������� ������� ����� �������� �������
'���������. ����� ����, ��������� ������� ������ � ������������: ��������� ���� ����� �������
'�� ����� �� ����� ������� � ����� ���������, � �������� ��� ������� �������������� �����������.
'-------------------------
Const alf$ = "�������������������������Ψ����", _
      cns1$ = "�����", _
      cns2$ = "�����", _
      cns3$ = "����������������", _
      ch$ = "����ߨ�", _
      ct$ = "�������"
'alf - ������� ����� ����������� ����, cns1 � cns2 - ������� � ������
'���������, cns3 - ���������, ����� �������� ������� ����������,
'ch, ct - ������� � ������ �������
Dim s$, v&, i&, b&, c$, old_c$, new_c$
'S - ������������� ������, V�� ����, ������� �������� �����
'������ ���������, i - ������� �����, B - ������� ����������
'��������, c$ - ������� ������, c_old$ - ���������� ������,
'new_c$�� ��������������� ������� ������.

'��������� � ������� �������, ��������� ������
'������� �� alf � �������� � S:
    Word = UCase$(Word)
    For i = 1 To Len(Word)
        c = Mid$(Word, i, 1)
        If InStr(alf, c) Then s = s + c
    Next i
    If Len(s) = 0 Then Exit Function
    '������� ���������:
    Select Case Right$(s, 6)
    Case "������": s = Left$(s, Len(s) - 6): v = -1
    Case "������": s = Left$(s, Len(s) - 6): v = -2
    Case "������": s = Left$(s, Len(s) - 6): v = -3
    Case "������": s = Left$(s, Len(s) - 6): v = -4
    Case Else
        Select Case Right$(s, 4)
        Case "����", "����": s = Left$(s, Len(s) - 4): v = 9
        Case Else
            Select Case Right$(s, 3)
            Case "���", "���": s = Left$(s, Len(s) - 3): v = 9
            Case "���": s = Left$(s, Len(s) - 3): v = 1
            Case "���", "���": s = Left$(s, Len(s) - 3): v = 4
            Case "���": s = Left$(s, Len(s) - 3): v = 3
            Case Else
                Select Case Right$(s, 2)
                Case "��", "��": s = Left$(s, Len(s) - 2): v = 4
                Case "��": s = Left$(s, Len(s) - 2): v = 6
                Case "��", "��": s = Left$(s, Len(s) - 2): v = 7
                Case "��", "��": s = Left$(s, Len(s) - 2): v = 5
                Case "��": s = Left$(s, Len(s) - 2): v = 8
                Case "��", "��": s = Left$(s, Len(s) - 2): v = 2
                Case "��", "��": s = Left$(s, Len(s) - 2): v = -5
                End Select
            End Select
        End Select
    End Select
    '�������� ��������� ������, ���� �� - ������� ���������:
    b = InStr(cns1, Right$(s, 1))
    If b Then Mid$(s, Len(s), 1) = Mid$(cns2, b, 1)
    old_c = " "
    s = Left$(s, 6)
    '�������� ����:
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        b = InStr(ch, c)
        If b Then '���� �������
            If old_c = "�" Or old_c = "�" Then
                If c = "�" Or c = "�" Then '�������������� � �������
                    old_c = "�"
                Else '���� �� �������������� � �������, � ������ �������
                    If c <> old_c Then new_c = Mid$(ct, b, 1)
                End If
            Else '���� �� �������������� � �������, � ������ �������
                If c <> old_c Then new_c = Mid$(ct, b, 1)
            End If
        Else '���� ���������
            If c <> old_c Then '��� ����������
                If InStr(cns3, c) Then '��������� ���������
                    b = InStr(cns1, old_c)
                    If b Then old_c = Mid$(cns2, b, 1)
                End If
                If c <> old_c Then new_c = c '��� ������
            End If
        End If
        old_c = c
        v = v * 24                '����� ����� � 24-������ ����� V��
        v = v + InStr(alf, new_c) '���������� ����� ����� new_c �alf$.
        
        '������ 24 ������� �alf�� ��� ��������� �������� ������
        '����� �����. ����� ��� �alf ��������� �� �������, �������
        '���������� ������ � ��������� ����� �� ������������.
    Next i
    MetaPhoneRu3 = v
End Function
Public Function SoundEx(ByVal Word As String) As String
' ������� Russell (NARA) Soundex, �������� � ������ c������ W � H
'-------------------------
' ��������: http://forum.aeroion.ru/topic443.html
' �����:    http://www.source-code.biz/snippets/vbasic/4.htm
'-------------------------
Dim s As String, l As Long
Dim Result As String
    s = Trim$(UCase$(Word)): l = Len(s)
    If l = 0 Then Result = String$(4, 0): GoTo HandleExit
Const RusTab = "�����Ũ��������������������������"  ' ��� ���������� ��������������
Const LatTab = "ABVGDEEGZIIKLMNOPRSTUFHCHHHIAWAXY"  '(��� �������������� �� ���� ��.TextTranslit)
Const CodeTab = "01230120022455012623010202"        ' ��� ����������� �� �������� Soundex
'               "ABCDEFGHIJKLNMOPQRSTUVWXYZ"        ' ��������������� ����� �������
'' �������������� ��� ��������� ��� �������������� ��: https://cyberleninka.ru/article/n/obzor-algoritmov-foneticheskogo-kodirovaniya
'Const CodeTab = "012460033074788019360235555000000"
''               "�����Ũ��������������������������"
Dim h As Long:  h = 0   ' ��� �������� �������
Dim lH As Long: lH = -1 ' ��� ����������� �������. -1 - ������ ������ (��� �����������)
Dim i As Long:  i = 1   ' ������� ������� �� ������� ������
Dim c As String         ' ������� ������
Dim p As Integer        ' ������� �������� ������� � ������ ��������������
    Do Until i > l
        c = Mid$(s, i, 1)
    ' ������� ��������������
        p = InStr(1, RusTab, c): If p > 0 Then c = Mid$(LatTab, p, 1)
    ' �������� ���������� ��������
        If InStr(1, LatTab, c) = 0 Then GoTo HandleNext
    ' �������� Soundex ��� �������
        h = Mid$(CodeTab, Asc(c) - 64, 1)
        If lH <> h Then
    ' c������ �������, ��� �������, ����������� ������� H ��� W,
    ' �������� � ���� � �� �� ������, ������������ ��� ����
            If lH = -1 Then
    ' ������ ������ � ������
                Result = c: lH = h
            Else
                If h = 0 Then
    ' �� ����� ������: "HWAEIOUY" - ������� (� ������� ���������� "�Ũ���������")
                    If InStr(1, "HW", c) = 0 Then lH = h '
                    GoTo HandleNext
                End If
    ' ���������� ���������� Soundex ��� � ��������� ���������� � �������������� ������
                lH = h: Result = Result & h
            End If
        End If
    ' ���� ����� Soundex ���� >4 - �������
        If Len(Result) >= 4 Then Exit Do
HandleNext: i = i + 1    ' ��������� ������
    Loop
    Result = Result & String$(4 - Len(Result), "0")
HandleExit: SoundEx = Result
End Function
Public Function SoundEx2(ByVal Word As String) As String
' ������� Refined Soundex
'-------------------------
' ��������: https://habr.com/ru/post/114947/
'-------------------------
Dim s As String, l As Long
Dim Result As String
    s = Trim$(UCase$(Word)): l = Len(s)
    If l = 0 Then Result = String$(4, 0): GoTo HandleExit
Const RusTab = "�����Ũ��������������������������"  ' ��� ��������������
Const LatTab = "ABVGDEEGZIIKLMNOPRSTUFHCHHHIAWAXY"
'               "ABCDEFGHIJKLNMOPQRSTUVWXYZ"
Const CodeTab = "01360240043788015936020505"
' ����� ����������� �� �������� Refined Soundex ������� �������� ��� ��������������
' �� https://cyberleninka.ru/article/n/obzor-algoritmov-foneticheskogo-kodirovaniya
'Const CodeTab = "01246003307478801936023555500000"
''               "�����Ũ�������������������������"
Dim h As Long:  h = 0   ' ��� �������� �������
Dim lH As Long: lH = -1 ' ��� ����������� �������. -1 - ������ ������ (��� �����������)
Dim i As Long:  i = 1   ' ������� ������� �� ������� ������
Dim c As String         ' ������� ������
Dim p As Integer        ' ������� �������� ������� � ������ ��������������
    Do Until i > l
        c = Mid$(s, i, 1)
    ' ������� ��������������
        p = InStr(1, RusTab, c): If p > 0 Then c = Mid$(LatTab, p, 1)
    ' �������� ���������� ��������
        If InStr(1, LatTab, c) = 0 Then GoTo HandleNext
    ' �������� Soundex ��� �������
        h = Mid$(CodeTab, Asc(c) - 64, 1)
    ' �������� � ���� � �� �� ������, ������������ ��� ����
        If lH = h Then GoTo HandleNext
    ' ������ ������ � ������
        If lH = -1 Then Result = c
    ' ���������� ���������� Soundex ��� � ��������� ���������� � �������������� ������
        lH = h: Result = Result & h
HandleNext: i = i + 1  ' ��������� ������
    Loop
HandleExit: SoundEx2 = Result
End Function
Public Function SoundExDM(ByVal Word As String) As String
' SoundEx - Daitch-Mokotoff
'-------------------------
' �������� ���������: http://www.avotaynu.com/soundex.htm
' ��� �������� ����������� ������: https://stevemorse.org/census/soundex.html
' ���� ������� ��� "rs":    "Halberstadt"   ��� "587943 587433", ������ - "587943"
'                           "Peters"        ��� "739400 734000", ������ - "739400"
' �.�. ������ ������ ������� ������ ��� ����� "rtz"
' � ���� ����� � https://cyberleninka.ru/article/n/obzor-algoritmov-foneticheskogo-kodirovaniya
' ��� ���������� ������� ���� ���������� "739400 734000"
' ����� ��� ������� ������������� rs ��� ���� 94 � 4 , � ������ ��� 9,4 (� 4)
' �.� ��� ��� ������� �������� ��� ����� ���������� ��������� � 94, � ������ � 4
' � ������ ������� �����������
'-------------------------
Const jMax = 6      ' ������������ ���������� �������� � �������� ����
Const nMax = 7      ' ������������ ����� ��������������� ��������
Const Delim = " "   ' ����������� ����������� � �������������� ������
Dim Result As String: Result = vbNullString 'String$(jMax, "0")
    On Error GoTo HandleError
    Word = Trim$(Word): If Len(Word) = 0 Then Err.Raise vbObjectError + 512 'GoTo HandleExit
    
    Word = LCase$(Replace(TextTranslit(Word), " ", ""))     ' �������������� ����� - ��������������, ������� ������� � ������ �������
Dim bolFound As Boolean                                     ' ������� ���������� ��������
Dim i As Integer, iMax As Integer: i = 1: iMax = Len(Word)  ' ������ ������� ������
Dim jEnd As Integer, jBeg As Integer                        ' ������� � �������������� ������
Dim r As Integer, rMax As Integer: rMax = 1                 ' ������ �������� ����������
Dim rOld As Integer: rOld = 1                               ' ���������� �������� ���������� �����������
Dim sPart As String, n As Integer                           ' �������������� �������� � ��� �����
Dim Code, sCode As String, � As Integer, cMax As Integer: cMax = 0  ' �������(��) ����
Dim Prev, sPrev As String, p As Integer, pMax As Integer: pMax = 0  ' ����������(��) ����
    Prev = vbNullString
    Do Until (i > iMax) 'Or (jEnd > jMax)
        If Len(Result) = rMax * (jMax + 1) - 1 Then Exit Do ' �������� �������� ��������, ����� ��� ���� ��� �������� ������������ ��������� �����. ������-�� ������ ���� �� ������
        cMax = 0                                ' ���������� ��������� ��������� ���� -1
        bolFound = False: Code = vbNullString   ' ���������� �������� ��������
        n = iMax - i + 1: If n > nMax Then n = nMax   ' �������� ������������ ��������� ����� ���������
        Do Until n < 1
    ' ���������� ���������� ��������� �� ����� �� ��������
        ' ��� ���������� ���������:
            '1. ������������� ���� bolFound=True
            '2. ���������� ����� ���������:
            '       [�] ������� ����� (i=1);
            '       [�] ����� ������� (InStr(1, c_strSymbEngVowel, Mid$(Word, i + n, 1)) > 0),
            '       [�] ��� ��������� ������
            '3. � ����������� �� ����� ���������� ��� ��������� (Code) ����������
            ' ��� �������� � �������������� (��. "ch","ck" � �.�.) ������ �����������
            
            sPart = Mid$(Word, i, n)
            Select Case n
            Case 7: Select Case sPart
                Case "schtsch": bolFound = True: If i > 1 Then Code = "4" Else Code = "2"
                End Select
            Case 6: Select Case sPart
                Case "schtsh", "schtch": bolFound = True: If i > 1 Then Code = "4" Else Code = "2"
                End Select
            Case 5: Select Case sPart
                Case "shtch", "shtsh", "stsch", "zhdzh": bolFound = True: If i > 1 Then Code = "4" Else Code = "2"
                Case "ttsch": bolFound = True: Code = "4"
                End Select
            Case 4: Select Case sPart
                Case "shch", "stch", "strz", "strs", "stsh", "szcz", "szcs", "zdzh": bolFound = True: If i > 1 Then Code = "4" Else Code = "2"
                Case "scht", "schd": bolFound = True: If i > 1 Then Code = "43" Else Code = "2"
                Case "ttch", "tsch", "ttsz", "zsch": bolFound = True: Code = "4" '4,4,4
                End Select
            Case 3: Select Case sPart
                Case "zdz": bolFound = True: If i > 1 Then Code = "4" Else Code = "2"
                Case "sht", "szt", "shd", "szd", "zhd": bolFound = True: If i > 1 Then Code = "43" Else Code = "2"
                Case "csz", "czs", "drz", "drs", "dsh", "dsz", "dzh", "dzs", "sch", "tch", "trz", "trs", "tsh", "tts", "ttz", "tzs", "tsz", "zsh": bolFound = True: Code = "4"
                Case "chs": bolFound = True: If i > 1 Then Code = "54" Else Code = "5"
                End Select
            Case 2: Select Case sPart
                Case "ai", "aj", "ay", "ei", "ej", "ey", "oi", "oj", "oy", "ui", "uj", "uy", "ue": bolFound = True: If i = 1 Then Code = "0" Else If InStr(1, c_strSymbEngVowel, Mid$(Word, i + n, 1)) > 0 Then Code = "1" 'Else Code = vbNullString
                Case "ia", "ie", "io", "iu": bolFound = True: If i = 1 Then Code = "1" 'Else Code = vbNullString
                Case "dt", "th": bolFound = True: Code = "3"
                Case "sc": bolFound = True: If i > 1 Then Code = "4" Else Code = "2"
                Case "st", "sd", "zd": bolFound = True: If i > 1 Then Code = "43" Else Code = "2"
                Case "cz", "cs", "ds", "dz", "sh", "ts", "tc", "tz", "sz", "zh", "zs", "rz": bolFound = True: Code = "4"
                Case "rs", "rz": bolFound = True: Code = Array("94", "4"): cMax = UBound(Code)  '"rs">"rtz","zh"
                Case "ch": bolFound = True: Code = Array("5", "4"): cMax = UBound(Code)   '"ch">"kh","tch"
                Case "ck": bolFound = True: Code = Array("5", "45"): cMax = UBound(Code)  '"ck">"k","tsk"
                Case "kh": bolFound = True: Code = "5"
                Case "ks": bolFound = True: If i > 1 Then Code = "54" Else Code = "5"
                Case "mn", "nm": bolFound = True: Code = "66"
                Case "pf", "ph", "fb": bolFound = True: Code = "7"
                Case "au": bolFound = True: If i = 1 Then Code = "0" Else If InStr(1, c_strSymbEngVowel, Mid$(Word, i + n, 1)) > 0 Then Code = "7"  'Else Code = vbNullString
                Case "eu": bolFound = True: If i = 1 Or InStr(1, c_strSymbEngVowel, Mid$(Word, i + n, 1)) > 0 Then Code = "1" 'Else Code = vbNullString
                End Select
            Case 1: Select Case sPart
                Case "a", "e", "i", "o", "u": bolFound = True: If i = 1 Then Code = "0" 'Else Code = vbNullString
                Case "y": bolFound = True: If i = 1 Then Code = "1" 'Else Code = vbNullString
                Case "j": bolFound = True: If i = 1 Then Code = Array("4", "1"): cMax = UBound(Code) '"j" >"dzh","y"
                Case "d", "t": bolFound = True: Code = "3"
                Case "s", "z": bolFound = True: Code = "4"
                Case "h": bolFound = True: If i = 1 Or InStr(1, c_strSymbEngVowel, Mid$(Word, i + n, 1)) > 0 Then Code = "5" 'Else Code = vbNullString  '5,5,-1
                Case "g", "k", "q": bolFound = True: Code = "5"
                Case "c": bolFound = True: Code = Array("5", "4"): cMax = UBound(Code) '"c" >"tz", "k"
                Case "x": bolFound = True: If i > 1 Then Code = "54" Else Code = "5" '5,54,54
                Case "m", "n": bolFound = True: Code = "6"
                Case "b", "f", "p", "v", "w": bolFound = True: Code = "7"
                Case "l": bolFound = True: Code = "8"
                Case "r": bolFound = True: Code = "9"
                End Select
            Case Else: Err.Raise vbObjectError + 512
            End Select
'!!! ��� ���� ���������� ���� ���������� ������ !!!
' ������ ���� � ���������.
    ' ������������ ������� �� ������� ���������� �������� ��� ������������ ��� ����������� �����������
    ' (������������ - ��������� �����), �� ������� ����� �������� ���� ��� �������� �������
    ' �, ��������, ���������� �������� � �������� ������ (�� �����������)
    ' ����� ������������ ��������� ��� ������ ��� �������� �� ������ ������ ��� ������/����� ���������
    ' ������������ ����������� ��� ����������� �������������� ��������� ���� �������� ������ ���������
    ' ���������, �����, �������� ��������� �������� � ������ ������� ��� ������������
    ' � ��������� ������������� �������� � ���������� ������� ������ ��� ���������� �������,
    ' ��.. - �� �� ���� ����� �����)
            If bolFound Or ((n = 1) And (i = iMax)) Then Else GoTo HandleNext
    ' ���� �������� ������ � ��������� ���������� ��� ��� � ���������
    ' ���� ��� �� ������������ ������ � ����� ������ - ��������� ������ �� ����� ���� (jMax)
        
        ' ���� ���������� ��� ���� ������������ ������� ���� ��� ����� ���� ������ �� �������
            If pMax = 0 Then sPrev = Prev Else sPrev = Prev(p)
        ' ���� ������������ ������� ���� ���� ��� � ��������� � ������������ �������������� ������
            If cMax = 0 Then sCode = Code: GoTo HandleMakeResult
        ' ���� ���� �������������� �������� ���� �������
            sCode = Code(0)          ' ���� ������ �� ������� �����
            rOld = rMax              ' ���������� ���������� ���������� ����������� (����� ��� ����������� ����� ������ ������������������)
            rMax = rOld * (cMax + 1) ' ���������� ����� ������������ ���������� �������������� �����������
        ' ��������� �������������� ������ � ������ �������
Dim c As Long: For c = 1 To cMax: Result = Result & Delim & Result: Next
HandleMakeResult:
        ' ���������� ��� ����������
            jBeg = 1: jEnd = 0  ' ������/����� �������� �������� ���� � �������������� ������
            r = 1: c = 0: p = 0 ' ������� ����������, ���� � ����������� ����
            Do
        ' �������� ������� ��������� �������� �������� ���������� (��������� ����������� + ����� �����������)
            ' ���� ��������� ����������� - ���� � �� ������� ���������� �����������
            ' ���� ���� ��������� ��� ����������� �� ������ - ����� ����� ������+1
                If rMax > 1 Then jEnd = VBA.InStr(jBeg, Result, Delim) - 1
                If jEnd <= 0 Then jEnd = Len(Result) + 1
            ' ���������� �������� ������������� ���� ����� MN/NM (66)
                If (sCode = sPrev) And (sCode <> "66") Then sCode = vbNullString
        ' ��������� r-��� ��������� � �������� ��� �� jMax
                If rMax = 1 Then sCode = Result & sCode Else sCode = Mid$(Result, jBeg, jEnd - jBeg + 1) & sCode
                sCode = Left$(sCode, jMax)
            ' ���� ����� ���������� ������� ����� ������ ����� ����� (i+n>=iMax)
            ' ��������� ������� ������� ���������� ������ �� ����������� ������(jMax)
                If (i + n) > iMax Then sCode = sCode & String(jMax - Len(sCode), "0")
        ' � ������������� ����������� ����� ��������� �������
            ' ��� ����� ������ �������� �������� Result �� jBeg � ������ �� �������������
            ' ����������� ���������� ���� � ������� ���������� sCode
            ' ���� ������� ��������� ����������� � ������� - ������� ������� ��������,
            '!��� ��������� ����� ��������� ������ �������� ���������!
            ' �� ����� ������ �� �����
        ' ���������� r-��� ���������.
                If rMax = 1 Then Result = sCode Else Result = VBA.Left$(Result, jBeg - 1) & sCode & VBA.Mid$(Result, jEnd + 1)
                ' �������� ������� ������ �������� �������� ����������
                jBeg = jBeg + Len(sCode) + Len(Delim) '- 1
' ����� ������������������� - ����������� ������� ��������� � ���������� ���������
    ' �.�. ����� �������� ���������� ����� ���������� ������ � ����� ������ -
    ' ��� �������� ����������� �������� ����� ���������:
        ' ��������������� ��������� �������� ����������,
        ' ��� ������� �������� ���� ���������� ��� �������� �����������,
        ' � ����� ��������� � ���������� �������� ��������
        
            ' ���� �������� ����� ������ (�������������) ������������������ -
            ' ����������� ������ �������� ���� (������� ����������) � �������� ��� ��������
                If r Mod rOld = 0 Then c = c + 1: c = IIf(c > cMax, 0, c): If cMax = 0 Then sCode = Code Else sCode = Code(c)    '
            ' ����������� ������ ����������� ���� (� ������� ����������) � �������� ��� ��������
                p = p + 1: p = IIf(p > pMax, 0, p): If pMax = 0 Then sPrev = Prev Else sPrev = Prev(p)
            ' ����������� ������ ����������, ���� ��������� ��� - �������
                r = r + 1: If r > rMax Then Exit Do
            Loop
            ' ���������� ����������(��) ���(�) � ������� �� �����
            Prev = Code: pMax = cMax: Exit Do
HandleNext: If n > 1 Then n = n - 1 Else Exit Do
        Loop
        i = i + n ' ��������� ������
    Loop
HandleExit:  SoundExDM = Result: Exit Function
HandleError: Result = String$(jMax, "0"): Err.Clear: Resume HandleExit
End Function

Public Function SimilarityLCS(ByVal Word1 As String, ByVal Word2 As String) As Double
' ���������� ����� ��������������������� - ��������� ������ ������� � ��������, �� �� ������
'-------------------------
' Longest Common Subsequence by George Brown - 11/14/04
' http://forums.devarticles.com/microsoft-access-development-49/approximate-string-matching-fyi-46951.html
'-------------------------
Dim i As Integer, j As Integer, m As Integer, n As Integer
Dim c() As Integer, b() As Integer, x$(), y$()
Dim Len1 As Integer, Len2 As Integer
Dim Result As Double

    Result = False
    On Error GoTo HandleError
    
    Word1 = UCase$(Trim$(Word1)): Word2 = UCase$(Trim$(Word2))
    Len1 = Len(Word1): Len2 = Len(Word2)
    
    n = IIf(Len1 <= Len2, Len1, Len2) 'n = min(Len1, Len2)
    m = IIf(Len1 >= Len2, Len1, Len2) 'm = max(Len1, Len2)
    
    ReDim x(m): ReDim y(m)
    ReDim c(m, m): ReDim b(m, m)
    '
    If Len1 > Len2 Then
        For i = 1 To Len1: x(i) = Mid$(Word1, i, 1): Next i
        For i = 1 To Len2: y(i) = Mid$(Word2, i, 1): Next i
    Else
        For i = 1 To Len2: x(i) = Mid$(Word2, i, 1): Next i
        For i = 1 To Len1: y(i) = Mid$(Word1, i, 1): Next i
    End If
    
    For i = 1 To m
        For j = 1 To n
            If x(i - 1) = y(j - 1) Then
                c(i, j) = c(i - 1, j - 1) + 1:  b(i, j) = 1
            ElseIf c(i - 1, j) >= c(i, j - 1) Then
                c(i, j) = c(i - 1, j):          b(i, j) = 2
            Else
                c(i, j) = c(i, j - 1):          b(i, j) = 3
            End If
        Next j
    Next i
    'If c(m, n)>0 Then Result = c(m, n)  ' ���������� ����� ���������������������
    If c(m, n) > 0 Then Result = 2 * c(m, n) / (Len1 + Len2) ' c(m, n) / IIf(Len1 > Len2, Len1, Len2)
    Erase x: Erase y: Erase c: Erase b
HandleExit:  SimilarityLCS = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function SimilarityLev(ByVal Word1 As String, ByVal Word2 As String) As Double
' ���������� ����������� - ��������� �������, �������� � ������ ��������
'-------------------------
' ��������: http://qaru.site/questions/390285/finding-similar-sounding-text-in-vba
' �����:    https://ru.wikibooks.org/wiki/����������_����������/����������_�����������#Visual_Basic_6.0
Dim Result As Double

    Result = False
    On Error GoTo HandleError
    If Word1 = vbNullString Or Word2 = vbNullString Then GoTo HandleExit
    Word1 = UCase$(Trim$(Word1)): Word2 = UCase$(Trim$(Word2))
    If Word1 = Word2 Then Result = 1: GoTo HandleExit
    Dim m As Integer: m = Len(Word1)
    Dim n As Integer: n = Len(Word2)
    Dim d As Variant: ReDim d(m, n)
    
    Dim i As Integer, j As Integer, k As Integer
    For i = 1 To m: d(i, 0) = i: Next i
    For j = 1 To n: d(0, j) = j: Next j
    
    Dim a(2), r As Integer, cost As Integer
    For i = 1 To m
        For j = 1 To n
            If Mid$(Word1, i, 1) = Mid$(Word2, j, 1) Then cost = 0 Else cost = 1
            a(0) = d(i - 1, j) + 1          ' ��������
            a(1) = d(i, j - 1) + 1          ' �������
            a(2) = d(i - 1, j - 1) + cost   ' �����������
            ' �������� ����������
            r = a(0)
            For k = 1 To 2
                If a(k) < r Then r = a(k)
            Next k
            d(i, j) = r
        Next j
    Next i
'    Result = d(m, n)                       ' ���������� �����������
    Result = 1 - d(m, n) / IIf(m > n, m, n) ' ���������� ������������� �� ����� ������
HandleExit:  SimilarityLev = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function SimilarityDL(ByVal Word1 As String, ByVal Word2 As String) As Double
' ���������� �������-����������� - ��������� �������, ��������, ������ � ������������ ���� �������� ��������
'-------------------------
' ��������: https://ru.wikipedia.org/wiki/����������_�������_�_�����������
'-------------------------
Dim Result As Double ': Result = False
    On Error GoTo HandleError
    If Word1 = vbNullString Or Word2 = vbNullString Then GoTo HandleExit
    Word1 = UCase$(Trim$(Word1)): Word2 = UCase$(Trim$(Word2))
    If Word1 = Word2 Then Result = 1: GoTo HandleExit
Dim m As Integer: m = Len(Word1)
Dim n As Integer: n = Len(Word2)
Dim d As Variant: ReDim d(m, n)

Dim i As Integer, j As Integer, k As Integer
    For i = 1 To m: d(i, 0) = i: Next i
    For j = 1 To n: d(0, j) = j: Next j
    
Dim a(3), r As Integer, cost As Integer
    For i = 1 To m
        For j = 1 To n
            If Mid$(Word1, i, 1) = Mid$(Word2, j, 1) Then cost = 0 Else cost = 1
            a(0) = d(i - 1, j) + 1          ' ��������
            a(1) = d(i, j - 1) + 1          ' �������
            a(2) = d(i - 1, j - 1) + cost   ' �����������
                                            ' ������������
            If i And j And Mid$(Word1, i + 1, 1) = Mid$(Word2, j, 1) _
                 And Mid$(Word1, i, 1) = Mid$(Word2, j + 1, 1) Then _
            a(3) = d(i - 2, j - 2) + cost Else a(3) = &H7FFF
            
            ' �������� ����������
            r = a(0)
            For k = 1 To 3
                If a(k) < r Then r = a(k)
            Next k
            d(i, j) = r
        Next j
    Next i
'    Result = d(m, n) ' ���������� �������-�����������
    Result = 1 - d(m, n) / IIf(m > n, m, n)
HandleExit:  SimilarityDL = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function SimilarityDice(ByVal Word1 As String, ByVal Word2 As String) As Double
' �������� �����
'-------------------------
' https://en.wikibooks.org/wiki/Algorithm_Implementation/Strings/Dice%27s_coefficient
'-------------------------
Const n = 2 ' ���������� ��������
Dim Result As Double ': Result = False
    On Error GoTo HandleError
    If Word1 = vbNullString Or Word2 = vbNullString Then GoTo HandleExit
    Word1 = UCase$(Trim$(Word1)): Word2 = UCase$(Trim$(Word2))
    If Word1 = Word2 Then Result = 1: GoTo HandleExit
Dim i As Integer, j As Integer
Dim iMax As Integer: iMax = Len(Word1) - (n - 1)    ' n-����� � Word1
Dim jMax As Integer: jMax = Len(Word2) - (n - 1)    ' n-����� � Word2
    If (iMax < 1) Or (jMax < 1) Then GoTo HandleExit    ' ����� ������ ����� n-������
' ��������� ��������� �������� n-����� ��� Word2 (�� ������� ������������ �������, �� �� ������� ������ �������)
Dim Col As New Collection, � As Long: For j = 1 To jMax: Col.Add j: Next j
' ������������ n-������ Word1 � Word2
Dim x As Integer: x = 0
    For i = 1 To iMax
        For j = 1 To Col.Count
' ��� ���������� n-�����:
    ' ����������� ������� ����������,
    ' ����������� ������ �� �������� n-����� Word2
    ' � ��������� � ��������� n-������ Word1
            If Mid$(Word1, i, n) = Mid$(Word2, Col(j), n) Then x = x + 1: Col.Remove j: Exit For
    Next j, i
    Result = 2 * x / (iMax + jMax)  ' ���������� �� ������� ���������� n-����� � ������
HandleExit:  SimilarityDice = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function SimilarityJaro(ByVal Word1 As String, ByVal Word2 As String) As Double
' �������� �����-�������� - ����������� ����� �������������� ��������������, ������� ���������� ��� ����, ����� �������� ���� ����� � ������
'-------------------------
' By: Ernanie F. Gregorio Jr. (from psc cd)
' ��������: https://planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=73978&lngWId=1
' https://blog.developpez.com/philben/p12207/vba-access/vba-distance-de-jaro-winkler
' https://ru.wikipedia.org/wiki/��������_�����_�_��������
'-------------------------
Dim Result As Double ': Result = False
    On Error GoTo HandleError
Dim iProximity As Integer ' set the number of adjacent characters to compare to
Dim i As Integer, x As Integer
Dim iFr As Integer, iTo As Integer
Dim iMatch As Integer, iTrans As Integer
Dim iJaro As Double
    Word1 = UCase$(Trim$(Word1)): Word2 = UCase$(Trim$(Word2))
    If Word1 = Word2 Then Result = 1: GoTo HandleExit ' check if the two words are the same
    If InStr(1, Word1, Word2) > 0 Then Result = 1: GoTo HandleExit
    ' compute for the proximity of character checking
    ' allows matching characters to be up to X number of characters away.
    If Len(Word1) >= Len(Word2) Then
        iProximity = (Len(Word1) / 2) - 1
    Else
        iProximity = (Len(Word2) / 2) - 1
    End If
    For i = 1 To Len(Word2)
    ' this is the index of the character to be compared to
        iTo = (i + iProximity) - 1
        ' get the left most side character based on the iProximity
        If i <= iProximity Then iFr = 1 Else iFr = i - iProximity + 1
        ' start the letter by letter comparison
        For x = iFr To iTo
            If Mid$(Word2, i, 1) = Mid$(Word1, x, 1) Then
                If i = x Then iMatch = iMatch + 1: GoTo HandleNext
                iMatch = iMatch + 1: iTrans = iTrans + 1: Exit For
            End If
        Next x
HandleNext:
    Next i
    
    iTrans = iTrans \ 2
    If iMatch <= 0 Then Result = 0: GoTo HandleExit
    x = 0
    For i = 1 To 4
        If Mid$(Word2, i, 1) = Mid$(Word1, i, 1) Then x = x + 1 Else Exit For
    Next i
    iJaro = ((iMatch / Len(Word2)) + (iMatch / Len(Word1)) + ((iMatch - iTrans) / iMatch)) / 3
    If x > 0 Then Result = iJaro + 0.1 * x * (1 - iJaro) Else Result = iJaro
HandleExit:  SimilarityJaro = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

'=========================
' ������� ��������� ������
'=========================
Public Function GenPassword( _
    Optional PassLen As Integer = 12, _
    Optional Symbols As String = vbNullString, _
    Optional bRepeats As Boolean = True, _
    Optional NewSeed As Boolean = True _
    ) As String
' ���������� "���������" ����� ��������� �����
'-------------------------
' PassLen - ����� ������
' Symbols - ���������� ����� ��������. ���� �� ����� ����� �����������
' bRepeats = True - ��������� ���������� ������� ������
' NewSeed = True - �������� Randomize ��� �������� ����� "���������" ������������������
'-------------------------
' ��� ��������� ������� � �����, Randomize �������� � ������� ������������������ ����� ~100-200 ���������� �������
' ������� ��� ��������� ���������� ������ ����� �������� True, ��� ��������� ����� ������� � ����� - False
'-------------------------
Const cMin = 0              ' ����������� ���������� ����� �������� ������ (0-������������)
Dim sMax As Byte: sMax = 3  ' ������������ ���������� ���������� �������� ������ (0-������������)
Const bDigits = True        ' ��������� �����
Const bLatin = True         ' ��������� ������� ���������� ��������
Const bCyrillic = False     ' ��������� ������� �������������� ��������
Const bOthers = False       ' ��������� ������� �� ���. ������
'Const bRepeats = True       ' ��������� ���������� ������� ������
Const sCase = 0             ' ������� �������� ����������� ������
                            ' 0-��������� ������� � ������� � ������ ���������
                            ' 1-������ � �������, 2-������ � ������
Dim Result As String ': Result = vbNullString
    On Error GoTo HandleError
    If PassLen < cMin Then Err.Raise vbObjectError + 512
    If PassLen < 1 Then GoTo HandleExit
    If Len(Symbols) = 0 Then
' ���� �� ������ - ��������� ������������������ ���������� ��������
'    ' ���� ��� ��������� "��������" ������� - ������ ������ ���������� ��������
'    ' ������������ �� ������ ������� ������ � ���������� � ������� ���������
'    ' ����������� ����� ����� �������. ������� �� ����� ����� ��� ������ HyphenateWord
' � ������� �����:
'    ' ����� ����������� "��������" ��������� ��������� ������ ������ ����� (N*Rnd)
'    ' ������ ������������ ����� ���������������� �����
'    ' ���-�� ����: Replace(cCur(Rnd*1E15),",","") - ��������� Rnd � ������ �� 19 ����
'    ' �� ������� ����� ����� ������������ ��������
        If bDigits Then Symbols = Symbols & c_strSymbDigits
        If bLatin Then If sCase = 0 Then Symbols = Symbols & UCase$(c_strSymbEngAll) & LCase$(c_strSymbEngAll) Else If sCase = 1 Then Symbols = Symbols & UCase$(c_strSymbEngAll) Else Symbols = Symbols & LCase$(c_strSymbEngAll)
        If bCyrillic Then If sCase = 0 Then Symbols = Symbols & UCase$(c_strSymbRusAll) & LCase$(c_strSymbRusAll) Else If sCase = 1 Then Symbols = Symbols & UCase$(c_strSymbRusAll) Else Symbols = Symbols & LCase$(c_strSymbRusAll)
        If bOthers Then Symbols = Symbols & c_strSymbDigits & c_strSymbMath & c_strSymbPunct & c_strSymbCommas & c_strSymbParenth & c_strSymbOthers
    End If
Dim sLen As Long: sLen = Len(Symbols)
    If sLen < 1 Then GoTo HandleExit
' ��������� ���������� ����� � ������������������ ���� �� ��������� ���� �� ������� �� ���������� ���������� ��������
    ' sType �.�. ���� ������ ����� ���� ����� ������ �������� ����� ������/��������� �����
Dim sTemp As String: sTemp = vbNullString
Dim sType As Integer
Dim i As Long
    For i = 1 To sLen
        sType = GetCharType(Mid$(Symbols, i, 1))
        If InStr(1, sTemp, sType) = 0 Then sTemp = sTemp & sType
    Next i
' ��������� �����������
    If sLen = 1 Then bRepeats = True ' ���� � ������ ���������� ����� ���� ������ ������� ������ ��������
    If Len(sTemp) <= 1 Then sMax = 0 ' ���� ��� ������� ������ ������ ���� ������� ������� �� ������� ���������� ��������
' ���������� ���������
Dim sChar As String * 1
Dim sPrev As Integer, sCount As Integer
    ' �������� ����� "���������" ������������������
    sType = 0: sPrev = -1
    If NewSeed Then Randomize Timer
    Do Until Len(Result) = PassLen
    ' �������� ������
HandleNewSymb: sChar = VBA.Mid$(Symbols, CLng((sLen - 1) * Rnd) + 1, 1)
    ' ��������� ������������ �������������� ����������� � ����������� ������:
        '1. �� ����� sMax ���������� �������� ������
        If sMax > 0 Then sType = GetCharType(sChar): If sType <> sPrev Then sPrev = sType: sCount = 1 Else If sCount >= sMax Then GoTo HandleNewSymb Else sCount = sCount + 1
        '2. ��������� � �������������� ������ ����������� ������ ���������� �������
        If Not bRepeats Then If LCase$(sChar) = LCase$(Right$(Result, 1)) Then GoTo HandleNewSymb
    ' ��������� �������������� ������
        Result = Result & sChar
    Loop
HandleExit:  GenPassword = Result: Exit Function
HandleError: Result = vbNullString
    Select Case Err.Number
    Case vbObjectError + 512: MsgBox "������� �������� ������." & vbCrLf & "������ ���� �� ������ " & cMin & " ��������.", vbOKOnly Or vbExclamation, "������!"
    End Select
    Err.Clear: Resume HandleExit
End Function
Public Function HyphenateWord( _
    ByVal Text As String, _
    Optional Delimiter As String = "�") As String
' ����������� �������� � ������
'-------------------------
' ��������: http://www.cyberforum.ru/vba/thread792944.html
' �������� � ���������� ��������� ��������� ����� https://habr.com/post/138088/
' ������� �� ������ ��������� ������ �������������� ������������� �������� �������� �����
Const cstrTemp = "xgg xgs xsg xss sggsg gsssssg gssssg gsssg sgsg gssg sggg sggs"
Dim sPattern() As String
Dim i As Long, j As Long, k As Long
Dim m As String, sText As String
    
    On Error GoTo HandleError
' ������ ���������� �������� �� �����: 0-�����(x), 1-�������(g), 2-���������(s)
    ' ������������ �������� - ��� ��������� ����������, ����� "�" ���� ������, � �� ���������
    ' ������ ��������������� �������� � ������ �������� �������
Dim sArr: sArr = Array(c_strSymbRusSign & c_strSymbEngSign, c_strSymbRusVowel & c_strSymbEngVowel, c_strSymbRusConson & c_strSymbEngConson)
' ������ �������������� ��������� � �����
Dim sTemp() As String: sTemp = Split(cstrTemp) 'Call xSplit(cstrTemp, sTemp)
' ������� ��������� ����� �������� - ����� ������� �������� (��. ������ ����) ����� �������� ���������� ��������� �����������
Dim sPos: sPos = Array(1, 1, 1, 1, 3, 3, 2, 2, 2, 2, 2, 2)

    sText = Text
' �������� ������� �������� ������ �� ������������� � �������� (x, g, s)
    For i = 1 To Len(Text)
        m = LCase$(Mid$(Text, i, 1))
        For j = 0 To UBound(sArr)
            If InStr(sArr(j), m) Then Mid$(Text, i, 1) = Mid$("xgs", j + 1, 1): Exit For
    Next j, i
    
' �������� �������� � ��������� ����������� � ������� ���������
' � ��������������� � �������� ������. ������ � ��������������� ������
' ����� ����� ��������� ��� ������������ ������� ���������
    For i = 0 To UBound(sTemp)
        j = 0
        Do
            k = InStr(j + 1, Text, sTemp(i))
            If k Then
                j = k + sPos(i)
                Text = VBA.Left$(Text, j - 1) & Delimiter & Mid$(Text, j)
                sText = VBA.Left$(sText, j - 1) & Delimiter & Mid$(sText, j)
            End If
        Loop While k
    Next i
HandleExit:  HyphenateWord = sText: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function NumToWords( _
    ByVal Number As String, _
    Optional ByVal NewCase As DeclineCase = DeclineCaseImen, _
    Optional ByRef NewNumb As DeclineNumb = DeclineNumbSingle, _
    Optional ByRef NewGend As DeclineGend = DeclineGendUndef, _
    Optional ByVal Animate As Boolean = False, _
    Optional ByRef NewType As NumeralType = NumeralOrdinal, _
    Optional ByVal Unit As String, Optional ByVal SubUnit As String, _
    Optional ByVal DecimalPlaces As Byte = 2, _
    Optional ByVal TranslateFrac As Boolean = False _
    ) As String
' �������������� ����� � ����� � ��������� �� �������
'-------------------------
' Number  - ������������� ����� (����� �����, ���������� ��� ����������� �����, ����� � ���������������� ����� �� ������������)
' NewCase - ����� ��������� (��","���","���","���","��","����)
' NewNumb - ����� ("��","��")
' NewGend - ��� ("�","�") ���� ������ ������� ��������� ������������ �� ���
' Animate - ������� ����, ��� ����������� ������ ��������� ���� �������� ��� �����������
' NewType - ��� ������������� (��������������","����������) ���� ������ ������� ����� - �.�. ������ ��������������
' Unit    - ������� ��������� - ����������� ����� ����� (��.�., ��.�.)
' SubUnit - ��������������� ������� ��������� - ����������� ������� ����� (��.�., ��.�.)
' DecimalPlaces - ����������� ��������������� ������� (���������� ������ ����� ������� � ���������� �����)
' TranslateFrac - (���� �� ������������) ���� True ����� ����������� � ����� ������� �����
'-------------------------
' ToDo: !!! ���������� ������� ��������� !!! - ��� ������� ���������� ���� ������������
'-------------------------
' ��� ��� ������ ����:
    ' 23,50     - �������� ��� ����� ��������� ������
    ' 23 1/2    - �������� ��� ����� ���� ������ �����
    '(�������)  - �������� ��� � ���� ������ �����
' SubUnit � DecimalPlaces � ���������� ����� ������ ��������������� ���� �����.
' �.�. ���� Unit = "�����"     � SubUnit = "�������",  DecimalPlaces �.�. = 2 (1/100 ���.),
'    � ���� Unit = "���������" � SubUnit = "�����",    DecimalPlaces �.�. = 3 (1/1000 ��)
' !!! �� ������ ����� �����: NumToWords(Day(Now), Unit:=LCase(Format(Now, "mmmm")), NewType:=NumeralCardinal, NewCase:=DeclineCaseImen, NewGend:=DeclineGendNeut)
'     �� ������ �� ������� ���-�� ����� "������ ������", � �������: "������ ������", ������ ��� ��������� ��������� ���������, ��� �� ������������� ���������� �������,
'     � ������ �� ��, ��� �� �������� ����� �����: "������ (�����) ������", ����� ���������� ���� �������� ����������� �����, �������� ����� (� ���.���.)
'     ��������� ����: NumToWords(Day(Now), NewType:=NumeralCardinal, NewCase:=DeclineCaseImen, NewGend:=DeclineGendNeut) & " " & DeclineWord(LCase(Format(Now, "mmmm")),DeclineCaseRod)
' !!! �������� ������������ ����������� ����������� ������� ����� ����������� ����� �� ������������
    ' ��������, - ���� DecimalPlaces �� ���������� (=0)
    ' ��� �������� ����: "0,1";"0,01" � "0,001" "�����"/"�������"
    ' ����� �������� ��� "���� ������ ���� �������"
' ����� ��������� ������ ���� ������ SubUnit � �������� ������ Unit:
    ' "������ ����� ������ ������" ??? - ���� �� ���� ��� ��� ��������� ����� ����������...
' ��� ����������� ������� ����� ��� ����������
    ' "������ � ���������� ����� �����" ??? - ��� ������ ��� ������ ��������� �������?
'-------------------------
Const �ShowNullInWhole = True   ' �������� ������� ����� �����
'Const �ShowOnesInDemom = True   ' �������� ��������� ������� � ����������� (���� ������������","��������)
Const �EmptyWholeUnit = "�����" ' ����������� ����� ����� ���� �� ������ ������� ���������
Const �EmptyWholeUnit2 = "�"    ' ��������� ����������� ����� � ������� ����� ����� ���� �� ������ ������� ���������

Const cWhlDelim = " "   ' Chr(32)  - ����������� �����/������� ����� ����������� �����
Const cNatDelim = "/"   ' Chr(47)  - ����������� ���������/����������� ����������� �����
Dim cDecDelim As String * 1: cDecDelim = p_GetLocaleInfo(LOCALE_SDECIMAL)   ' Chr(44)  - ����������� �����/������� ����� ���������� �����
Dim cPosDelim As String * 1: cPosDelim = p_GetLocaleInfo(LOCALE_STHOUSAND)  ' Chr(160) - ����������� �������� ����� �����

Dim strWhole As String, strNomin As String, strDenom As String
Dim bolWhole As Boolean, bolNomin As Boolean, bolDenom As Boolean
Dim bytStep As Byte     ' ������� ��� ���������: 1-����� �����,2-���������,3-�����������,0-�� ����������
Dim Result As String
    Result = vbNullString
    bolWhole = False: bolNomin = False: bolDenom = False
    On Error GoTo HandleError
    Number = Trim$(Number): Unit = Trim$(Unit): SubUnit = Trim$(SubUnit)
' ����� �� ����� � ������� �����
    ' ������� ���������� ����� �� ������������ �������� � ��������
    ' Replace ������ CLng ������ ��� �������� ������ -
    ' ������������ ����� � ��� ����� ��������� �� ����������� ���� Long
Dim tmpSymPos As Long: tmpSymPos = Nz(InStrRev(Number, cDecDelim), 0)
    If (tmpSymPos > 0) Then
' ���������� �����
    ' ����������� ���������� ����� ������� ���� �� ������ ����������� ������� �����
        bolDenom = Len(SubUnit) = 0
        strWhole = VBA.Left$(Number, tmpSymPos - 1)
        strNomin = VBA.Mid$(Number, tmpSymPos + Len(cDecDelim), Len(Number) - tmpSymPos)
    ' ��������� ����������� ���������� �����
    ' ���������� ���������� ������ � ����������� �.�. �� 1 �� ���� ���������� ��������� 10^33. ����������� �� ������� ������� �� i=37 (������) � �����
        If bolDenom Then
        ' ���� ����������� ������� �� ������� - ������� �� ���������� �������� � ���������
            DecimalPlaces = Len(strNomin)
        Else
        ' ���� ����������� ������� �������
            Select Case Len(strNomin)
            Case Is < DecimalPlaces
            ' ���� ����������� ��������������� ������� ������ ����� ������ ����� ������� -
                ' ��������� ��������� ������ �� ������������ ������������
                strNomin = strNomin & String(DecimalPlaces - Len(strNomin), "0")
            Case Is > DecimalPlaces:
            ' ���� ����������� ��������������� ������� ������ ����� ������ ����� ������� -
                ' ����� �������� ����� ����� ����� � ���������
                Result = NumToWords(strWhole, NewType:=NewType, NewCase:=NewCase, Unit:=Unit, Animate:=Animate, DecimalPlaces:=0)     ' ����� �����
                ' ��������� ���������� � ���������� ����� �� ���������� ���������� ��������
                strNomin = Left$(strNomin, DecimalPlaces) & cDecDelim & Mid$(strNomin, DecimalPlaces + 1)
                Result = Result & " " & NumToWords(strNomin, NewType:=NewType, NewCase:=NewCase, Unit:=SubUnit, Animate:=Animate, DecimalPlaces:=0)  ' ������� �����
                GoTo HandleExit
            End Select
        End If
        strDenom = 1 & String(DecimalPlaces, "0")
    Else
        tmpSymPos = Nz(InStrRev(Number, cNatDelim), 0)
        If tmpSymPos > 0 Then
' ����������� �����
            bolDenom = True ' ������ ������� ��� ��� ����������� �����
            SubUnit = vbNullString ' ��� ����������� ����� ��������������� ������� ��������� �� ����� ������
            ' �����������
            strDenom = VBA.Mid$(Number, tmpSymPos + Len(cNatDelim))
            Number = VBA.Left$(Number, tmpSymPos - 1): tmpSymPos = Nz(InStrRev(Number, cWhlDelim), 0)
            ' ���������
            strNomin = VBA.Mid$(Number, tmpSymPos + Len(cWhlDelim), Len(Number) - tmpSymPos)
            ' ����� �����
            If tmpSymPos > 0 Then strWhole = VBA.Left$(Number, tmpSymPos - 1)
        Else
' ����� ����� ��� �� �����
            strWhole = Number: strNomin = 0: strDenom = 1
        End If
    End If
' ���������� ������������� ������ ������ �����
    ' ��������� ������� ���� �� ��������
    ' ����������� ������� ���� �� ��������
        ' � ����� ������ ��� �������� (����������� ����� ��� ���������� ��� ��������������� �������)
    ' ����� ����� ������� ���� �� ��������
        ' ��� ���� ����� ����� ������� ����� ����� �����
        ' ��� ���� �� ����� ����� ��������� (��������� ������)
    On Error Resume Next
    bolNomin = p_NumType(strNomin) > 0
    bolDenom = p_NumType(strDenom) > 0 And bolDenom
    bolWhole = p_NumType(strWhole) > 0 Or �ShowNullInWhole
    On Error GoTo HandleError
    
' ���������� ��� �������� ����� ������ �����
Dim strNumb As String   ' ����������� ����� ����� (�����/���������/�����������)
Dim strWord As String   ' ����� ������� ����� ������������ �����
Dim strDelim As String  ' ����������� ���� (������ = Chr(32)) � ����������
' ���������� ��� �������� ���������� �������� �������� �����
Dim bytNumb As Byte     ' ��� ������������ ����� (��� ���������, ��. p_NumType)
Dim intTrip As Integer  ' ���������� �������� �������� ����������� ����� �����
Dim bytTrip As Byte     ' ���������� ����� �������� ����� (� �����, ������� � 0)
Dim bolNull As Boolean  ' ������� ������� ���������� ��������. (����� ��� ����������� ���������)
    ' True  - �������� ���������� ������ ��� ����� ������� �������� ��� �� ��������
    '   �.�. ��� True ��� ������������� ������� ������� ���������,
    '   ����� � ������ ����������� ��� True ��� �������� >0 �������� ��������� ���� ���� ��������, ��� 1 ������ �������
    ' False - ������ ��������� ������� ��� �������, ��� ����� ������� ���������
' ���������� ��� �������� ���������� ��������� strWord
Dim tmpType  As NumeralType, tmpCase As DeclineCase, tmpNumb As DeclineNumb, tmpGend As DeclineGend  ' ��������� (��������������� ������������ ��� ���������)
'Dim NewGend As DeclineGend: NewGend = DeclineGendUndef ' ��� �����������
' �������� ������ ������
    ' ��������� � ��� �������: ��� ����� �����, ��� ��������� � ��� �����������
    bytStep = 2                     ' �������� � �����������
    Do
        bytTrip = 0                 ' ���������� ����� �������� (������)
        bolNull = True              ' ������� ������� ������
        strWord = vbNullString      ' ����������� ��.���/������� ��� ����� �������� �����
    ' ���� ��� ����� ����� 0 � ��� ����� � ������� ������ �������� ������� �������� ����� �����
    ' ��.���. Unit ������ ��������� ����� ����� ����� ���� ���� ������� �����, ��� ���������� ����� � ����� SubUnit
    ' ����� �������� ����� ����������� � ���������� ������������ ������� (1)
    ' �.�. "���� ����� ��������� ������", �� "���� �����(�) ��������� ����� �����" � "���� �����(�) ���� ������ �����"
    '      "������ ������ ��������� ������", �� "������ �����(�) ��������� ����� �����" � "������ �����(�) ���� ������ �����"
    ' ���  "���� ������ ������� �������", �� "���� �����(�) ������� �������� ������" ��� "���� �����(�) ���� ������ ������"
    '      "������ ���� ������� �������", �� "������ �����(�) ������� �������� ������" ��� "������ �����(�) ���� ������ ������"
        Select Case bytStep
        Case 0: If bolWhole Then strNumb = strWhole: strWord = IIf((Len(Unit) = 0) Or bolDenom Or (bolNomin And (Len(SubUnit) = 0)), IIf(bolNomin, IIf(NewType = NumeralCardinal, �EmptyWholeUnit2, �EmptyWholeUnit), vbNullString), Unit): GoTo HandleBegin
        Case 1: If bolNomin Then strNumb = strNomin: strWord = IIf(bolDenom, vbNullString, SubUnit): GoTo HandleBegin
        Case 2: If bolDenom Then strNumb = strDenom: strWord = Unit: GoTo HandleBegin
        Case Else: Exit Do
        End Select
        GoTo HandleNextPart
HandleBegin:
' ������ ��������� ����� �����
    ' ���������� ������ ����������� ����� �����
        If Len(strNumb) = 0 Then strNumb = 0        ' ������ ������ = 0
        ' ������� �� ������������ �������� � ��������
        strNumb = Replace$(Replace(strNumb, cPosDelim, vbNullString), Space(1), vbNullString)
        ' ������� ���� ������� (����� ����� ������� ���������� �� �����)
        tmpSymPos = 1: Do While VBA.Mid$(strNumb, tmpSymPos, 1) = "0": tmpSymPos = tmpSymPos + 1: Loop: If (tmpSymPos > 1) Then If (tmpSymPos > Len(strNumb)) Then strNumb = "0" Else strNumb = VBA.Mid$(strNumb, tmpSymPos)
    ' ���������� ��� ����� (��.p_NumType) ���������� ��� ����������� ���������
        Select Case bytStep
        Case 0: bytNumb = p_NumType(strNumb):   tmpType = NewType
        Case 1: bytNumb = p_NumType(strNumb):   If bolDenom Then tmpType = NumeralOrdinal Else tmpType = NewType
        Case 2: bytNumb = p_NumType(strNomin):  tmpType = NumeralCardinal  ': tmpCase = DeclineCaseImen
        End Select
        
' ������ ��������� �����
        intTrip = Abs(CInt(VBA.Right$(strNumb, 3))) ' ���� ������� (0) ������� �����
HandleUnits:
' ����������� ������� ���������, �������� ��� ��� � ��������� ��� � �������������� ������ ����� �����
    ' ���� ������ ��� �� ���� (��� ������ �� ������ �������) � strWord ������ ������� ��������� ����� ��� �����
        tmpGend = DeclineGendUndef
        If Len(strWord) > 0 Then
    ' �������� ����������� ������� ��������� � ���������� ��� ���
        ' ����� �����.  ��.� ��� ����� �� 1 (����� 11) � ����� �� 2-4 (����� 12-14) � ��.,���. � ���.�., ��������� - �� ��.�.
        ' ����� �����.  NewCase, ����� ��. � ���.�. ��� ����� �� ��������������� �� 1 (����� 11) ��� � ���.�.
            tmpCase = NewCase: tmpNumb = DeclineNumbPlural
            If (bytStep = 2) Or (tmpType = NumeralCardinal) Then
                tmpNumb = DeclineNumbSingle: If (bytStep = 2) Then tmpCase = DeclineCaseRod
            Else
                If (intTrip = 0) Or (bytTrip <> 0) Then tmpCase = DeclineCaseRod
                Select Case bytNumb
                Case 1:     If (bytTrip = 0) Then tmpNumb = DeclineNumbSingle
                Case 2:     If (tmpCase = DeclineCaseImen) Or (tmpCase = DeclineCaseVin) Then tmpCase = DeclineCaseRod: If (bytTrip = 0) Then tmpNumb = DeclineNumbSingle
                Case Else:  If (tmpCase = DeclineCaseImen) Or (tmpCase = DeclineCaseVin) Then tmpCase = DeclineCaseRod
                End Select
            End If
            ' �������� ��� ����������� - ��������������
            If bytNumb = 2 Then If p_GetWordSpeechPartType(strWord) = SpeechPartTypeAdject Then tmpNumb = DeclineNumbPlural
        ' �������� � ������ � ���������
            strWord = DeclineWord(strWord, tmpCase, tmpNumb, tmpGend, Animate): If tmpGend <> DeclineGendUndef Then NewGend = tmpGend
            If Len(Result) > 0 Then Result = strWord & strDelim & Result Else Result = strWord
        End If
        ' ��� ������������� - �� ���� ��.���������/�������/���� �� ��������� - ���.��� (�������� ����), ���.��� (���� ����� ��� �����)
        If tmpGend = DeclineGendUndef Then If NewGend = DeclineGendUndef Then tmpGend = DeclineGendMale Else tmpGend = NewGend
        'If tmpGend = DeclineGendUndef Then If NewGend = DeclineGendUndef Then tmpGend = DeclineGendFem Else tmpGend = NewGend
        Do
' ����� ����� ����� �� �������� �������� (���������� ����� ������ �� 3).
' ������ ��� ������� � ��������� � intTrip, � bytTrip - ���������� ����� �������� ��������
HandleThousands:
' ����������� �������, �������� ��� ��� � ��������� ��� � �������������� ������ ����� �����
            strDelim = Space(1)         ' ����������� ��������� ����� - ������ (����� ��������� ����������)
    ' ���������� ������ �������� (����� ��������. 0 � ������� - �������� 0 �����)
            If intTrip = 0 Then If (Not bolWhole) Or (bytStep <> 0) Or (bytNumb <> 0) Then GoTo HandleNextTriplet
    ' � �������� �������� ��� ������ ����������� �������, ��� - ����������
            If bytTrip = 0 Then GoTo HandleDigits
    ' ��� ������� ��������� �������� ����������� �������, �������� ��� � ���������� ��� ���
        ' ����������� ������� - ����������, ����� ��������������� �������
        ' ��� ����������� ������� ������������ ��� ��������� ������� ��������� �������������
    ' �������� ����������� ������� � ���������� ��� ���
        ' ����� �����.  ��.� ��� ����� �� 1 (����� 11) � ����� �� 2-4 (����� 12-14) � ��.,���. � ���.�., ��������� - �� ��.�.
        ' ����� �����.  NewCase, ����� ��.,���. � ���.�. ��� ����� �� ��������������� �� 1 (����� 11) ��� � ���.�.
            strWord = vbNullString
            tmpCase = NewCase: tmpNumb = DeclineNumbPlural
            If ((bytStep = 2) Or (NewType = NumeralCardinal)) And Not bolNull Then tmpCase = DeclineCaseImen
            If (tmpType = NumeralCardinal) And (bytStep = 0) Then
                tmpNumb = DeclineNumbSingle
            Else
                Select Case bytNumb
                Case 1:     tmpNumb = DeclineNumbSingle
                Case 2:     If (bytStep <> 2) Then If (tmpCase = DeclineCaseImen) Or (tmpCase = DeclineCaseVin) Then tmpCase = DeclineCaseRod: tmpNumb = DeclineNumbSingle
                Case Else:  If (tmpCase = DeclineCaseImen) Or (tmpCase = DeclineCaseVin) Then tmpCase = DeclineCaseRod
                End Select
            End If
        ' �������� � ������ � ���������
            strWord = p_NumDecline(intTrip, bytTrip, tmpCase, tmpNumb, tmpGend, tmpType, Animate)
            If Len(strWord) = 0 Then GoTo HandleDigits
            If Len(Result) > 0 Then Result = strWord & strDelim & Result Else Result = strWord
        ' ���� ����������� ������� ������ ����������, �.�. ���� ������� ������ (�������) ������
            ' ������ �������� ������� ���������������
            ' ������ ����������� ������ �.�. ���������� �� -��������/-��������� � �.�. ������� ������
            ' ��� ����������� - � ���������� �������� �������� ������� � ����������� �� ���� ����� � �����������
            If tmpType = NumeralCardinal Then tmpType = NumeralOrdinal: strDelim = vbNullString: If bytStep = 2 Then bytNumb = p_NumType(intTrip)
HandleDigits:
'    ' ���������� ������ �������� (����� ��������. 0 � ������� - �������� 0 �����)
'            If intTrip = 0 Then If (Not bolWhole) Or (bytStep <> 0) Or (bytNumb <> 0) Then GoTo HandleNextTriplet
            Do
' ���������� ��������������� �������� ��������:
    ' �����, ������� (����� �������), ������ ������� (10-19), ������� � ����
    ' ��������� ������ �������� � ��������� �������
    ' �������� ����� ������������ �������� � ������ ���� ����������� �������/������� ���������
                tmpCase = NewCase: If tmpCase = DeclineCaseUndef Then tmpCase = DeclineCaseImen
                Select Case bytNumb
                Case 0, 1:  tmpNumb = DeclineNumbSingle
                Case Else:  tmpNumb = DeclineNumbPlural
                End Select
                If tmpType = NumeralCardinal Then
                ' ������ ����� � ������ �������� �������� ��� ���������� � ��������� (NewCase) ������,
                ' ����� ��. � ���., �� - � ���.�. ��� ����� �� �� 1 � 2-4
                    Select Case bytStep
                    Case 0: tmpNumb = NewNumb
                    'Case 1: If Not bolDenom Then tmpNumb = DeclineNumbSingle
                    Case 2: If (bytNumb <> 1 And bytNumb <> 2) And ((tmpCase = DeclineCaseImen) Or (tmpCase = DeclineCaseVin)) Then tmpCase = DeclineCaseRod
                    End Select
                ElseIf ((NewType = NumeralCardinal) And (bytStep = 0)) Or (bytStep = 2) Then
                ' ������� ������� � ������ ��������
                ' � ������� �������� (����� ������� ����������) �������� � ��.�.
                    tmpCase = DeclineCaseImen
                    If bolNull And bytTrip > 0 Then
                ' ������� �������� (� ������ ���������) ��������:
                '   ��� ������� �������� � �.�., ����.: ����- � ���-
                '   -�������� -���������� � �.�., ���� ������ �����������
                        If (intTrip Mod 10 = 1) And ((intTrip Mod 100) \ 10 <> 1) Then
                            tmpGend = DeclineGendNeut ': tmpNumb = DeclineNumbSingle
                        ElseIf intTrip <> 100 Then
                            tmpCase = DeclineCaseRod
                        End If
                    End If
                End If
        ' �������� � ������ � ���������
                strWord = p_NumDecline(intTrip, , tmpCase, tmpNumb, tmpGend, tmpType, Animate, NumbRest:=intTrip)
                If Len(strWord) > 0 Then If Len(Result) > 0 Then Result = strWord & strDelim & Result Else Result = strWord
            ' �������������� ��� ������������� (������ ������ ������� ���������� ������������� �.�. ����������)
                tmpNumb = DeclineNumbUndef ': tmpGend = DeclineGendUndef
                If tmpType = NumeralCardinal Then tmpType = NumeralOrdinal
            Loop While intTrip > 0
            bolNull = False
HandleNextTriplet:
        ' �������� ����������� �������
            If Len(strNumb) > 2 Then strNumb = VBA.Left$(strNumb, Len(strNumb) - 3) Else strNumb = vbNullString
        ' ��������� ���� �� ��������� �������� ��������
            If Len(strNumb) = 0 Then Exit Do
        ' ����� ��������� ������� ����� ����� � (���� ��� �����) ��� ���
            intTrip = Abs(CInt(VBA.Right$(strNumb, 3))): If Not (bolNull And (tmpType = NumeralCardinal)) Then bytNumb = p_NumType(intTrip)
        ' ����������� ������� ����������� ���������
            bytTrip = bytTrip + 1
        Loop 'While Len(strNumb)>0
HandleNextPart:
        ' ��������� � ��������� ��������� ����� �����
        If bytStep = 0 Then Exit Do Else bytStep = bytStep - 1
    Loop While bytStep >= 0
'    ' ��������� �����
'    If VBA.Left$(Number, 1) = "-" And Len(Result)>0 Then Result = "����� " & Result
HandleExit:  NumToWords = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function

Private Function p_NumType(ByVal Numb As Variant, Optional InNomin As Boolean = False) As Byte
' ��� NumToWords. �������� ��� ����� 0-999
'-------------------------
' InNomin - ���� ��������������� ��� �������� ��������� ��� ����������� ��������� �����������
' � ���� ������ ����� �� 2 ������ ���������� � ��.�. �� ���� ��������� - �� ��.�.
' ����������:
'   0 - ���� Numb = 0
'   1 - ���� Numb = xx1 � <>x11         (��� InNomin=True ����� xx2 � <>x12)
'   2 - ���� Numb = xx2-4 � <>x12-x14   (��� InNomin=True ����� xx2 � <>x12)
'  255 - ��� ���������
'-------------------------
' ��� ����������� ��������� ����� ��������� ����� ��������
'-------------------------
Dim Result As Byte
    Result = 255
    On Error GoTo HandleError
    If Numb = 0 Then Result = 0:   GoTo HandleExit      ' 0
    If Numb < 0 Then Err.Raise vbObjectError + 512
    If Numb > 999 Then Numb = VBA.Right$(Numb, 3)
    If (Numb \ 10) Mod 10 = 1 Then GoTo HandleExit      ' 10-19
    Select Case Numb Mod 10
    Case 1:         Result = 1
    Case 2:         Result = IIf(InNomin, 1, 2)
    Case 3 To 4:    Result = 2
    End Select
HandleExit:  p_NumType = Result: Exit Function
HandleError: Result = 255: Err.Clear: Resume HandleExit
End Function


Public Function DeclineWord( _
    ByVal Word As String, _
    Optional NewCase As DeclineCase = DeclineCaseRod, _
    Optional NewNumb As DeclineNumb = DeclineNumbSingle, _
    Optional ByRef NewGend As DeclineGend = DeclineGendUndef, _
    Optional ByRef Animate As Boolean = False, _
    Optional IsFio As Byte = 0, _
    Optional ByRef SymbCase As Integer = 0, Optional ByRef Template As String _
    ) As String
' ��������� �����. �������� ������ �������
'-------------------------
' ������ ������� �������� ��������� ���, �����, ����� ����, �������� ���������
' ����� �� ��������� ������������/�������������� (��-�� �����, � ���������, ����������� �������� � ��.�. ���.�.)
' ��� ������ ����������� ����������� Morpher'�� � http://morpher.ru
' ��� Padej'�� http://www.delphikingdom.com/asp/viewitem.asp?catalogid=412
' Word - ��������������� � ������������ ����� ������������ ������
' NewCase - ����� ("�","�","�","�","�")
' NewNumb - ����� ("��","��")
' NewGend - ��� ("�","�")
' Animate - ������� (������������","��������������)
' IsFio - ������� ��� (0-�� ���, 1-�������, 2-���, 3-��������)
' SymbCase, Template - ��������� �������� �������� ��������� ����� � ������
'-------------------------
' v.1.0.1       : 06.07.2019 - ��������� ������������ �������� � ��������� �������
'-------------------------
Dim WordBeg As String, WordEnd As String
Dim WordType As SpeechPartType
Dim sChar As String '* 1
Dim i As Long, iMax As Long
Dim Result As String
' ���� ���������� �� ����� �������� ���������� �������
    On Error GoTo HandleError
    If NewCase = DeclineCaseUndef Then NewCase = DeclineCaseImen Else If NewCase > DeclineCasePred Or NewCase < DeclineCaseImen Then Err.Raise vbObjectError + 512
    If NewNumb = DeclineNumbUndef Then NewNumb = DeclineNumbSingle
    If IsFio Then Animate = True ' ���������� ������������ �� �����, ������ ��� - ���������� ������������
' �������� ������ �������� �������� �����
    Word = Trim$(Word): If SymbCase = 0 Then SymbCase = p_GetSymbCase(Word, Template)
    
    Result = LCase$(Word)
' ��������� ����������
'' ������ ����������
'    Case ""
'        Select Case NewCase
'        Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = ""   ' ��.�. ��.�    (���/���)       Nominative
'        Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "", "")  ' �.�. ��/��.�  (����/����)     Genitive
'        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "", "")  ' �.�. ��/��.�  (����/����)     Dative
'        Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "", "")  ' �.�. ��/��.�  (����/���)      Accusative
'        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "", "")  ' �.�. ��/��.�  (���/���)       Ablative
'        Case DeclineCasePred: WordEnd = Choose(NewNumb, "", "")  ' �.�. ��/��.�  (� ���/� ���)   Prepositional
'        'Case Else 'DeclineCaseUndef = 0
'        End Select
'        i = Len(Result)-2: GoTo HandleExit ' i = ���-�� ������ �� ������ ����� � ������� ����������� ���������
' ���������� �������� � ��. ������������ �����
    Select Case Result
    Case "�", "�", "�", "�", "�", "�", "�", "��", "��", "��", "��", "��", "��", "��", "��", "��", _
         "���", "���", "���", "���", "���", "���", "�����", "�����", "�����", "���������", "�����", _
         "������", "������", "�����", "�����", "�����", "�����", "������", "������", "�������", "�����", _
         "������", "���������", "�����", "������", "�����", "���������", "����������"
        i = Len(Result): GoTo HandleExit
    End Select
' ������ ����������
    If IsFio Then
    ' �� ���������� ������� ��:
    Select Case Right(Result, 1)
    Case "�": i = Len(Result): GoTo HandleExit
    End Select
    End If
    Select Case Result
    ' ������ ������� (� -> �)
    Case "���", "��", "��": If NewCase <> DeclineCaseImen Or NewNumb = DeclineNumbPlural Then Mid$(Result, 2, 1) = "�"
    ' ��������� ������� (2-� � �����)
    Case "�����", "����", "�����", "���", "���", "���", "�������": If Not (NewCase = DeclineCaseImen Or (Not Animate And NewCase = DeclineCaseVin)) Or NewNumb = DeclineNumbPlural Then i = Len(Result) - 2: Result = Left$(Result, i) & Mid(Result, i + 2)
    ' ��������� ������� (3-� � �����)
    Case "����", "�����", "����", "�����": If Not (NewCase = DeclineCaseImen Or (Not Animate And NewCase = DeclineCaseVin)) Or NewNumb = DeclineNumbPlural Then i = Len(Result) - 3: Result = Left$(Result, i) & Mid(Result, i + 2)
    ' ��������� ������� (������ ������)
    Case "����": NewNumb = DeclineNumbSingle: If Not (NewCase = DeclineCaseImen Or NewCase = DeclineCaseVin Or NewCase = DeclineCaseTvor) Then i = Len(Result) - 3: Result = Left$(Result, i) & Mid(Result, i + 2)
    ' �����������
    Case "�": Animate = True
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = Choose(NewNumb, "�", "��")
        Case DeclineCaseRod, DeclineCaseVin: WordEnd = Choose(NewNumb, "����", "���")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "���", "���")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "����", "����")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, "���", "���")
        End Select
        i = Len(Result) - 1: GoTo HandleExit
    Case "��": Animate = True
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = Choose(NewNumb, "��", "��")
        Case DeclineCaseRod, DeclineCaseVin: WordEnd = Choose(NewNumb, "����", "���")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "����", "���")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "�����", "����")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, "����", "���")
        End Select
        i = Len(Result) - 2: GoTo HandleExit
    Case "��": Animate = True
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = "�"
        Case DeclineCaseRod, DeclineCaseVin: WordEnd = "��"
        Case DeclineCaseDat:  WordEnd = "��"
        Case DeclineCaseTvor: WordEnd = "���"
        Case DeclineCasePred: WordEnd = "��"
        End Select
        i = Len(Result) - 1: GoTo HandleExit
    Case "��": 'Animate = True
        If NewGend = DeclineGendUndef Then NewGend = DeclineGendMale
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = Choose(NewNumb, Choose(NewGend, "��", "���", "���"), "���")
        Case DeclineCaseRod, DeclineCaseVin: WordEnd = Choose(NewNumb, Choose(NewGend, "����", "��", "���"), "��")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, Choose(NewGend, "����", "���", "���"), "���")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, Choose(NewGend, "���", "���", "����"), "����")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, Choose(NewGend, "��", "���", "���"), "���")
        End Select
        '"�" ����������� ����� ��������� ��� ��������� ���� �� �����
        i = Len(Result) - 2: GoTo HandleExit
    Case "��": Animate = False: NewGend = DeclineGendNeut
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = Choose(NewNumb, "�", "�")
        Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "���", "��")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "���", "��")
        Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "��", "��")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, "��", "��")
        End Select
        i = Len(Result) - 1: GoTo HandleExit
    Case "���": Animate = False: NewGend = DeclineGendNeut
        NewGend = DeclineGendNeut
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = Choose(NewNumb, "�", "�")
        Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "���", "��")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "���", "��")
        Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "��", "��")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, "��", "��")
        End Select
        i = Len(Result) - 1: GoTo HandleExit
    Case "�����": Animate = False: NewGend = DeclineGendMale
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: If NewNumb = DeclineNumbPlural Then WordEnd = "�"
        Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "��")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "�", "��")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
        End Select
        i = Len(Result): GoTo HandleExit
    End Select
    
'    SymbCase = p_GetSymbCase(Word, Template) ' �������� ������ ��������
    iMax = Len(Result): WordEnd = vbNullString: i = iMax ' ������� ������ ���������
' ����������� ����� ���� (����� ��������������)
    WordType = p_GetWordSpeechPartType(LCase$(Word))
    Select Case WordType
    Case SpeechPartTypeNoun, _
         SpeechPartTypeAdject   ' ��������������� � �������������� �������� �� �������� ����
    Case SpeechPartTypeNumeral  ' ������������ �������� ��������
        Result = p_NumDecline(Word, , NewCase, NewNumb, NewGend, , Animate): i = Len(Result): GoTo HandleExit
    Case Else: GoTo HandleExit  ' ��� ��������� (����������������) - ����������
'    Case SpeechPartTypePronoun ' �����������
'    Case SpeechPartTypeVerb    ' �������
'    Case SpeechPartTypePretext ' ��������
'    Case SpeechPartTypeUndef   ' �����������
    End Select
' ���������� ������, ��������� ����� � ����� ����� ����������
    ' ������� ������ �>� ���� ������ ����� ����������� ���������
    ' ��� ��������� -��, -�� � -��, -��
    'Result = Replace(Result, "�", "�")
    Call p_GetWordParts(Result, WordBeg, WordEnd, Template)
    i = iMax - Len(WordEnd)                     ' ������� ������ ���������
' �������������� ������������: ���� ����� ����� < 3 � ��������� ������� - ����� �� ��������
    If i < 3 And InStr(1, c_strSymbRusVowel, WordEnd) Then GoTo HandleExit
    sChar = LCase$(Mid$(Result, i, 1))          ' ����� ����� ����������
' ����������� ���� (����� ��������������)
    If NewGend = DeclineGendUndef Then _
       NewGend = p_GetWordGender(Word, WordEnd) 'And NewNumb <> DeclineNumbPlural
    If NewGend = DeclineGendNeut Then Animate = False ' ��.� ������� ��������������
' ��������� ��������� � ���������
    Select Case WordEnd
    'Case "��"
    '' ��.�����
    'Case "�"
    '' ��.�����
    Case "�"
        Select Case sChar
        Case "�", "�", "�"
    ' -��, -��, -��
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "�"
            Case DeclineCaseRod: If NewNumb <> 2 Then WordEnd = "�": GoTo HandleExit
                    Select Case Mid$(WordBeg, i - 1, 1)
                    Case "�":                   WordEnd = "�" & sChar: i = i - 2 '-���, -���
                    Case "�", "�", "�", "�":    WordEnd = "�" & sChar: i = i - 1 '-���, -���
                    Case Else:                  WordEnd = ""
                    End Select
            Case DeclineCaseDat: WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseVin: If NewNumb <> DeclineNumbPlural Then WordEnd = "�": GoTo HandleExit
                    If Animate Then
                    Select Case Mid$(WordBeg, i - 1, 1)
                    Case "�":                   WordEnd = "�" & sChar: i = i - 2 '-���, -���
                    Case "�", "�", "�", "�":    WordEnd = "�" & sChar: i = i - 1 '-���, -���
                    Case Else:                  WordEnd = ""
                    End Select
                    Else: WordEnd = "�"
                    End If
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
            End Select
        Case "�", "�"
    ' -��, -��
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "�"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "�", IIf(Animate, "��", "�"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
            End Select
        Case "�"
    ' -��
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "�"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "�", IIf(Animate, "", "�"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
            End Select
        Case "�", "�"
            If IsFio = 1 And NewGend = 2 Then
    ' ������� ������� �� -��, -��
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "�"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "��", "��")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "��", "��")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "��", "��")
            End Select
    ' ������ �� -��, -��
            Else
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "�"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseDat, DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "�", IIf(Animate, "", "�"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���")
            End Select
            End If
        Case Else
    ' ������ �� -�
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "�"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "�", IIf(Animate, "", "�"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
            End Select
        End Select
    Case "�"
        Select Case sChar
        Case "�"
    ' �� -��
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "���"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "���", "��")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "���", "����")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "�", IIf(Animate, "��", "���"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "����", "�����")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "���", "����")
            End Select
        Case Else
    ' ������ �� -�
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "�"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "�", IIf(Animate, "", "�"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
            End Select
        End Select
    Case "�"
    ' �� -�
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "�"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "�", IIf(Animate, "", "�"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
            End Select
    Case "�"
    ' �� -�
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "�"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "�", IIf(Animate, "��", "�"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
            End Select
    Case "��"
    ' �� -��
            If InStr(1, "���", sChar) Then sChar = "�" Else sChar = "�"
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = sChar & "�"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "���", sChar & "�")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "���", sChar & "�")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "��", sChar & IIf(Animate, "�", "�")) '
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, sChar & "�", sChar & "��")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "��", sChar & "�")
            End Select
    Case "��"
    ' �� -��
            If sChar <> "�" Then sChar = "�" Else sChar = "�"
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = sChar & "�"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "���", sChar & "�")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "���", sChar & "�")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "��", IIf(Animate, "��", sChar & "�"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, sChar & "�", sChar & "��")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "��", sChar & "�")
            End Select
    Case "��"
    ' �� -��
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "��"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "��", "��")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "��", "���")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "��", IIf(Animate, "��", "��"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "���", "����")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "��", "���")
            End Select
    Case "��"
    ' �� -��
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "��"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "��", "��")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "��", "���")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "��", IIf(Animate, "��", "��"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "���", "����")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "��", "���")
            End Select
    Case "��"
    ' �� -��
        If InStr(1, "�����", sChar) Then
            If NewNumb = DeclineNumbPlural And sChar = "�" Then sChar = "�" Else sChar = "�"
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "��"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "��", sChar & "�")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "��", sChar & "�")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "��", IIf(Animate, sChar & "�", "��"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", sChar & "��")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "��", sChar & "�")
            End Select
        Else
            If NewNumb = DeclineNumbPlural Then
            If InStr(1, "���", sChar) Then sChar = "�" Else sChar = "�"
            End If
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = sChar & "�"
            Case DeclineCaseRod, DeclineCasePred: WordEnd = Choose(NewNumb, "��", sChar & "�")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "��", sChar & "�")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "��", sChar & IIf(Animate, "�", "�"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", sChar & "��")
            End Select
        End If
    Case "��", "��"
        i = i + 2
        If IsFio = 1 Then
    ' ������� �� -��, -��
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, vbNullString, "�")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
            End Select
        Else
    ' ������ �� -��, -��
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = WordEnd & "�"
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "�", "")
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
            End Select
        End If
    Case "��", "��"
            If (NewNumb = DeclineNumbPlural) And (WordEnd = "��") And (sChar = "�" Or sChar = "�") Then sChar = "�" Else sChar = "�"
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, Choose(NewGend, WordEnd, "��", "��"), sChar & "�")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "���", "��"), sChar & "�")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "���", "��"), sChar & "�")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, IIf(Animate, "���", WordEnd), "��"), sChar & "�")
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "��", "��"), sChar & "��")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "��", "��"), sChar & "�")
            End Select
    Case "��"
        Select Case sChar
        Case "�"    ' -���
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, Choose(NewGend, "��", "��", "��"), "��")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "���", "��"), "��")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "���", "��"), "��")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "���", "��"), "��")
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "��", "��"), "���")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "��", "��"), "��")
            End Select
        Case "�"    ' -���
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, Choose(NewGend, "��", "��", "��"), "��")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "��", "��"), "���")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "��", "��"), "���")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "��", "��"), "���")
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "���", "���"), "����")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "��", "��"), "���")
            End Select
        Case "�"    ' -���
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, Choose(NewGend, "��", "��", "��"), "��")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "����", "���"), "���")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "����", "���"), "���")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "����", "���"), "���")
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "���", "���"), "����")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "���", "���"), "���")
            End Select
        Case "�", "�", "�"   ' -��� � -���
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, Choose(NewGend, "��", IIf(sChar = "�", "�", "�") & "�", "��"), "��")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "���", "��"), "��")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "���", "��"), "��")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, Choose(NewGend, IIf(Animate, "���", "��"), IIf(sChar = "�", "�", "�") & "�", "��"), IIf(Animate, "��", "��"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "��", "��"), "���")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, IIf(NewGend <> DeclineGendFem, "��", "��"), "��")
            End Select
        End Select
    Case "��", "��", "��", "��", "��", "��" ', "��"
        If sChar <> "�" Then
    ' ������ �� -��, -��, -��, -��, -��, -�� ', -��
            i = i + 1
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, "�", "�")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "�")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(Animate, "�", "�"), IIf(Animate, "��", "�"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "�")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
            End Select
        End If
    Case "��"
    ' �� -��
            i = i + 2
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, vbNullString, "�")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(Animate, "�", ""), IIf(Animate, "��", "�"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
            End Select
    Case "��", "��"
    ' �� -��
            If Left$(WordEnd, 1) = "�" Then sChar = "�" Else sChar = ""
            Select Case NewCase
            Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = IIf(Left$(WordEnd, 1) = "�", "�", "") & "��"
            Case DeclineCaseRod:  WordEnd = sChar & Choose(NewNumb, "��", "���")
            Case DeclineCaseDat:  WordEnd = sChar & Choose(NewNumb, "��", "���")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(Animate, sChar & "��", WordEnd), IIf(Animate, sChar & "���", sChar & "��"))
            Case DeclineCaseTvor: WordEnd = sChar & Choose(NewNumb, "���", "����")
            Case DeclineCasePred: WordEnd = sChar & Choose(NewNumb, "��", "���")
            End Select
            i = Len(Result) - 2
'    Case "��" '>��� � �.�.
'    Case "��", "��"
'    Case "��", "��"
    Case Else
    ' ��� ���������
        WordBeg = LCase$(Result): WordEnd = vbNullString: i = iMax
        sChar = Right$(WordBeg, 1)
        Select Case LCase$(Right$(Template, 1))
        Case "s" ' ������������� �� ���������
        ' ������� ������� �� ��������� ��������� ��� ����
            If IsFio = 1 And NewGend = 2 Then GoTo HandleExit
            Select Case sChar
            Case "�"
        ' �� -�
                Select Case NewCase
                Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "�":
                Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "��")
                Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "�", "��")
                Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(Animate, "�", WordEnd), IIf(Animate, "��", "�"))
                Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���")
                Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
                End Select
                i = i - 1
            Case "�", "�", "�", "�"
        ' �� -�,-�,-�,-�
                Select Case NewCase
                Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "�"
                Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "��")
                Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "�", "��")
                Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(Animate, "�", WordEnd), IIf(Animate, "��", "�"))
                Case DeclineCaseTvor: WordEnd = Choose(IsFio + 1, Choose(NewNumb, "��", "���"), Choose(NewNumb, "��", "���"), Choose(NewNumb, "��", "���"), Choose(NewNumb, "��", "���"))
                Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
                End Select
            Case "�"
      ' �� -�
                Select Case NewCase
                Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = "�"
                Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "��")
                Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "�", "��")
                Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(Animate, "�", ""), IIf(Animate, "��", "�"))
                Case DeclineCaseTvor: WordEnd = Choose(IsFio + 1, Choose(NewNumb, "��", "���"), Choose(NewNumb, "��", "���"), Choose(NewNumb, "��", "���"), Choose(NewNumb, "��", "���"))
                Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
                End Select
            Case Else
      ' �� ��������� ���������
                Select Case NewCase
                Case DeclineCaseImen: If NewNumb = DeclineNumbPlural Then WordEnd = IIf(InStr(1, "�������", sChar), "�", "�")
                Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "��")
                Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "�", "��")
                Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(Animate, "�", ""), IIf(Animate, "��", "�")) ' "�"",""��"
                Case DeclineCaseTvor: WordEnd = Choose(IsFio + 1, Choose(NewNumb, "��", "���"), Choose(NewNumb, "��", "���"), Choose(NewNumb, "��", "���"), Choose(NewNumb, "��", "���"))
                Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
                End Select
            End Select
        Case "x"
    ' ������������� �� -�
        i = i - 1
        Select Case VBA.Left$(VBA.Right$(Result, 2), 1)
        Case "�", "�", "�":
            If IsFio = 1 Then GoTo HandleExit
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, "�", "�")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "�", "��")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "�", IIf(Animate, "��", "�")) ' ���� ���� -> �������
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
            End Select
        Case Else:
            If IsFio = 1 And NewGend <> 1 Then GoTo HandleExit
            Select Case NewCase
            Case DeclineCaseImen: WordEnd = Choose(NewNumb, "�", "�")
            Case DeclineCaseRod:  WordEnd = Choose(NewNumb, IIf(NewGend = 1, "�", "�"), "��")
            Case DeclineCaseDat:  WordEnd = Choose(NewNumb, IIf(NewGend = 1, "�", "�"), "��")
            Case DeclineCaseVin:  WordEnd = Choose(NewNumb, IIf(Animate, "�", "�"), IIf(Animate, "��", "�"))
            Case DeclineCaseTvor: WordEnd = Choose(NewNumb, IIf(NewGend = 1, "��", "��"), "���")
            Case DeclineCasePred: WordEnd = Choose(NewNumb, IIf(NewGend = 1, "�", "�"), "��")
            End Select
        End Select
        End Select
    End Select
HandleExit:
    Result = p_SetSymbCase(Left$(Result, i) & WordEnd, SymbCase, Template)
    DeclineWord = Result
    Exit Function
HandleError: i = iMax: WordEnd = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function DeclineWords( _
    ByVal Words As String, _
    Optional NewCase As DeclineCase = DeclineCaseRod, _
    Optional NewNumb As DeclineNumb = DeclineNumbSingle, _
    Optional ByVal NewGend As DeclineGend = DeclineGendUndef, _
    Optional ByRef Animate As Boolean = False, _
    Optional ByRef IsFio As Boolean = False, _
    Optional SkipWords As String _
    ) As String
' ��������� ����. �������� ������ �������
'-------------------------
' Word - ��������������� � ������������ ����� ������������ ������
' NewCase - ����� ("�","�","�","�","�","�")
' NewNumb - ����� ("��","��")
' NewGend - ��� ("�","�")
' Animate - ������� ������������� / ���������������
' IsFIO - ������� ���
' SkipWords - ������ ������� ����, ������� ���������� ���������� ��� ��������� ��������������
'   ������� � 1, �� ����������� ����� ",", �������� ����� "-"
'   ��������: "2,4-5,7-" - ��� ��������� ����� ��������� ��� ����� ����� 1,3 � 6
'-------------------------
' ��� ������ ����������� ����������� Morpher'�� � http://morpher.ru
' ��� Padej'�� http://www.delphikingdom.com/asp/viewitem.asp?catalogid=412
' ��� ������������� ������� � ���������� ���������� ��� ����
'-------------------------
' v.1.0.2       : 06.07.2019 - � SkipWords ��������� ������������ �������� ���������
' v.1.0.1       : 05.07.2019 - ���������� ������ ��� ���������� ������ ������ ������ � Replace �������� � ������ ��� ������� �������� � ������
'-------------------------
Const cstrListDelim = ",", cstrDiapDelim = "-"
Dim strWord As String, strTail As String, aWords() As String, aSkip() As String
Dim i As Long, iMax As Long ' ������� � ������
Dim j As Long, jMin As Long ' ����� ����������� �����
Dim n As Long, nMin As Long ' ����� �������� � ������ ����������� ����
Dim s As String, sArr() As String ' ������� ������ ����������� ����
Dim sNum As Byte, sMin As Byte ' ������� ������ �������� ������
Dim S1 As Long, S2 As Long  ' ������� �������� ������
Dim tmpGender As DeclineGend
Dim Result As String
'
    On Error GoTo HandleError
    Result = vbNullString
    strTail = Trim$(Words): iMax = Len(strTail)
    Call Tokenize(Words, aWords): jMin = LBound(aWords): j = UBound(aWords)
    ' �������� ������ ���� ������� ���������� ����������
    n = -1: SkipWords = Trim$(SkipWords)
    If Len(SkipWords) > 0 Then aSkip = Split(SkipWords, cstrListDelim): nMin = LBound(aSkip): n = UBound(aSkip) + 1
    'If Len(SkipWords) > 0 Then Call xSplit(SkipWords, aSkip, cstrListDelim): nMin = LBound(aSkip): n = UBound(aSkip) + 1
    S1 = j - jMin + 1: S2 = 0
    Do While j >= jMin
    ' ���������� ����� �� ����� � ������
    ' ��������� �������������� ����� �������� ���������
        If n < nMin Then GoTo HandleText
        If S1 <= j - jMin + 1 And S1 <> 0 And S2 <> 0 Then GoTo HandleText
HandleDiap:
        n = n - 1: If n < nMin Then S1 = 0: S2 = 0: GoTo HandleText
    ' �������� ������� ������ ��������� ��������
        ' ��������� � �������� ������ �������� ������� ������� ���������
        s = Trim$(aSkip(n))
        sArr() = Split(s, cstrDiapDelim) 'Call xSplit(S, sArr(), cstrDiapDelim)
        sMin = LBound(sArr): sNum = UBound(sArr) - sMin + 1
        Select Case sNum
        Case 1
        ' "s1" - � ��������� ���� �������� �������
            If IsNumeric(sArr(sMin)) Then
                If S1 <> 0 Then S2 = CLng(Trim$(sArr(sMin)))
                S1 = CLng(Trim$(sArr(sMin)))
            End If
        Case 2 '
            If IsNumeric(sArr(sMin)) And IsNumeric(sArr(sMin + 1)) Then
        ' "s1-s2" -  � ��������� ��� �������� ��������
            ' ����� ��������� �������� ������� � ������ �������
                If S1 <> 0 Then S2 = CLng(Trim$(sArr(sMin + 1)))
                S1 = CLng(Trim$(sArr(sMin)))
            ElseIf IsNumeric(sArr(sMin)) Then
        ' "s1-" -    � ��������� ���� �������� �������, ������� ������� �������
            ' ����� ������� ������� ����� ���������� ������
                If S1 <> 0 Then S2 = S1
                S1 = CLng(Trim$(sArr(sMin)))
            ElseIf IsNumeric(sArr(sMin + 1)) Then
        ' "-s2" -    � ��������� ���� �������� �������, ������ ������� �������
            ' ���� ������ ������� ���������� ��������� � ������
            ' ���� �� ������ �������� � ��������� ������� ��� �� ��������� ������
                S2 = CLng(Trim$(sArr(sMin + 1)))
                If n = nMin Then S1 = 1 Else S1 = 0: GoTo HandleDiap
            Else
        ' �������� �����
                Stop
            End If
        Case Else: Stop ' "-n-","--n" � �.�. � ��������� ��� ��������� ��� ������ 2 ???
        End Select
HandleText:
    ' ���� ���������� �����
        strWord = aWords(j)
    ' ���� ������ ����������� ����� � ������ (� ����� ������)
        i = InStrRev(strTail, strWord)
    ' ��������� � ������ ������ ����������� �� �������� ������
        Result = Right$(strTail, iMax - i - Len(strWord) + 1) & Result
    ' �������� �������� ������ �� ������ ����������� �����
        iMax = i - 1: strTail = Left$(strTail, iMax)
    ' ��������� ����� ����� �� ������ ��������
        Select Case j - jMin + 1 ' ������ �������� �����
        Case S1 To S2:  ' �������� � ������� ��������� �������� - ����������
            'newWord = strWord
        Case Else:      ' �� �������� - �������� ����� � ������
            tmpGender = NewGend
            'newWord = DeclineWord(strWord, NewCase, NewNumb, tmpGender, Animate, IsFio:=IIf(IsFio, j - jMin + 1, 0))
            strWord = DeclineWord(strWord, NewCase, NewNumb, tmpGender, Animate, IsFio:=IIf(IsFio, j - jMin + 1, 0))
        End Select
    ' ��������� � ������ ������ ����� ������������ ����� ��������� ���������
        Result = strWord & Result 'Result = newWord & Result
    ' ��������� � ���������� �����
HandleNext:
        j = j - 1
    Loop
    ' ��������� ���������� �����������
    Result = strTail & Result
    Erase aWords: Erase aSkip
HandleExit:  DeclineWords = Result: Exit Function
HandleError: i = iMax: Err.Clear: Resume HandleExit ': WordEnd = vbNullString
End Function

Private Function p_NumDecline( _
    ByRef Numb As Variant, Optional Triplet As Byte = 0, _
    Optional ByVal NewCase As DeclineCase = DeclineCaseRod, _
    Optional ByVal NewNumb As DeclineNumb = DeclineNumbSingle, _
    Optional ByRef NewGend As DeclineGend = DeclineGendUndef, _
    Optional ByVal NewNumeralType As NumeralType = NumeralOrdinal, _
    Optional ByRef Animate As Boolean = False, _
    Optional ByRef SymbCase As Integer = 0, Optional ByRef Template As String, _
    Optional ByRef NumbRest As Integer = 0 _
    ) As String
' ��������������� ��� ��������� ���� �������� ������������: 0-9,10-19,x0,x00,1000,100000...
'-------------------------
' Numb - ������ ���������� ����� ����� (��� �������� � ��.��������), ���� ���� �� ��������� �������� ����� (��.p_NumWordsArray)
' NewCase - ����� ("�","�","�","�","�","�")
' NewNumb - ����� ("��","��")
' NewGend - ��� ("�","�")
' NewNumeralType - ��� ������������� ("���","���") - ��������� ������ ��� ������������
' Animate - ������� ��������������, ����� ��� ����������� ��������� � ���.�.
' SymbCase, Template - ��������� �������� �������� ��������� ����� � ������
' Triplet - ����� �������������� ��������. ���� =0 ��� >len(Numb)\3-1 - ������������� ������� �������
' NumbRest - ���� ������� �������� ��� ����� �������� �������,- ���������� ������������� ������� �������������� ��������
'-------------------------
' �������� � ��������� ������� ����� ��������� ���������� �������� ����� ����� ������ ����� (�������� ��� ������ �� NumToWords)
' ������� ���������� �������������� ����� � ��������� ����, � ����� ���������
' ��������� ������ �������:
'   Numb=123, Triplet=0 - ��������� ����� 123       >> Result="���",NumbRest=23
'   Numb=123, Triplet=1 - ��������� ����� 123000    >> Result="������",NumbRest=123
'   Numb=123000 - ��������� ����� 123000            >> Result="������",Triplet=1,NumbRest=123
'-------------------------
Dim WordWhole As String, WordBeg As String, WordEnd As String
Dim i As Long, iMax As Long
Dim Result As String

    On Error GoTo HandleError
    If NewGend = DeclineGendNeut Then Animate = False
    If NewNumeralType = NumeralUndef Then NewNumeralType = NumeralOrdinal
    If IsNumeric(Numb) Then
' ���� �������� ����� ������� - �������� ������ � ������� � ��������������� �����
    On Error Resume Next
        NumbRest = CLng(Numb): i = Err.Number: Err.Clear
    On Error GoTo HandleError
        If i = 0 And NumbRest < 1000 And Triplet = 0 Then
        ' ��� ������, ����� � ��������� 0..999 � ����� �������� �� �����
'    ' ������ �������� ����� ������� (�����>�������>�������)
'            Select Case NumbRest
'            Case 0 To 19:       i = NumbRest:               NumbRest = 0
'            Case 20 To 99:      i = 18 + NumbRest \ 10:     NumbRest = NumbRest Mod 10
'            Case 100 To 999:    i = 27 + NumbRest \ 100:    NumbRest = NumbRest Mod 100
'            End Select
    ' ������ �������� ������ ������ (�������>�������>�����)
            i = NumbRest
            If i > 0 Then ' �� ������ �������
                i = NumbRest Mod 100 ' ������� �����
            If i = 0 Then                                   ' x00 - ����� (x=1-9)
                i = 27 + NumbRest \ 100: NumbRest = 0
            ElseIf i >= 20 And (i Mod 10 = 0) Then     ' xy0 - ������� (y=2-9)
                i = 18 + i \ 10:  NumbRest = 100 * (NumbRest \ 100)
            ElseIf i < 20 Then                              ' x1z - ������ ������� (z=0-9)
                NumbRest = 100 * (NumbRest \ 100)
            Else                                            ' xyz - ������ ������� (z=1-9)
                i = i Mod 10: NumbRest = 10 * (NumbRest \ 10)
            End If
            End If
        Else
        ' ����� >1000 ( >Long, >1000 � <Long, ������ �������)
            iMax = Len(Numb)
            If Triplet = 0 Then Triplet = (iMax - 1) \ 3    ' ����� �������� ��������
            i = iMax - 3 * Triplet: If i < 1 Then i = iMax  '
            NumbRest = Right$(Left$(Numb, i), 3)                 ' ��������� �������� ����� ��������
            i = 36 + Triplet  ' ������ ���������� �������� (������� ��������) � ������� ��������� ��������
        End If
    ' �������� �������� ��������� ��������� ��������������� ����� �����
        Result = p_NumWordsArray(i): WordWhole = Result: iMax = Len(Result)
    Else
' ���� �������� ����� - �������� ������ � �������
        Result = Trim$(Numb): WordWhole = LCase$(Result) ': WordWhole = Replace(WordWhole, "�", "�")
        If SymbCase = 0 Then SymbCase = p_GetSymbCase(Result, Template)
        i = 0: iMax = Len(Result)
        Do Until p_NumWordsArray(i) = WordWhole
            If i <= UBound(p_NumWordsArray) Then i = i + 1 Else Err.Raise 6
        Loop
    End If
' �������������� ����� � ���������
    If NewCase > DeclineCasePred Then Err.Raise vbObjectError + 512
    If NewNumb = DeclineNumbUndef Then NewNumb = DeclineNumbSingle
    If NewGend = DeclineGendUndef Then NewGend = DeclineGendMale
    If NewNumeralType = NumeralOrdinal Then
    ' �������������� ������������
        Select Case i
        Case 0
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = Choose(NewNumb, "�", "�")
        Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "��")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "�", "��")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
        End Select
        i = Len(Result) - 1: GoTo HandleExit
        Case 1
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = Choose(NewNumb, Choose(NewGend, "��", "��", "��"), "��")
        Case DeclineCaseRod:  WordEnd = Choose(NewNumb, Choose(NewGend, "����", "���", "����"), "���")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, Choose(NewGend, "����", "���", "����"), "���")
        Case DeclineCaseVin:  WordEnd = Choose(NewNumb, Choose(NewGend, IIf(Animate, "����", "��"), "��", "��"), "���")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, Choose(NewGend, "���", "���", "���"), "����")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, Choose(NewGend, "���", "���", "���"), "���")
        End Select
        i = Len(Result) - 2: GoTo HandleExit
        Case 2
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = Choose(NewGend, "�", "�", "�")
        Case DeclineCaseRod, DeclineCasePred: WordEnd = "��"
        Case DeclineCaseVin:  WordEnd = IIf(Animate, "��", Choose(NewGend, "�", "�", "�"))
        Case DeclineCaseDat: WordEnd = "��"
        Case DeclineCaseTvor: WordEnd = "���"
        End Select
        i = Len(Result) - 1: GoTo HandleExit
        Case 3
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = "�"
        Case DeclineCaseRod, DeclineCasePred: WordEnd = "��"
        Case DeclineCaseVin: WordEnd = IIf(Animate, "��", "�")
        Case DeclineCaseDat: WordEnd = "��"
        Case DeclineCaseTvor: WordEnd = "���"
        End Select
        i = Len(Result) - 1: GoTo HandleExit
        Case 4
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = "�"
        Case DeclineCaseRod, DeclineCasePred: WordEnd = "��"
        Case DeclineCaseVin: WordEnd = IIf(Animate, "��", "�")
        Case DeclineCaseDat: WordEnd = "��"
        Case DeclineCaseTvor: WordEnd = "���"
        End Select
        i = Len(Result) - 1: GoTo HandleExit
        Case 5 To 7, 9, 10 To 21 '5-7,9,10-20,30
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = "�"
        Case DeclineCaseRod, DeclineCaseDat, DeclineCasePred: WordEnd = "�"
        Case DeclineCaseTvor: WordEnd = "��"
        End Select
        i = Len(Result) - 1: GoTo HandleExit
        Case 8
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = "���"
        Case DeclineCaseRod, DeclineCaseDat, DeclineCasePred: WordEnd = "���"
        Case DeclineCaseTvor: WordEnd = "����"
        End Select
        i = Len(Result) - 3: GoTo HandleExit
        Case 22 '40
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = ""
        Case DeclineCaseRod, DeclineCaseDat, DeclineCaseTvor, DeclineCasePred: WordEnd = "�"
        End Select
        i = Len(Result): GoTo HandleExit
        Case 23 To 26 '50-80 (��� 80 - ����������� -�-/-�-)
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = "������"
        Case DeclineCaseRod, DeclineCaseDat, DeclineCasePred: WordEnd = "�������":  If i = 26 Then Mid$(Result, 4, 1) = "�"
        Case DeclineCaseTvor: WordEnd = "���������":      If i = 26 Then Mid$(Result, 4, 1) = "�"
        End Select
        i = Len(Result) - 6: GoTo HandleExit
        Case 27, 28 '90,100
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = "�"
        Case DeclineCaseRod, DeclineCaseDat, DeclineCaseTvor, DeclineCasePred: WordEnd = "�"
        End Select
        i = Len(Result) - 1: GoTo HandleExit
        Case 29 '200
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = "����"
        Case DeclineCaseRod:  WordEnd = "�����"
        Case DeclineCaseDat:  WordEnd = "������"
        Case DeclineCaseTvor: WordEnd = "��������"
        Case DeclineCasePred: WordEnd = "������"
        End Select
        i = Len(Result) - 4: GoTo HandleExit
        Case 30, 31 '300,400
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = VBA.Right$(Result, 4)
        Case DeclineCaseRod:  WordEnd = "�����"
        Case DeclineCaseDat:  WordEnd = "������"
        Case DeclineCaseTvor: WordEnd = IIf(i = 30, "�", "�") & "�������"
        Case DeclineCasePred: WordEnd = "������"
        End Select
        i = Len(Result) - 4: GoTo HandleExit
        Case 32 To 36 '500-900 (��� 800 - ����������� -�-/-�-)
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = "����"
        Case DeclineCaseRod:  WordEnd = "����":    If i = 35 Then Mid$(Result, 4, 1) = "�"
        Case DeclineCaseDat:  WordEnd = "�����":   If i = 35 Then Mid$(Result, 4, 1) = "�"
        Case DeclineCaseTvor: WordEnd = "�������": If i = 35 Then Mid$(Result, 4, 1) = "�"
        Case DeclineCasePred: WordEnd = "�����":   If i = 35 Then Mid$(Result, 4, 1) = "�"
        End Select
        i = Len(Result) - 4: GoTo HandleExit
        ' � �������� 1E3,1E6 � �.�. ��������� ���������� DeclineWord
        Case 37 '1000
        NewGend = DeclineGendFem
        Select Case NewCase
        Case DeclineCaseImen: WordEnd = Choose(NewNumb, "�", "�")
        Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "�", "��")
        Case DeclineCaseVin:  WordEnd = Choose(NewNumb, "�", "�")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���") ', "��", "���")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
        End Select
        i = Len(Result) - 1: GoTo HandleExit
        Case 38 To 47 '1E6,1E9,...
        NewGend = DeclineGendMale
        Select Case NewCase
        Case DeclineCaseImen, DeclineCaseVin: WordEnd = Choose(NewNumb, "", "�")
        Case DeclineCaseRod:  WordEnd = Choose(NewNumb, "�", "��")
        Case DeclineCaseDat:  WordEnd = Choose(NewNumb, "�", "��")
        Case DeclineCaseTvor: WordEnd = Choose(NewNumb, "��", "���")
        Case DeclineCasePred: WordEnd = Choose(NewNumb, "�", "��")
        End Select
        i = Len(Result): GoTo HandleExit
        Case Else: Err.Raise 6
        End Select
        Result = p_SetSymbCase(Left$(Result, i) & WordEnd, SymbCase, Template)
    Else
    ' ���������� ������������
        ' ����������� � ���������� � �������� ��� ��������������
        Select Case i
        Case 0: WordEnd = "�������":    i = 0: Result = vbNullString
        Case 1: WordEnd = "������":     i = 0: Result = vbNullString
        Case 2: WordEnd = "������":     i = 0: Result = vbNullString
        Case 3: WordEnd = "����":       i = 2
        Case 4: WordEnd = "�����":     i = 3
        Case 6: WordEnd = "��":         i = Len(Result) - 1
        Case 7: WordEnd = "�����":      i = 2
        Case 8: WordEnd = "����":       i = 3
        Case 5, 9 To 21, 27: WordEnd = "��": i = Len(Result) - 1        ' 5,9-19,20,30,90
        Case 22: WordEnd = "����":      i = Len(Result)                 ' 40
        Case 23 To 26: WordEnd = "��":  i = Len(Result)                 ' 50-80
        Case 28: WordEnd = "����":      i = 1                           ' 100
        Case 29: WordEnd = "�������":   i = 2                           ' 200
        Case 30, 31: WordEnd = "�������":   i = Len(Result) - 4         ' 300,400
        Case 32 To 34, 36: WordEnd = "������":   i = Len(Result) - 4    ' 500-700,900
        Case 35: WordEnd = "��������":  i = 3                           ' 800
        Case 37: WordEnd = "���":       i = Len(Result) - 1             ' 1000
        Case 38 To 47:  WordEnd = "���": i = Len(Result)                ' 10^6, 10^9 etc
        Case Else: Err.Raise 6
        End Select
        Result = p_SetSymbCase(Left$(Result, i) & WordEnd, SymbCase, Template)
    ' �������� ��� ��������������
        Result = DeclineWord(Result, NewCase, NewNumb, NewGend, Animate)  ', SymbCase:=SymbCase, Template:=Template)
        WordEnd = vbNullString: i = Len(Result)
    End If
HandleExit:  p_NumDecline = Left$(Result, i) & WordEnd: Exit Function
HandleError: i = iMax: WordEnd = vbNullString: Err.Clear: Resume HandleExit
End Function
Private Function p_NumWordsArray()
' ������ ��������� �������� ��� ������������
'-------------------------
' i = 00..09    -   �������      x,     ��� x=0-9
' i = 10..19    -   1� �������  1x,     ��� x=0-9
' i = 20..27    -   �������     x0,     ��� x=2-9
' i = 28..36    -   �����       x00,    ��� x=1-9
' i = 37..47    -   ������ � �. 10^(3*x), ��� x=1-11
'-------------------------
On Error Resume Next
Static arrData(), iMin As Long: iMin = LBound(arrData): If Err Then Err.Clear Else p_NumWordsArray = arrData: Exit Function
    arrData = Array( _
        "����", "����", "���", "���", "������", "����", "�����", "����", "������", "������", _
        "������", "�����������", "����������", "����������", "������������", "����������", "�����������", "����������", "������������", "������������", _
        "��������", "��������", "�����", "���������", "����������", "���������", "�����������", "���������", _
        "���", "������", "������", "���������", "�������", "��������", "�������", "���������", "���������", _
        "������", "�������", "��������", "��������", "�����������", "�����������", "�����������", "����������", "���������", "���������", "���������")
HandleExit: p_NumWordsArray = arrData
End Function
Private Function p_GetWordTemplate(ByVal Word As String, Optional CheckCase As Boolean = False) As String
' ������� �� ������ ����� ��� ������������ �������
'-------------------------
' CheckCase - ���������� ����� �� ��� �������� ������� ����������� ������� �������
'-------------------------
Dim Result As String: Result = vbNullString
    On Error GoTo HandleError
Dim sArr:   sArr = Array(c_strSymbRusSign & c_strSymbEngSign, c_strSymbRusVowel & c_strSymbEngVowel, c_strSymbRusConson & c_strSymbEngConson)
' �������� ������� �������� ������ �� ������������� � �������� (x, g, s)
Dim i As Long, j As Long, m As String, s As String
    s = "xgs" ' ������� ��� �����������
    For i = 1 To Len(Word)
    ' ������� �������� �����
        m = Mid$(Word, i, 1)
        ' ���� ��������� ������� ������� ������ ������� �������� ����������� � ������������ � ��������� �������
        If CheckCase Then If m = LCase(m) Then s = LCase(s) Else s = UCase(s)
        m = LCase$(m)
        For j = 0 To UBound(sArr)
        ' ������� ��������� ������� �����������
            If InStr(sArr(j), m) Then Mid$(Word, i, 1) = Mid$(s, j + 1, 1): Exit For
    Next j, i
    Result = Word
HandleExit:  p_GetWordTemplate = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function

Private Function p_GetWordParts(ByVal Word As String, _
    Optional ByRef WordBeg As String, Optional ByRef WordEnd As String, _
    Optional ByRef Template As String) As Boolean
' �������� � ����� ������ � ��������� (�������)
'-------------------------
Dim Result As Boolean ' Result = False
    On Error GoTo HandleError
    If Len(Template) = 0 Then Template = p_GetWordTemplate(Word)
    Dim i As Long: i = Len(Template)
' ���������� ������� ��� ��� ���� ������ �� ������ (�����
' ���������, ������� � ����� ����� �����) ��������� � �����
    i = InStrRev(LCase$(Template), "s", i - 1)
    WordBeg = Left$(Word, i): WordEnd = Mid$(Word, i + 1)
    Result = True
HandleExit:  p_GetWordParts = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function p_GetWordParts2(ByVal Word As String, _
    Optional ByRef WordBeg As String, Optional ByRef WordEnd As String _
    ) As Boolean
' �������� � ����� ������ � ��������� (������ �������)
'-------------------------
Dim sChar As String * 1
Dim fStop As Boolean, c As Long, v As Long
Dim i As Long, iMax As Long
Dim Result As Boolean

    Result = False
    On Error GoTo HandleError
    Word = LCase$(Word)
    iMax = Len(Word): i = iMax: c = 0: v = 0: fStop = False
    WordBeg = Word: WordEnd = vbNullString
    Do Until fStop
        i = i - 1
        sChar = LCase$(Right$(WordBeg, 1))
        Select Case GetCharType(Right$(WordBeg, 1))
        Case SymbolTypeVowel: WordEnd = sChar & WordEnd: WordBeg = Left$(Word, i): v = v + 1
        Case SymbolTypeCons: If v > 0 Or c = 2 Then fStop = True Else WordEnd = sChar & WordEnd: WordBeg = Left$(Word, i): c = c + 1
        Case SymbolTypeSign: WordBeg = Left$(Word, i): If v = 0 And c = 0 Then WordEnd = sChar & WordEnd: Exit Do
        Case Else: Err.Raise vbObjectError + 512
        End Select
        If i < 1 And Not fStop Then WordBeg = Word: WordEnd = vbNullString: fStop = True
    Loop
    Result = True
HandleExit:  p_GetWordParts2 = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Private Function p_GetWordGender(ByVal Word As String, Optional ByRef WordEnd As String) As DeclineGend
' ���������� ��� �� ��������� ����� (�������)
'-------------------------
Dim Result As DeclineGend

    Result = DeclineGendUndef
    On Error GoTo HandleError
    Word = LCase$(Word)
    If Len(WordEnd) = 0 Then Call p_GetWordParts(Word, WordEnd:=WordEnd)
    Select Case WordEnd
    '������� ����� ��������� -�, -�, � ������� (����, ����, ���, ����, ������)
    '������� ����� ��������� -�, -�, � ������� (����, ����, ����, �����, �������)
    Case "�"
        Select Case LCase$(Word)
        Case "����", "�����", "����", _
             "�����", "�����", "�������": Result = 1   '������� ���
        Case "����", "����", "����", _
             "������": Result = 2           '������� ���
        Case Else:
            Select Case VBA.Left$(VBA.Right$(Word, 2), 1)
            Case "�", "�", "�", "�":    Result = 2  '������� ��� (����,����,�����,����...)
            Case "�", "�", "�":         Result = 2  '������� ��� (����,����,���� �� - ����)
            Case "�":
                Select Case VBA.Left$(VBA.Right$(Word, 3), 1)
                Case "�":               Result = 1  '������� ��� (-���)
                Case "�":               Result = 2  '������� ��� (-���)
                Case Else:              Result = 1  '������� ���
                End Select
            Case Else:                  Result = 1  '������� ��� (����,����...)
            End Select
        End Select
    Case "�", "�"
        Select Case Word
        Case "����", "����", "�������": Result = 1  '������� ���
        Case Else:                      Result = 2  '������� ���
        End Select
'    ''����� ��� - � ����������� �� ���������, ����� ������������� � � �������, � � ������� ����
'    ''    (������, �������, ������, ������, ������).
'    Case "��", "��":                    NewGend = 1 '������� ���
    Case "��", "��":                    Result = 2  '������� ���
    Case "���", "���", "���":           Result = 2  '������� ���
    Case "�", "�", "��", "��", _
         "�", "�", "�": Result = 3               '������� ���
    Case Else
'    ''������� ��� (��������� ���������)
        If GetCharType(Right$(WordEnd, 1)) = SymbolTypeCons Then Result = 1
    End Select
HandleExit:  p_GetWordGender = Result: Exit Function
HandleError: Result = DeclineGendUndef: Err.Clear: Resume HandleExit
End Function

Private Function p_GetWordSpeechPartType(ByVal Word As String) As SpeechPartType
' ���������� ����� ���� �� ��������� ����� (�������)
'-------------------------
' �� ���� ����� � ��� ������, - ����� ���, �� �������...
' ����� ����� � �������� ��������� ��������� � ������ ����� ����
'-------------------------
Dim Result As SpeechPartType

    On Error GoTo HandleError
    Select Case Word
    ' �����������
    Case "�", "��", "��", "���", "���", "��", "���", "���", "����", _
        "��", "��", "���", "��", "���":
            Result = SpeechPartTypePronoun
    ' ��������
    Case "�", "�", "�", "�", "�", "�", "�", "��", "��", "��", "��", "��", "��", "��", "��", "��", _
         "���", "���", "���", "���", "���", "���", "�����", "�����", "�����", "���������", "�����", _
         "������", "������", "�����", "�����", "�����", "�����", "������", "������", "�������", "�����", _
         "������", "���������", "�����", "������", "�����", "���������", "����������"
            Result = SpeechPartTypePreposition
    ' ��������������� �� -�� � ��,�� � �.�.
    Case "����", "����", "����", "����", "����", "����", "����", "������", "������", "����", "����", "������", "������", _
         "����" ', "�������", "�����������", "���������", "�����������", "�������������", "������", _
         "��������"
            Result = SpeechPartTypeNoun
    Case Else
'        If Len(WordEnd) = 0 Then Call p_GetWordParts(Word, WordEnd:=WordEnd)
        Select Case Right$(Word, 2)
    ' ��������������
        Case "��", "��", "��", "��", "��", "��", "�", "�", "�", "��", "��", "��" ', "��" ' - ������� ����� ���. �� -��
            Result = SpeechPartTypeAdject: GoTo HandleExit
        End Select
    ' ������������
        Dim tmp: For Each tmp In p_NumWordsArray
            If Word = tmp Then _
            Result = SpeechPartTypeNumeral: GoTo HandleExit
        Next tmp
        Select Case Right$(Word, 3)
    ' �������
        Case "���", "���", "���", "���", "���", "���", "���"
        ' ��� ����: ����,���� � �.�. - ���.,� ���� � �� -����� - ������������
            Result = SpeechPartTypeVerb: GoTo HandleExit
        End Select
    ' ��������� ������� ����������������
            Result = SpeechPartTypeNoun
    End Select
HandleExit:  p_GetWordSpeechPartType = Result: Exit Function
HandleError: Result = SpeechPartTypeUndef: Err.Clear: Resume HandleExit
End Function

Private Function p_GetSymbCase(Word As String, Optional Template As String) As Integer
' ���������� ��������� �������� �������� �����
'-------------------------
Dim Result As Integer
    Result = False
    On Error GoTo HandleError
' 0 - �� ����������
    If Word = UCase$(Word) Then
' 1 (vbUpperCase) - ��� ������� � ������� ��������
        Result = vbUpperCase
    ElseIf Word = LCase$(Word) Then
' 2 (vbLowerCase) - ��� ������� � ������ ��������
        Result = vbLowerCase
    ElseIf Word = StrConv(Word, vbProperCase) Then
' 3 (vbProperCase) - ������ ������ � ������� ���������, - � ������ ��������
        Result = vbProperCase
    Else
'-1 - ������� �������� ������������ �� ������� (����� ���� � �������, ����� - � ������ ��������)
    ' ��������� ������ �������� �� �����
        Template = p_GetWordTemplate(Word, True): Result = -1
    End If
HandleExit:  p_GetSymbCase = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Private Function p_SetSymbCase(Word As String, SymbCase As Integer, Optional ByVal Template As String) As String
' ������������� ��������� �������� �������� �����
'-------------------------
Dim Result As String
    Result = Word
    On Error GoTo HandleError
    Select Case SymbCase
    Case vbUpperCase:   Result = UCase$(Word)
    Case vbLowerCase:   Result = LCase$(Word)
    Case vbProperCase:  Result = StrConv(Word, vbProperCase)
    Case -1             ' ����������� �� �������
' ������ ���?? - ����... - ������ ����������� - ��� �� ����, � �����? )
' ������� ��� ����� ��� ��� ���� ������ � �����������,
' � ������ � ���������������� ���������� � ����� ����� ������ ������������
' � ���� ������ ����������� ��������� ����������� ��������� ����� �������� �� ���������
        ' ��������� ������ �������� ����� � ���������� ��
        Dim s As String * 1, � As Integer
        Dim i As Long, iMax As Long: i = 1: iMax = Len(Template)
        Dim NewTemp As String: NewTemp = p_GetWordTemplate(Word, True)
        If LCase$(Template) = LCase$(NewTemp) Then
        ' ���� ��������� - ����������� �� �������
            Do Until i > iMax
                s = Mid$(Template, i, 1)
                If LCase(s) = s Then
                    Mid$(Result, i, 1) = LCase$(Mid$(Result, i, 1))
                Else
                    Mid$(Result, i, 1) = UCase$(Mid$(Result, i, 1))
                End If
                i = i + 1
            Loop
        Else
        ' ���� ��� �������� �������� ������ � ������� �������� �����
        Dim j As Long, jMax As Long: j = 1: jMax = Len(NewTemp)
        ' ����������� �� ���������� �������
            ' ���� ��� ������� (xgs) ����������� (���������) ������� ���������
            ' � ����� ������� ������ ����� ����������� ��� � ���������� �������,
            ' ���� �� ��������� - ��������� ������� ����������� �������
            ' ����� �� ����������� ��� �������
            Template = Replace$(Template, "x", "g"): NewTemp = Replace$(NewTemp, "x", "g")
            ' ������ ������ ����� � �������� ��������� �������
            s = Mid$(Template, i, 1)
            If LCase$(s) = s Then
                Mid$(Result, j, 1) = LCase(Left$(Result, 1))
            Else
                Mid$(Result, j, 1) = UCase(Left$(Result, 1))
            End If
            i = i + 1: j = j + 1
            Do Until j > jMax
                If i > iMax Then i = 2 ' ���� ����������� ����� ����������� ����������� - �������� �������
                s = Mid$(Template, i, 1)
                ' ���� ��� ������� ����������� ������� ��������� � ����� ������� ������ �����
                ' ����� ������� ������� ����������� ������� � ����������� ������� ������
                If LCase(Mid$(NewTemp, j, 1)) = LCase(s) Then
                    i = i + 1
                    If LCase(s) = s Then
                        Mid$(Result, j, 1) = LCase$(Mid$(Result, j, 1))
                        Mid$(NewTemp, j, 1) = LCase$(Mid$(NewTemp, j, 1))
                    Else
                        Mid$(Result, j, 1) = UCase$(Mid$(Result, j, 1))
                        Mid$(NewTemp, j, 1) = UCase$(Mid$(NewTemp, j, 1))
                    End If
                End If
                j = j + 1
            Loop
            Template = NewTemp
        End If
        ' ���������� ����������� �����
    Case Else ' ��������� ��� ����
    End Select
HandleExit:  p_SetSymbCase = Result: Exit Function
HandleError: Result = Word: Err.Clear: Resume HandleExit
End Function

'=========================
' ���������������
'=========================
Private Function Pwr2(Index) As Long
' ���������� ������� ����� 2. ����� ��� ������� ��������
'-------------------------
    On Error GoTo HandleError
    If Index < 0 Or Index > 31 Then Err.Raise vbObjectError
    Select Case Index
    Case 0:     Pwr2 = &H1&
    Case 1:     Pwr2 = &H2&
    Case 2:     Pwr2 = &H4&
    Case 3:     Pwr2 = &H8&
    Case 4:     Pwr2 = &H10&
    Case 5:     Pwr2 = &H20&
    Case 6:     Pwr2 = &H40&
    Case 7:     Pwr2 = &H80&
    Case 8:     Pwr2 = &H100&
    Case 9:     Pwr2 = &H200&
    Case 10:    Pwr2 = &H400&
    Case 11:    Pwr2 = &H800&
    Case 12:    Pwr2 = &H1000&
    Case 13:    Pwr2 = &H2000&
    Case 14:    Pwr2 = &H4000&
    Case 15:    Pwr2 = &H8000&
    Case 16:    Pwr2 = &H10000
    Case 17:    Pwr2 = &H20000
    Case 18:    Pwr2 = &H40000
    Case 19:    Pwr2 = &H80000
    Case 20:    Pwr2 = &H100000
    Case 21:    Pwr2 = &H200000
    Case 22:    Pwr2 = &H400000
    Case 23:    Pwr2 = &H800000
    Case 24:    Pwr2 = &H1000000
    Case 25:    Pwr2 = &H2000000
    Case 26:    Pwr2 = &H4000000
    Case 27:    Pwr2 = &H8000000
    Case 28:    Pwr2 = &H10000000
    Case 29:    Pwr2 = &H20000000
    Case 30:    Pwr2 = &H40000000
    Case 31:    Pwr2 = &H80000000
    Case Else:  Pwr2 = 0
    End Select
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Private Function p_GetLocaleInfo(LCType As Long) As String
' ���������� ������������ ���������
'-------------------------
' lcType - ��������� LOCALE_
Dim lpBuffer As String * 100
    On Error GoTo HandleError
    If GetLocaleInfo(LOCALE_USER_DEFAULT, LCType, lpBuffer, 99) = 0 Then Err.Raise vbObjectError + 512
    p_GetLocaleInfo = Left$(lpBuffer, InStr(lpBuffer, Chr$(0)) - 1)
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
'    Select Case Err.Number
'    Case vbObjectError + 512: MsgBox "Unknown constant", vbOKOnly Or vbExclamation, "Error!"
'    End Select
End Function
Private Function p_GetCollKeys(oColl As Collection, Optional oIdxs) As String()
' ���������� ������ ������ ��������� (Base=1)
'-------------------------
' �������� ���� ������: https://stackoverflow.com/questions/5702362/vba-collection-list-of-keys
'-------------------------
Dim CollPtr As LongPtr, KeyPtr As LongPtr, ItemPtr As LongPtr, Address As LongPtr
Dim bIdxs As Boolean: bIdxs = Not IsMissing(oIdxs): If bIdxs Then Set oIdxs = New Collection
Dim Result() As String, Length As Long
Dim i As Long, iMax As Long
    CollPtr = VBA.ObjPtr(oColl)                             ' ����� ��������� � ������
    Address = CollPtr + 3 * PTR_LENGTH + 4                  ' ����� ���������� ��������� ���������
    If Address <> 0 Then Call CopyMemory(ByVal VarPtr(iMax), ByVal Address, 4)
    If iMax <> oColl.Count Then Stop                        ' �� ��������� � ����������� ������������ �������� - ������!
    ReDim Result(1 To iMax)                                 ' ��������� ������ ��� �������� ������
    Address = CollPtr + 4 * PTR_LENGTH + 8                  ' ����� ������� �������� ���������
    If Address <> 0 Then Call CopyMemory(ByVal VarPtr(ItemPtr), ByVal Address, PTR_LENGTH)
    For i = 1 To iMax
        If ItemPtr = 0 Then Exit For
        Address = ItemPtr + 2 * PTR_LENGTH + 8              ' ����� ����� �������� ���������
        If Address <> 0 Then Call CopyMemory(ByVal VarPtr(KeyPtr), ByVal Address, PTR_LENGTH)
        If KeyPtr <> 0 Then                                 ' ��������� ���� �������� ���������
        Call CopyMemory(ByVal VarPtr(Length), ByVal KeyPtr - 4, PTR_LENGTH)
        Result(i) = Space(Length \ 2)
        Call CopyMemory(ByVal StrPtr(Result(i)), ByVal KeyPtr, ByVal Length)
        End If
        Address = ItemPtr + 4 * PTR_LENGTH + 8              ' ����� ���������� �������� ���������
        If Address <> 0 Then Call CopyMemory(ByVal VarPtr(ItemPtr), ByVal Address, PTR_LENGTH)
        If bIdxs Then oIdxs.Add i, Result(i)                ' ���� ����� ���� �������� ��������� ������������ ����� ��������
    Next i
    p_GetCollKeys = Result
End Function
Private Function p_HFontByControl(Optional ctl As Variant, Optional FontName, Optional FontSize, _
    Optional FontColor, Optional FontWeight, Optional FontUnderline, Optional FontStrikeOut, Optional FontItalic, Optional hdc As LongPtr = 0) As LongPtr
' ������� hFont �� ���������� ��������
'-------------------------
    'If Not TypeOf ctl Is Access.Control Then Err.Raise vbObjectError + 512
Dim tDC As LongPtr, hFont As LongPtr
    If hdc = 0 Then tDC = GetDC(0) Else tDC = hdc
' ������ �����
On Error Resume Next
Dim fName As String:    fName = IIf(IsMissing(FontName), ctl.FontName, FontName): If Err Then fName = "Arial": Err.Clear
Dim fSize As Long:      fSize = IIf(IsMissing(FontSize), ctl.FontSize, FontSize): If Err Then fSize = 10: Err.Clear
'Dim fColor As Long:     fColor = IIf(IsMissing(FontColor), ctl.ForeColor, FontColor): If Err Then fColor = vbBlack: Err.Clear
Dim fWeight As Long:    fWeight = IIf(IsMissing(FontWeight), ctl.FontWeight, FontWeight): If Err Then fWeight = 0: Err.Clear
Dim fItalic As Long:    fItalic = IIf(IsMissing(FontItalic), ctl.FontItalic, FontItalic): If Err Then fItalic = False: Err.Clear
Dim fUnderline As Long: fUnderline = IIf(IsMissing(FontUnderline), ctl.FontUnderline, FontUnderline): If Err Then fUnderline = False: Err.Clear
'Dim fStrikeOut As Long: fStrikeOut = IIf(IsMissing(FontStrikeOut), ctl.FontStrikeOut, FontStrikeOut): If Err Then fStrikeOut = False: Err.Clear
On Error GoTo HandleError
    'FontSize = -(FontSize * PT / TwipsPerPixels)
    'fSize = -Int(fSize * GetDeviceCaps(tDC, LOGPIXELSY) / PointsPerInch)
    fSize = -MulDiv(fSize, GetDeviceCaps(tDC, LOGPIXELSY), PointsPerInch)
    hFont = CreateFont(fSize, 0, 0, 0, _
        fWeight, fItalic, fUnderline, 0, _
        RUSSIAN_CHARSET, 0, 0, ANTIALIASED_QUALITY, 0, fName)  ' PROOF_QUALITY | CLEARTYPE_QUALITY | ANTIALIASED_QUALITY
    If hdc = 0 Then ReleaseDC 0, tDC
HandleExit:  p_HFontByControl = hFont: Exit Function
HandleError: hFont = False: Err.Clear: Resume HandleExit
End Function
Private Function p_IsEvalutable(ByRef Expr As String, Optional ByRef Value) As Boolean
' ��������� ������������� ���������� ���������, � Value ���������� ��������� ����������
'-------------------------
    On Error GoTo HandleError
    If IsNumeric(Expr) Then p_IsEvalutable = False: Exit Function ' ��� ������������� ��������� �����
#If APPTYPE = 0 Then ' Access
    Value = Application.Eval(Expr)
#ElseIf APPTYPE = 1 Then ' Excel
    Value = Application.Evaluate(Expr)
#Else
    Err.Raise 2438
#End If
' ��� ����������� ������������ ���������� ��������� � ����������� �������
' ����� �������� ������������ ���������
    If IsNumeric(Value) Then
'Dim cPosDelim As String * 1: cPosDelim = p_GetLocaleInfo(LOCALE_STHOUSAND)  ' Chr(160) - ����������� �������� ����� �����
'        Value = Replace(Value, cPosDelim, vbNullString)
Dim cDecDelim As String * 1: cDecDelim = "," ' p_GetLocaleInfo(LOCALE_SDECIMAL)   ' Chr(44)  - ����������� �����/������� ����� ���������� �����
        Value = Replace(Value, cDecDelim, ".")
    End If
HandleExit:  p_IsEvalutable = True:  Exit Function
HandleError: p_IsEvalutable = False: Err.Clear
End Function
Private Function p_IsExist(Key As String, Coll As Collection, Optional ByRef Value) As Boolean
' ��������� ������� �������� � ���������
'-------------------------
    On Error GoTo HandleError
    Value = Coll(Key)
HandleExit:  p_IsExist = True:  Exit Function
HandleError: p_IsExist = False: Err.Clear
End Function
#If APPTYPE = 1 Then ' ��� Excel ����� ������ Nz
Private Function Nz(p1, Optional p2) As Variant
    Select Case True
    Case Not IsNull(p1): Nz = p1
    Case IsMissing(p2):  Nz = Empty
    Case Else: Nz = p2
    End Select
' Nz(Null) return Empty in MS Access, so the following Excel vba matches MS Access perfectly.

' to test it open vba immediate window and type    ?(nz(null) = 0) & " " & (nz(null) = "")
' You will get True True
End Function
#End If

