Attribute VB_Name = "modForms"
Option Explicit
Option Compare Database
'=========================
Private Const c_strModule As String = "modForms"
'=========================
' ��������      : ������� ��� ������ � �������/�������� � ���������� Access
' ������        : 1.0.20.453555937
' ����          : 04.03.2024 14:14:56
' �����         : ������ �.�. (KashRus@gmail.com)
' ����������    :
' v.1.0.20      : 08.08.2021 - ��������� � AccControlLocation - ��������� ����������� ���������� ����������/�������� ����������
' v.1.0.19      : 18.07.2019 - ��������� CreateHFontByControl - �������� hFont �� ������ ���������� ��������
' v.1.0.18      : 16.04.2019 - ��������� AccControlLocation - ��������� ������� �������� �� ������ ��� � ����������  ������� ���� (� ��������)
' v.1.0.17      : 19.10.2012 -
'=========================
' ToDo: ����� ����� �������� ������ ������� ������ ������� ����������
'-------------------------
Public Const twMinLim = 10  ' �������� ������ ������� ��������� ����������, ������ - �������

Public Const pxWinGap = 16                          ' �������� ���� � px ��� ���������� ��������
' ��� �������� Width/Height >> InsideWidth/InsideHeight
Public Const pxScrBar = 17                          ' 17px ������ ScrollBars
Public Const pxNavBut = 18                          ' 18px ������ NavigationButtons
Public Const pxRecSel = 19                          ' 19px ������ RecordSelectors
'-------------------------
' ��������� ��� ����� ��������� ������
'-------------------------
Public Const c_strKey = "ID"            '
Public Const c_strParent = "PARENT"     '
Public Const c_strName = "NAME"         '
Public Const c_strCName = "CNAME"       '
Public Const c_strFaceKey = "FaceId"    ' ��� ���� ���� �������� ��� PictureData_SetToControl ������������ � ��������� ��� �������� ���� � ������� CreateItemsFromSQL
'-------------------------
' ��������� �������������� ��������
'-------------------------
Public Const DotsPerInch = 96
Public Const PointsPerInch = 72
Public Const TwipsPerInch = 1440
Public Const CentimitersPerInch = 2.54              '1 ���� = 127 / 50 ��
Public Const HimetricPerInch = 2540                 '1 ���� = 1000 * 127/50 himetrix
'
Public Const cm = TwipsPerInch / CentimitersPerInch '1 �� = 566.929133858 twips
Public Const inch = TwipsPerInch                    '1 ���� = 1440 twips
Public Const pt = TwipsPerInch / PointsPerInch      '1 ����� = 20 twips
Public Const px = TwipsPerInch / DotsPerInch        '1 ������� = 15 twips
'-------------------------
Public Enum eDirection
    DIRECTION_HORIZONTAL = 0
    DIRECTION_VERTICAL = 1
End Enum
Public Enum acFormState         ' ��������� �����
    acObjStateClosed = 0            ' Closed
    acObjStateOpen = 1              ' Open
    acObjStateDirty = 2             ' Changed but not saved
    acObjStateNew = 4               ' New
End Enum
Public Enum ePosition           ' ��������� ��������� �� ������� (��� ��������������� ������������ � ������������)
    ePosUndef = 0                   ' �� ������
    eLeft = 1                       ' �� ������ ����
    eRight = 2                      ' �� ������� ����
    eTop = 4                        ' �� �������� ����
    eBottom = 8                     ' �� ������� ����
    eWidth = eLeft + eRight         ' �� ������ (����� �� �����������)
    eCenterHorz = eWidth
    eHeight = eTop + eBottom        ' �� ������ (����� �� ���������)
    eCenterVert = eHeight
    eCascade = &H100                ' ���������� (������ ��� ����� ??)
End Enum
Public Enum eAlign              ' ������������ ������ �������
    eAlignUndef = 0                 ' �� ������
    ' 2 ����������� �� 3 ��������� ����� �������
    ' �����: 3x3 = 9 ����� ������������.
    eAlignLeftTop = eLeft + eTop                ' �� ������ �������� ����
    eAlignRightTop = eRight + eTop              ' �� ������� �������� ����
    eAlignLeftBottom = eLeft + eBottom          ' �� ������ ������� ����
    eAlignRightBottom = eRight + eBottom        ' �� ������� ������� ����
    eCenterHorzTop = eCenterHorz + eTop         ' �� �������� ���� ������������ �� �����������
    eCenterHorzBottom = eCenterHorz + eBottom   ' �� ������� ���� ������������ �� �����������
    eCenterVertLeft = eLeft + eCenterVert       ' �� ������ ���� ������������ �� ���������
    eCenterVertRight = eRight + eCenterVert     ' �� ������� ���� ������������ �� ���������
    eCenter = eCenterHorz + eCenterVert         ' ������������ ���������� �������
End Enum
Public Enum ePlace              ' ���������� Obj2 ������������ Obj1
    ' 2 ������� �� 9 ����� �������� �� ������: LT,LC,LB,CB,RB,RC,RT,CT,CC
    ' �����: 9x9 = 81 ������� ��������.
    ' ����������� �� ��� ������������, ������� �������� ��� ��� ���,
    ' �� �������� ����� ���������� �������� �� �����:
    ' =H2+V2+H1+V1, ���:
    ' Obj1 (� �������� �����������) - ���� 0-3:  L1=1,  R1=2,  T1=4,  B1=8
    '   H1 - ��������� �� ����������� ����� �������� �� Obj1
    '       ={eLeft|eRight|eCenterHorz}
    '   V1 - ��������� �� ��������� ����� �������� �� Obj1
    '       ={eTop|eBottom|eCenterVert}
    ' Obj2 (������� �����������)    - ���� 4-8:  L2=16, R2=32, T2=64, B2=128
    '   H2 - ��������� �� ����������� ����� �������� �� Obj2
    '       ={eLeft|eRight|eCenterHorz} * 16
    '   V2 - ��������� �� ��������� ����� �������� �� Obj2
    '       ={eTop|eBottom|eCenterVert} * 16
    ePlaceUndef = 0     ' ��-��������� = 222 -> ePlaceOnRight - ������� ������ �� ������
' ������ �� ������
    ePlaceCenter = eCenter * 16 + eCenter                           ' �� ������ (������)
    ePlaceToLeft = eCenterVertLeft * 16 + eCenterVertLeft           ' ������ ����� �� ������
    ePlaceToRight = eCenterVertRight * 16 + eCenterVertRight        ' ������ ������ �� ������
    ePlaceToTop = eCenterHorzTop * 16 + eCenterHorzTop              ' ������ �� ������ ������
    ePlaceToBottom = eCenterHorzBottom * 16 + eCenterHorzBottom     ' ������ �� ������ �����
' ������� �� ������
    ePlaceOnLeft = eCenterVertRight * 16 + eCenterVertLeft          ' ������� ����� �� ������
    ePlaceOnRight = eCenterVertLeft * 16 + eCenterVertRight         ' ������� ������ �� ������
    ePlaceOnTop = eCenterHorzBottom * 16 + eCenterHorzTop           ' ������� �� ������ ������
    ePlaceOnBottom = eCenterHorzTop * 16 + eCenterHorzBottom        ' ������� �� ������ �����
' ������ �� ����
    ePlaceToLeftTop = eAlignLeftTop * 16 + eAlignLeftTop            ' ������ ����� ������
    ePlaceToRightTop = eAlignRightTop * 16 + eAlignRightTop         ' ������ ������ ������
    ePlaceToLeftBottom = eAlignLeftBottom * 16 + eAlignLeftBottom   ' ������ ����� �����
    ePlaceToRightBottom = eAlignRightBottom * 16 + eAlignRightBottom ' ������ ������ �����
' ������� �� ����
    ePlaceOnLeftToTop = eAlignRightTop * 16 + eAlignLeftTop         ' ������� ����� � �������� ����
    ePlaceOnLeftToBottom = eAlignRightBottom * 16 + eAlignLeftBottom ' ������� ����� � ������� ����
    ePlaceOnRightToTop = eAlignLeftTop * 16 + eAlignRightTop        ' ������� ������ � �������� ����
    ePlaceOnRightToBottom = eAlignLeftBottom * 16 + eAlignRightBottom ' ������� ������ � ������� ����
    ePlaceOnTopToLeft = eAlignLeftBottom * 16 + eAlignLeftTop       ' ������� � ������ ���� ������
    ePlaceOnTopToRight = eAlignRightBottom * 16 + eAlignRightTop    ' ������� � ������� ���� ������
    ePlaceOnBottomToLeft = eAlignLeftTop * 16 + eAlignLeftBottom    ' ������� � ������ ���� �����
    ePlaceOnBottomToRight = eAlignRightTop * 16 + eAlignRightBottom ' ������� � ������� ���� �����
' ���������� (������ ��� ����� ??)
    eCascadeFromLeftTop = eCascade + ePlaceToLeftTop                ' ���������� �������� ������-����
    eCascadeFromRightTop = eCascade + ePlaceToRightTop              ' ���������� �������� �����-����
    eCascadeFromLeftBottom = eCascade + ePlaceToLeftBottom          ' ���������� �������� ������-�����
    eCascadeFromRightBottom = eCascade + ePlaceToRightBottom        ' ���������� �������� �����-�����
End Enum

Public Enum eObjSizeMode                    ' ��������������� ��������
    apObjSizeZoomDown = -1                  '-1 - ���������������� ��������������� (������ ����������)
    apObjSizeClip = acOLESizeClip           ' 0 - �� ������ ������. ���� ������ ������ ������� ������ - �������
    apObjSizeStretch = acOLESizeStretch     ' 1 - ������/���������� (�������� ���������)
    'apObjSizeAutoSize = acOLESizeAutoSize   ' 2 - ???
    apObjSizeZoom = acOLESizeZoom           ' 3 - ���������������� ���������������
End Enum

Public Enum eControlScale   ' ��������������� ��������
    csYes = -1                  ' �������������� �������
    csNo = 0                    ' �� �������������� �������
    csDefault = 1               ' ������������ �������� ��-��������� ��� �����
End Enum
Public Enum eScaleWhen      ' ����� �������������� �������
    scNo = 0                    ' �������
    scYes = -1                  ' ��� ��������� ��������
    scAtLoad = 1                ' ������ ��� �������������� ��������
End Enum
Public Enum eScaleType      ' ��� ��������������� �����
    sfNo = 0                    ' �� ������������� ����� ��� ��������� ���� �����
    sfSequent = -1              ' ��������������� �������� ������� ������ ��� ��������� ���� �����
    sfProp = 1                  ' �������� ������� ������ ��������������� ��������� ���� �����
End Enum
Public Enum eControlSize    ' ��� ��������� �������� ��������
    czNone = 0                  ' �� �����������
    czRight = 1                 ' ����������� ������
    czBottom = 2                ' ����������� ����
    czBoth = 3                  ' ����������� ������-����
End Enum
Public Enum eControlFloat   ' ��� �������� ��������
    cfNone = 0                  ' ��� �������� (�������� � ������ ��������)
    cfRight = 1                 ' �������� � ������� (��������) ����
    cfBottom = 2                ' �������� � (������) ������� ����
    cfBoth = 3                  ' �������� � �������-������� ����
End Enum

Public Enum eObjectStyle   ' ����� ����������� �������� ��������
    lsNone = 0                  ' ������� � ������ �� �������� (��� ������������� � �������������)
' �������� ����� ��������
    lsLeft = 1                  ' �������� � ����� ������� �������
    lsRight = 2                 ' �������� � ������ ������� �������
    lsTop = 4                   ' �������� � ������ ������� �������
    lsBottom = 8                ' �������� � ������ ������� �������
' �������������� ����� ��������
    lsHorz = 3                  ' ������������ ����� ������� � ������ ��������� ������� (������������� �������������)
    lsLeftRight = 3
    lsLeftTop = 5               ' �������� � ������ �������� ���� �������
    lsRightTop = 6              ' �������� � ������� �������� ���� �������
    lsHorzTop = 7               ' �������� � ������� ������� � ������������ ����� ������� � ������ ��������� ������� (������������� �������������)
    lsLeftBottom = 9            ' �������� � ������ ������� ���� �������
    lsRightBottom = 10          ' �������� � ������� ������� ���� �������
    lsHorzBottom = 11           ' �������� � ������ ������� � ������������ ����� ������� � ������ ��������� ������� (������������� �������������)
    lsVert = 12                 ' ������������ ����� ������� � ������ ��������� ������� (������������� �����������)
    lsTopBottom = 3
    lsVertLeft = 13             ' �������� � ����� ������� � ������������ ����� ������� � ������ ��������� ������� (������������� �����������)
    lsVertRight = 14            ' �������� � ������ ������� � ������������ ����� ������� � ������ ��������� ������� (������������� �����������)
    lsFull = 15                 ' ������������� ����������� � ������������� �� ������ � ������ ������� �������
    lsLeftRightTopBottom = 15
' ����� ������������� ��������
    lsXProp = 16                ' ������� Left ������� �� ������ ������� (������������� �� ������)
    lsRProp = 32                ' ������� Right ������� �� ������ ������� (������������� �� ������)
    lsYProp = 64                ' ������� Top ������� �� ������ ������� (������������� �� ������)
    lsBProp = 128               ' ������� Bottom ������� �� ������ ������� (������������� �� ������)
    lsWProp = lsXProp + lsRProp ' ������ ��������������� ������ �������
    lsHProp = lsYProp + lsBProp ' ������ ��������������� ������ �������
' ����� ������ �����������/������
    lsShowIcon = 1024           ' �������� ������ ������
    lsShowText = 2048           ' �������� ������ �������
    lsShowIconText = 3072       ' �������� ������ � �������
End Enum
Public Enum eControlSplit   ' ��� �����������
    cdNone = 0                  ' �� �������� ����������
    cdVert = 1                  ' ������������ ��������
    cdHorz = 2                  ' �������������� ��������
'    cdBoth = 3                 ' ??? ��� ��� ???
End Enum
Public Enum eAlignText ' ������������ ������
    TA_LEFT = 0                 ' ������� ����� ��������� �� ����� ������ �������� ��������������.
    TA_RIGHT = 2                ' ������� ����� ��������� �� ������ ������ �������� ��������������.
    TA_CENTER = 6               ' ������� ����� ������������� ������������� �� ������ �������� ��������������.
    TA_TOP = 0                  ' ������� ����� �� ������� ������ �������� ��������������.
    TA_BOTTOM = 8               ' ������� ����� �� ������ ������ �������� ��������������.
    TA_BASELINE = 24            ' ������� ����� ��������� �� ������� ����� ������.
    TA_RTLREADING = 256         ' �������� Windows �� ������ �������� �������: ����� ����������� ��� ������� ������ ������ ������ , � ����������������� ������� ������ �� ��������� ����� �������. ��� ����������� ������ �����, ����� �����, ��������� � �������� ���������� ������������ ��� ��� ���������� ��� ��� ��������� �����.
'    TA_NOUPDATECP               ' ������� ������� �� �������������� ����� ������� ������ ������ ������.
'    TA_UPDATECP                 ' ������� ������� �������������� ����� ������� ������ ������ ������.
'    TA_MASK  = (TA_BASELINE + TA_CENTER + TA_UPDATECP)
End Enum
'Public Enum eObjSizeMode    ' ����� ��������� �������� �����������
'    apObjSizeClip = acOLESizeClip           ' 0 - �� ������ ������. ���� ������ ������ ������� ������ - �������
'    apObjSizeStretch = acOLESizeStretch     ' 1 - ������/���������� (�������� ���������)
'    ' 2 'acOLESizeAutoSize
'    apObjSizeZoom = acOLESizeZoom           ' 3 - ���������������� ���������������
'    apObjSizeZoomDown = -1                  '-1 - ���������������� ���������������, ������ ���������
'End Enum

Public Enum eFieldFormat
    vbInteger = 2
    vbLong = 3
    vbByte = 17
    vbDate = 7
    vbDateTimeJ = 77
    vbSingle = 4
    vbDouble = 5
    vbCurrency = 6
    vbString = 8
    vbBoolean = 11
End Enum

''--------------------------------------------------------------------------------
'' POINTER
''--------------------------------------------------------------------------------
'#If VBA7 = 0 Then       'LongPtr trick by @Greedo (https://github.com/Greedquest)
'Private Enum LongPtr
'    [_]
'End Enum
'#End If
'#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>  PtrSafe, LongPtr and LongLong
'Private Const PTR_LENGTH As Long = 8
'Private Const VARIANT_SIZE As Long = 24
'#Else                   '<OFFICE97-2010>        Long
'Private Const PTR_LENGTH As Long = 4
'Private Const VARIANT_SIZE As Long = 16
'#End If                 '<WIN32>

Private Type RECT ' Store rectangle coordinates.
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINT
    x As Long
    y As Long
End Type
'--------------------------------------------------------------------------------
' ���������� ��������
'--------------------------------------------------------------------------------
Public Const IDC_ARROW = 32512&        ' �������� �����
'Public Const IDC_IBEAM = 32513&        ' ��������� ������
'Public Const IDC_WAIT = 32514&         ' ������� ����������
'Public Const IDC_CROSS = 32515&        ' ����������� ���������
'Public Const IDC_UPARROW = 32516&      ' ����������� ���������
'
'Public Const IDC_SIZE = 32640&         '
'Public Const IDC_ICON = 32641&         '
'Public Const IDC_SIZENWSE = 32642&     ' ��������� �������� �� ���������1
'Public Const IDC_SIZENESW = 32643&     ' ��������� �������� �� ���������2
Public Const IDC_SIZEWE = 32644&       ' ��������� �������������� ��������
Public Const IDC_SIZENS = 32645&       ' ��������� ������������ ��������
Public Const IDC_SIZEALL = 32646&      ' �����������
'Public Const IDC_NO = 32648&           ' �������� ����������
'Public Const IDC_HAND = 32649&         ' ����������� ���� (�����������)
'Public Const IDC_APPSTARTING = 32650&  ' ������� ����� (�������� ����)

#If Win64 Then
Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Public Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr
Public Declare PtrSafe Function LoadCursorByNum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As LongPtr, ByVal lpCursorName As Long) As LongPtr
#Else
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr
Public Declare Function LoadCursorByNum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As LongPtr, ByVal lpCursorName As Long) As LongPtr
#End If

'--------------------------------------------------------------------------------
' KERNEL32
'--------------------------------------------------------------------------------
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>      PtrSafe, LongPtr and LongLong
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
#Else                   '<WIN32>                    Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
#End If                 '<WIN32>
'--------------------------------------------------------------------------------
' USER32
'--------------------------------------------------------------------------------
'Private Type SCROLLBARINFO
'    cbSize As Long
'    rcScrollBar As RECT
'    dxyLineButton As Long
'    xyThumbTop As Long
'    xyThumbBottom As Long
'    reserved As Long
'    rgstate(0 To 5) As Long
'End Type
'Private Type SCROLLINFO
'    cbSize As Long
'    fMask As Long
'    nMin As Long
'    nMax As Long
'    nPage As Long
'    nPos As Long
'    nTrackPos As Long
'End Type
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>  PtrSafe, LongPtr and LongLong
Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hWndLock As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function ClientToScreen Lib "user32.dll" (ByVal hwnd As LongPtr, ByRef lpPoint As POINT) As Long
Private Declare PtrSafe Function ScreenToClient Lib "user32.dll" (ByVal hwnd As LongPtr, ByRef lpPoint As POINT) As Long
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Boolean) As Long
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As apiGetWindowLongIndex) As Long
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpTextString As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare PtrSafe Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As RECT, ByVal fuWinIni As Long) As Long
'������� ���������� �����������
Private Declare PtrSafe Function GetScrollPos Lib "user32" (ByVal hwnd As LongPtr, ByVal nBar As Long) As Long
'Private Declare PtrSafe Function SetScrollPos Lib "user32" (ByVal hWnd As LongPtr, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
'Private Declare PtrSafe Function GetScrollBarInfo Lib "user32" (ByVal hWnd As LongPtr, ByVal idObject As Long, psbi As SCROLLBARINFO) As Long
'Private Declare PtrSafe Function GetScrollInfo Lib "user32" (ByVal hWnd As LongPtr, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
'Private Declare PtrSafe Function SetScrollInfo Lib "user32" (ByVal hWnd As LongPtr, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
#Else                   '<OFFICE97-2010>        Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINT) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Boolean) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As apiGetWindowLongIndex) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpTextString As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As RECT, ByVal fuWinIni As Long) As Long
'������� ���������� �����������
Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
'Private Declare Function SetScrollPos Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
'Private Declare Function GetScrollBarInfo Lib "user32" (ByVal hWnd As Long, ByVal idObject As Long, psbi As SCROLLBARINFO) As Long
'Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
'Private Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
#End If                 '<WIN32>
'--------------------------------------------------------------------------------
' GDI32
'--------------------------------------------------------------------------------
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>  PtrSafe, LongPtr and LongLong
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal e As Long, ByVal o As Long, ByVal w As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal q As Long, ByVal PAF As Long, ByVal f As String) As LongPtr
#Else                   '<OFFICE97-2010>        Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal e As Long, ByVal o As Long, ByVal w As Long, ByVal i As Long, ByVal U As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal cp As Long, ByVal q As Long, ByVal PAF As Long, ByVal f As String) As Long
#End If                 '<WIN32>
'--------------------------------------------------------------------------------
Private Enum apiDeviceCapability
    HORZSIZE = 4
    VERTSIZE = 6
    HORZRES = 8
    VERTRES = 10
    LOGPIXELSX = 88        '  Logical pixels/inch in X
    LOGPIXELSY = 90        '  Logical pixels/inch in Y
End Enum

Public Const HWND_DESKTOP = &H0
Public Const SPI_GETWORKAREA = &H30

Public Const ANSI_CHARSET As Long = &H0
Public Const DEFAULT_CHARSET As Long = &H1
Public Const SYMBOL_CHARSET As Long = &H2
Public Const RUSSIAN_CHARSET As Long = &HCC
Public Const OEM_CHARSET As Long = &HFF

Public Const DEFAULT_QUALITY As Long = 0
Public Const DRAFT_QUALITY  As Long = 1
Public Const PROOF_QUALITY  As Long = 2
Public Const NONANTIALIASED_QUALITY  As Long = 3
Public Const ANTIALIASED_QUALITY As Long = 4
Public Const CLEARTYPE_QUALITY As Long = 5

'------------------------------
' ������ ���� Access
'------------------------------
Public Const accClass = "OMain"                 ' ����� ���� Access
Public Const accClassChild = "MDIClient"        ' ���������� ������� ���� Access
Public Const accClassBD = "ODb"                 ' ���� ���� ������
Public Const accClassFormWindow = "OForm"       ' ����� ���� ����� Access
Public Const accClassFormClient = "OFormSub"    ' ����� ����� Access
Public Const accClassFormPopup = "OFormPopup"   ' ����� ����������� ����� Access
Public Const accClassFormChild = "OFormChild"   ' ����� ����������� ����� Access
Public Const accClassFormNoClose = "OFormNoClose"
Public Const accClassFormClientChild = "OFEDT"  ' ����� ������������ ���� ����� Access
Public Const accClassTableClientChild = "OGNUM" ' ����� ������������ ���� ��������� ����� Access
Public Const accClassRecordSlector = "OSUI"     ' ����� ��������� ������� ��������� ����� Access
Public Const accClassTextbox = "OKttbx"         ' ����� ���������� ����� Access

Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2
Public Const MAX_LEN = 255
'The window position
Public Const SWP_NOSIZE = &H1 ' Retains the current size (ignores the cx and cy members).
Public Const SWP_NOMOVE = &H2 ' Retains the current position (ignores the x and y members).
Public Const SWP_NOZORDER = &H4 ' Retains the current Z order (ignores the hwndInsertAfter member).
Public Const SWP_NOREDRAW = &H8 ' Does not redraw changes. If this flag is set, no repainting of any kind occurs. This applies to the client area, the nonclient area (including the title bar and scroll bars), and any part of the parent window uncovered as a result of the window being moved. When this flag is set, the application must explicitly invalidate or redraw any parts of the window and parent window that need redrawing.
Public Const SWP_NOACTIVATE = &H10 ' Does not activate the window. If this flag is not set, the window is activated and moved to the top of either the topmost or non-topmost group (depending on the setting of the hwndInsertAfter member).
Public Const SWP_DRAWFRAME = &H20 ' Draws a frame (defined in the window's class description) around the window. Same as the SWP_FRAMECHANGED flag.
Public Const SWP_FRAMECHANGED = &H20 ' Sends a WM_NCCALCSIZE message to the window, even if the window's size is not being changed. If this flag is not specified, WM_NCCALCSIZE is sent only when the window's size is being changed.
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80 ' Hides the window.
Public Const SWP_NOCOPYBITS = &H100 ' Discards the entire contents of the client area. If this flag is not specified, the valid contents of the client area are saved and copied back into the client area after the window is sized or repositioned.
Public Const SWP_NOOWNERZORDER = &H200 ' Does not change the owner window's position in the Z order.
Public Const SWP_NOREPOSITION = &H200  ' Does not change the owner window's position in the Z order. Same as the SWP_NOOWNERZORDER flag.
Public Const SWP_NOSENDCHANGING = &H400 ' Prevents the window from receiving the WM_WINDOWPOSCHANGING message.

' ��������� ��� ����������
Public Const SBS_HORZ = &H0&
Public Const SBS_VERT = &H1&
Public Const SBS_SIZEBOX = &H8&
Public Const SB_CTL = 2
Public Const SB_THUMBPOSITION = 4
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115

Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10

Enum apiGetWindowLongIndex
    GWL_WNDPROC = -4&
    GWL_HINSTANCE = -6&
    GWL_HWNDPARENT = -8&
    GWL_ID = -12&
    GWL_STYLE = -16&
    GWL_EXSTYLE = -20&
    GWL_USERDATA = -21&
End Enum
Enum apiMouseKeys
    MK_LBUTTON = &H1
    MK_RBUTTON = &H2
    MK_SHIFT = &H4
    MK_CONTROL = &H8
    MK_MBUTTON = &H10
End Enum
Enum apiWheelDeltaConst
    WHEEL_DELTA = 120
End Enum
' ==================================
' ��������� ��� ����� �������������� �������
' ----------------------------------
Public Enum eObjectProps   ' ���� �������������� ����������, ������� �.�. ���������/�������� ��/� Tag
    ltDefault = 0                   ' ��-��������� �����������/����������� ��� ��������� ��� ��������
    ltAll = &HFFFFFFFF              ' ���
' ----------------------------------
' �������/�������
' ----------------------------------
    ltLeft = 1                      ' ������� ����� �������                                         (adhcDefLeft)
    ltRight = 2                     ' ������� ������ �������                                         ...
    ltTop = 4                       ' ������� ������� �������
    ltBottom = 8                    ' ������� ������ �������
    ltWidth = &H10                  ' ������
    ltHeight = &H20                 ' ������
    ltSizeAll = &H3F                ' ��� �������
    
    ltLeftRight = ltLeft + ltRight  ' ������ ������������ ����� � ������ �������
    ltLeftWidth = ltLeft + ltWidth  ' ������ ������������ ����� ������� � ������
    ltRightWidth = ltRight + ltWidth ' ������ ������������ ������ ������� � ������
    
    ltBadHorz = ltLeft + ltRight + ltWidth ' �� ����� ���� ������ ������������ Left, Right and Width
    
    ltTopBottom = ltTop + ltBottom  ' ������ ������������ ������� � ������ �������
    ltTopHeight = ltTop + ltHeight  ' ������ ������������ ������� ������� � ������
    ltBottomHeight = ltBottom + ltHeight ' ������ ������������ ������ ������� � ������
    
    ltBadVert = ltTop + ltBottom + ltHeight ' �� ����� ���� ������ ������������ Top, Bottom and Height
' ----------------------------------
' �����/����������
' ----------------------------------
    ltStyle = &H40                  ' ����� �����������                                             (adhcSizeIt/adhcFloatIt/adhcStyleIt)
    ltColors = &H80                 ' �������� ���������� (Back/Fore/Font/TextColor)                (adhcColor/adhcBackColor/etc)
'' ----------------------------------
'' ������������� (����������������) �������/�������
'    ' ����������� ������������� ���������� (������� �����)
'    ' ����� ����������� ��� �������� ������������ �������/�������� �������� ������ �������/������
'    ' ���� ��� �� ������ � �������� ���� ������� ������� �����/������
'' ----------------------------------
'    ltPropLeft = ltLeft * &H100     '                                                               (adhcDefLeft  [& adhcBoundLeft])
'    ltPropRight = ltRight * &H100   '                                                                ...
'    ltPropTop = ltTop * &H100
'    ltPropBottom = ltBottom * &H100
'    ltPropWidth = ltWidth * &H100
'    ltPropHeight = ltHeight * &H100
'' ----------------------------------
'' ������������� (��������) �������/�������
'    ' ����������� ������������� ���������� (+- ����� �����)
'    ' ����� ����������� ��� �������� ������������ �������/�������� �������� ������ �������/������
'    ' ���� ��� �� ������ � �������� ���� ������� ������� �����/������
'' ----------------------------------
'    ltShiftLeft = ltLeft * &H101    ' = ltPropLeft + ltLeft                                         (adhcDefLeft  [& adhcBoundLeft])
'    ltShiftRight = ltRight * &H101  '                                                                ...
'    ltShiftTop = ltTop * &H101
'    ltShiftBottom = ltBottom * &H101
'    ltShiftWidth = ltWidth * &H101
'    ltShiftHeight = ltHeight * &H101
' ----------------------------------
' ��������
' ----------------------------------
    ltSplit = &H4000                ' ������� �����������                                           (adhcSplitIt)
    ltAction = &H8000               ' �������� ��� �������� ��������                                (adhcAction)
' ----------------------------------
' �������� ���������� �����������
' ----------------------------------
    ltPictAll = &HFFFF0000
    ltPictName = &H10000            '                                                               (adhcObjectName)
    ltPictSize = &H20000            '                                                               (adhcObjectSize)
    ltPictMode = &H40000            '                                                               (adhcObjectMode)
    ltPictPlace = &H80000           '                                                               (adhcObjectAlign)
    ltPictAngle = &H100000          '                                                               (adhcObjectAngle)
    ltPictGray = &H200000           '                                                               (adhcObjectGray)
    ltPictText = &H1000000          '                                                               (adhcObjectText)
    ltTextAlign = &H4000000         '                                                               (adhcObjectTextAlign)
    ltTextPlace = &H8000000         '                                                               (adhcObjectTextPlace)
    ltTextAngle = &H10000000        '                                                               (adhcObjectTextAngle)
    ltFontName = &H40000000         '                                                               (adhcDefFontName)
    ltFontSize = &H80000000         '                                                               (adhcDefFontSize)
    ltPictShow = ltPictName + ltPictText
' ----------------------------------
End Enum
' ----------------------------------
' ��������� ���������
' ----------------------------------
Public Const adhcNone = "None"
Public Const adhcDefault = "Default", adhcDef = "Def"
Public Const adhcLeft = "Left", adhcLeft1 = "L", adhcLeft2 = "X"
Public Const adhcRight = "Right", adhcRight1 = "R"
Public Const adhcTop = "Top", adhcTop1 = "T", adhcTop2 = "Y"
Public Const adhcBottom = "Bottom", adhcBottom1 = "B"
Public Const adhcWidth = "Width", adhcWidth1 = "W"
Public Const adhcHeight = "Height", adhcHeight1 = "H"                       '!!! adhcHeight1 ��������� � adhcHorz2 - ������� �� ����������
Public Const adhcHorz = "Horizontal", adhcHorz1 = "Horz", adhcHorz2 = "H"
Public Const adhcVert = "Vertical", adhcVert1 = "Vert", adhcVert2 = "V"
Public Const adhcFull = "Full", adhcFull1 = "F"
Public Const adhcBoth = "Both"
Public Const adhcYes = "Yes", adhcNo = "No"
Public Const adhcTrue = "True", adhcFalse = "False"
Public Const adhcOff = "Off"
Public Const adhcCenter = "Center", adhcCenter1 = "C"                       ' ������������ �� ������
Public Const adhcTo = "To"                                                  ' ��� �������� � Place ���������� ������ (�..)
Public Const adhcOn = "On"                                                  ' ��� �������� � Place ���������� ������� (��..), ����� ���� �������������� ��� ������ � �������� True
Public Const adhcCascade = "Cascade"                                        ' ��� �������� � Place ����������� ���������� (������ ��� ����)
Public Const adhcFrom = "From"                                              ' ��� ���������� (������ ��� ����)
Public Const adhcProp = "Prop"                                              ' ��� �������� ����� ����������������� �������
Public Const adhcMin = "Min", adhcMax = "Max"
Public Const adhcFont = "Font"
Public Const adhcSize = "Size"
Public Const adhcScale = "Scale"
Public Const adhcMode = "Mode"
Public Const adhcAlign = "Align"
Public Const adhcPlace = "Place"
Public Const adhcAngle = "Angle"
Public Const adhcName = "Name"
Public Const adhcBack = "Back"
Public Const adhcBorder = "Border"
Public Const adhcFore = "Fore"
Public Const adhcAction = "Action"                                          ' ��������
' ������� ���������
Public Const adhcCm = "cm", adhcCm1 = "��"                                  ' ����������
Public Const adhcInch = "in", adhcInch1 = "'"                               ' �����
Public Const adhcPoints = "pt"
Public Const adhcPixels = "px"
Public Const adhcTwips = "tw"
' ----------------------------------
Public Const adhcStyleIt = "StyleIt"            ' c���� ��������. �������� ����� ����������.
Public Const c_strStyleDelims = "+&,"           ' ���������� ����������� �������� �������� StyleIt

Public Const adhcLeftTop = adhcLeft & adhcTop, adhcLeftTop1 = adhcLeft1 & adhcTop1                    ' 5  Left-Top
Public Const adhcRightTop = adhcRight & adhcTop, adhcRightTop1 = adhcRight1 & adhcTop1                ' 6  Right-Top
Public Const adhcLeftBottom = adhcLeft & adhcBottom, adhcLeftBottom1 = adhcLeft1 & adhcBottom1        ' 9  Left-Bottom
Public Const adhcRightBottom = adhcRight & adhcBottom, adhcRightBottom1 = adhcRight1 & adhcBottom1    ' 10 Right-Bottom
Public Const adhcHorTop = adhcHorz & adhcTop, adhcHor1Top = adhcHorz1 & adhcTop, adhcHor2Top1 = adhcHorz2 & adhcTop1 ' 7  Horz-Top
Public Const adhcHorBottom = adhcHorz & adhcBottom, adhcHor1Bottom = adhcHorz1 & adhcBottom, adhcHor2Bottom1 = adhcHorz2 & adhcBottom1 ' 11 Horz-Bottom
Public Const adhcVerLeft = adhcVert & adhcLeft, adhcVer1Left = adhcVert1 & adhcLeft, adhcVer2Left1 = adhcVert2 & adhcLeft1 ' 13 Left-Vert
Public Const adhcVerRight = adhcVert & adhcRight, adhcVer1Right = adhcVert1 & adhcRight, adhcVer2Right1 = adhcVert2 & adhcRight1  ' 14 Right-Vert

Public Const adhcCenterHor = adhcCenter & adhcHorz, adhcCenterHor1 = adhcCenter & adhcHorz1, adhcCenterHor2 = adhcCenter1 & adhcHorz2   ' �������������� ������������ �� ������
Public Const adhcCenterHorTop = adhcCenterHor & adhcTop, adhcCenterHor1Top = adhcCenterHor1 & adhcTop, adhcCenterHor2Top1 = adhcCenterHor2 & adhcTop1
Public Const adhcCenterHorBottom = adhcCenterHor & adhcBottom, adhcCenterHor1Bottom = adhcCenterHor1 & adhcBottom, adhcCenterHor2Bottom1 = adhcCenterHor2 & adhcBottom1
Public Const adhcCenterVer = adhcCenter & adhcVert, adhcCenterVer1 = adhcCenter & adhcVert1, adhcCenterVer2 = adhcCenter1 & adhcVert2   ' ������������ ������������ �� ������
Public Const adhcCenterVerLeft = adhcCenterVer & adhcLeft, adhcCenterVer1Left = adhcCenterVer1 & adhcLeft, adhcCenterVer2Left1 = adhcCenterVer2 & adhcLeft1
Public Const adhcCenterVerRight = adhcCenterVer & adhcRight, adhcCenterVer1Right = adhcCenterVer1 & adhcRight, adhcCenterVer2Right1 = adhcCenterVer2 & adhcRight1

Public Const adhcXProp = adhcLeft2 & adhcProp, adhcLProp = adhcLeft1 & adhcProp   '16  ������� ����� ������� �� ������ ������� (������������� �� ������)
Public Const adhcRProp = adhcRight1 & adhcProp                                         '32  ������� ������ ������� �� ������ ������� (������������� �� ������)
Public Const adhcYProp = adhcTop2 & adhcProp, adhcTProp = adhcTop1 & adhcProp     '64  ������� ������ ������� �� ������ ������� (������������� �� ������)
Public Const adhcBProp = adhcBottom1 & adhcProp                                        '128 ������� ����� ������� �� ������ ������� (������������� �� ������)
Public Const adhcWProp = adhcWidth1 & adhcProp                                         '48  ������ ��������������� ������ �������
Public Const adhcHProp = adhcHeight1 & adhcProp                                        '192 ������ ��������������� ������ �������

Public Const adhcPict = "Pict"                          '1024 �������� ������ ��������
Public Const adhcText = "Text"                          '2048 �������� ������ �����
Public Const adhcIconAndText = adhcPict & adhcText      '3072 �������� �������� � �����
' ----------------------------------
Public Const adhcSizeIt = "SizeIt"              ' �������� ������ (��������� �����-������/�������-������ �������)
Public Const adhcSizeRight = adhcRight          ' ������ �� ���� �����
Public Const adhcSizeBottom = adhcBottom        ' ���� �� ���� �����
Public Const adhcSizeBoth = adhcBoth            ' ������ � ���� �� ���� �����
Public Const adhcSizeNone = adhcNone
' ----------------------------------
Public Const adhcFloatIt = "FloatIt"            ' ���������� (��������� ������/������ ������� � ������)
Public Const adhcFloatRight = adhcRight         ' ������ �� ���� �����
Public Const adhcFloatBottom = adhcBottom       ' ���� �� ���� �����
Public Const adhcFloatBoth = adhcBoth           ' ������ � ���� �� ���� �����
Public Const adhcFloatNone = adhcNone
' ----------------------------------
Public Const adhcScaleIt = "ScaleIt"            ' ��������������
' ----------------------------------
Public Const adhcSplitNone = adhcNone           ' �� �������� ������������
Public Const adhcSplitIt = "SplitIt"            ' ����������� (�������� ������������) (���������� ������� �������)
Public Const adhcDelimIt = "Delimit"            ' ����������� (�������������) (������� ������� ��� �������� ���������)
Public Const adhcSplitVer = adhcVert, adhcSplitVer1 = adhcVert1, adhcSplitVer2 = adhcVert2 ' ������������ �����������
Public Const adhcSplitHor = adhcHorz, adhcSplitHor1 = adhcHorz1, adhcSplitHor2 = adhcHorz2 ' �������������� �����������
'Public Const adhcSplitBoth = adhcBoth           ' ��� (������������ � ��������������) �����������
'Public Const adhcDefault = adhcDefault          ' ��-��������� - �� �������� ������������
' ----------------------------------
' ��� ��������/������� ��������
' ----------------------------------
' adhcDef &
Public Const adhcDefLeft = adhcLeft, adhcDefLeft1 = adhcLeft1, adhcDefLeft2 = "X"   ' ��������� �������� ��-��������� (�.�. ������ � ���������� ��������� ������������ ����, ���� � ����� ������������ �������� �����/������)
Public Const adhcDefTop = adhcTop, adhcDefTop1 = adhcTop1, adhcDefTop2 = "Y"
Public Const adhcDefRight = adhcRight, adhcDefRight1 = "R"
Public Const adhcDefBottom = adhcBottom, adhcDefBottom1 = "B"
Public Const adhcDefWidth = adhcWidth, adhcDefWidth1 = adhcWidth1                   ' ������� �������� ��-��������� (�.�. ������ � ���������� ���������, ���� � ����� ������������ �������� �����/������)
Public Const adhcDefHeight = adhcHeight, adhcDefHeight1 = adhcHeight1
'
Public Const adhcMinWidth = adhcMin & adhcWidth, adhcMinWidth1 = adhcMin & adhcWidth1, adhcMinWidth2 = adhcWidth1 & "0"     ' ���������� ������� (�.�. ������ � ���������� ���������, ���� � ����� ������������ �������� �������� ��-���������)
Public Const adhcMaxWidth = adhcMax & adhcWidth, adhcMaxWidth1 = adhcMax & adhcWidth1, adhcMaxWidth2 = adhcWidth1 & "1"
Public Const adhcMinHeight = adhcMin & adhcHeight, adhcMinHeight1 = adhcMin & adhcHeight1, adhcMinHeight2 = adhcHeight1 & "0"
Public Const adhcMaxHeight = adhcMax & adhcHeight, adhcMaxHeight1 = adhcMax & adhcHeight1, adhcMaxHeight2 = adhcHeight1 & "1"
' ----------------------------------
Private Const adhcBond = "Bond"           ' �������� �������� (��� �������� ������������ ������ �������� ���������� ����������)
Public Const adhcBondLeft = adhcBond & adhcDefLeft, adhcBondLeft1 = "BL", adhcBondLeft2 = "BX"
Public Const adhcBondTop = adhcBond & adhcDefTop, adhcBondTop1 = "BT", adhcBondTop2 = "BY"
Public Const adhcBondRight = adhcBond & adhcDefRight, adhcBondRight1 = "BR"
Public Const adhcBondBottom = adhcBond & adhcDefBottom, adhcBondBottom1 = "BB"
Public Const adhcBondWidth = adhcBond & adhcDefWidth, adhcBondWidth1 = "BW"
Public Const adhcBondHeight = adhcBond & adhcDefHeight, adhcBondHeight1 = "BH"
' ----------------------------------
' ��� ��������������� ����������� �� ���������
' ----------------------------------
Public Const adhcClip = "Clip" ' 0 - �� ������ ������. ���� ������ ������ ������� ������ - �������
Public Const adhcStretch = "Stretch" ' 1 - ������/���������� (�������� ���������)
Public Const adhcZoom = "Zoom" ' 3 - ���������������� ���������������
Public Const adhcDown = "Down", adhcDown1 = adhcZoom & adhcDown ' -1 - ���������������� ���������������, ������ ���������
' ----------------------------------
' ��� ������� �������������� ���������� ����������� � ������
' ----------------------------------
' ������ �� ������
Public Const adhcPlaceCenter = adhcCenter                       ' �� ������ (������)
Public Const adhcPlaceToLeft = adhcTo & adhcLeft                ' ������ ����� �� ������
Public Const adhcPlaceToRight = adhcTo & adhcRight              ' ������ ������ �� ������
Public Const adhcPlaceToTop = adhcTo & adhcTop                  ' ������ �� ������ ������
Public Const adhcPlaceToBottom = adhcTo & adhcBottom            ' ������ �� ������ �����
' ������� �� ������
Public Const adhcPlaceOnLeft = adhcOn & adhcLeft                ' ������� ����� �� ������
Public Const adhcPlaceOnRight = adhcOn & adhcRight              ' ������� ������ �� ������
Public Const adhcPlaceOnTop = adhcOn & adhcTop                  ' ������� �� ������ ������
Public Const adhcPlaceOnBottom = adhcOn & adhcBottom            ' ������� �� ������ �����
' ������ �� ����
Public Const adhcPlaceToLeftTop = adhcTo & adhcLeftTop          ' ������ ����� ������
Public Const adhcPlaceToRightTop = adhcTo & adhcRightTop        ' ������ ������ ������
Public Const adhcPlaceToLeftBottom = adhcTo & adhcLeftBottom    ' ������ ����� �����
Public Const adhcPlaceToRightBottom = adhcTo & adhcRightBottom  ' ������ ������ �����
' ������� �� ����
Public Const adhcPlaceOnLeftToTop = adhcOn & adhcLeft & adhcTo & adhcTop          ' ������� ����� � �������� ����
Public Const adhcPlaceOnLeftToBottom = adhcOn & adhcLeft & adhcTo & adhcBottom    ' ������� ����� � ������� ����
Public Const adhcPlaceOnRightToTop = adhcOn & adhcRight & adhcTo & adhcTop        ' ������� ������ � �������� ����
Public Const adhcPlaceOnRightToBottom = adhcOn & adhcRight & adhcTo & adhcBottom  ' ������� ������ � ������� ����
Public Const adhcPlaceOnTopToLeft = adhcOn & adhcTop & adhcTo & adhcLeft          ' ������� � ������ ���� ������
Public Const adhcPlaceOnTopToRight = adhcOn & adhcTop & adhcTo & adhcRight        ' ������� � ������� ���� ������
Public Const adhcPlaceOnBottomToLeft = adhcOn & adhcBottom & adhcTo & adhcLeft    ' ������� � ������ ���� �����
Public Const adhcPlaceOnBottomToRight = adhcOn & adhcBottom & adhcTo & adhcRight  ' ������� � ������� ���� �����
'' ���������� (������ ��� ����� ??)
Public Const adhcCascadeFromLeftTop = adhcCascade & adhcFrom & adhcLeftTop               ' ���������� �������� ������-����
Public Const adhcCascadeFromRightTop = adhcCascade & adhcFrom & adhcRightTop             ' ���������� �������� �����-����
Public Const adhcCascadeFromLeftBottom = adhcCascade & adhcFrom & adhcLeftBottom         ' ���������� �������� ������-�����
Public Const adhcCascadeFromRightBottom = adhcCascade & adhcFrom & adhcRightBottom       ' ���������� �������� �����-�����
' ----------------------------------
' ��� ��������� ���������
' ----------------------------------
Public Const adhcColor = "Color"
Public Const adhcBackColor = adhcBack & adhcColor
Public Const adhcForeColor = adhcFore & adhcColor
Public Const adhcFontColor = adhcFont & adhcColor
Public Const adhcTextColor = adhcText & adhcColor
Public Const adhcBorderColor = adhcBorder & adhcColor
Public Const adhcColorBlack = "Black"
Public Const adhcColorWhite = "White"
Public Const adhcColorGray = "Gray"
Public Const adhcColorDark = "Dark"        ' =appColorDark
Public Const adhcColorDark2 = adhcColorDark & 2
Public Const adhcColorDark3 = adhcColorDark & 3
Public Const adhcColorBright = "Bright"    ' =appColorBright
Public Const adhcColorBright2 = adhcColorBright & 2
Public Const adhcColorBright3 = adhcColorBright & 3
Public Const adhcColorLight = "Light"      ' =appColorLight
Public Const adhcColorLight2 = adhcColorLight & 2
Public Const adhcColorLight3 = adhcColorLight & 3
' ----------------------------------
' ��� ���������� �����������
' ----------------------------------
Public Const adhcObjectName = adhcPict 'PictName
Public Const adhcObjectSize = adhcPict & adhcSize 'PictSize
Public Const adhcObjectMode = adhcPict & adhcMode, adhcObjectMode1 = adhcObjectSize & adhcMode 'PictMode
Public Const adhcObjectAlign = adhcPict & adhcPlace 'PictAlign
Public Const adhcObjectAngle = adhcPict & adhcAngle 'PictAngle
Public Const adhcObjectText = adhcText 'TextString
Public Const adhcObjectTextAlign = adhcText & adhcAlign
Public Const adhcObjectTextPlace = adhcText & adhcPlace 'TextPlace
Public Const adhcObjectTextAngle = adhcText & adhcAngle 'TextAngle
Public Const adhcObjectGray = adhcText & adhcColorGray, adhcObjectGray1 = adhcObjectGray & adhcScale 'GrayScale
'
Public Const adhcFontName = adhcFont & adhcName
Public Const adhcFontSize = adhcFont & adhcSize
' ----------------------------------
Public Const COLORNOTSET = &HFFFFFFFF

' ==================================
Public Function AccWinHide()
On Error Resume Next
    DoCmd.SelectObject acTable, , True
    DoCmd.RunCommand acCmdWindowHide
    Err.Clear
End Function
Public Function AccWinShow()
On Error Resume Next
    DoCmd.SelectObject acTable, , True
    'DoCmd.RunCommand acCmdWindowUnhide
    Err.Clear
End Function

Public Function IsActiveControl(ctl As Control) As Boolean
Const PTR_SIZE = 4
Dim o As Object
Dim ptrAct As Long 'active control ptr
Dim ptrCtl As Long 'control ptr
Dim Ret As Boolean
    On Error Resume Next
    Set o = Screen.ActiveControl
    If Not (o Is Nothing) Then
        CopyMemory ptrAct, o, PTR_SIZE
        CopyMemory ptrCtl, ctl, PTR_SIZE
        If ptrAct = ptrCtl Then Ret = True
    End If
    'IsActiveControl = Screen.ActiveControl Is ctl
    IsActiveControl = Ret
End Function

Public Function IsSubForm(frm As Form) As Boolean
' ��������� ������� �� ����� ��� ��������
    On Error Resume Next
Dim strName As String: strName = frm.PARENT.Name
    IsSubForm = (Err.Number = 0): Err.Clear
End Function

Public Function IsSubformFocus() As Boolean
' ��������� ����� � ��������
    On Error Resume Next
Dim ctl As Object: Set ctl = Screen.ActiveControl.PARENT
    If Not TypeOf ctl Is Access.Form Then Set ctl = ctl.PARENT
    Set ctl = ctl.PARENT
    IsSubformFocus = Not CBool(Err.Number)
    Err.Clear
End Function

Public Function IsFormExists(FormName As String) As Boolean
' ��������� ���������� �� �����
Dim Result As Boolean:  Result = False
    On Error GoTo HandleError
    Result = (CurrentProject.AllForms(FormName).Name = FormName) '
HandleExit:  IsFormExists = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function IsFormLoaded(FormName As String) As Boolean
' ��������� ��������� �� �����
Dim Result As Boolean:  Result = False
    On Error GoTo HandleError
    If SysCmd(acSysCmdGetObjectState, acDefault, FormName) = acObjStateClosed Then Err.Raise vbObjectError + 512
    Result = (Application.Forms(FormName).CurrentView <> 0)  ' <> Design
HandleExit:  IsFormLoaded = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function IsFormOpened(FormName As String, _
    Optional View As AcFormView, Optional state As acFormState) As Boolean
' ��������� ������� �� �����
' State =   0 - Closed
'           1 - (acObjStateOpen)    Open
'           2 - (acObjStateDirty)   Changed but not saved
'           4 - (acObjStateNew)     New
' View  =   0 - Normal View
'           1 - Design View
'           3 - Datasheet View
Dim Result As Boolean:  Result = False
    On Error GoTo HandleError
    state = SysCmd(acSysCmdGetObjectState, acDefault, FormName): If state = acObjStateClosed Then Err.Raise vbObjectError + 512
    Select Case Application.Forms(FormName).CurrentView
    Case 0: View = acDesign
    Case 1: View = acNormal
    Case 2: View = acFormDS
    Case Else: View = -1
    End Select
HandleExit:  IsFormOpened = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function IsReportExists(FormName As String) As Boolean
' ��������� ���������� �� �����
Dim Result As Boolean:  Result = False
    On Error GoTo HandleError
    Result = (CurrentProject.AllReports(FormName).Name = FormName) '
HandleExit:  IsReportExists = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function IsReportLoaded(ReportName As String) As Boolean
' ��������� �������� �� �����
Dim Result As Boolean:  Result = False
    On Error GoTo HandleError
    If SysCmd(acSysCmdGetObjectState, acDefault, ReportName) = acObjStateClosed Then Err.Raise vbObjectError + 512
    Result = (Application.Reports(ReportName).CurrentView <> 0) ' <> Design
HandleExit:  IsReportLoaded = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function IsReportOpened(ReportName As String, _
    Optional View As AcFormView, Optional state As acFormState) As Boolean
' ��������� ������ �� �����
' State =   0 - Closed
'           1 - Open
' View  =   0 - Normal View
'           1 - Design View
Dim Result As Boolean:  Result = False
    On Error GoTo HandleError
    state = SysCmd(acSysCmdGetObjectState, acDefault, ReportName): If state = acObjStateClosed Then Err.Raise vbObjectError + 512
    Select Case Application.Reports(ReportName).CurrentView
    Case 0: View = acDesign
    Case 1: View = acNormal
    Case Else: View = -1
    End Select
HandleExit:  IsReportOpened = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function FormOpen( _
    FormName As String, _
    Optional View As AcFormView = acNormal, _
    Optional WhereCondition, _
    Optional DataMode As AcFormOpenDataMode = acFormPropertySettings, _
    Optional WindowMode As AcWindowMode = acWindowNormal, _
    Optional OpenArgs, _
    Optional PARENT, _
    Optional x, Optional y, _
    Optional Placement As ePlace = eCascadeFromLeftTop, _
    Optional Icon, _
    Optional Visible As Boolean = True, _
    Optional NewForm As Access.Form _
    ) As Boolean
'    Optional FilterName,
' ��������� ����� � �������� ������
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
'----------------------------------
' ������������ ����� ��������
'----------------------------------
'        Select Case View
'        Case acNormal: .DefaultView = 0            ' ����������
'        Case acFormDS: .DefaultView = 2            ' ����� �������
'        Case acFormPivotTable: .DefaultView = 3    ' ����� ������� �������
'        Case acFormPivotChart: .DefaultView = 4    ' ����� ������� ���������
'        Case acPreview ' ���� �� ��������������    ' ����� ��������� ������
'        Case acDesign  ' ���� �� ��������������    ' ����� ��������� ������
'        End Select
    ' ��������� ������� ��������� �����
Dim acState As acFormState: acState = SysCmd(acSysCmdGetObjectState, acForm, FormName)
    If (acState <> acObjStateClosed) Then
    ' ����� ������������ ����� ������ ��� �������� ������ ��� ��������
        Set NewForm = Application.Forms(FormName)
        If (View = acDesign) Xor (NewForm.CurrentView = 0) Then
    ' ����� �������, �� �������� �����, ���� ������� - ����� ������������
        ' ����� ������������ ����� ������ ��� �������� ������ ��� �������� - ���������
            DoCmd.Close acForm, FormName, Save:=acSaveYes: acState = acObjStateClosed
        End If
    End If
    If acState = acObjStateClosed Then
' ��������� ����� �� ����� � ������ ������
Dim tmpWinMode As AcWindowMode: tmpWinMode = acHidden '  WindowMode
        DoCmd.OpenForm FormName, View, _
                WhereCondition:=WhereCondition, DataMode:=DataMode, _
                WindowMode:=tmpWinMode, OpenArgs:=OpenArgs ',FilterName:=FilterName
    End If
'----------------------------------
' ���������� ������ �� ������
'----------------------------------
Dim i As Long
    For i = Application.Forms.Count - 1 To 0 Step -1
        Result = (Application.Forms(i).Name = FormName):   If Result Then Set NewForm = Application.Forms(i):   Exit For
    Next i
    If Not Result Then: Err.Raise vbObjectError + 512
'----------------------------------
' ����� �������� �������
'----------------------------------
    Result = AccObjectSet(NewForm, _
        View:=View, _
        WhereCondition:=WhereCondition, _
        DataMode:=DataMode, _
        WindowMode:=WindowMode, _
        OpenArgs:=OpenArgs, _
        x:=x, y:=y, _
        ObjectParent:=PARENT, _
        Icon:=Icon, _
        Visible:=Visible, _
        Placement:=Placement)
    NewForm.SetFocus
    'WindowUnFreeze  ' ������������ ����������
HandleExit:  DoCmd.Echo True: FormOpen = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function FormOpenDrop( _
    FormName As String, _
    Optional FormVal, _
    Optional PARENT, _
    Optional x, Optional y, _
    Optional Placement As ePlace = ePlaceOnBottomToLeft, _
    Optional Icon, _
    Optional Visible As Boolean = True, _
    Optional ByRef NewForm As Object _
    )
Const c_strProcedure = "FormOpenDrop"
' ��������� ���������� ����� ���� FormMode
' � NewForm ���������� ������ �� �������� ����� ()
Dim tmpPoint As POINT ' ���������� ������� �����
Dim Result

    Result = False
    On Error GoTo HandleError
    If IsMissing(PARENT) Then
        GetCursorPos tmpPoint: x = tmpPoint.x: y = tmpPoint.y ' �� ��������� ������� � ���������� �������
    ElseIf PARENT Is Nothing Then
        GetCursorPos tmpPoint: x = tmpPoint.x: y = tmpPoint.y ' �� ��������� ������� � ���������� �������
    End If
    Result = FormOpen(FormName, _
        WindowMode:=acDialog, x:=x, y:=y, Icon:=Icon, Placement:=Placement, _
        PARENT:=PARENT, NewForm:=NewForm, _
        Visible:=False): If Result = 0 Then Err.Raise vbObjectError + 512
    On Error Resume Next ':GoTo HandleError
    With NewForm
Dim bolModal As Boolean: bolModal = (.ModalResult = .ModalResult)
        If IsMissing(FormVal) Then Result = PARENT.Value Else Result = FormVal
        .Value = Result
        Err.Clear
'        On Error GoTo HandleError
' ������� �� �����
        .Visible = Visible: If Visible Then .SetFocus
' ������ � ����� ���� �� ������� ������ ������������
    If bolModal Then
        Do While .Visible: DoEvents: Loop: If .ModalResult = vbOK Then Result = .Value
        DoCmd.Close acForm, NewForm.Name, acSaveNo: Set NewForm = Nothing ' ���� ��������� ��������� - ����� ��������� ������ ���������
    End If
    End With
' ���������� ��������� � �������� ����
    If TypeOf PARENT Is Access.Control Then PARENT.Value = Result
    Err.Clear
HandleExit:  FormOpenDrop = Result: Exit Function
HandleError: Result = False:  Err.Clear: Resume HandleExit
End Function

Public Function FormOpenContext( _
    ContextData As String, _
    Optional ByRef ContextMenu As Object, _
    Optional ByRef ContextVal, _
    Optional ByRef ContextName As String, _
    Optional PARENT, _
    Optional x, Optional y, _
    Optional Arrange As eAlign = eAlignLeftTop, _
    Optional Visible As Boolean = True)
' ������� � ��������� ����������� ����
Const c_strProcedure = "FormOpenContext"
' ContextData - �������� ��������� ���� ��� ��� ������� ��������� ���������
' ContextMenu - ������ �� ����������� (��������) ����
' ContextVal - �������� ������������ ���� (������������ ��-��������� ��� ������������)
' ContextName - ��� ������������ ���� (��-��������� - "~tmpContextMenu")
' Parent  - ������ �� ������������ ������
' X, Y - ���������� ������ ����
' Visible - ���������� ��������� ����������� ���� ������� ��� �� �������
' Arrange - ��� ������������ ���� ���-�� ��������� (��-��������� ���������� ������ ������� ����� ���� ����)
Const cstrContextName = "~tmpContextMenu"
Dim mnu As clsContextMenu
Dim strWhere As String
Dim Ret As Long
    ContextName = Trim$(ContextName)
    'If Len(ContextName) = 0 Then ContextName = cstrContextName
' ������� ����������� ����
    Set mnu = New clsContextMenu 'Set ContextMenu = Application.CommandBars.Add(Name:=ContextName, Position:=msoBarPopup)
    With mnu ' ContextMenu
        .CreateContextMenu ContextName
    ' ��������� ContextData
'Stop
On Error Resume Next
    ' ��� ������������ ���� �� ������� SysMenu
        If IsNumeric(ContextData) Then Ret = .CreateItemsFromSQL(c_strTableMenu, WhereCond:=c_strParent & sqlEqual & ContextData): GoTo HandleShow
Dim dbs As DAO.Database: Set dbs = CurrentDb
Dim rst As DAO.Recordset
Dim strSQL As String
    ' ������� ��� ������������ ���� �� ������� SysMenu
        strSQL = sqlSelectAll & c_strTableMenu & sqlWhere & c_strCName & sqlEqual & "'" & ContextData & "'"
        Set rst = dbs.OpenRecordset(strSQL): If Err Then Err.Clear Else Ret = .CreateItemsFromSQL(c_strTableMenu, WhereCond:=c_strParent & sqlEqual & rst.Fields(c_strKey)): GoTo HandleShow
    ' ��� �������/�������/�������� �������
        Set rst = dbs.OpenRecordset(ContextData): If Err Then Err.Clear Else Ret = .CreateItemsFromSQL(ContextData): GoTo HandleShow
    ' ������ ���������
        Ret = .CreateItemsFromString(ContextData)
HandleShow:
' ������� � ��� ������
        If Ret Then
            .ShowMenu x, y ', Arrange ' ���������� ����
            ContextVal = .Value
        Else
            .RemoveContextMenu ContextName: Set mnu = Nothing
        End If
    End With
    'Result = ContextVal
    Set ContextMenu = mnu
    'DoCmd.Echo True' �������� ����������� �� ������
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Public Function ReportOpen( _
    ReportName As String, _
    Optional View As AcFormView = acViewPreview, _
    Optional WhereCondition, _
    Optional WindowMode As AcWindowMode = acWindowNormal, _
    Optional OpenArgs, _
    Optional PARENT, _
    Optional Placement As ePlace = eCascadeFromLeftTop, _
    Optional NewReport As Access.Report _
    ) As Boolean
' ��������� ����� � ��������� �����������
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
'----------------------------------
' ������������ ����� ��������
'----------------------------------
'        Select Case View
'        Case acNormal: .DefaultView = 0            ' ����������
'        Case acFormDS: .DefaultView = 2            ' ����� �������
'        Case acFormPivotTable: .DefaultView = 3    ' ����� ������� �������
'        Case acFormPivotChart: .DefaultView = 4    ' ����� ������� ���������
'        Case acPreview ' ���� �� ��������������    ' ����� ��������� ������
'        Case acDesign  ' ���� �� ��������������    ' ����� ��������� ������
'        End Select
    ' ��������� ������� ��������� �����
Dim acState As acFormState: acState = SysCmd(acSysCmdGetObjectState, acReport, ReportName)
    If (acState <> acObjStateClosed) Then
    ' ����� ������������ ����� ������ ��� �������� ������ ��� ��������
        Set NewReport = Application.Reports(ReportName)
        If (View = acDesign) Xor (NewReport.CurrentView = 0) Then
    ' ����� ������, �� �������� �����, ���� ������� - ����� ������������
        ' ����� ������������ ����� ������ ��� �������� ������ ��� �������� - ���������
            DoCmd.Close acReport, ReportName, Save:=acSaveYes: acState = acObjStateClosed
        End If
    End If
    If acState = acObjStateClosed Then
' ��������� ����� �� ����� � ������ ������
        DoCmd.OpenReport ReportName, View, _
            WhereCondition:=WhereCondition, _
            WindowMode:=WindowMode, OpenArgs:=OpenArgs ',FilterName:=FilterName
    End If
'----------------------------------
' ���������� ������ �� ������
'----------------------------------
Dim i As Long
    For i = Application.Reports.Count - 1 To 0 Step -1
        Result = (Application.Reports(i).Name = ReportName): If Result Then Set NewReport = Application.Reports(i): Exit For
    Next i
'----------------------------------
' ����� �������� �������
'----------------------------------
    If Not Result Then: Err.Raise vbObjectError + 512
    Result = AccObjectSet(NewReport, _
        View:=View, _
        WhereCondition:=WhereCondition, _
        WindowMode:=WindowMode, _
        OpenArgs:=OpenArgs, _
        ObjectParent:=PARENT, _
        Placement:=Placement)
    'WindowUnFreeze  ' ������������ ����������
HandleExit:  DoCmd.Echo True: ReportOpen = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function AccObjectSet(AccObject As Object, _
    Optional View As AcFormView = acNormal, _
    Optional WhereCondition, _
    Optional DataMode As AcFormOpenDataMode = acFormPropertySettings, _
    Optional WindowMode As AcWindowMode = acWindowNormal, _
    Optional OpenArgs, _
    Optional ObjectParent, _
    Optional x, Optional y, _
    Optional Placement As ePlace = eCascadeFromLeftTop, _
    Optional Icon, _
    Optional Visible As Boolean = True, _
    Optional ModalResult As VbMsgBoxResult, _
    Optional FormResult) As Boolean
' ����� �������� ������� Access (����� ��� �����)
'---------------------
'���������� ������ �� �������� ������
'accObject - ������ �� ������ Access ����� ��� �����
'View - (��������������) ����� ��������� �����
'FilterName -
'WhereCondition - (��������������) ��������� ��������� ��������������� SQL WHERE ��� WHERE.
'DataMode - (��������������) ����� ������� � ������ �����. �������������� �������� ������� AllowEdits, AllowDeletions, AllowAdditions � DataEntry
'WindowMode - (��������������) ����� �������� ����
'OpenArgs - (��������������) ������ ���������� ������������ ����������� �����
'ObjectParent - (��������������) ������ �� ������������ ����� ��� ������ �����
'ModalResult - (��������������) ���������� �������� (VbMsgBoxResult) ������� ������ � ����� �������� ��� ���� �������
'ObjectResult - (��������������) ��������� ������������ ������ ����� ���������� ������
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    
    If TypeOf AccObject Is Access.Form Then
    ElseIf TypeOf AccObject Is Access.Report Then
    Else: Err.Raise vbObjectError + 512
    End If
    
'    'WindowFreeze Application.hWndAccessApp
    DoCmd.Echo False
    On Error Resume Next
    With AccObject '
    ' ����� ���� ��� ������ ����������, � ���� ���:
    ' ������� �� �����, �������� ��������� � ������ �������� ���� �����
'----------------------------------
' ��������� � ����������� �����
'----------------------------------
        '.Parametres = OpenArgs  ' ��������� ����� (�������� ������ ����������)
        If Not IsMissing(WhereCondition) Then If Len(WhereCondition) > 0 Then .Filter = WhereCondition: .FilterOn = True
        Err.Clear
'----------------------------------
' ������������ ����� ������ ������
'----------------------------------
        Select Case DataMode
        Case acFormPropertySettings    ' ������������ ���������� �����
        Case acFormEdit        ' ����� �������������� ������
        '.DataEntry = False
            .RecordsetType = 0  ' Dynaset
            .AllowEdits = True: .AllowAdditions = True: .AllowDeletions = True
        Case acFormAdd         ' ����� ����� ������
            .DataEntry = True   ' ������ ������
            .RecordsetType = 1  ' Dynaset
            .AllowEdits = False: .AllowAdditions = True: .AllowDeletions = False
        Case acFormReadOnly    ' ����� ��������� ������
        '.DataEntry = False
            .RecordsetType = 2  ' Snapshot
            .AllowEdits = False: .AllowAdditions = False: .AllowDeletions = False
        End Select
'----------------------------------
' ������������ ����� �������� ����
'----------------------------------
        Select Case WindowMode
        Case acWindowNormal:    WindowSet .hwnd, SW_SHOWNORMAL: .Visible = True ' ������� ����          '.SetFocus
        Case acIcon:            WindowSet .hwnd, SW_SHOWMINIMIZED               ' ��������� � ������    '.Visible = True: .SetFocus
        Case acDialog:          .Visible = False: .PopUp = True: .Modal = True  ' ���������� ����       '.Visible = Visible ': .SetFocus:WindowSet .hWnd, SW_SHOWNORMAL'.BorderStyle = 3   ' ������ � ������������
        Case acHidden:          .Visible = False                                ' �������               'WindowSet .hWnd, SW_HIDE
        End Select
'----------------------------------
' ������������ ���������� ����
'----------------------------------
        If WindowMode = acDialog Then   ' ���������� ����
        '
        End If
HandlePlaceForm:
'----------------------------------
' ������������� ���� ������������ ��������� ��������
'----------------------------------
        
        On Error GoTo HandleObjectDesign
' �������� ����� �������� � ��������� ���������� ��������� ����� ������������ ��������
Dim Cascade As Boolean: Cascade = ((Placement And eCascade) = eCascade)
Dim rXpar As Single, rYpar As Single: Call GetAlignPoint(Placement Mod &H10, rXpar, rYpar)    ' �� �������� � �������
Dim rX�li As Single, rY�li As Single: Call GetAlignPoint(Placement \ &H10, rX�li, rY�li)      ' �� ������� � ��������
Dim Xcli As Long, Ycli As Long, dX As Long, dY As Long
' �������� ����������/������� ��������� �����
Dim cliRect As RECT: GetWindowRect AccObject.hwnd, cliRect
' �������� ����������/������� ������������� �������
Dim parRect As RECT:
        If IsMissing(ObjectParent) Then
    ' �������� �� ����� - ��������� ������������ ������� ������� Access
            GetWindowRect FindWindowEx(hWndAccessApp, 0&, accClassChild, vbNullString), parRect ' MDIClient
        ElseIf ObjectParent Is Nothing Then
            GetWindowRect FindWindowEx(hWndAccessApp, 0&, accClassChild, vbNullString), parRect ' MDIClient
        ElseIf TypeOf ObjectParent Is Access.Form Or TypeOf ObjectParent Is Access.Report Then
    ' �������� ����� ��� ����� - ��������� ������������ ������� ���� �������� �����/������
            GetWindowRect ObjectParent.hwnd, parRect                                            ' Access Form/Report
        ElseIf TypeOf ObjectParent Is Access.Control Then
    ' �������� ������� - ��������� ������������ ������� ��������� ��������
            With parRect: AccControlLocation ObjectParent, .Left, .Top, .Right, .Bottom: .Right = .Left + .Right: .Bottom = .Top + .Bottom: End With
        Else: Err.Raise vbObjectError + 512
        End If
    ' ��������� �������� ��� ���������� ���������� ����
        If Cascade Then dX = pxWinGap: dY = pxWinGap
    ' x, y ������������� ��� �������� �� ����� ��������
        If Not IsMissing(x) Then dX = dX + x
        If Not IsMissing(y) Then dY = dY + y
    ' � ����������� �� Placement ��������� ����� ������������ ��������
        Xcli = parRect.Left + rXpar * (parRect.Right - parRect.Left) - rX�li * (cliRect.Right - cliRect.Left) + dX
        Ycli = parRect.Top + rYpar * (parRect.Bottom - parRect.Top) - rY�li * (cliRect.Bottom - cliRect.Top) + dY
' ���� �������� ������ �������� ������������� ���� �� ��������� ������� ���������
        FormMove AccObject, Xcli, Ycli ' �������������
'----------------------------------
' ����� ���������� �����
'----------------------------------
HandleObjectDesign:
' ...
    On Error Resume Next
Dim tmp, i As Long
        For i = acDetail To acFooter
            tmp = GetColorFromText(TaggedStringGet(.Section(i).Tag, adhcBackColor)): Err.Clear
            If IsNull(tmp) Then tmp = GetColorFromText(TaggedStringGet(.Tag, adhcBackColor)): Err.Clear
            If IsNumeric(tmp) Then .Section(i).BackColor = tmp
        Next i
        If Not IsMissing(Icon) Then
            With AccObject: Call PictureData_SetIcon(.hwnd, Icon): End With
        End If
' ...
' ToDo: ������� ���������� �������������� ����������/���������������� ��������� �����
' !!! ���� �������� ������������ ���������� ����� � ������ ���������
Dim ctl As Control
        For Each ctl In .Controls
            AccControlSet ctl, Init:=True
        Next ctl
'----------------------------------
    End With
    Result = True
    'WindowUnFreeze 'AccObject.hwnd ' ������������ ����������
HandleExit:  DoCmd.Echo True: AccObjectSet = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function AccControlSet( _
    ctl As Access.Control, _
    Optional x, Optional y, _
    Optional w, Optional h, _
    Optional r, Optional b, _
    Optional Show, _
    Optional Init As Boolean = False, _
    Optional frm As Access.Form) As Boolean
' ������� ������� � �������� �������, � ����������� ��� ������� ���
Const c_strProcedure = "AccControlSet"
' ctl       - ������������� �������
' x/y/r/b/w/h - �������/������� ��� ������ �������� ��������
' Relative  - ������� ������������� �������/�������� ��������
'             ����� ����� ������ ��� Show=False - ����� ���������� ��������� ������� �������/������� � ��� ����� ����������
' Show      - ������� ��������� ��������
' Init      - ������� ������������� ��������
' frm       - ������ �� ������������ �����
'-------------------------
' v.1.0.0       : 17.07.2023 - �������� ������
'-------------------------
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
Const cShift = px \ 3       ' �������� �������� ������� ���� ������� �� ������ �� ������
Dim bTry As Byte: bTry = 0  ' ������� ������� �������� �������
    If frm Is Nothing Then Set frm = GetTopParent(ctl, True)  ' ???
Dim sec As Access.Section: Set sec = frm.Section(ctl.Section)
'Dim bInit As Boolean: bInit = bInit And Show
    
    With ctl
'----------------------------------
' ������ � ��������� ��� �������� �������� ���������� � ����
'----------------------------------
    If Init Then
    If Len(.Tag) > 0 Then
'Stop
Dim cTags As New Collection, aKeys() As String
        'Set cTags = New Collection
        Call TaggedString2Collection(.Tag, cTags, aKeys)
    On Error Resume Next
Dim Key: For Each Key In aKeys
            Select Case Key
            Case adhcBackColor:     .BackColor = GetColorFromText(cTags(Key)) ': Err.Clear
            Case adhcBorderColor:   .BorderColor = GetColorFromText(cTags(Key)) ': Err.Clear
            Case adhcForeColor, adhcFontColor, adhcTextColor: _
                                    .ForeColor = GetColorFromText(cTags(Key))
                                    If Err Then Err.Clear:  .FontColor = GetColorFromText(cTags(Key))
                                    If Err Then Err.Clear:  .TextColor = GetColorFromText(cTags(Key))
        ' etc
            End Select
            Err.Clear
        Next Key
    End If
    End If
'----------------------------------
    On Error GoTo HandleError
'----------------------------------
' ����� ������� ��������
'----------------------------------
Dim xCtl As Long, yCtl As Long, wCtl As Long, hCtl As Long
        If Show Then
        ' �������� ��������� ������� ��������
        ' ����  >1      - ���� ��� ���������� �������� � tw,
        ' ���� [0..1]   - ���� ������������ �������� �����/������ ��������
            If IsMissing(x) Then xCtl = .Left Else If x > twMinLim Then xCtl = x Else xCtl = frm.Width * x
            If IsMissing(y) Then yCtl = .Top Else If y > twMinLim Then yCtl = y Else yCtl = sec.Height * y
            If IsMissing(w) Then wCtl = .Width Else If w > twMinLim Then wCtl = w Else wCtl = frm.Width * w
            If IsMissing(h) Then hCtl = .Height Else If h > twMinLim Then hCtl = h Else hCtl = sec.Height * h
        ' ���� ������� �� ������ �������� ������� �����/������
            If xCtl + wCtl > frm.Width Then frm.Width = xCtl + wCtl
            If yCtl + hCtl > sec.Height Then sec.Height = yCtl + hCtl
        Else
' !!! ����� - ���� ������ ����� �������� ������� �������� ����� ��� ��� �������������� ��� ���������
'  �������� ���� ���������� ��� ������� ��������� ������ � ���, � ����� ���������������
            xCtl = 0: yCtl = 0: wCtl = 0: hCtl = 0
        End If
        ' ����� ��������� ��������
        If Not IsMissing(Show) Then .Visible = Show
        ' ����� ������� ��������
        .Width = wCtl: .Height = hCtl: .Left = xCtl:
HandleNextTry: .Top = yCtl: On Error Resume Next
    End With
    Result = True
HandleExit:  AccControlSet = Result: Exit Function
HandleError: ' ���� ������� �� ����� ���� �������� � ��������� ����� - ������� ��������� ������
    If Err.Number = 2100 And bTry < 3 Then Err.Clear: sec.Height = sec.Height + cShift: bTry = bTry + 1: Resume HandleNextTry
    Result = False: Err.Clear: Resume HandleExit
End Function

Public Function FormMove( _
    AccForm, _
    ByRef x As Long, ByRef y As Long, _
    Optional Width, Optional Height, _
    Optional Arrange As eAlign = eAlignLeftTop, _
    Optional Inscribe As Boolean = True) As Long
' ������������� �����c������ ����� � �������� ����������
' � ��������� �� ���������� ���������� ������� ���� Access
Const c_strProcedure = "FormMove"
Dim accHwnd As LongPtr ', frmHwnd As LongPtr
Dim accPoint As POINT 'accX As Long, accY As Long
Dim frmRect As RECT, accRect As RECT

    On Error GoTo HandleError
' �������� ���������� ������� ����
    With AccForm
        GetWindowRect .hwnd, frmRect
    ' ��������� ����� �� �������� �� ���������� ������� ������� Access
        ' ����������� ����� �������� �� �����, - ��� ��������� ������ Access
        If Not .PopUp Then
'Stop    ' ����� ���� ����������� �� �� ����� ��������� PopUp - ���? - ���� ���������
        ' ������ �������� �� ���������� ������� ������� Access
            ' ���� ���������� ������� ���� Access
            accHwnd = FindWindowEx(hWndAccessApp, 0&, accClassChild, vbNullString) ' MDIClient
            GetWindowRect accHwnd, accRect
            ' �������� �������� ���������� �������� ������ ���� ���������� ����� ������� ������� ������� Access
            accPoint.x = 0: accPoint.y = 0
            ClientToScreen accHwnd, accPoint
            ' ���������� ������ Access ��� ������� ���������������� �� ������
            x = x - accPoint.x: y = y - accPoint.y
'        Else
        End If
    ' ���� �� ������� ������� ����� ���������� �������
        If IsMissing(Height) Or IsMissing(Width) Then
            GetWindowRect .hwnd, frmRect
            With frmRect
                Width = .Right - .Left
                Height = .Bottom - .Top
            End With
        End If
    ' ������������ x (���������� �������������)
        If (Arrange And eCenterHorz) = eCenterHorz Then
            x = x + Width \ 2 ' ������������
         ElseIf (Arrange And eLeft) = eLeft Then
            'x = x           ' ��������� �� ������ (���������� ������ �� �����)
         ElseIf (Arrange And eRight) = eRight Then
            x = x - Width     ' ��������� �� ������� (���������� ����� �� �����)
        End If
    ' ������������ y (���������� �����������)
        If (Arrange And eCenterVert) = eCenterVert Then
            y = y + Height \ 2 ' ������������
         ElseIf (Arrange And eTop) = eTop Then
            'y = y           ' ��������� �� �������� (���������� ���� �����)
         ElseIf (Arrange And eBottom) = eBottom Then
            y = y - Height     ' ��������� �� ������� (���������� ���� �����)
        End If
        If Inscribe Then
        ' ���� ������� ������� � ������� �������
        ' �������� �������� Arrange ����� ����� ��� ���������� � �����
        End If
    ' ������������� �����
        MoveWindow .hwnd, x, y, Width, Height, 1
    End With

HandleExit:

    Exit Function
HandleError:
'    Dbg.Error Err.Number, Err.Description, Err.Source, Erl(), c_strModule & "." & c_strProcedure
    Err.Clear
    Resume HandleExit
End Function
Public Function PosAccForm2Client( _
    AccForm As Access.Form, _
    ByRef x As Long, ByRef y As Long, _
    Optional Section As AcSection = acDetail)
' ��������� ���������� ����� Access � ���������� ��������� ������� (� ��������)
' ����� AccFormCoords2Client
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
Const lSelectorWidth = 18
Dim v As Long, h As Long
Dim lBorder As Long
Dim accPoint As POINT
Dim hwnd As Long
' ��������� ���������� �������� � ����� � ������ � �������
    With AccForm
        hwnd = .hwnd
        Select Case .DefaultView
        Case 1, 2 '��������� ��� ���������
            x = x + .CurrentSectionLeft ', DIRECTION_HORIZONTAL)
            y = y + .CurrentSectionTop ', DIRECTION_VERTICAL)
        Case Else
        End Select
    End With
    accPoint.x = TwipsToPixels(x, DIRECTION_HORIZONTAL) '(X) / lTwipsPerInch * lLogPixelPerInchX
    accPoint.y = TwipsToPixels(y, DIRECTION_VERTICAL)   '(Y) / lTwipsPerInch * lLogPixelPerInchY

    '����� ���������
    ScrollbarGetPos AccForm.hwnd, v, h
    accPoint.x = accPoint.x - h
    If Section = acDetail Then accPoint.y = accPoint.y - v  '������������ ��������� ������ ������ �� Detail
    
    x = accPoint.x: y = accPoint.y
    Result = True
HandleExit:  PosAccForm2Client = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function AccControlLocation(AccObject As Variant, _
    Optional ByRef x, Optional ByRef y, Optional ByRef w, Optional ByRef h, _
    Optional ParentObject, _
    Optional ClientAreaPos As Boolean) As Boolean
' ���������� ������ � ������� (��������/����������) ������� Access (� ��������)
'----------------
' ctl - ������ �� ����������� ������ Access
' x/y - ������������ ���������� �������� (� ��������) � ����������� �� RetType ��������/����������
' w/h - ������������ ������/������ �������� (� ��������)
' ParentObject  - ������ �� ������� ������������ �����/����� ������� ����������� �������
' ClientAreaPos - ���������� ��� ������������ ��������� ��������
'       False - �������� (��-���������)
'       True  - ���������� � ���������� ������� �����
'----------------
' v.0.2.2       : 08.08.2021 - ��������� ����������� ���������� ��������/��������� ����������
' v.0.2.1       : 15.04.2019 - ���������� ���������������� � ������� �����
'----------------
' � ����������� ������� ������� c�������� ctl.accLocation X, Y, W, H, varChild
' �� - � ��������� ����� (��������� ���� � Access 2003) ��� ������� ������� �� ���������
' ������ �� � ������ ������ ��������� ������ ������ ���������
' ��-�� ���� �������� �� ������ � ������ ����������
' ������� ������� ��-�������:
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
Dim lpPoint As POINT
'Dim frm As Access.Form
Dim cX As Long, cY As Long, cW As Long, ch As Long
Dim dX As Long, dY As Long, dW As Long, dH As Long
' ���������� ������������ ����� � ����������/������� ����������� ��������
    If TypeOf AccObject Is Access.Control Then
        Set ParentObject = GetTopParent(AccObject)
        With AccObject: cX = .Left: cY = .Top: cW = .Width: ch = .Height: End With
        With ParentObject
        dX = .CurrentSectionLeft
        If AccObject.Section <> acHeader Then
    ' ���� ��� �� ��������� �����
        ' ��� ������� ����� ��������� ������ ����������� ������
        ' ��� ��������� ����� ��������� ���������� �� �������� ���� �����
            On Error Resume Next: dY = .Section(acHeader).Height: Err.Clear: On Error GoTo HandleError
            Select Case AccObject.Section
            Case acDetail: If (.DefaultView = 1 Or .DefaultView = 2) Then dY = .CurrentSectionTop
            Case acFooter: dY = dY + .Section(acDetail).Height
            End Select
        End If
        End With
    ElseIf TypeOf AccObject Is Access.Form Or TypeOf AccObject Is Access.Report Then
        Set ParentObject = AccObject
        cX = 0: cY = 0
        '' ������� ���� �����
        'cW = frm.InsideWidth: cH = frm.InsideHeight: End With
        '' ������� �����/ ���� ���� ��������� � ���������� ��������� �� ������ � ������ ������ Detail
        With ParentObject
            On Error Resume Next
            dH = .Section(acHeader).Height + .Section(acFooter).Height: Err.Clear
            On Error GoTo HandleError
            cW = .Width: ch = dH + .Section(acDetail).Height
        End With
    Else: Err.Raise vbObjectError + 512
    End If
' ���� ���� - �������� ����� ������� (�������-����� ���� ���������� ����� �����)
    If Not ClientAreaPos Then ClientToScreen ParentObject.hwnd, lpPoint    ' ���������� ������� ����� � �������� �����������
' ���������� ���������� �� ���� ����� �� ��������
    If Not IsMissing(x) Then x = lpPoint.x + TwipsToPixels(cX + dX, DIRECTION_HORIZONTAL)
    If Not IsMissing(y) Then y = lpPoint.y + TwipsToPixels(cY + dY, DIRECTION_VERTICAL)
' �������� ������� ��������
    If Not IsMissing(w) Then w = TwipsToPixels(cW, DIRECTION_HORIZONTAL)
    If Not IsMissing(h) Then h = TwipsToPixels(ch, DIRECTION_VERTICAL)
    Result = True
HandleExit:  AccControlLocation = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Function PosAccForm2Screen( _
    AccForm As Access.Form, _
    ByRef x As Long, ByRef y As Long, _
    Optional Section As AcSection = acDetail)
' ��������� ���������� ����� Access � ���������� ������ (� ��������)
' ����� AccFormCoords2Screen
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
Const lSelectorWidth = 18
Dim v As Long, h As Long
Dim lBorder As Long
Dim accPoint As POINT
Dim hwnd As Long
' ��������� ���������� �������� � ����� � ������ � �������
' ���� ������� � ������ �������� � ��������� �����
    With AccForm
        hwnd = .hwnd
        Select Case .DefaultView
         Case 1, 2 '��������� ��� ���������
            x = x + .CurrentSectionLeft ', DIRECTION_HORIZONTAL)
            y = y + .CurrentSectionTop ', DIRECTION_VERTICAL)
         Case Else
        End Select
    End With
    accPoint.x = TwipsToPixels(x, DIRECTION_HORIZONTAL) '(X) / lTwipsPerInch * lLogPixelPerInchX
    accPoint.y = TwipsToPixels(y, DIRECTION_VERTICAL)   '(Y) / lTwipsPerInch * lLogPixelPerInchY

    '���� ���� �������������� - ���� ��������� ��� ������
    'If frm.RecordSelectors Then cltPoint.X = cltPoint.X + lSelectorWidth
    
    ClientToScreen& hwnd, accPoint '��������� ���������� ���������� � ��������
    
    'am 030407_10:39:00  --begin-- **************
    '����� ���������
    ScrollbarGetPos AccForm.hwnd, v, h
    accPoint.x = accPoint.x - h
    If Section = acDetail Then accPoint.y = accPoint.y - v  '������������ ��������� ������ ������ �� Detail
    
    x = accPoint.x: y = accPoint.y
    Result = True
HandleExit:  PosAccForm2Screen = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function FormGetClientSize( _
    ByRef aForm As Access.Form, _
    ByRef cliWidth As Long, ByRef cliHeight As Long, _
    Optional sec As AcSection = -1, _
    Optional ByVal InTwips As Boolean = True) _
    As Boolean
' ���������� ������� ������� ������� ������� ����� Access/������ ����� � �������� ��� ������
Const c_strProcedure = "FormGetClientSize"
' Sec = acHeader, acDetail, acFooter - ������ ������, ����� �� �������� - ������ ���������� ������� �����

' ������ ���������� ������� ����� �������� �������� ���:
' RecordSelectors(w=19px) � NavigationButtons (h=18px), ScrollBars (h/w=17px)
' ������� ������ �������� ���������� ������� ����� ��� �������� ����� ����� ����� �������� ������
' ������ ���������� ������� ����� = ����� ����� ������ + (���-�� ������� ������-1)*1px
' ������ ��� ���� ���� "OFormSub" ����������� ���� ����� hWnd � ��������� �������:
' 0-Header; 1-Detail; 2-Footer. ������������� ������ �� ����� ����� ���� ������� = 0
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    cliWidth = 0: cliHeight = 0
Dim hwnd As LongPtr: hwnd = aForm.hwnd
Dim hSec As LongPtr: hSec = FindWindowEx(hwnd, 0, accClassFormClient, vbNullString): If hSec = 0 Then Err.Raise vbObjectError + 512
Dim s As Byte, sMax As Byte: s = 0: sMax = 2
Dim hRect As RECT, tmp As Long
    With hRect
        Do
            GetWindowRect hSec, hRect: tmp = (.Bottom - .Top)
            Select Case sec
            Case acHeader: If s = 0 Then cliHeight = tmp: Exit Do
            Case acDetail: If s = 1 Then cliHeight = tmp: Exit Do
            Case acFooter: If s = 2 Then cliHeight = tmp: Exit Do
            Case Else
                If cliHeight > 0 And tmp > 0 Then tmp = tmp + 1 '1px - ����� ��������?
                cliHeight = cliHeight + tmp
            End Select
            hSec = FindWindowEx(hwnd, hSec, accClassFormClient, vbNullString): If hSec = 0 Then Exit Do
            s = s + 1
        Loop Until s > sMax
        cliWidth = (.Right - .Left)
    End With
' ��������� ������� � �����
    If InTwips Then cliWidth = cliWidth * px: cliHeight = cliHeight * px
    Result = True
HandleExit:  FormGetClientSize = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function

Public Sub GetScreenInfo( _
    Optional ScrResX As Long, Optional ScrResY As Long, _
    Optional ScrDpiX As Long, Optional ScrDpiY As Long, _
    Optional DesResX As Long, Optional DesResY As Long, _
    Optional DesDpiX As Long, Optional DesDpiY As Long)
' ������ �������/���������� �������� ������
    On Error GoTo HandleError
Dim lpDC As LongPtr: lpDC = GetDC(HWND_DESKTOP): If lpDC = 0 Then GoTo HandleExit
Dim rcScr As RECT: Call SystemParametersInfo(SPI_GETWORKAREA, 0, rcScr, 0)
' ������ ������� ������� ������ � �������� (�������)
    With rcScr: ScrResX = (.Right - .Left): ScrResY = (.Bottom - .Top): End With
' ���������� ������ � �������� (�������)
    If Not IsMissing(ScrDpiX) Then ScrDpiX = GetDeviceCaps(lpDC, LOGPIXELSX)
    If Not IsMissing(ScrDpiY) Then ScrDpiY = GetDeviceCaps(lpDC, LOGPIXELSY)
    Call ReleaseDC(HWND_DESKTOP, lpDC)
'    ScrTppX = TwipsPerInch / ScrDpiX: ScrTppY = TwipsPerInch / ScrDpiY
On Error Resume Next
Dim tmpString As String
' ������ ������� ������� ������ � �������� (��� ����������)
    If Not IsMissing(DesResX) Or Not IsMissing(DesResY) Then
        CustomPropertyGet c_strDesignRes, tmpString, CurrentProject  ', dbText
        DesResX = Split(tmpString, c_strResDelim)(0): If DesResX = 0 Then DesResX = ScrResX
        DesResY = Split(tmpString, c_strResDelim)(1): If DesResY = 0 Then DesResY = ScrResY
    End If
' ���������� ������ � �������� (��� ����������)
    If Not IsMissing(DesDpiX) Or Not IsMissing(DesDpiY) Then
        CustomPropertyGet c_strDesignDpi, tmpString, CurrentProject ', dbText
        DesDpiX = Split(tmpString, c_strResDelim)(0): If DesDpiX = 0 Then DesDpiX = ScrDpiX
        DesDpiY = Split(tmpString, c_strResDelim)(1): If DesDpiY = 0 Then DesDpiY = ScrDpiY
    End If
    Err.Clear
HandleExit:  Exit Sub
HandleError: Err.Clear: Resume HandleExit
End Sub
Public Sub FixScreenInfo()
' ��������� �������/���������� �������� ������ � �������� ���������� ��� ��������� ����������
    On Error GoTo HandleError
Dim lpDC As LongPtr: lpDC = GetDC(HWND_DESKTOP): If lpDC = 0 Then GoTo HandleExit
Dim tmpString As String
' ��������� ���������� ������ � �������� (�������)
Dim rcScr As RECT: Call SystemParametersInfo(SPI_GETWORKAREA, 0, rcScr, 0)
    With rcScr
    tmpString = (.Right - .Left) & c_strResDelim & (.Bottom - .Top)
    CustomPropertySet c_strDesignRes, tmpString, CurrentProject ', dbText
    End With
' ��������� ������ ������� ������� ������ � �������� (�������)
    tmpString = GetDeviceCaps(lpDC, LOGPIXELSX) & c_strResDelim & GetDeviceCaps(lpDC, LOGPIXELSY)
    CustomPropertySet c_strDesignDpi, tmpString, CurrentProject ', dbText
    Call ReleaseDC(HWND_DESKTOP, lpDC)
HandleExit:  Exit Sub
HandleError: Err.Clear: Resume HandleExit
End Sub
Public Function ScreenRes(Optional xRes, Optional yRes)
' �������� ���������� ������
Dim hdc As LongPtr: hdc = GetDC(GetDesktopWindow()): If hdc = 0 Then Exit Function
    xRes = GetDeviceCaps(hdc, HORZRES)
    yRes = GetDeviceCaps(hdc, VERTRES)
End Function
Public Function ScreenSize(Optional xSize, Optional ySize)
' �������� ������� ������
Dim hdc As LongPtr: hdc = GetDC(GetDesktopWindow()): If hdc = 0 Then Exit Function
    xSize = GetDeviceCaps(hdc, HORZSIZE)
    ySize = GetDeviceCaps(hdc, VERTSIZE)
End Function

Public Function FormGetSectionHwnd(hWndForm As LongPtr, Seciton As Long, Optional twWidth As Long, Optional twHeight As Long) As LongPtr
'*********************************************************
'����������:����� Access ������� �� �������� ����,
'���������� �������� �������� ����� �������� hwnd �����.
'� ����� ���� ���������� �������� ���� - �� ������ �� ������ ��
'Section �����. ��� ������ ���� ���� - OFormSub
'��� �������� ���� ������ �� ������ ����� �� ���� - ������� �������
'�� �� ������������- �� ���� ����� ������� -�����,
'����� ������ - ����� �����
'�����:hwndForm - hwnd ����� (�������� ���� �����)
'Seciton - ����� ������, ����� ���� ����� 0,1,2
'am v1.0.0_030407_11:47:16
'http://am.rusimport.ru
'appto:a_mitin@app.ru
'*********************************************************
Const c_strProcedure = "FormGetSectionHwnd"
On Error GoTo HandleError
Dim hwnd As LongPtr, sec(2) As LongPtr
Dim Result As LongPtr
Dim Top(2) As Long
Dim minVal As Long, maxVal As Long
Dim minIdx As Long, maxIdx As Long
Dim avgIdx As Long
Dim i As Long
Dim r As RECT
    minIdx = -1
    hwnd = GetWindow(hWndForm, GW_CHILD)
    Do While hwnd <> 0
        If StrComp(GetWinClass(hwnd), accClassFormClient, vbTextCompare) = 0 Then
            sec(i) = hwnd
            GetWindowRect hwnd, r
            Top(i) = r.Top
            If minIdx = -1 Or minVal >= Top(i) Then
                minIdx = i
                minVal = Top(i)
                maxIdx = i
                maxVal = Top(i)
            End If
            If maxVal < Top(i) Then
                maxIdx = i
                maxVal = Top(i)
            End If
            i = i + 1
        End If
        hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    Loop
    avgIdx = 3 - (maxIdx + minIdx)
    Select Case Seciton
    Case acDetail:  Result = sec(avgIdx)
    Case acFooter:  Result = sec(maxIdx)
    Case acHeader:  Result = sec(minIdx)
    Case Else:      Err.Raise vbObjectError + 512, , "��� ������ " & Seciton & " �� ��������������!"
    End Select
    GetWindowRect Result, r
    twWidth = PixelsToTwips(r.Right - r.Left, DIRECTION_HORIZONTAL)
    twHeight = PixelsToTwips(r.Bottom - r.Top, DIRECTION_VERTICAL)
HandleExit: FormGetSectionHwnd = Result: Exit Function
HandleError: Result = 0: Err.Clear: Resume HandleExit
End Function

Public Function AdjustHeight(obj As Object) As Boolean
' ����������� ������������ ������ �����/������ �� �������� ������
Const c_strProcedure = "AdjustHeight"
Dim o As Object
    On Error GoTo HandleError
    If TypeOf obj Is Access.Form Or TypeOf obj Is Access.Report Then
        Set o = obj
    ElseIf TypeOf obj.PARENT Is Access.Form Or TypeOf obj.PARENT Is Access.Report Then
        Set o = obj.PARENT
    Else
        GoTo HandleExit
    End If
    
Dim tmp As Long: tmp = 0
Dim i As Byte
    On Error Resume Next
    For i = acDetail To acGroupLevel2Footer:  tmp = tmp + o.Section(i).Height:  Next i
    o.InsideHeight = tmp
    Err.Clear
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Public Function FormGetInsideHeight(frm As Form)
' �������� ���������� ������ �����
    On Error Resume Next
Dim i&, Result&
    Result = 0
    With frm
        For i = acDetail To acGroupLevel2Footer
            Result = Result + .Section(i).Height
            Err.Clear
        Next i
    End With
    FormGetInsideHeight = Result
End Function

Public Function FormSectionsCount(frm As Access.Form) As Long
Dim i&, Result&, s$
    Result = 0
    With frm
        For i = acDetail To acGroupLevel2Footer
            Err.Clear
            s = .Section(i).Name
            If Err = 0 Then Result = Result + 1
        Next i
    End With
    FormSectionsCount = Result
End Function

Public Function GetOpenedObjects(Optional ListOfObjectNames As String)
' ���������� ������ �������� �������� Access
'Const c_strProcedure = "GetOpenedObjects"
'Dim arrNames() As String
'Dim i As Long
'Dim obj, nam
'Dim ObjName As String
'Dim bolCheckList As Boolean
'Dim bolAddName As Boolean
'Dim Result() As Access.Form
'
'    On Error GoTo HandleError
'    bolCheckList = Len(Trim(ListOfObjectNames))
'    ' ���� ������ ������� �� ������
'    If bolCheckList Then arrNames = Split(ListOfObjectNames, ";")
'    For Each obj In Application.Forms
'' ��������� �������� �����
'        bolAddName = False
'        ObjName = obj.Name
''        If IsFormOpened(strFrmName) Then
''            IsFormLoaded = Forms(strFrmName).CurrentView <> conDesignView
''        End If
'
'        If IsFormLoaded(ObjName) Then
'            If bolCheckList Then
'    ' ��������� ����� �� ������
'                For Each nam In arrNames
'                    bolAddName = ObjName = nam
'                    If bolAddName Then Exit For
'                Next nam
'            Else
'                bolAddName = True
'            End If
'        End If
'        If bolAddName Then Call InsertElementIntoArray(Result, obj) 'ObjName)
'    Next obj
'    For Each obj In Application.Reports
'' ��������� �������� ������
'        bolAddName = False
'        ObjName = obj.Name
'        If IsReportLoaded(ObjName) Then
'            If bolCheckList Then
'    ' ��������� ����� �� ������
'                For Each nam In arrNames
'                    bolAddName = ObjName = nam
'                    If bolAddName Then Exit For
'                Next nam
'            Else
'                bolAddName = True
'            End If
'        End If
'        If bolAddName Then Call InsertElementIntoArray(Result, obj) 'ObjName)
'    Next obj
'HandleExit:
'    GetOpenedObjects = Result
'    Exit Function
'HandleError:
'    'Result = Null
''    Dbg.Error Err.Number, Err.Description, Err.Source, Erl(), c_strModule & "." & c_strProcedure
'    Err.Clear
'    Resume HandleExit
End Function

Public Sub ListboxUnSelectAll(lst As ListBox)
'Precondition: MultiSelect > 0
'None: 0 (Default) Multiple selection isn't allowed.
'Simple: 1
'Extended: 2
Dim varItem As Variant

    For Each varItem In lst.ItemsSelected
        lst.Selected(varItem) = False
    Next
End Sub

Public Sub ListboxSelectAll(lst As ListBox)
'Precondition: MultiSelect > 0
'None: 0 (Default) Multiple selection isn't allowed.
'Simple: 1
'Extended: 2
    Dim lngRow As Long

    For lngRow = 0 To lst.ListCount - 1
        lst.Selected(lngRow) = True
    Next
End Sub

Public Function ListboxToInStr( _
    ListCtrl As ListBox, _
    Optional dtFormat As eFieldFormat = vbLong, _
    Optional ListDelim As String = ",") As String
'���������� ������ �������� ������ � ����������������
'����������� �������� ������ ��������� In(...) ��� Not In (...)
'  � ����������� �� ����������� ����������� � �� �����������
' � �����
'   - "True" ���� �������� ��� ��������
'   - "False" ���� ��� ��������� �������
Const c_strProcedure = "ListboxToInStr"

'Dim UseSelected As Boolean  '��� �������� - ���������� ��� �� ����������
'Dim strIn As String         '��������� - In (...) ��� Not In (...)
'Dim i As Integer, iMax As Integer, iSel As Integer '������ �������� ������
'Dim tmpVal
'Dim Result As String
'On Error GoTo HandleError
''�������� �� ������� ��������� ����. ������
'    strIn = sqlIn
'    Result = vbNullString 'sqlFalse
'    UseSelected = True
'    With ListCtrl
'        i = 0: iMax = .ListCount: iSel = .ItemsSelected.Count
'        If iSel = 0 Then
'            Result = sqlFalse
'            GoTo HandleExit
'         ElseIf iSel = iMax Then
'            Result = sqlTrue
'            GoTo HandleExit
'        End If
'    '���������� ���� ������ ����������� ��� �� �����������
'    ' (�� ���� ������ - ������ ����� ���� ������������ (��������)
'    ' ����� ��������� �� Not IN () ������ �� ����� ������������� ������)
''        UseSelected = (iSel < CInt(0.5*iMax ))
'    '������ ������
'        Do While i < iMax
'            If .Selected(i) = UseSelected Then
'                tmpVal = .ItemData(i)
'                If IsNull(tmpVal) Then
'                    Result = Result & sqlNull & ListDelim
'                'ElseIf Len(tmpVal) = 0 Then
'                '    Result = Result & sqlNull & ListDelim
'                 Else
'                    Select Case dtFormat
'                     Case vbString  '��������� ��������
'                        Result = Result & "'" & Replace(CStr(tmpVal), """", """""") & "'" & ListDelim
'                     Case vbInteger, vbLong, vbByte '�������� ��������
'                        Result = Result & CStr(tmpVal) & ListDelim
'                     Case vbDate 'date ��������
'                        Result = Result & DateToSQL(CDate(tmpVal)) & ListDelim
'                     Case vbDateTimeJ 'dateTime ��������
'                        Result = Result & DateTimeToSQL(CDate(tmpVal)) & ListDelim
'                     Case vbBoolean 'boolean ��������
'                        Result = Result & IIf(CBool(tmpVal), sqlTrue, sqlFalse) & ListDelim    '& "#"
'                     Case vbSingle, vbDouble, vbCurrency '� ���������� ������
'                        Result = Result & Replace(CStr(tmpVal), ",", ".") & ListDelim
'                     Case Else
'                        Result = vbNullString '" ����������� In (...) ��� ������ ���� ������ �� �������������"
'                        GoTo HandleExit
'                    End Select
'                End If
'            End If
'            i = i + 1
'        Loop
'    End With
'    If Right$(Result, Len(ListDelim)) = ListDelim Then Result = Trim$(Left$(Result, Len(Result) - Len(ListDelim)))
'    If Len(Result) = 0 Then GoTo HandleExit
'    If Not UseSelected Then strIn = sqlNot & LTrim$(strIn)
'    Result = strIn & "(" & Result & ")"
'HandleExit:
'    ListboxToInStr = Result
'    Exit Function
'HandleError:
'    Dbg.Error Err.Number, Err.Description, Err.Source, Erl(), c_strModule & "." & c_strProcedure
'    Err.Clear
'    Result = vbNullString
'    Err.Clear
'    Resume HandleExit
End Function

Public Function IsControlExists(frm As Form, ctlName As String) As Boolean
    Const c_strProcedure = "IsControlExists"
    On Error Resume Next
Dim strValue As String
    ' If you can retrieve the value, the such control exists.
    strValue = frm.Controls(ctlName).Name
    IsControlExists = (Err.Number = 0)
    Err.Clear
End Function

Public Function GetTopParent(AccObject, Optional AllowSubForms As Boolean) As Object 'Access.Form
' ���������� ������ �� ������� ������������ �����(�����) ��� �������
' AllowSubForms - ���� True ����� ��������������� �� �����/������ � ��� ����� �������� ��� ��������(-�����),
'                 ����� ����� ����������� ���� ���� �� ����� ����� (�����) ��� ���������
Dim Result As Object
    On Error GoTo HandleError
' ���������� ��������� �������� ���� �� �������� �� �������� (����� ��� ������)
    Set Result = AccObject
    Do While (TypeOf Result Is Access.Control): Set Result = Result.PARENT: Loop
    If AllowSubForms Then GoTo HandleExit
    On Error Resume Next
' ���������� ��������� ����� ��� ������ ���� �� �� ������ �� ������� ����� (��� ������)
Dim strName As String
    Do
        strName = Result.ParentName
        If Err Then Exit Do
        Set Result = Result.PARENT
    Loop
    Err.Clear
HandleExit:  Set GetTopParent = Result: Exit Function
HandleError: Resume HandleExit
End Function

Public Function ActiveControlHwndGet() As LongPtr
Dim CtlHwnd As LongPtr
' Windows ���������� ��������� ��������
' ��������� hWnd � ���������� ���������� app.AccessCtlHwnd
' ������ ��������� � ������� GotFocus �������� ����� ���� �������
    CtlHwnd = GetFocus() ' hWnd ���������� ������ � ������ ��������� ������
    'app.AccessCtlHwnd = ctlHwnd
    ActiveControlHwndGet = CtlHwnd
End Function

Public Function ControlHwndGet(ctl As Access.Control) As LongPtr
' Windows ���������� ��������
    ctl.SetFocus ' ������� Access ����� hWnd ������ ���� � ��� �����
    ControlHwndGet = ActiveControlHwndGet
End Function

Public Function ControlRectGet(ctl As Access.Control) As RECT
' Windows ���������� ��������
Dim ControlRect As RECT
    'ctl.SetFocus
    'GetWindowRect ControlHwndGet(ctl), ControlRect
    ' ��������� Access hWnd ������������� ����
    ' � ������ ������� GotFocus.
    GetWindowRect ControlHwndGet(ctl), ControlRect
    ControlRectGet = ControlRect
End Function

'=================================
Public Function WindowMove(hwnd As LongPtr, Optional x, Optional y, Optional Arrange As eAlign = eAlignLeftTop, Optional Inscribe As Boolean = True) As Long
' ���������������� ���� �� ������
' frm - ��������������� �����
' x,y - �������� ���������� ����� ������������ ������� ��������������� �����
' Arrange - ����� ����������������
' Inscribe - ���� ����������, - ��� ����������������
'   ����� ����������� ��� ����� ����� ���������� � ������� ������� Access
Dim accHwnd As LongPtr
Dim accRect As RECT, accHeight As Long, accWidth As Long    ' ������� ���������� ������� Access
Dim winRect As RECT, winHeight As Long, winWidth As Long    ' ������� ���������� �����
Dim xPos As Long, yPos As Long
    On Error GoTo HandleError
    
' ���� ���������� ������� ���� Access
    accHwnd = FindWindowEx(hWndAccessApp, 0&, accClassChild, vbNullString) ' MDIClient
    GetWindowRect accHwnd, accRect
    With accRect
        accHeight = (.Bottom - .Top)
        accWidth = (.Right - .Left)
    End With
' �������� ���������� ������� ����
    GetWindowRect hwnd, winRect
    With winRect
        winHeight = (.Bottom - .Top)
        winWidth = (.Right - .Left)
    End With
' ������������ x (���������� �������������)
    If (Arrange And eCenterHorz) = eCenterHorz Then
    ' ������������
        xPos = x + winWidth \ 2
    ElseIf (Arrange And eLeft) = eLeft Then
    ' ��������� �� ������ (���������� ������ �� �����)
        xPos = x
    ElseIf (Arrange And eRight) = eRight Then
    ' ��������� �� ������� (���������� ����� �� �����)
        xPos = x - winWidth
    End If
' ������������ y (���������� �����������)
    If (Arrange And eCenterVert) = eCenterVert Then
    ' ������������
        yPos = y + winHeight \ 2
    ElseIf (Arrange And eTop) = eTop Then
    ' ��������� �� �������� (���������� ���� �����)
        yPos = y
    ElseIf (Arrange And eBottom) = eBottom Then
    ' ��������� �� ������� (���������� ���� �����)
        yPos = y - winHeight
    End If
' �������� �������� ����� ����� ��� ���������� � �����
    If Inscribe Then
    ' ���� ������� ������� � ������� �������
        If xPos < accRect.Left Then xPos = accRect.Left
        If xPos + winWidth > accRect.Right Then xPos = accRect.Right - winWidth
        If yPos < accRect.Top Then yPos = accRect.Top
        If yPos + winHeight > accRect.Bottom Then yPos = accRect.Bottom - winHeight
    End If
    MoveWindow hwnd, xPos, yPos, winWidth, winHeight, 1
HandleExit:  Exit Function 'x = xPos: y = yPos
HandleError: Resume HandleExit
End Function
Public Function WindowSet(hwnd As LongPtr, nCmdShow As Long) As Long
' ������������� ����� ����
Dim loX  As Long
Dim loForm As Object
Dim Result As Boolean: Result = False
    On Error Resume Next
    If hwnd = 0 Then GoTo HandleExit
    Set loForm = Screen.ActiveForm
    If Err <> 0 Then '��� �������� �����
        If nCmdShow = SW_HIDE Then
            MsgBox "�� ���� ������ ���� ���� ����������� ����� �� ������"
        Else
            loX = ShowWindow(hwnd, nCmdShow): Err.Clear
        End If
     Else
        If nCmdShow = SW_SHOWMINIMIZED And loForm.Modal = True Then
            MsgBox "�� ���� �������������� ����" & (loForm.Caption + " ") & "���� ����� �� ������"
        ElseIf nCmdShow = SW_HIDE And loForm.PopUp <> True Then
           MsgBox "�� ���� ������ ���� " & (loForm.Caption + " ") & "���� ����� �� ������"
        Else
            loX = ShowWindow(hwnd, nCmdShow)
        End If
    End If
    Result = (loX <> 0)
HandleExit:  WindowSet = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
'==============================
' �������������� ��������
'==============================
Public Function AccClientSize(Optional Width As Long, Optional Height As Long) As RECT
' ���� ���������� ������� ���� Access
Dim tmpHwnd As LongPtr
Dim tmpRect As RECT
    tmpHwnd = FindWindowEx(hWndAccessApp, 0&, accClassChild, vbNullString)
    GetWindowRect tmpHwnd, tmpRect
'    GetClientRect hWndAccessApp, tmpRect
    With tmpRect
        Width = PixelsToTwips(.Right - .Left, DIRECTION_HORIZONTAL)
        Height = PixelsToTwips(.Bottom - .Top, DIRECTION_VERTICAL)
    End With
    AccClientSize = tmpRect
End Function

Public Function ScrollbarGetPos(FormHwnd As LongPtr, ByRef VSB_Pos As Long, ByRef HSB_Pos As Long)
'����������:�������� ������� ������������ ��������� (���� ����)
On Error GoTo HandleError
Dim hWndVSB As LongPtr, hWndHSB As LongPtr
    hWndHSB = ScrollbarGetHwnd(FormHwnd, SBS_HORZ): If hWndHSB = 0 Then HSB_Pos = 0 Else HSB_Pos = GetScrollPos(hWndHSB, SB_CTL)
    hWndVSB = ScrollbarGetHwnd(FormHwnd, SBS_VERT): If hWndVSB = 0 Then VSB_Pos = 0 Else VSB_Pos = GetScrollPos(hWndVSB, SB_CTL)
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function

Public Function ScrollbarSetPos(FormHwnd As LongPtr, VSB_Pos As Long, HSB_Pos As Long)
' ������������� ��������� ��������� (���� ����)
On Error GoTo HandleError
Dim hWndVSB As LongPtr, hWndHSB As LongPtr
    hWndHSB = ScrollbarGetHwnd(FormHwnd, SBS_HORZ)
    hWndVSB = ScrollbarGetHwnd(FormHwnd, SBS_VERT)
    If hWndHSB <> 0 And HSB_Pos >= 0 Then Call SendMessage&(FormHwnd, WM_HSCROLL, (HSB_Pos * 2 ^ 16) Or SB_THUMBPOSITION, 0) 'SetScrollPos hwndHSB, SB_CTL, HSB_Pos, 1
    If hWndVSB <> 0 And VSB_Pos >= 0 Then Call SendMessage&(FormHwnd, WM_VSCROLL, (VSB_Pos * 2 ^ 16) Or SB_THUMBPOSITION, 0)  'SetScrollPos hwndVSB, SB_CTL, VSB_Pos, 1
HandleExit:  Exit Function
HandleError: Resume HandleExit
End Function
Public Function ScrollbarGetHwnd(FormHwnd As LongPtr, ScrollbarType As Integer) As LongPtr
'����������:�������� hwnd ��������� �����
On Error GoTo HandleError
Dim hWndChild As LongPtr
Dim s As String
Dim Style&
    '������� ���� ����� ����� - � ���� ����������
    hWndChild = GetWindow(FormHwnd, GW_CHILD)
    If hWndChild = 0 Then
        ScrollbarGetHwnd = 0
    Else
        Do
            s = GetWinClass(hWndChild)
            If StrComp(s, "SCROLLBAR", vbTextCompare) = 0 Then
                '��� �������� - �������� ���
                Style& = GetWindowLong&(hWndChild, GWL_STYLE)
                If (Style& And SBS_SIZEBOX) = False _
                        And (Style& And &H1) = SBS_HORZ Then
                    '��������������
                    If ScrollbarType = 0 Then
                        '�����
                        ScrollbarGetHwnd = hWndChild ' GetScrollPos(hwndChild, SB_CTL)
                        Exit Function
                    End If
                End If
                If (Style& And &H1) = SBS_VERT Then
                    '������������
                    If ScrollbarType = 1 Then
                        '�����
                        ScrollbarGetHwnd = hWndChild 'GetScrollPos(hwndChild, SB_CTL)
                        Exit Function
                    End If
                End If
            End If
            hWndChild = GetWindow(hWndChild, GW_HWNDNEXT)
        Loop While hWndChild <> 0
    End If
HandleExit:  Exit Function
HandleError: Resume HandleExit
End Function
Private Function StrZ(Par As String) As String
Dim nSize As Long, i As Long ', Rez As String
   nSize = Len(Par)
   i = InStr(1, Par, Chr(0)) - 1
   If i > nSize Then i = nSize
   If i < 0 Then i = nSize
   StrZ = Mid$(Par, 1, i)
End Function
Public Function GetWinClass(hwnd As LongPtr) As String
' �������� ����� ����
Dim Buff As String, BuffSize As Long
    Buff = Space(255)
    BuffSize = GetClassName(hwnd, Buff, 255)
    GetWinClass = StrZ(Buff)
End Function
Public Function GetWinText(hwnd As LongPtr) As String
' �������� ����� ����
Dim Buff As String, BuffSize As Long
    BuffSize = GetWindowTextLength(hwnd) + 1            ' ����� ���������
    Buff = Space(BuffSize)                              ' ����������� ��������� �����
    BuffSize = GetWindowText(hwnd, Buff, BuffSize)      ' �������� ���������
    GetWinText = StrZ(Buff)
End Function
Public Function ChildWindowProc(ByVal hwnd As LongPtr, ByVal lp As Long) As LongPtr
' ���������� �������� EnumChildWindows
' ��������� hWnd � lp ���������� �������� EnumChildWindows
Dim RetVal As Long
'Static i As Long
'Dim strClass As String
'Dim strText As String
    
    RetVal = 1
'    i = i + 1
'
'    strClass = GetWinClass(hWnd) ' ��� ������ ����
'    strText = GetWinText(hWnd)
'    Select Case strClass
'     Case accClassFormWindow = "OForm"       ' ����� ���� ����� Access
'     Case accClassFormClient = "OFormSub"    ' ����� ����� Access
'        ' ������� ������ ����� ����, ��������
'     Case accClassFormPopup = "OFormPopup"   ' ����� ����������� ����� Access
'     Case accClassFormChild = "OFormChild"   ' ����� ����������� ����� Access
'     Case accClassFormNoClose = "OFormNoClose"
'     Case accClassFormClientChild = "OFEDT"  ' ����� ������������ ���� ����� Access
'     Case accClassTableClientChild = "OGNUM" ' ����� ������������ ���� ��������� ����� Access
'     Case accClassRecordSlector = "OSUI"     ' ����� ����� ��������� ������� Access
'     Case accClassTextbox = "OKttbx"         ' ����� ���������� ����� Access
'    End Select
' Debug.Print i; "Window: "; hWnd; " Class: "; strClass; " Text: "; strText
    ChildWindowProc = RetVal ' 1 - ���������� ������������ 0 - ��������
End Function

'=========================
' �������������� ��������
'=========================
Public Function TwipsToPixels(ByVal lngTwips As Long, lngDirection As Long) As Long
Const c_strProcedure = "PixelsToTwips"
On Error GoTo HandleError
'   Function to convert Twips to pixels for the current screen resolution
'   Accepts:
'       lngTwips - the number of twips to be converted
'       lngDirection - direction (x or y - use either DIRECTION_VERTICAL or DIRECTION_HORIZONTAL)
'   Returns:
'       the number of pixels corresponding to the given twips
Dim hdc As LongPtr
Dim lngPixelsPerInch As Long
    hdc = GetDC(0)
    If lngDirection = DIRECTION_HORIZONTAL Then
        lngPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSX)
     Else
        lngPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSY)
    End If
    hdc = ReleaseDC(0, hdc)
    TwipsToPixels = lngTwips / TwipsPerInch * lngPixelsPerInch
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function PixelsToTwips(ByVal lngPixels As Long, Optional lngDirection As Long = DIRECTION_HORIZONTAL) As Long
'   Function to convert pixels to twips for the current screen resolution
'   Accepts:
'       lngPixels - the number of pixels to be converted
'       lngDirection - direction (x or y - use either DIRECTION_VERTICAL or DIRECTION_HORIZONTAL)
'   Returns:
'       the number of twips corresponding to the given pixels
On Error GoTo HandleError
Dim hdc As LongPtr
Dim lngPixelsPerInch As Long
    hdc = GetDC(0)
    If lngDirection = DIRECTION_HORIZONTAL Then
        lngPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSX)
     Else
        lngPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSY)
    End If
    hdc = ReleaseDC(0, hdc)
    PixelsToTwips = lngPixels * TwipsPerInch / lngPixelsPerInch
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function HimetrixToPixels(ByVal lngHiMetrix As Long, lngDirection As Long) As Long
Dim Result As Long
    On Error GoTo HandleError
    ' conversion from Himetrics to Pixels when ScaleX/Y is not available
    If lngDirection = DIRECTION_HORIZONTAL Then
        Result = lngHiMetrix * TwipsPerInch / 2540 / TwipsPerPixel(LOGPIXELSX) 'Screen.TwipsPerPixelX
    Else
        Result = lngHiMetrix * TwipsPerInch / 2540 / TwipsPerPixel(LOGPIXELSY) 'Screen.TwipsPerPixelY
    End If
HandleExit:  HimetrixToPixels = Result: Exit Function
HandleError: Result = False: Resume HandleExit
End Function
Public Function TwipsPerPixel(Optional ByVal Dimension As Long = LOGPIXELSX) As Long
Dim hdc As LongPtr
    On Error GoTo HandleError
    hdc = GetDC(Application.hWndAccessApp) 'DesktopDC = GetDC(HWND_DESKTOP)
    TwipsPerPixel = TwipsPerInch / GetDeviceCaps(hdc, Dimension)
HandleError:
    Call ReleaseDC(Application.hWndAccessApp, hdc) 'Call ReleaseDC(HWND_DESKTOP, DesktopDC)
  'Call Exception.RaiseAgain
End Function
Public Function CreateHFontByControl(Optional ctl As Variant, _
    Optional FontName, Optional FontSize, Optional FontColor, Optional FontWeight, Optional FontUnderline, Optional FontStrikeOut, Optional FontItalic, _
    Optional hdc As LongPtr) As LongPtr
' ������� hFont �� ���������� ��������
Const c_strProcedure = "CreateHFontbyControl"
Dim Result As LongPtr ': Result = 0
On Error Resume Next
Dim tDC As LongPtr
    'If Not TypeOf ctl Is Access.Control Then Err.Raise vbObjectError + 512
    If hdc = 0 Then tDC = GetDC(0) Else tDC = hdc
' ������ �����
Dim fName As String:    fName = IIf(IsMissing(FontName), ctl.FontName, FontName): If Err Then fName = "Arial": Err.Clear
Dim fSize As Long:      fSize = IIf(IsMissing(FontSize), ctl.FontSize, FontSize): If Err Then fSize = 10: Err.Clear
Dim fColor As Long:     fColor = IIf(IsMissing(FontColor), ctl.FontColor, FontColor): If Err Then fColor = vbBlack: Err.Clear
Dim fWeight As Long:    fWeight = IIf(IsMissing(FontWeight), ctl.FontWeight, FontWeight): If Err Then fWeight = 0: Err.Clear
Dim fItalic As Long:    fItalic = IIf(IsMissing(FontItalic), ctl.FontItalic, FontItalic): If Err Then fItalic = False: Err.Clear
Dim fUnderline As Long: fUnderline = IIf(IsMissing(FontUnderline), ctl.FontUnderline, FontUnderline): If Err Then fUnderline = False: Err.Clear
Dim fStrikeOut As Long: fStrikeOut = IIf(IsMissing(FontStrikeOut), ctl.FontStrikeOut, FontStrikeOut): If Err Then fStrikeOut = False: Err.Clear
On Error GoTo HandleError
    'FontSize = -(FontSize * PT / TwipsPerPixel)
    'fSize = -Int(fSize * GetDeviceCaps(tDC, LOGPIXELSY) / 72)
    fSize = -MulDiv(fSize, GetDeviceCaps(tDC, LOGPIXELSY), 72)
    Result = CreateFont(fSize, 0, 0, 0, _
        fWeight, fItalic, fUnderline, fStrikeOut, _
        RUSSIAN_CHARSET, 0, 0, ANTIALIASED_QUALITY, 0, fName)  ' PROOF_QUALITY | CLEARTYPE_QUALITY | ANTIALIASED_QUALITY
    If hdc = 0 Then ReleaseDC 0, tDC
HandleExit:  CreateHFontByControl = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
'=========================
' ������� ������/������ ������� �� ����
'=========================
Public Function TagsRead(objFrom As Object, Optional objTo As Object) As Boolean '(Optional TagTypes As eObjectProps = ltAll) As Boolean
' ������ ��� �������� �� Tag ������� � �������������� ��������
'-------------------------
' objFrom   - ������ �� ������ (Form,Section,Control) �� ���� �������� �������� �������������� ��������
' objTo     - ������ �� ������ � ������� �������� �������������� �������� �� ���� objFrom
'-------------------------
Dim Result As Boolean: Result = False
    On Error Resume Next 'GoTo HandleError
Dim Tags As New Collection, Keys As Variant
Dim i As Long, iMax As Long
    With objFrom
    If Len(Trim(.Tag)) = 0 Then Exit Function
    Call TaggedString2Collection(.Tag, Tags, Keys): i = LBound(Keys): iMax = UBound(Keys)
    End With
    If (objTo Is Nothing) Then Set objTo = objFrom
Dim tmpVal
    With objTo
    For i = LBound(Keys) To UBound(Keys) 'Each sKey In Keys
        Select Case Keys(i)
' ----------------------------------
' �����/����������
' ----------------------------------
' �������� ���������� (Back/Fore/Font/TextColor)                (adhcColor/adhcBackColor/etc)
    ' ����������� �� �������� ������ (objFrom)
        Case adhcBackColor: tmpVal = GetColorFromText(UCase(Tags(i))):
            If (tmpVal <> COLORNOTSET) Then .BackColor = tmpVal
        Case adhcTextColor, adhcFontColor, adhcForeColor: tmpVal = GetColorFromText(UCase(Tags(i))):
            If (tmpVal <> COLORNOTSET) Then .ForeColor = tmpVal: If Err Then Err.Clear: .FontColor = tmpVal: If Err Then Err.Clear: .TextColor = tmpVal
' ----------------------------------
' ����� �����������                                             (adhcSizeIt/adhcFloatIt/adhcStyleIt)
' ----------------------------------
        Case adhcStyleIt: .StyleIt = GetStyleFromText(Tags(i))
' ���������� (������������� StyleIt)
        Case adhcSizeIt: Select Case Tags(i)
            Case adhcSizeRight: .StyleIt = .StyleIt Or lsHorz      ' ����������� �� ����������� (������)
            Case adhcSizeBottom: .StyleIt = .StyleIt Or lsVert     ' ����������� �� ��������� (����)
            Case adhcSizeBoth: .StyleIt = .StyleIt Or lsFull       ' ����������� �� ����������� � ��������� (������-����)
            End Select
' �������� (������������� StyleIt)
        Case adhcFloatIt: Select Case Tags(i)
            Case adhcFloatRight: .StyleIt = .StyleIt Or lsRight    ' ������� �� ����������� (������)
            Case adhcFloatBottom: .StyleIt = .StyleIt Or lsBottom  ' ������� �� ��������� (������)
            Case adhcFloatBoth: .StyleIt = .StyleIt Or lsRightBottom ' ������� �� ����������� � ��������� (������-����)
            End Select
' ���������������
        Case adhcScaleIt:   .ScaleIt = GetBoolFromText(Tags(i))     ' ��������������
' ----------------------------------
' ���������
' ----------------------------------
        Case adhcSplitIt:   .SplitIt = GetSplitFromText(Tags(i))    ' ������� ����������� (���������)
        Case adhcAction:    .Action = Tags(i)                       ' �������� ��� �������� ��������
' ----------------------------------
' ���������
' ----------------------------------
    ' �������� �������� (�������� �� ����� �������� ������������ ������ �������� ���������� ���������� ������ �� ����)
        Case adhcBondLeft, adhcBondLeft1, adhcBondLeft2:    Set .BondLeft = objFrom.Form.Controls(Tags(i))
        Case adhcBondTop, adhcBondTop1, adhcBondTop2:       Set .BondTop = objFrom.Form.Controls(Tags(i))
        Case adhcBondRight, adhcBondRight1:                 Set .BondRight = objFrom.Form.Controls(Tags(i))
        Case adhcBondBottom, adhcBondBottom1:               Set .BondBottom = objFrom.Form.Controls(Tags(i))
        Case adhcBondWidth, adhcBondWidth1:                 Set .BondWidth = objFrom.Form.Controls(Tags(i))
        Case adhcBondHeight, adhcBondHeight1:               Set .BondHeight = objFrom.Form.Controls(Tags(i))
    ' ��������� ��-���������
        Case adhcDefLeft, adhcDefLeft1, adhcDefLeft2:       .DefLeft = GetSizeFromText(Tags(i), .GetBoundInTwips(eLeft))
        Case adhcDefTop, adhcDefTop1, adhcDefTop2:          .DefTop = GetSizeFromText(Tags(i), .GetBoundInTwips(eTop))
        Case adhcDefRight, adhcDefRight1:                   .DefRight = GetSizeFromText(Tags(i), .GetBoundInTwips(eRight))
        Case adhcDefBottom, adhcDefBottom1:                 .DefBottom = GetSizeFromText(Tags(i), .GetBoundInTwips(eBottom))
        Case adhcDefWidth, adhcDefWidth1:                   .DefWidth = GetSizeFromText(Tags(i), .GetBoundInTwips(eWidth))
        Case adhcDefHeight, adhcDefHeight1:                 .DefHeight = GetSizeFromText(Tags(i), .GetBoundInTwips(eHeight))
    ' ���������� �������
        Case adhcMinWidth, adhcMinWidth1:                   .MinWidth = GetSizeFromText(Tags(i), .GetBoundInTwips(eWidth))
        Case adhcMaxWidth, adhcMaxWidth1:                   .MaxWidth = GetSizeFromText(Tags(i), .GetBoundInTwips(eWidth))
        Case adhcMinHeight, adhcMinHeight1:                 .MinHeight = GetSizeFromText(Tags(i), .GetBoundInTwips(eHeight))
        Case adhcMaxHeight, adhcMaxHeight1:                 .MaxHeight = GetSizeFromText(Tags(i), .GetBoundInTwips(eHeight))
' ----------------------------------
' �����������/�����
' ----------------------------------
' �������� ���������� �����������
        Case adhcObjectName:        .ObjName = Tags(i)
        Case adhcObjectSize:        .ObjSize = Tags(i)
        Case adhcObjectMode:        .ObjMode = GetPictModeFromText(Tags(i))
        Case adhcObjectAlign:       .ObjAlign = GetAlignFromText(Tags(i))
        Case adhcObjectAngle:       .ObjAngle = Tags(i)
        Case adhcObjectGray:        .ObjGray = GetBoolFromText(Tags(i))
' �������� ���������� ������
        Case adhcObjectText:        .ObjText = Tags(i)
        Case adhcObjectTextAlign:   .TxtAlign = GetAlignFromText(Tags(i))
        Case adhcObjectTextPlace:   .TxtPlace = GetPlaceFromText(Tags(i))
        Case adhcObjectTextAngle:   .TxtAngle = Tags(i)
' �������� ������ ���������� ������
        Case adhcFontName:          .FontName = Tags(i)
        Case adhcFontSize:          .FontSize = Tags(i)
' ----------------------------------
        End Select
HandleNext: Err.Clear
    Next
    End With
    Result = True
HandleExit:  TagsRead = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function TagsSave(objFrom As Object, Optional objTo As Object) As Boolean '(Optional TagTypes As eObjectProps = ltAll) As Boolean
' ��������� �������������� �������� � ��� ��������
'-------------------------
' objFrom   - ������ �� ������ �������������� �������� �������� �����������
' objTo     - ������ �� ������ � �������� Tag �������� ����������� �������������� ��������
' TagTypes  - ���� ��� ������� ������� ���������� ��������� (����� � ����� �������������)
'-------------------------
Dim Result As Boolean: Result = False
    On Error Resume Next 'GoTo HandleError
'    With Control
'Dim strTag As String, strVal As String
'' ----------------------------------
'' ������ �������� �������� � �������������� ��������
'' ----------------------------------
'    If GetFromCtrl Then
'' ----------------------------------
'' �����/����������
'' ----------------------------------
'' �������� ���������� (Back/Fore/Font/TextColor)                (adhcColor/adhcBackColor/etc)
'' ----------------------------------
'
'' ----------------------------------
'' ���������
'' ----------------------------------
'    'If (SplitIt <> cdNone) Then      ' ������������� ���������
'' ----------------------------------
'' ���������
'' ----------------------------------
'    'If () Then      '
'' ----------------------------------
'' �����������/�����
'' ----------------------------------
'    'If () Then      '
'' ----------------------------------
'    End If
''' ----------------------------------
''' ��������� �� � Tag
''' ----------------------------------
''' ----------------------------------
''' ����� �����������                                             (adhcSizeIt/adhcFloatIt/adhcStyleIt)
''' ----------------------------------
''    If (TagTypes And ltStyle) = ltStyle Then
'''   ��������� ScaleIt
'        strTag = adhcScaleIt: strVal = Choose(ScaleIt + 2, adhcYes, adhcNo, vbNullString): Call TaggedStringSet(.Tag, strTag, strVal)
'''   ��������� StyleIt ��� SizeIt/FloatIt
'''        If SizeIt <> czNone Then Call TaggedStringSet(.Tag, adhcSizeIt, Choose(SizeIt, adhcSizeRight, adhcSizeBottom, adhcSizeBoth))
'''        If FloatIt <> cfNone Then Call TaggedStringSet(.Tag, adhcFloatIt, Choose(SizeIt, adhcFloatRight, adhcFloatBottom, adhcFloatBoth))
'''   ��������� StyleIt
''    ' ���� �������� �������� ������� ������ ������, �� ������ ��� ������������ - ������������� ��������������� �����
'        StyleIt = SetBits(StyleIt, lsXProp, Abs(DefLeft) <= twMinLim)
'        StyleIt = SetBits(StyleIt, ls�Prop, Abs(DefTop) <= twMinLim)
'        StyleIt = p_SetBits(StyleIt, lsRProp, Abs(DefRight) <= twMinLim)
'        StyleIt = p_SetBits(StyleIt, lsBProp, Abs(DefBottom) <= twMinLim)
'        StyleIt = p_SetBits(StyleIt, lsWProp, Abs(DefWidth) <= twMinLim)
'        StyleIt = p_SetBits(StyleIt, lsHProp, Abs(DefHeight) <= twMinLim)
'        strTag = adhcStyleIt: strVal = GetStyleText(StyleIt): Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''' ----------------------------------
''' ����������
''' ----------------------------------
''' �������� ���������� (Back/Fore/Font/TextColor)                (adhcColor/adhcBackColor/etc)
''    If (TagTypes And ltColors) = ltColors Then
'        strTag = adhcBackColor: strVal = GetColorText(BackColor, False): Call TaggedStringSet(.Tag, strTag, strVal)
'        strTag = adhcForeColor: strVal = GetColorText(ForeColor, False): Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''' ----------------------------------
''' ���������
''' ----------------------------------
''' ������� ����������� (���������)
''    If (TagTypes And ltSplit) = ltSplit Then
'        strTag = adhcSplitIt: strVal = GetSplitText(SplitIt): Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''' �������� ��� �������� ��������
''    If (TagTypes And ltAction) = ltAction Then
'        strTag = adhcAction: strVal = Trim(Action): Call TaggedStringSet(.Tag, strTag, strVal) ', adhcSplitBoth))
''    End If
''' ----------------------------------
''' ���������
''' ----------------------------------
'Dim BondName As String
''    If (TagTypes And ltLeft) = ltLeft Then
'        strTag = adhcDefLeft1: If DefLeft = 0 Then strVal = vbNullString Else strVal = GetSizeText(DefLeft, adhcCm, p_GetBoundInTwips(eLeft, BondName), ((StyleIt And lsXProp) = lsXProp))
'        Call TaggedStringSet(.Tag, strTag, strVal): If Len(strVal) > 0 Then strVal = BondName
'        strTag = adhcBondLeft1: Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltTop) = ltTop Then
'        strTag = adhcDefTop1: If DefTop = 0 Then strVal = vbNullString Else strVal = GetSizeText(DefTop, adhcCm, p_GetBoundInTwips(eTop, BondName), ((StyleIt And lsYProp) = lsYProp))
'        Call TaggedStringSet(.Tag, strTag, strVal): If Len(strVal) > 0 Then strVal = BondName
'        strTag = adhcBondTop1: Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltRight) = ltRight Then
'        strTag = adhcDefRight1: If DefRight = 0 Then strVal = vbNullString Else strVal = GetSizeText(DefRight, adhcCm, p_GetBoundInTwips(eRight, BondName), ((StyleIt And lsRProp) = lsRProp))
'        Call TaggedStringSet(.Tag, strTag, strVal): If Len(strVal) > 0 Then strVal = BondName
'        strTag = adhcBondRight1: Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltBottom) = ltBottom Then
'        strTag = adhcDefBottom1: If DefBottom = 0 Then strVal = vbNullString Else strVal = GetSizeText(DefBottom, adhcCm, p_GetBoundInTwips(eBottom, BondName), ((StyleIt And lsBProp) = lsBProp))
'        Call TaggedStringSet(.Tag, strTag, strVal): If Len(strVal) > 0 Then strVal = BondName
'        strTag = adhcBondBottom1: Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltWidth) = ltWidth Then
'        strTag = adhcDefWidth1: If DefWidth = 0 Then strVal = vbNullString Else strVal = GetSizeText(DefWidth, adhcCm, p_GetBoundInTwips(eWidth, BondName), ((StyleIt And lsWProp) = lsWProp))
'        Call TaggedStringSet(.Tag, strTag, strVal): If Len(strVal) > 0 Then strVal = BondName
'        strTag = adhcBondWidth1: Call TaggedStringSet(.Tag, strTag, strVal)
'
'        strTag = adhcMinWidth1: If MinWidth = 0 Then strVal = vbNullString Else strVal = GetSizeText(MinWidth, adhcCm, p_GetBoundInTwips(eWidth), ((StyleIt And lsWProp) = lsWProp))
'        Call TaggedStringSet(.Tag, strTag, strVal)
'
'        strTag = adhcMaxWidth1: If MaxWidth = 0 Then strVal = vbNullString Else strVal = GetSizeText(MaxWidth, adhcCm, p_GetBoundInTwips(eWidth), ((StyleIt And lsWProp) = lsWProp))
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltHeight) = ltHeight Then
'        strTag = adhcDefHeight1: If DefHeight = 0 Then strVal = vbNullString Else strVal = GetSizeText(DefHeight, adhcCm, p_GetBoundInTwips(eHeight, BondName), ((StyleIt And lsHProp) = lsHProp))
'        Call TaggedStringSet(.Tag, strTag, strVal): If Len(strVal) > 0 Then strVal = BondName
'        strTag = adhcBondHeight1: Call TaggedStringSet(.Tag, strTag, strVal)
'
'        strTag = adhcMinHeight1: If MinHeight = 0 Then strVal = vbNullString Else strVal = GetSizeText(MinHeight, adhcCm, p_GetBoundInTwips(eHeight), ((StyleIt And lsHProp) = lsHProp))
'        Call TaggedStringSet(.Tag, strTag, strVal)
'
'        strTag = adhcMaxHeight1: If MaxHeight = 0 Then strVal = vbNullString Else strVal = GetSizeText(MaxHeight, adhcCm, p_GetBoundInTwips(eHeight), ((StyleIt And lsHProp) = lsHProp))
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''' ----------------------------------
''' �������� ���������� �����������
''' ----------------------------------
''    If (TagTypes And ltPictName) = ltPictName Then
'        strTag = adhcObjectName: strVal = Trim(ObjName): Call TaggedStringSet(.Tag, strTag, strVal) ' ���/��� ���������� �������
''    End If
'        If Len(strVal) > 0 Then
''    If (TagTypes And ltPictSize) = ltPictSize Then
'        strTag = adhcObjectSize: If ObjSize > 0 Then strVal = ObjSize Else strVal = vbNullString    ' ������ ���������� ������� � ��������
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltPictMode) = ltPictMode Then
'        strTag = adhcObjectMode: strVal = GetPictModeText(ObjMode)                                  ' ����� ��������������� �������
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltPictPlace) = ltPictPlace Then
'        strTag = adhcObjectAlign: strVal = GetAlignText(ObjAlign)                                   ' ����� ������������ �������
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltPictAngle) = ltPictAngle Then
'        strTag = adhcObjectAngle: If ObjAngle = 0 Then strVal = vbNullString Else strVal = ObjAngle ' ���� �������� �����������
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltPictGray) = ltPictGray Then
'        strTag = adhcObjectGray: If Not ObjGray Then strVal = vbNullString Else strVal = GetBoolText(ObjGray, 1)   ' �����/������� �����������
'        Call TaggedStringSet(.Tag, strTag, strVal)
'        End If
''    End If
''' ----------------------------------
''' ���������� ������� �� ��������
''' ----------------------------------
''    If (TagTypes And ltPictText) = ltPictText Then
'        strTag = adhcObjectText: strVal = Trim(ObjText): Call TaggedStringSet(.Tag, strTag, strVal) ' ����� ��������� ������ � ������������
''    End If
'        If Len(strVal) > 0 Then
''    If (TagTypes And ltTextPlace) = ltTextPlace Then
'        strTag = adhcObjectTextPlace: strVal = GetPlaceText(TxtPlace)                               ' ���������� ������ ������������ �����������
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltTextAlign) = ltTextAlign Then
'        strTag = adhcObjectTextAlign: strVal = GetAlignText(TxtAlign)                               ' ������������ ������
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltTextAngle) = ltTextAngle Then
'        strTag = adhcObjectTextAngle: If TxtAngle = 0 Then strVal = vbNullString Else strVal = TxtAngle ' ���� ������� ������
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltFontName) = ltFontName Then
'        strTag = adhcDefFontName: strVal = Trim(DefFontName)                                        ' ����� ������
'        Call TaggedStringSet(.Tag, strTag, strVal)
''    End If
''    If (TagTypes And ltFontSize) = ltFontSize Then
'        strTag = adhcDefFontSize: If DefFontSize > 0 Then strVal = DefFontSize Else strVal = vbNullString ' ������ ������
'        Call TaggedStringSet(.Tag, strTag, strVal)
'        End If
''    End If
''' ----------------------------------
'    End With
    Err.Clear: Result = True
HandleExit:  TagsSave = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function GetUnitByName(Unit As String, Optional UnitName As String) As Single
' ���������� ��������� ��� ������� ��������� �������
    Select Case Unit
    Case adhcCm, adhcCm1:       GetUnitByName = cm:     UnitName = adhcCm
    Case adhcInch, adhcInch1:   GetUnitByName = inch:   UnitName = adhcInch
    Case adhcPoints:            GetUnitByName = pt:     UnitName = adhcPoints
    Case adhcPixels:            GetUnitByName = px:     UnitName = adhcPixels
    'Case adhcTwips:             GetUnitByName = 1:      UnitName = vbNullString
    Case Else:                  GetUnitByName = 1:      UnitName = vbNullString
    End Select
End Function
Public Function GetSizeText(twSize, Optional Unit As String, _
        Optional twBase As Long, Optional Prop As Boolean = False) As String
' ��������� ������� �������� ������ � ������
Const cDecimails = 3 ' ���������� ������ ����� ������� � ����������
' twSize - ������/������� (twip) ��� ���������������� ��������
' Unit - ��������� ����������� ������ ��������� ����������
' twBase - ������/������� (twip) ���� ��� ������������� �������
' Prop - ������� ������������� ������� ��������� ����� ���������������� �������� � twBase
Dim Result As Single
    If Abs(twSize) <= twMinLim Then
' ������� ������������� ������ (����), ��������� ����� �������������� (�����������������) �������
        Result = twSize: Unit = vbNullString            ' ��������� ��� ����
    ElseIf Prop And twBase <> 0 Then
' ������� ���������� ������ (twip), ��������� ����� �������������� (�����������������) �������
        Result = twSize / twBase: Unit = vbNullString   ' ���������� ��������� ����������� � ����
    Else
' ������� ���������� ������ (twip), ��������� ����� ������� � ���������� ��������� (Unit)
Dim sngUnit As Single: sngUnit = GetUnitByName(Unit, Unit)
        Result = (twSize - twBase) / sngUnit              ' ���������� ��������� ����������� ������������ ���� � ��������� ��������
    End If
' ����������� ���������
    Result = Round(Result, cDecimails)
    GetSizeText = Format(Result, "#0" & IIf(Result <> Fix(Result), "." & String$(cDecimails, "#"), vbNullString))
    If Len(Unit) > 0 Then GetSizeText = GetSizeText & Unit  ' ��������� ������� ���������
End Function
Public Function GetSizeFromText(SizeText As String, Optional strBase As String) As Single
' ��������� ������� �������� ������� � ����� ����� (�����), ���� ����� (������������� ��������)
' ���� ������ �������� ���� - ����� ���������� ��� ��������, ������� ��� ��������� � ���� � ��������� � ���������� (�����)
' twMinLim - ����� ���� �������� �������� ��������� ����������
' SizeText - ������ �������/������� (����� ������������ ������� ���������)
' strBase - ������ �������/������� ������������ ������� ������ ��������
'-------------------------
' ! ���������� ����������� � ������������ � ������������� ����������� (,)
'-------------------------
' IsNumeric ��������� ����� � ������������ � ������������� �����������
' cDbl ����������� � ������������ � ������������� �����������
' Val ���������� ������������ ���������
Dim sngSize As Single
Dim twBase As Long: If Len(strBase) > 0 Then twBase = GetSizeFromText(strBase)
    If IsNumeric(SizeText) Then
' ����� (��� ������ ���������)
        sngSize = CSng(SizeText)
    ' ����� - ������������� ��������
        If Abs(sngSize) <= twMinLim Then If twBase > 0 Then GetSizeFromText = sngSize * twBase: Exit Function
    ' ����� - ���������� �������� �������� � twip
        GetSizeFromText = twBase + sngSize: Exit Function
    End If
' ����� - ���������� �������� � �������� ���������
' ��������� �������������� �� �������� ��������
Dim arrRules: arrRules = Array("��", "cm", "in", "'", "pt", "px", "tw")
Dim Pos As Long, Sym As String, tmp As String
Dim i As Long
    For i = LBound(arrRules) To UBound(arrRules)
        Pos = InStrRev(SizeText, arrRules(i))
        If Pos > 1 Then
            tmp = Left$(SizeText, Pos - 1)
            If IsNumeric(tmp) Then
                sngSize = CSng(tmp) * GetUnitByName(CStr(arrRules(i)))
                GetSizeFromText = twBase + sngSize: Exit Function
            End If
        End If
    Next i
End Function
Public Function GetColorText(vb�olor As Long, Optional SchemeColorsOnly As Boolean = False) As String
    On Error GoTo HandleError
    GetColorText = vbNullString
' ���� ����� ������
    Select Case vb�olor
' �������� �������� ������ (������������� � �������� �����)
' �������� ����1, ����2, ����3
    Case appColorDark:      GetColorText = StrConv(adhcColorDark, vbProperCase)  ' ������ ����
    Case appColorDark2:     GetColorText = StrConv(adhcColorDark2, vbProperCase)
    Case appColorDark3:     GetColorText = StrConv(adhcColorDark3, vbProperCase)
    Case appColorBright:    GetColorText = StrConv(adhcColorBright, vbProperCase) ' ����� ����
    Case appColorBright2:   GetColorText = StrConv(adhcColorBright2, vbProperCase)
    Case appColorBright3:   GetColorText = StrConv(adhcColorBright3, vbProperCase)
    Case appColorLight:     GetColorText = StrConv(adhcColorLight, vbProperCase)  ' ������� ����
    Case appColorLight2:    GetColorText = StrConv(adhcColorLight2, vbProperCase)
    Case appColorLight3:    GetColorText = StrConv(adhcColorLight3, vbProperCase)
    End Select
    If SchemeColorsOnly Then Err.Raise vbObjectError + 512
    Select Case vb�olor
' ����� �������� ������ (��������)
    Case vbBlack:           GetColorText = adhcColorBlack   ' ������ (&H0)
    Case vbWhite:           GetColorText = adhcColorWhite   ' ����� (&HFFFFFF)
    Case &H808080:          GetColorText = adhcColorGray    ' �����
    Case &H333333:          GetColorText = adhcColorDark & adhcColorGray     ' ����� �����
    Case &HC0C0C0:          GetColorText = "Silver"         ' ����������
    Case vbRed:             GetColorText = "Red"            ' ������� (&HFF)
    Case vbBlue:            GetColorText = "Blue"           ' ����� (&HFF0000)
    Case &H8000:            GetColorText = "Green"          ' �������
    Case vbYellow:          GetColorText = "Yellow"         ' ������ (&HFFFF)
    Case vbMagenta:         GetColorText = "Magenta"        ' �������/������ (&HFF00FF)
    Case &H80:              GetColorText = "Navy"           ' �������� �����
    Case &H808000:          GetColorText = "Teal"           ' ����-������
    Case vbCyan:            GetColorText = "Cyan"           ' ������-�������/��������� (&HFFFF00)
    Case &HFF00:            GetColorText = "Lime"           ' �����-������� (��������)
    Case &H8080:            GetColorText = "Olive"          ' ����� ���������-������ (���������)
    Case &H800000:          GetColorText = "Maroon"         ' ����-�������� (������)
    Case &H800080:          GetColorText = "Purple"         ' ���������
    End Select
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function GetColorFromText(ColorName)  'As Long
    On Error GoTo HandleError
    GetColorFromText = Null
' ���� ����� ������
    If IsNumeric(ColorName) Then GetColorFromText = CLng(ColorName): GoTo HandleExit
' ���� ����� ������
    Select Case ColorName
' �������� �������� ������ (������������� � �������� �����)
' �������� ����1, ����2, ����3
    Case adhcColorDark:     GetColorFromText = appColorDark   ' ������ ����
    Case adhcColorDark2:    GetColorFromText = appColorDark2
    Case adhcColorDark3:    GetColorFromText = appColorDark3
    Case adhcColorBright:   GetColorFromText = appColorBright ' ����� ����
    Case adhcColorBright2:  GetColorFromText = appColorBright2
    Case adhcColorBright3:  GetColorFromText = appColorBright3
    Case adhcColorLight:    GetColorFromText = appColorLight  ' ������� ����
    Case adhcColorLight2:   GetColorFromText = appColorLight2
    Case adhcColorLight3:   GetColorFromText = appColorLight3
' ����� �������� ������ HTML: https://colorscheme.ru/html-colors.html?ysclid=lrrntfgkb9889651755
' �������� �����:
    ' HTML Color Name       ' VB Color (BGR)                        ' HTML Color (RGB)
    Case "Black":           GetColorFromText = vbBlack              ' #000000   ' ������
    Case "White":           GetColorFromText = vbWhite              ' #FFFFFF   ' �����
    Case "Gray", "Grey":    GetColorFromText = rgb(128, 128, 128)   ' #808080   ' �����
    Case "Silver":          GetColorFromText = rgb(192, 192, 192)   ' #C0C0C0   ' ����������
    Case "Fuchsia", "Magenta": GetColorFromText = vbMagenta         ' #FF00FF   ' �������/������
    Case "Purple":          GetColorFromText = rgb(128, 0, 128)     ' #800080   ' ���������
    Case "Red":             GetColorFromText = vbRed                ' #FF0000   ' �������
    Case "Maroon":          GetColorFromText = rgb(128, 0, 0)       ' #800000   ' ����-�������� (������)
    Case "Yellow":          GetColorFromText = rgb(255, 255, 0)     ' #FFFF00   ' �����
    Case "Olive":           GetColorFromText = rgb(128, 128, 0)     ' #808000   ' ����� ���������-������ (���������)
    Case "Lime":            GetColorFromText = rgb(0, 255, 0)       ' #00FF00   ' �����-������� (��������)
    Case "Green":           GetColorFromText = rgb(0, 128, 0)       ' #008000   ' �������
    Case "Aqua", "Cyan":    GetColorFromText = rgb(0, 255, 255)     ' #00FFFF   ' ������-�������/���������
    Case "Teal":            GetColorFromText = rgb(0, 128, 128)     ' #008080   ' ����-������
    Case "Blue":            GetColorFromText = vbBlue               ' #0000FF   ' �����
    Case "Navy":            GetColorFromText = rgb(0, 0, 128)       ' #000080   ' �������� �����
' ������� ����:
    Case "IndianRed":       GetColorFromText = rgb(205, 92, 92)     ' #CD5C5C
    Case "LightCoral":      GetColorFromText = rgb(240, 128, 128)   ' #F08080
    Case "Salmon":          GetColorFromText = rgb(250, 128, 114)   ' #FA8072
    Case "DarkSalmon":      GetColorFromText = rgb(233, 150, 122)   ' #E9967A
    Case "LightSalmon":     GetColorFromText = rgb(255, 160, 122)   ' #FFA07A
    Case "Crimson":         GetColorFromText = rgb(220, 20, 60)     ' #DC143C
    'Case "Red":             GetColorFromText = rgb(255, 0, 0)       ' #FF0000
    Case "FireBrick":       GetColorFromText = rgb(178, 34, 34)     ' #B22222
    Case "DarkRed":         GetColorFromText = rgb(139, 0, 0)       ' #8B0000
' ������� ����:
    Case "Pink":            GetColorFromText = rgb(255, 192, 203)   ' #FFC0CB
    Case "LightPink":       GetColorFromText = rgb(255, 182, 193)   ' #FFB6C1
    Case "HotPink":         GetColorFromText = rgb(255, 105, 180)   ' #FF69B4
    Case "DeepPink":        GetColorFromText = rgb(255, 20, 147)    ' #FF1493
    Case "MediumVioletRed": GetColorFromText = rgb(199, 21, 133)    ' #C71585
    Case "PaleVioletRed":   GetColorFromText = rgb(219, 112, 147)   ' #DB7093
' ��������� ����:
    Case "LightSalmon":     GetColorFromText = rgb(255, 160, 122)   ' #FFA07A
    Case "Coral":           GetColorFromText = rgb(255, 127, 80)    ' #FF7F50
    Case "Tomato":          GetColorFromText = rgb(255, 99, 71)     ' #FF6347
    Case "OrangeRed":       GetColorFromText = rgb(255, 69, 0)      ' #FF4500
    Case "DarkOrange":      GetColorFromText = rgb(255, 140, 0)     ' #FF8C00
    Case "Orange":          GetColorFromText = rgb(255, 165, 0)     ' #FFA500
' Ƹ���� ����:
    Case "Gold":            GetColorFromText = rgb(255, 215, 0)     ' #FFD700
    'Case "Yellow":          GetColorFromText = rgb(255, 255, 0)     ' #FFFF00
    Case "LightYellow":     GetColorFromText = rgb(255, 255, 224)   ' #FFFFE0
    Case "LemonChiffon":    GetColorFromText = rgb(255, 250, 205)   ' #FFFACD
    Case "LightGoldenrodYellow":  GetColorFromText = rgb(250, 250, 210) ' #FAFAD2
    Case "PapayaWhip":      GetColorFromText = rgb(255, 239, 213)   ' #FFEFD5
    Case "Moccasin":        GetColorFromText = rgb(255, 228, 181)   ' #FFE4B5
    Case "PeachPuff":       GetColorFromText = rgb(255, 218, 185)   ' #FFDAB9
    Case "PaleGoldenrod":   GetColorFromText = rgb(238, 232, 170)   ' #EEE8AA
    Case "Khaki":           GetColorFromText = rgb(240, 230, 140)   ' #F0E68C
    Case "DarkKhaki":       GetColorFromText = rgb(189, 183, 107)   ' #BDB76B
' ���������� ����:
    Case "Lavender":        GetColorFromText = rgb(230, 230, 250)   ' #E6E6FA
    Case "Thistle":         GetColorFromText = rgb(216, 191, 216)   ' #D8BFD8
    Case "Plum":            GetColorFromText = rgb(221, 160, 221)   ' #DDA0DD
    Case "Violet":          GetColorFromText = rgb(238, 130, 238)   ' #EE82EE
    Case "Orchid":          GetColorFromText = rgb(218, 112, 214)   ' #DA70D6
    'Case "Fuchsia", "Magenta": GetColorFromText = rgb(255, 0, 255)  ' #FF00FF
    Case "MediumOrchid":    GetColorFromText = rgb(186, 85, 211)    ' #BA55D3
    Case "MediumPurple":    GetColorFromText = rgb(147, 112, 219)   ' #9370DB
    Case "BlueViolet":      GetColorFromText = rgb(138, 43, 226)    ' #8A2BE2
    Case "DarkViolet":      GetColorFromText = rgb(148, 0, 211)     ' #9400D3
    Case "DarkOrchid":      GetColorFromText = rgb(153, 50, 204)    ' #9932CC
    Case "DarkMagenta":     GetColorFromText = rgb(139, 0, 139)     ' #8B008B
    Case "Purple":          GetColorFromText = rgb(128, 0, 128)     ' #800080
    Case "Indigo":          GetColorFromText = rgb(75, 0, 130)      ' #4B0082
    Case "SlateBlue":       GetColorFromText = rgb(106, 90, 205)    ' #6A5ACD
    Case "DarkSlateBlue":   GetColorFromText = rgb(72, 61, 139)     ' #483D8B
' ���������� ����:
    Case "Cornsilk":        GetColorFromText = rgb(255, 248, 220)   ' #FFF8DC
    Case "BlanchedAlmond":  GetColorFromText = rgb(255, 235, 205)   ' #FFEBCD
    Case "Bisque":          GetColorFromText = rgb(255, 228, 196)   ' #FFE4C4
    Case "NavajoWhite":     GetColorFromText = rgb(255, 222, 173)   ' #FFDEAD
    Case "Wheat":           GetColorFromText = rgb(245, 222, 179)   ' #F5DEB3
    Case "BurlyWood":       GetColorFromText = rgb(222, 184, 135)   ' #DEB887
    Case "Tan":             GetColorFromText = rgb(210, 180, 140)   ' #D2B48C
    Case "RosyBrown":       GetColorFromText = rgb(188, 143, 143)   ' #BC8F8F
    Case "SandyBrown":      GetColorFromText = rgb(244, 164, 96)    ' #F4A460
    Case "Goldenrod":       GetColorFromText = rgb(218, 165, 32)    ' #DAA520
    Case "DarkGoldenRod":   GetColorFromText = rgb(184, 134, 11)    ' #B8860B
    Case "Peru":            GetColorFromText = rgb(205, 133, 63)    ' #CD853F
    Case "Chocolate":       GetColorFromText = rgb(210, 105, 30)    ' #D2691E
    Case "SaddleBrown":     GetColorFromText = rgb(139, 69, 19)     ' #8B4513
    Case "Sienna":          GetColorFromText = rgb(160, 82, 45)     ' #A0522D
    Case "Brown":           GetColorFromText = rgb(165, 42, 42)     ' #A52A2A
    'Case "Maroon":          GetColorFromText = rgb(128, 0, 0)       ' #800000
' ������ ����:
    Case "GreenYellow":     GetColorFromText = rgb(173, 255, 47)    ' #ADFF2F
    Case "Chartreuse":      GetColorFromText = rgb(127, 255, 0)     ' #7FFF00
    Case "LawnGreen":       GetColorFromText = rgb(124, 252, 0)     ' #7CFC00
    'Case "Lime":            GetColorFromText = rgb(0, 255, 0)       ' #00FF00
    Case "LimeGreen":       GetColorFromText = rgb(50, 205, 50)     ' #32CD32
    Case "PaleGreen":       GetColorFromText = rgb(152, 251, 152)   ' #98FB98
    Case "LightGreen":      GetColorFromText = rgb(144, 238, 144)   ' #90EE90
    Case "MediumSpringGreen": GetColorFromText = rgb(0, 250, 154)   ' #00FA9A
    Case "SpringGreen":     GetColorFromText = rgb(0, 255, 127)     ' #00FF7F
    Case "MediumSeaGreen":  GetColorFromText = rgb(60, 179, 113)    ' #3CB371
    Case "SeaGreen":        GetColorFromText = rgb(46, 139, 87)     ' #2E8B57
    Case "ForestGreen":     GetColorFromText = rgb(34, 139, 34)     ' #228B22
    'Case "Green":           GetColorFromText = rgb(0, 128, 0)       ' #008000
    Case "DarkGreen":       GetColorFromText = rgb(0, 100, 0)       ' #006400
    Case "YellowGreen":     GetColorFromText = rgb(154, 205, 50)    ' #9ACD32
    Case "OliveDrab":       GetColorFromText = rgb(107, 142, 35)    ' #6B8E23
    'Case "Olive":           GetColorFromText = rgb(128, 128, 0)     ' #808000
    Case "DarkOliveGreen":  GetColorFromText = rgb(85, 107, 47)     ' #556B2F
    Case "MediumAquamarine": GetColorFromText = rgb(102, 205, 170)  ' #66CDAA
    Case "DarkSeaGreen":    GetColorFromText = rgb(143, 188, 143)   ' #8FBC8F
    Case "LightSeaGreen":   GetColorFromText = rgb(32, 178, 170)    ' #20B2AA
    Case "DarkCyan":        GetColorFromText = rgb(0, 139, 139)     ' #008B8B
    'Case "Teal":            GetColorFromText = rgb(0, 128, 128)     ' #008080
' ����� ����:
    'Case "Aqua", "Cyan":    GetColorFromText = rgb(0, 255, 255)     ' #00FFFF
    Case "LightCyan":       GetColorFromText = rgb(224, 255, 255)   ' #E0FFFF
    Case "PaleTurquoise":   GetColorFromText = rgb(175, 238, 238)   ' #AFEEEE
    Case "Aquamarine":      GetColorFromText = rgb(127, 255, 212)   ' #7FFFD4
    Case "Turquoise":       GetColorFromText = rgb(64, 224, 208)    ' #40E0D0
    Case "MediumTurquoise": GetColorFromText = rgb(72, 209, 204)    ' #48D1CC
    Case "DarkTurquoise":   GetColorFromText = rgb(0, 206, 209)     ' #00CED1
    Case "CadetBlue":       GetColorFromText = rgb(95, 158, 160)    ' #5F9EA0
    Case "SteelBlue":       GetColorFromText = rgb(70, 130, 180)    ' #4682B4
    Case "LightSteelBlue":  GetColorFromText = rgb(176, 196, 222)   ' #B0C4DE
    Case "PowderBlue":      GetColorFromText = rgb(176, 224, 230)   ' #B0E0E6
    Case "LightBlue":       GetColorFromText = rgb(173, 216, 230)   ' #ADD8E6
    Case "SkyBlue":         GetColorFromText = rgb(135, 206, 235)   ' #87CEEB
    Case "LightSkyBlue":     GetColorFromText = rgb(135, 206, 250)  ' #87CEFA
    Case "DeepSkyBlue":     GetColorFromText = rgb(0, 191, 255)     ' #00BFFF
    Case "DodgerBlue":      GetColorFromText = rgb(30, 144, 255)    ' #1E90FF
    Case "CornflowerBlue":  GetColorFromText = rgb(100, 149, 237)   ' #6495ED
    Case "MediumSlateBlue": GetColorFromText = rgb(123, 104, 238)   ' #7B68EE
    Case "RoyalBlue":       GetColorFromText = rgb(65, 105, 225)    ' #4169E1
    'Case "Blue":            GetColorFromText = rgb(0, 0, 255)       ' #0000FF
    Case "MediumBlue":      GetColorFromText = rgb(0, 0, 205)       ' #0000CD
    Case "DarkBlue":        GetColorFromText = rgb(0, 0, 139)       ' #00008B
    'Case "Navy":            GetColorFromText = rgb(0, 0, 128)       ' #000080
    Case "MidnightBlue":    GetColorFromText = rgb(25, 25, 112)     ' #191970
' ����� ����:
    'Case "White":           GetColorFromText = rgb(255, 255, 255)   ' #FFFFFF
    Case "Snow":            GetColorFromText = rgb(255, 250, 250)   ' #FFFAFA
    Case "Honeydew":        GetColorFromText = rgb(240, 255, 240)   ' #F0FFF0
    Case "MintCream":       GetColorFromText = rgb(245, 255, 250)   ' #F5FFFA
    Case "Azure":           GetColorFromText = rgb(240, 255, 255)   ' #F0FFFF
    Case "AliceBlue":       GetColorFromText = rgb(240, 248, 255)   ' #F0F8FF
    Case "GhostWhite":      GetColorFromText = rgb(248, 248, 255)   ' #F8F8FF
    Case "WhiteSmoke":      GetColorFromText = rgb(245, 245, 245)   ' #F5F5F5
    Case "Seashell":        GetColorFromText = rgb(255, 245, 238)   ' #FFF5EE
    Case "Beige":           GetColorFromText = rgb(245, 245, 220)   ' #F5F5DC
    Case "OldLace":         GetColorFromText = rgb(253, 245, 230)   ' #FDF5E6
    Case "FloralWhite":     GetColorFromText = rgb(255, 250, 240)   ' #FFFAF0
    Case "Ivory":           GetColorFromText = rgb(255, 255, 240)   ' #FFFFF0
    Case "AntiqueWhite":    GetColorFromText = rgb(250, 235, 215)   ' #FAEBD7
    Case "Linen":           GetColorFromText = rgb(250, 240, 230)   ' #FAF0E6
    Case "LavenderBlush":   GetColorFromText = rgb(255, 240, 245)   ' #FFF0F5
    Case "MistyRose":       GetColorFromText = rgb(255, 228, 225)   ' #FFE4E1
' ����� ����:
    Case "Gainsboro":       GetColorFromText = rgb(220, 220, 220)   ' #DCDCDC
    Case "LightGrey", "LightGray": GetColorFromText = rgb(211, 211, 211) ' #D3D3D3
    'Case "Silver":          GetColorFromText = rgb(192, 192, 192)   ' #C0C0C0
    Case "DarkGray", "DarkGrey": GetColorFromText = rgb(169, 169, 169) ' #A9A9A9
    'Case "Gray", "Grey":         GetColorFromText = rgb(128, 128, 128) ' #808080
    Case "DimGray", "DimGrey": GetColorFromText = rgb(105, 105, 105) ' #696969
    Case "LightSlateGray", "LightSlateGrey":   GetColorFromText = rgb(119, 136, 153) ' #778899
    Case "SlateGray", "SlateGrey": GetColorFromText = rgb(112, 128, 144) ' #708090
    Case "DarkSlateGray", "DarkSlateGrey": GetColorFromText = rgb(47, 79, 79) ' #2F4F4F
    'Case "Black":           GetColorFromText = rgb(0, 0, 0)         ' #000000
    End Select
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function GetStyleText(lngStyle As eObjectStyle, Optional Short As Boolean = False) As String
' ����������� �������� �������� �����/������������/���������� � ���������
Dim strVal As String, strDelim As String * 1:   strDelim = c_strStyleDelims  ' �������� ������ ����� ������������ ������� '+&,'
Dim tmp As Long: tmp = lngStyle
' �������� ��������/���������
    If (tmp And lsFull) = lsFull Then strVal = strVal & IIf(Short, adhcFull1, adhcFull): tmp = tmp And Not lsFull
    If (tmp And lsHorz) = lsHorz Then strVal = strVal & IIf(Short, adhcHorz2, adhcHorz1): tmp = tmp And Not lsHorz
    If (tmp And lsVert) = lsVert Then strVal = strVal & IIf(Short, adhcVert2, adhcVert1): tmp = tmp And Not lsVert
    If (tmp And lsLeft) = lsLeft Then strVal = strVal & IIf(Short, adhcLeft1, adhcLeft): tmp = tmp And Not lsLeft
    If (tmp And lsRight) = lsRight Then strVal = strVal & IIf(Short, adhcRight1, adhcRight): tmp = tmp And Not lsRight
    If (tmp And lsTop) = lsTop Then strVal = strVal & IIf(Short, adhcTop1, adhcTop): tmp = tmp And Not lsTop
    If (tmp And lsBottom) = lsBottom Then strVal = strVal & IIf(Short, adhcBottom1, adhcBottom): tmp = tmp And Not lsBottom
' �������� ������������� ��������
    If (tmp And lsWProp) = lsWProp Then strVal = strVal & strDelim & adhcWProp: tmp = tmp And Not lsWProp
    If (tmp And lsHProp) = lsHProp Then strVal = strVal & strDelim & adhcHProp: tmp = tmp And Not lsHProp
    If (tmp And lsXProp) = lsXProp Then strVal = strVal & strDelim & adhcXProp: tmp = tmp And Not lsXProp
    If (tmp And lsYProp) = lsYProp Then strVal = strVal & strDelim & adhcYProp: tmp = tmp And Not lsYProp
    If (tmp And lsRProp) = lsRProp Then strVal = strVal & strDelim & adhcRProp: tmp = tmp And Not lsRProp
    If (tmp And lsBProp) = lsBProp Then strVal = strVal & strDelim & adhcBProp: tmp = tmp And Not lsBProp
' ����� �����������/�������
    Select Case (tmp And lsShowIconText)
    Case lsShowIconText:    strVal = strVal & strDelim & adhcIconAndText: tmp = tmp And Not lsShowIconText
    Case lsShowIcon:        strVal = strVal & strDelim & adhcPict: tmp = tmp And Not lsShowIcon
    Case lsShowText:        strVal = strVal & strDelim & adhcText: tmp = tmp And Not lsShowText
    End Select
    If Left(strVal, Len(strDelim)) = strDelim Then strVal = Mid$(strVal, Len(strDelim) + 1)
    GetStyleText = strVal
End Function
Public Function GetStyleFromText(strStyle As String) As eObjectStyle
' ����������� ��������� �������� �����/������������/���������� � ��������
    ' �������� ������ ����� ������������ ������� '+&,'
Dim strVal As String, arrVal() As String, j As Long, tmp As eObjectStyle
    GetStyleFromText = lsNone
    Call Tokenize(strStyle, arrVal, c_strStyleDelims)     ' SplitMany(strVal, c_strStyleDelims)
    For j = LBound(arrVal) To UBound(arrVal)
        strVal = UCase(Trim(arrVal(j)))
        tmp = lsNone ' ������� � ������ ������������� (��� ������������� � �������������)
        Select Case strVal
        Case lsNone, lsLeft, lsRight, lsHorz, lsTop, lsLeftTop, _
            lsRightTop, lsHorzTop, lsBottom, lsLeftBottom, lsRightBottom, _
            lsHorzBottom, lsVert, lsVertLeft, lsVertRight, lsFull, _
            lsShowIcon, lsShowText, lsShowIconText, _
            lsXProp, lsYProp, lsRProp, lsBProp, lsWProp, lsHProp ', _
            'lsCenterHorz, lsCenterVert, lsCenter,
                                                    tmp = CLng(strVal)                  ' ������ ������
        Case adhcNone:                         tmp = lsNone                        '
' ����� ������������ ������� (��������)
        Case adhcLeft, adhcLeft1:         tmp = lsLeft                        ' �������� � ����� ������� �������
        Case adhcRight, adhcRight1:       tmp = lsRight                       ' �������� � ������ ������� �������
        Case adhcTop, adhcTop1:           tmp = lsTop                         ' �������� � ������ ������� �������
        Case adhcBottom, adhcBottom1:     tmp = lsBottom                      ' �������� � ������ ������� �������
        Case adhcVert, adhcVert1, adhcVert2: tmp = lsVert                   ' ������������ ����� ������� � ������ ��������� ������� (������������� �����������)
        Case adhcHorz, adhcHorz1, adhcHorz2: tmp = lsHorz                   ' ������������ ����� ������� � ������ ��������� ������� (������������� �������������)
' ����� ������������ ������� (�����������)
        Case adhcLeftTop, adhcLeftTop1:   tmp = lsLeftTop                     ' �������� � ������ �������� ���� �������
        Case adhcRightTop, adhcRightTop1: tmp = lsRightTop                    ' �������� � ������� �������� ���� �������
        Case adhcHorTop, adhcHor1Top, adhcHor2Top1: tmp = lsHorzTop      ' �������� � ������� ������� � ������������ ����� ������� � ������ ��������� ������� (������������� �������������)
        Case adhcLeftBottom, adhcLeftBottom1: tmp = lsLeftBottom              ' �������� � ������ ������� ���� �������
        Case adhcRightBottom, adhcRightBottom1: tmp = lsRightBottom           ' �������� � ������� ������� ���� �������
        Case adhcHorBottom, adhcHor1Bottom, adhcHor2Bottom1: tmp = lsHorzBottom ' �������� � ������ ������� � ������������ ����� ������� � ������ ��������� ������� (������������� �������������)
        Case adhcVerLeft, adhcVer1Left, adhcVer2Left1:   tmp = lsVertLeft ' �������� � ����� ������� � ������������ ����� ������� � ������ ��������� ������� (������������� �����������)
        Case adhcVerRight, adhcVer1Right, adhcVer2Right1: tmp = lsVertRight ' �������� � ������ ������� � ������������ ����� ������� � ������ ��������� ������� (������������� �����������)
        Case adhcFull, adhcFull1:         tmp = lsFull                        ' ������������� ����������� � ������������� �� ������ � ������ ������� �������
' ����� ������������� ��������
        Case adhcXProp:                   tmp = lsXProp                       ' Left ������� �� ������ ������� (������������� �� ������)
        Case adhcYProp:                   tmp = lsYProp                       ' Top ������� �� ������ ������� (������������� �� ������)
        Case adhcRProp:                   tmp = lsRProp                       ' Right ������� �� ������ ������� (������������� �� ������)
        Case adhcBProp:                   tmp = lsBProp                       ' Bottom ������� �� ������ ������� (������������� �� ������)
        Case adhcWProp:                   tmp = lsWProp                       ' Width ��������������� ������ �������
        Case adhcHProp:                   tmp = lsHProp                       ' Height ��������������� ������ �������
' ����� ������ �����������/������
        Case adhcPict:                    tmp = lsShowIcon                    ' �������� ������
        Case adhcText:                    tmp = lsShowText                    ' �������� �������
        Case adhcIconAndText:             tmp = lsShowIconText                ' �������� ������ � �������
        Case Else 'Exit for' ���� �� ��������������� �������� ���������� ����� � �������
        End Select
        GetStyleFromText = GetStyleFromText Or tmp ' ��������� �����
    Next j
End Function
Public Function GetSplitFromText(strSplit As String) As eControlSplit
    Select Case strSplit
    Case adhcSplitVer, adhcSplitVer1: GetSplitFromText = cdVert
    Case adhcSplitHor, adhcSplitHor1: GetSplitFromText = cdHorz
    'Case adhcSplitBoth: GetSplitFromText = cdBoth
    Case Else: GetSplitFromText = cdNone
    End Select
End Function
Public Function GetSplitText(lngSplit As eControlSplit) As String
    Select Case lngSplit
    Case cdVert: GetSplitText = adhcSplitVer1
    Case cdHorz: GetSplitText = adhcSplitHor1
    'Case cdBoth: GetSplitText = adhcSplitBoth
    End Select
End Function
Public Function GetBoolFromText(strBool As String) As Boolean
    Select Case strBool
    Case adhcTrue, adhcYes, adhcOn: GetBoolFromText = True
    Case Else: GetBoolFromText = False
    End Select
End Function
Public Function GetBoolText(bBool As Boolean, Optional Mode As Long) As String
    If bBool Then
        GetBoolText = Choose(Mode + 1, adhcTrue, adhcYes, adhcOn)
    Else
        GetBoolText = Choose(Mode + 1, adhcFalse, adhcNo, adhcOff)
    End If
End Function
Public Function GetAlignText(lngAlign As eAlign, Optional Short As Boolean = False) As String
    If lngAlign = eAlignUndef Then Exit Function
'    If lngAlign = eAlignLeftTop Then Exit Function
Dim strVal As String
Dim tmp As Long: tmp = lngAlign
' ����� ������������ �������
    If (tmp And lsFull) = lsFull Then strVal = strVal & IIf(Short, adhcCenter1, adhcCenter): tmp = tmp And Not lsFull
    If (tmp And lsHorz) = lsHorz Then strVal = strVal & IIf(Short, adhcHorz2, adhcCenterHor1): tmp = tmp And Not lsHorz
    If (tmp And lsVert) = lsVert Then strVal = strVal & IIf(Short, adhcVert2, adhcCenterVer1): tmp = tmp And Not lsVert
    If (tmp And lsLeft) = lsLeft Then strVal = strVal & IIf(Short, adhcLeft1, adhcLeft): tmp = tmp And Not lsLeft
    If (tmp And lsRight) = lsRight Then strVal = strVal & IIf(Short, adhcRight1, adhcRight): tmp = tmp And Not lsRight
    If (tmp And lsTop) = lsTop Then strVal = strVal & IIf(Short, adhcTop1, adhcTop): tmp = tmp And Not lsTop
    If (tmp And lsBottom) = lsBottom Then strVal = strVal & IIf(Short, adhcBottom1, adhcBottom): tmp = tmp And Not lsBottom
    GetAlignText = strVal
End Function
Public Function GetAlignFromText(strAlign As String) As eAlign
' ����� ������������ �������
    GetAlignFromText = eAlignUndef: If Len(Trim(strAlign)) = 0 Then Exit Function
Dim strVal As String, arrVal() As String, j As Long, tmp As eAlign
    Call Tokenize(strAlign, arrVal, c_strStyleDelims)     ' SplitMany(strVal, c_strStyleDelims)
    For j = LBound(arrVal) To UBound(arrVal)
        strVal = UCase(Trim(arrVal(j)))
        tmp = eAlignUndef
        Select Case strVal
        Case eAlignUndef, eLeft, eRight, eTop, eBottom, eCenterHorz, eCenterVert, _
            eAlignLeftTop, eAlignRightTop, eAlignLeftBottom, eAlignRightBottom, _
            eCenterHorzTop, eCenterHorzBottom, eCenterVertLeft, eCenterVertRight, eCenter, eCascade
                                        tmp = CLng(strVal)                  ' ������ ������
' ����� ������������ ������� (��������)
        Case adhcLeft, adhcLeft1:       tmp = eLeft                         ' �� ������ ����
        Case adhcRight, adhcRight1:     tmp = eRight                        ' �� ������� ����
        Case adhcTop, adhcTop1:         tmp = eTop                          ' �� �������� ����
        Case adhcBottom, adhcBottom1:   tmp = eBottom                       ' �� ������� ����
        Case adhcHorz, adhcHorz1, adhcHorz2, _
             adhcCenterHor, adhcCenterHor1, adhcCenterHor2: tmp = eCenterHorz        ' ������������ �� �����������
        Case adhcVert, adhcVert1, adhcVert2, _
             adhcCenterVer, adhcCenterVer1, adhcCenterVer2: tmp = eCenterVert         ' ������������ �� ���������
' ����� ������������ ������� (�����������)
'    ' 2 ����������� �� 3 ��������� ����� �������
'    ' �����: 3x3 = 9 ����� ������������.
        Case adhcLeftTop, adhcLeftTop1:   tmp = eAlignLeftTop                 ' �� ������ �������� ����
        Case adhcRightTop, adhcRightTop1: tmp = eAlignRightTop                ' �� ������� �������� ����
        Case adhcLeftBottom, adhcLeftBottom1: tmp = eAlignLeftBottom          ' �� ������ ������� ����
        Case adhcRightBottom, adhcRightBottom1: tmp = eAlignRightBottom       ' �� ������� ������� ����
        Case adhcHorTop, adhcHor1Top, _
             adhcCenterHorTop, adhcCenterHor1Top: tmp = eCenterHorzTop                  ' �� �������� ���� ������������ �� �����������
        Case adhcHorBottom, adhcHor1Bottom, adhcHor2Bottom1, _
             adhcCenterHorBottom, adhcCenterHor1Bottom, adhcCenterHor2Bottom1: tmp = eCenterHorzBottom ' �� ������� ���� ������������ �� �����������
        Case adhcVerLeft, adhcVer1Left, adhcVer2Left1, _
             adhcCenterVerLeft, adhcCenterVer1Left, adhcCenterVer2Left1: tmp = eCenterVertLeft ' �� ������ ���� ������������ �� ���������
        Case adhcVerRight, adhcVer1Right, adhcVer2Right1, _
             adhcCenterVerRight, adhcCenterVer1Right, adhcCenterVer2Right1: tmp = eCenterVertRight ' �� ������� ���� ������������ �� ���������
        Case adhcCenter, adhcCenter1, adhcFull, adhcFull1: tmp = eCenter      ' ������������ ���������� �������
        Case adhcCascade:       tmp = eCascade                      ' ���������� (������ ��� ����� ??)
        Case eAlignUndef:       tmp = eAlignUndef                   '
        End Select
        GetAlignFromText = GetAlignFromText Or tmp ' ���������
    Next j
End Function
Public Function GetPlaceText(lngPlace As ePlace) As String
    Select Case lngPlace
    ' ������ �� ������
    Case ePlaceCenter: GetPlaceText = adhcPlaceCenter                      ' �� ������ (������)
    Case ePlaceToLeft: GetPlaceText = adhcPlaceToLeft                      ' ������ ����� �� ������
    Case ePlaceToRight: GetPlaceText = adhcPlaceToRight                    ' ������ ������ �� ������
    Case ePlaceToTop: GetPlaceText = adhcPlaceToTop                        ' ������ �� ������ ������
    Case ePlaceToBottom: GetPlaceText = adhcPlaceToBottom                  ' ������ �� ������ �����
    ' ������� �� ������
    Case ePlaceOnLeft: GetPlaceText = adhcPlaceOnLeft                      ' ������� ����� �� ������
    Case ePlaceOnRight: GetPlaceText = adhcPlaceOnRight                    ' ������� ������ �� ������
    Case ePlaceOnTop: GetPlaceText = adhcPlaceOnTop                        ' ������� �� ������ ������
    Case ePlaceOnBottom: GetPlaceText = adhcPlaceOnBottom                  ' ������� �� ������ �����
    ' ������ �� ����
    Case ePlaceToLeftTop: GetPlaceText = adhcPlaceToLeftTop                ' ������ ����� ������
    Case ePlaceToRightTop: GetPlaceText = adhcPlaceToRightTop              ' ������ ������ ������
    Case ePlaceToLeftBottom: GetPlaceText = adhcPlaceToLeftBottom          ' ������ ����� �����
    Case ePlaceToRightBottom: GetPlaceText = adhcPlaceToRightBottom        ' ������ ������ �����
    ' ������� �� ����
    Case ePlaceOnLeftToTop: GetPlaceText = adhcPlaceOnLeftToTop            ' ������� ����� � �������� ����
    Case ePlaceOnLeftToBottom: GetPlaceText = adhcPlaceOnLeftToBottom      ' ������� ����� � ������� ����
    Case ePlaceOnRightToTop: GetPlaceText = adhcPlaceOnRightToTop          ' ������� ������ � �������� ����
    Case ePlaceOnRightToBottom: GetPlaceText = adhcPlaceOnRightToBottom    ' ������� ������ � ������� ����
    Case ePlaceOnTopToLeft: GetPlaceText = adhcPlaceOnTopToLeft            ' ������� � ������ ���� ������
    Case ePlaceOnTopToRight: GetPlaceText = adhcPlaceOnTopToRight          ' ������� � ������� ���� ������
    Case ePlaceOnBottomToLeft: GetPlaceText = adhcPlaceOnBottomToLeft      ' ������� � ������ ���� �����
    Case ePlaceOnBottomToRight: GetPlaceText = adhcPlaceOnBottomToRight    ' ������� � ������� ���� �����
    ' ���������� (������ ��� ����� ??)
    Case eCascadeFromLeftTop: GetPlaceText = adhcCascadeFromLeftTop          ' ���������� �������� ������-����
    Case eCascadeFromRightTop: GetPlaceText = adhcCascadeFromRightTop        ' ���������� �������� �����-����
    Case eCascadeFromLeftBottom: GetPlaceText = adhcCascadeFromLeftBottom    ' ���������� �������� ������-�����
    Case eCascadeFromRightBottom: GetPlaceText = adhcCascadeFromRightBottom  ' ���������� �������� �����-�����
    Case Else: GetPlaceText = vbNullString 'adhcUndef
'Dim StrVal As String
'    StrVal = GetAlignText(lngPlace And eCenter, Short)
'    StrVal = StrVal & IIf(Short, adhcTo1, adhcTo)
'    StrVal = StrVal & GetAlignText((lngPlace And eCenter * &H10) / &H10, Short)
'    GetPlaceText = StrVal
    End Select
End Function
Public Function GetPlaceFromText(strPlace As String) As ePlace
' ����� ���������� ������ ������� ������������ ������� (OnLeftToTop, Center, etc)
    Select Case strPlace
    ' ������ �� ������
    Case adhcPlaceCenter: GetPlaceFromText = ePlaceCenter                       ' �� ������ (������)
    Case adhcPlaceToLeft: GetPlaceFromText = ePlaceToLeft                       ' ������ ����� �� ������
    Case adhcPlaceToRight: GetPlaceFromText = ePlaceToRight                     ' ������ ������ �� ������
    Case adhcPlaceToTop: GetPlaceFromText = ePlaceToTop                         ' ������ �� ������ ������
    Case adhcPlaceToBottom: GetPlaceFromText = ePlaceToBottom                   ' ������ �� ������ �����
    ' ������� �� ������
    Case adhcPlaceOnLeft: GetPlaceFromText = ePlaceOnLeft                       ' ������� ����� �� ������
    Case adhcPlaceOnRight: GetPlaceFromText = ePlaceOnRight                     ' ������� ������ �� ������
    Case adhcPlaceOnTop: GetPlaceFromText = ePlaceOnTop                         ' ������� �� ������ ������
    Case adhcPlaceOnBottom: GetPlaceFromText = ePlaceOnBottom                   ' ������� �� ������ �����
    ' ������ �� ����
    Case adhcPlaceToLeftTop: GetPlaceFromText = ePlaceToLeftTop                 ' ������ ����� ������
    Case adhcPlaceToRightTop: GetPlaceFromText = ePlaceToRightTop               ' ������ ������ ������
    Case adhcPlaceToLeftBottom: GetPlaceFromText = ePlaceToLeftBottom           ' ������ ����� �����
    Case adhcPlaceToRightBottom: GetPlaceFromText = ePlaceToRightBottom         ' ������ ������ �����
    ' ������� �� ����
    Case adhcPlaceOnLeftToTop: GetPlaceFromText = ePlaceOnLeftToTop             ' ������� ����� � �������� ����
    Case adhcPlaceOnLeftToBottom: GetPlaceFromText = ePlaceOnLeftToBottom       ' ������� ����� � ������� ����
    Case adhcPlaceOnRightToTop: GetPlaceFromText = ePlaceOnRightToTop           ' ������� ������ � �������� ����
    Case adhcPlaceOnRightToBottom: GetPlaceFromText = ePlaceOnRightToBottom     ' ������� ������ � ������� ����
    Case adhcPlaceOnTopToLeft: GetPlaceFromText = ePlaceOnTopToLeft             ' ������� � ������ ���� ������
    Case adhcPlaceOnTopToRight: GetPlaceFromText = ePlaceOnTopToRight           ' ������� � ������� ���� ������
    Case adhcPlaceOnBottomToLeft: GetPlaceFromText = ePlaceOnBottomToLeft       ' ������� � ������ ���� �����
    Case adhcPlaceOnBottomToRight: GetPlaceFromText = ePlaceOnBottomToRight     ' ������� � ������� ���� �����
    ' ���������� (������ ��� ����� ??)
    Case adhcCascadeFromLeftTop: GetPlaceFromText = eCascadeFromLeftTop             ' ���������� �������� ������-����
    Case adhcCascadeFromRightTop: GetPlaceFromText = eCascadeFromRightTop       ' ���������� �������� �����-����
    Case adhcCascadeFromLeftBottom: GetPlaceFromText = eCascadeFromLeftBottom   ' ���������� �������� ������-�����
    Case adhcCascadeFromRightBottom: GetPlaceFromText = eCascadeFromRightBottom ' ���������� �������� �����-�����
    Case Else: GetPlaceFromText = ePlaceUndef
    'Dim Result As eAlign: Result = ePlaceUndef
    '    On Error GoTo HandleError
    '' ���� ��������
    '    If IsNumeric(strPlace) Then Result = CLng(strPlace): GoTo HandleExit
    '' ���� ����������� � ������� ��������� ������� ��������� ��� ���������
    '' ������� strPlace �� adhcTo �� 2 �����
    'Dim Pos As Long, i As Long, arrDelim()
    '    arrDelim = Array(adhcTo, adhcTo1)
    '    For i = LBound(arrDelim) To UBound(arrDelim)
    '        Pos = InStr(1, strPlace, arrDelim(i)): If Pos > 1 Then Exit For
    '    Next i
    '    If Pos < 1 Then Err.Raise vbObjectError + 512
    '' ������ �������� ����� GetAlignFromText � ���������� �� ������� &h10
    '    Result = GetAlignFromText(Left$(strPlace, Pos - 1))                                ' ������ ����� �����
    '    Result = Result Or &H10 * GetAlignFromText(Mid$(strPlace, Pos + Len(arrDelim(i))))    ' ������ ������ �����
    End Select
'HandleExit:  GetPlaceFromText = Result: Exit Function
'HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function GetPictModeText(ObjMode As eObjSizeMode) As String
    Select Case ObjMode
    Case acOLESizeStretch:      GetPictModeText = adhcStretch             ' ������/���������� (�������� ���������)
    Case acOLESizeZoom:         GetPictModeText = adhcZoom                ' ���������������� ���������������
    Case apObjSizeZoomDown:     GetPictModeText = adhcDown                ' ���������������� ���������������, ������ ���������
    'Case acOLESizeAutoSize:     GetPictModeText = adhcAuto
    'Case acOLESizeClip:         GetPictModeText = vbNullString 'adhcClip  ' �� ������ ������. ���� ������ ������ ������� ������ - �������
    Case Else:                  GetPictModeText = vbNullString 'adhcClip  ' �� ������ ������. ���� ������ ������ ������� ������ - �������
    End Select
End Function

Public Function GetPictModeFromText(ObjModeText As String) As eObjSizeMode
    Select Case ObjModeText
    Case adhcStretch:           GetPictModeFromText = acOLESizeStretch    ' ������/���������� (�������� ���������)
    Case adhcZoom:              GetPictModeFromText = acOLESizeZoom       ' ���������������� ���������������
    Case adhcDown, adhcDown1:   GetPictModeFromText = apObjSizeZoomDown   ' ���������������� ���������������, ������ ���������
    'Case adhcAuto, adhcAuto1:   GetPictModeFromText = acOLESizeAutoSize  '
    'Case adhcClip:              GetPictModeFromText = acOLESizeClip      ' �� ������ ������. ���� ������ ������ ������� ������ - �������
    Case Else:                  GetPictModeFromText = acOLESizeClip       ' �� ������ ������. ���� ������ ������ ������� ������ - �������
    End Select
End Function
' --------------------
Private Function CustomPropertyGet( _
    PropName As String, _
    PropValue As Variant, _
    Optional PropObject As Object _
    ) As Boolean
' ������ ���������������� �������� ������������� �������
Const c_strProcedure = "CustomPropertyGet"
' PropName      - ��� ��������
' PropValue     - �������� ��������
' PropObject    - ������ � �������� ����������� ��������
Dim prp As Property
Dim Result As Boolean
    Result = False
    On Error GoTo HandleError
    If PropObject Is Nothing Then Set PropObject = CurrentProject ' ��-���������
    PropValue = PropObject.Properties(PropName)
    Result = True
HandleExit: CustomPropertyGet = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Private Function CustomPropertySet( _
    PropName As String, _
    PropValue As Variant, _
    Optional PropObject As Object, _
    Optional PropType As DataTypeEnum = dbText _
    ) As Boolean
' ��������� ���������������� �������� � ������� DAO ��� AccessObject
Const c_strProcedure = "CustomPropertySet"
' PropName      - ��� ��������
' PropValue     - �������� ��������
' PropObject    - ������ � �������� ����������� ��������
' PropType      - ��� ������ ��������
    On Error Resume Next
    If PropObject Is Nothing Then Set PropObject = CurrentProject ' ��-���������
    ' �������� �������� ��������
    PropObject.Properties(PropName) = PropValue
    If Err.Number = 0 Then GoTo HandleExit
    Err.Clear: On Error GoTo HandleExit
    ' ��� ������ �������� - ���������
    If TypeOf PropObject.Properties Is DAO.Properties Then
    ' ��������� DAO ��������
    Dim daoProp As DAO.Property
        Set daoProp = PropObject.CreateProperty(PropName, PropType, PropValue)
        PropObject.Properties.Append daoProp
    ElseIf TypeOf PropObject.Properties Is AccessObjectProperties Then
    ' ��������� AccessObject ��������
        PropObject.Properties.Add PropName, PropValue
    Else
        Err.Raise vbObjectError + 512
    End If
HandleExit: CustomPropertySet = Err.Number = 0: Err.Clear: Exit Function
End Function
Public Function WindowFreeze(hwnd As LongPtr): LockWindowUpdate (hwnd): End Function
Public Function WindowUnFreeze(): LockWindowUpdate (0): End Function
'---------------------
Public Function GetAlignPoint(Alignment As eAlign, _
    cX As Single, cY As Single, Optional Cascade As Boolean)
' ���������� ���������������� ���������� ����� �������� � ����������� �� ��������� ������ ������������
'---------------------
' ��������:
'   Alignment - ����� ������������
' ����������:
'   cX,cY     - ������� ����� �������� ����� �������������
'---------------------
    '
    ' Horz region anchor point position
    Select Case (Alignment And eCenterHorz)
    Case eLeft:         cX = 0            ' Left-to-Left
    Case eRight:        cX = 1            ' Right-to-Right
    Case eCenterHorz:   cX = 1 / 2        ' CenterHorz-to-CenterHorz
    End Select
    ' Vert region anchor point position
    Select Case (Alignment And eCenterVert)
    Case eTop:          cY = 0            ' Top-to-Top
    Case eBottom:       cY = 1            ' Bottom-to-Bottom
    Case eCenterVert:   cY = 1 / 2        ' CenterVert-to-CenterVert
    End Select
End Function

