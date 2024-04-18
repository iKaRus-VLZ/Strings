Attribute VB_Name = "~Sources"
Option Compare Database
Option Explicit
Option Base 0
#Const APPTYPE = 0          '0|1        '0=ACCESS,1=EXCEL ' not yet
#Const USEZIPCLASS = False  'False|True
'=========================
Private Const c_strModule As String = "~Sources"
'=========================
' ��������      : ������ ��� ������ � �������� ������� ����������
' ������        : 1.9.1.453985865
' ����          : 16.04.2024 14:04:34
' �����         : ������ �.�. (KashRus@gmail.com)
' ����������    : ��� USEZIPCLASS = True, ����� ������ clzZipArchive
' v.1.9.1       : 24.10.2022 - ��������� � ZipPack - ������ ������ �������� ���������� ������������� ��������
' v.1.9.0       : 04.10.2022 - �������� ������ ����������/���������� �������� ����������
' v.1.8.5       : 12.09.2022 - ��������� ��������� SourceUpdateAll ��� ��������� ���������� ������� ������ �������� �� ���������� ���������
' v.1.8.3       : 24.04.2019 - ��������� ����������� ������ ������ ����������� ������ �������
' v.1.8.1       : 19.04.2019 - �������� ������ ���������� ������� - ������ ������ ����������� � ����������� bas � cls.
' v.1.7.12      : 02.04.2019 - ��������� ��������� UpdateFunc - ��������� ������ �������
' v.1.7.6       : 27.09.2018 - � UpdateModule ��������� ����������� ��������� ����������� � ������
'=========================
' ToDo: ������� ������������� ������������� ��� ObjectType (acObjectType|MSysObjectsType|vbext_ComponentType)
'=========================
' �������� �������� � ������ ������:
'-------------------------
' AskAuthor - ����������� ���������� � ������������ � ��������� � ��������� ������� (������������ � ������������ � �������)
' AskAppData - ����������� ���������� � ���������� � ��������� � ��������� �������
' ProjBackup/ProjRestore - ���������/��������������� ��� ������� ������� �/�� �����
' UpdateModule/UpdateFunc - ��������� ������ ������/�������
'=========================
' ���������� � ������ ���� A.B.C.D[r][n], ���
'   A � ������� ����� ������ (major version number).
'   B � ��������������� ����� ������ (minor version number).
'   C � ����� ������, ����� ���������� �������� �� ������ ��� ������������ ������ A.B (build number).
'   D � ����� �������, �������� ����� ����������� ������������� ����������� ������������ �������� ������ (SVN). ����� ������� SVN ������ ������������������ � ������� ������� � AssemblyInfo ��� ������ ������ ������ (revision number).
'       ��������� � Access ����� ������ �� ������������� - ����� ����� ������� ���� � ����:
'       DDDDDTTTT = CCur(Now)*10^c_bytTimeDig.
'       �������������� ���� �����: CDate(DDDDDTTTT/10^c_bytTimeDig)
'   [r] � �������� ����������� ������, [n] - ����� ������
'       Pre-alpha (pa) � ������������� ����� ������ ����� ��� �������. ��������������� �������� ����������� � ����������� � ������� ����������� ������. Pre-alpha ������ �� �������� ������ ���������� ��.
'       Alpha (a) � ������������� ����� ���������� ���������� ������ �����������. ������� � alpha ������ ����� ���������� �� ���������������, � ��� ������ �� ����� ���������� ������ � ���� ����� �� ��������� ������. ���� ��������������� ������� ����������� �� ������������ ������ ������������� ���������� �� � ���������� ������.
'       Beta (b) � ������������� ����� ���������� ������������. ��� ������ �����, ������� ������� �� ������� ������ ���������� ��. �� ���� ����� ����������� ��������� �� ������������� �� ���������� �������� � ������ ��������� �������������� ������� � �����������.
'       Release Candidate (rc) � ���� ���������� ���������� � ��������� ������������, ��� ��������� �� ���������� ������ ������ ����������. �� ���� ����� ����� ��������� ��������� � ������������ � ������������ ��������.
'       Release to manufacturing ��� Release to marketing (rtm) � ������ ��� ��������� ����, ��� �� ������������� ���� ����������� ��������, � ������ ��� ��������� ���������������. RTM �� ���������� ������� �������� ������ (���� ��� ��������) � ������ ���� ��� ��������� ����, ��� �������� ���������� ��� ��������� ���������������.
'       General availability (ga) � ��������� �����, ��������������� ���������� ���� ����� �� ���������������� ��������, ������� ��������� ����� � �������� ����� ��� ��� �� ���������� ���������.
'       End of life (eol) � ������ �� �������� � ��������� �������� ���������.
'-------------------------

'Private Const c_strLibPath = "%UserProfile%\Documents\VBA Code\" ' ���� � ���������� �������� ����������
Private Const c_strLibPath = "D:\Documents\VBA Code\" ' ���� � ���������� �������� ����������
Private Const c_strPrefModName = "Private Const c_strModule As String = "   ' ������ ������ ������ ���� ��� ������ � ������ � ������

' ��� �������� ����������� ��� ����������
' ������� ������ � ��������� ������� �������
Private Const strBegLineMarker = "'=== BEGIN INSERT ==="
Private Const strEndLineMarker = "'==== END INSERT ===="
' ��������� ������
    ' �� ������� "Attribute VB_"
    ' �� ������� "[Public | Private | Friend] [Static] [Function | Sub | Property]
Private Const c_strPrefLen As Byte = 17 ' ����� �������� �� ��������� ":"
'Private Const c_strPrefModAttr As String = "Attribute VB_"
Private Const c_strPrefModLine = "'========================="
Private Const c_strPrefModNone = "'               :"
Private Const c_strPrefModDesc = "' ��������      :"
Private Const c_strPrefModAuth = "' �����         :"
Private Const c_strPrefModVers = "' ������        :"
Private Const c_strPrefModDate = "' ����          :"
Private Const c_strPrefModComm = "' ����������    :"
Private Const c_strPrefModHist = "^\s*'\s*v\.\s*(\d{1,}?\.\d{1,}?\.\d{1,}?)\s*:\s*(\d{1,2}?\.\d{1,2}?\.\d{2,4}?)\s*-\s*(.*)"
Private Const c_strCodeProcBeg = "^\s*(Public\s+|Private\s+|Friend\s+)?(Static\s+)?(Function|Sub|Property)" '[Public | Private | Friend] [Static] [Function | Sub | Property]" ' ��������� ������ ������������� � ����������� ������ ���������
Private Const c_strPrefModDebg = "#Const DEBUGGING"

Private Const c_strLibPathLine = "Private Const c_strLibPath = "

Private Const c_strCodeHeadBeg = "CodeBehindForm"       ' ����� ������ ������ �����/������ ���������� ����� ���� "CodeBehindForm"
Private Const cEmptyVers = "0.0.0.0", cEmptyDate = #1/1/1980#

Private Const c_strMSysObjects = "MSysObjects"

Private Const c_strHyphen = " _"                ' ������� ������ � ����
Private Const c_strSpace = " "                  ' ������
Private Const c_strBrokenQuotes = """ & """     ' " & " - ����������� ���� ��������� �����

Private Const c_strDelim As String = ";"
Private Const c_strInDelim As String = ", " ' ����������� ��������� ��������� �������

' ��� ���������� �������
Private Const c_strSrcPath As String = "SRC"    ' ��� �������� ��� �������� ����������
Private Const c_strBreakProcessMessage = "�������� �������?"
' ������� ������������ ��� ������� (���� �� ���������� ������ ������������ ����)
Private Const c_strObjIgnore = "~Sources;clsProgress;frmSERV_Progress" '
' ���� � �������� ����� ����������� ��������
Private Const c_strObjTypModule = "Module"
#If APPTYPE = 0 Then        ' APPTYPE=Access
Private Const c_strObjTypAccFrm = "Form", c_strFrmModPref = c_strObjTypAccFrm & "_"
Private Const c_strObjTypAccRep = "Report", c_strRepModPref = c_strObjTypAccRep & "_"   ' ������� ������ ������
Private Const c_strObjTypAccMac = "Macro"
Private Const c_strObjTypAccQry = "Query"
Private Const c_strObjTypAccTbl = "Table"
#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
#Else                       ' APPTYPE=��������� ����� ���
#End If                     ' APPTYPE

' ���������� ����� �������� �������� �������
Private Const c_strObjExtZip = "zip"
' ����� � ���������� ����������� ��������
Private Const c_strAppNamPrj = "Project"  ' ��� ����� ������� �������
Private Const c_strObjExtUndef = "src"  ' ������ ���������� ����� ������ �������
' ���������� ����������� ������ �������� �������
Private Const c_strObjExtBas = "bas"    ' ����������� ������
Private Const c_strObjExtCls = "cls"    ' ������ ������
Private Const c_strObjExtXml = "xml"    ' ��������� ������� � XML
Private Const c_strObjExtTxt = "txt"    ' ��������� ������� � TXT
Private Const c_strObjExtCsv = "csv"    ' ��������� ������� � CSV (����� � �������������)
#If APPTYPE = 0 Then        ' APPTYPE=Access
Private Const c_strObjExtFrm = "accfrm" ' ����� Access (������� ������)
Private Const c_strObjExtRep = "accrep" ' ����� Access (������� ������)
Private Const c_strObjExtLnk = "acclnk" ' ��������� �������
Private Const c_strObjExtQry = "accqry" ' ������ Access
Private Const c_strObjExtDoc = "doccls" ' ������ ������ ��������� Access (Form ��� Report)
Private Const c_strObjExtMac = "accmac" ' ������ Access
Private Const c_strObjExtPrj = "accprj" ' ���������� ����� ������� ������� Access
#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
#Else                       ' APPTYPE=��������� ����� ���
#End If                     ' APPTYPE

' ����� ��������/������ � ������:
' ��� �������
Private Const c_strPrpSecProject = "VBProject"
Private Const c_strPrpKeyPrj = "ProjectName"    ' VBE.ActiveVBProject.Name

Private Const c_strPrpSecCustom = "Custom"      ' CurrentProject.Properties
Private Const c_strPrpKeyApp = "Application"
Private Const c_strPrpKeyVer = "Version"
Private Const c_strPrpKeyAuthor = "Author"
Private Const c_strPrpKeySupport = "Support"

' ��� ������� �������
Private Const c_strPrjSecName = "Project"
Private Const c_strPrjKeyName = "Name"
Private Const c_strPrjKeyDesc = "Description"
Private Const c_strPrjKeyHelp = "Help"
' ��� ������� ���� ������
Private Const c_strDbsSecName = "Database Properties"
' ��� ���������������� �������
Private Const c_strPrpSecName = "Custom Properties"
' ��� ������ (References)
Private Const c_strRefSecName = "References"
' ��� ������������� ������
Private Const c_strLnkSecParam = "Params"
Private Const c_strLnkKeyTable = "TableName"
Private Const c_strLnkKeyConnect = "ConnectString"
Private Const c_strLnkKeyLocal = "LocalName"
Private Const c_strLnkKeyAttribute = "Attributes"

' ��� ������
Private Const c_strPrefVerComm As String = " v."
Private Const c_strVerDelim = "."   ' ����������� ��������� ������
Private Const c_bytMajorDig = 2     ' ���������� �������� ��� A
Private Const c_bytMinorDig = 3     ' ���������� �������� ��� B
Private Const c_bytDateDig = 5      ' ���������� �������� � ��������� ������� DDDDDTTTT ������������ �� ��� ����
Private Const c_bytTimeDig = 4      ' ���������� �������� � ��������� ������� DDDDDTTTT ������������ �� ��� ������� (�� ������ 4 ����� � Long ������������)

Public Enum appRelType      ' ��� ������
    appReleaseNotDefine = 0     '�� �����
    appReleasePreAlpha = 1      'Pre-alpha (pa) � ������������� ����� ������ ����� ��� �������. ��������������� �������� ����������� � ����������� � ������� ����������� ������. Pre-alpha ������ �� �������� ������ ���������� ��.
    appReleaseAlpha = 2         'Alpha (a) � ������������� ����� ���������� ���������� ������ �����������. ������� � alpha ������ ����� ���������� �� ���������������, � ��� ������ �� ����� ���������� ������ � ���� ����� �� ��������� ������. ���� ��������������� ������� ����������� �� ������������ ������ ������������� ���������� �� � ���������� ������.
    appReleaseBeta = 3          'Beta (b) � ������������� ����� ���������� ������������. ��� ������ �����, ������� ������� �� ������� ������ ���������� ��. �� ���� ����� ����������� ��������� �� ������������� �� ���������� �������� � ������ ��������� �������������� ������� � �����������.
    appReleaseCandidate = 4     'Release Candidate (rc) � ���� ���������� ���������� � ��������� ������������, ��� ��������� �� ���������� ������ ������ ����������. �� ���� ����� ����� ��������� ��������� � ������������ � ������������ ��������.
    appReleaseToMarketing = 5   'Release to manufacturing ��� Release to marketing (rtm) � ������ ��� ��������� ����, ��� �� ������������� ���� ����������� ��������, � ������ ��� ��������� ���������������. RTM �� ���������� ������� �������� ������ (���� ��� ��������) � ������ ���� ��� ��������� ����, ��� �������� ���������� ��� ��������� ���������������.
    appReleaseGeneral = 6       'General availability (ga) � ��������� �����, ��������������� ���������� ���� ����� �� ���������������� ��������, ������� ��������� ����� � �������� ����� ��� ��� �� ���������� ���������.
    appReleaseEOL = 7           'End of life (eol) � ������ �� �������� � ��������� �������� ���������.
End Enum
Public Enum appVerType      ' ��� ������� ������
    appVerMajor = 0
    appVerMinor = 1
    appVerBuild = 2
    appVerRevis = 3
    appRelease = 4
    appRelSubNum = 5
End Enum
Public Type typVersion  ' ��� ������ ������
    Major As Long           ' A ������� ������
    Minor As Long           ' B ������� ������
    Build As Long           ' C ����
    Revision As Long        ' D ������� (String?)
    Release As appRelType   ' [r] ��� ���� ������
    RelSubNum As Integer    ' [n] ����� ������
    RelShort As String      ' ������� ����� ������
    RelFull As String       ' ������ ����� ������
    VerDate As Date         ' ���� ������ CDate(Revision/10^c_bytTimeDig)
    VerCode As Long         ' ��� ������ AABBB
End Type

Private Const c_bolDebugMode = True ' ��������������� ����� ������� ��� ������� CompileAll

Private Const c_strSysPath = "%SYSTEMROOT%", c_strPrgPath = "%PROGRAMFILES%"
Private Const c_strSys32Path = c_strSysPath & "\System32\", c_strSys64Path = c_strSysPath & "\SysWoW64\"
Private Const c_strRegKey = "HKEY_CLASSES_ROOT\TypeLib\"

' ��� �������������� ����
Private Const c_strCodeLen = 2
Private Const c_strCodeSym = "%"
Private Const c_strHexPref = "&H"

Private Const c_strSymbDigits = "0123456789"
Private Const c_strSymbRusAll = "�����Ũ��������������������������"
Private Const c_strSymbEngAll = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Private Const c_strOthers = " -~_"

Private Const msoFileDialogViewList = 1
Private Const msoFileDialogFolderPicker = 4

Public Enum eDataType
    dbBoolean = 1
    dbByte = 2
    dbInteger = 3
    dbLong = 4
    dbCurrency = 5
    dbSingle = 6
    dbDouble = 7
    dbDate = 8
    dbBinary = 9
    dbText = 10
    dbLongBinary = 11
    dbMemo = 12
    dbGUID = 15
    dbBigInt = 16
    dbVarBinary = 17
    dbChar = 18
    dbNumeric = 19
    dbDecimal = 20
    dbFloat = 21
    dbTime = 22
    dbTimeStamp = 23
End Enum

' ���������� ����� � API �������
'--------------------------------------------------------------------------------
' POINTER
'--------------------------------------------------------------------------------
#If VBA7 = 0 Then       'LongPtr trick by @Greedo (https://github.com/Greedquest)
Private Enum LongPtr
    [_]
End Enum
#End If
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>  PtrSafe, LongPtr and LongLong
Private Const PTR_LENGTH As Long = 8
#Else                   '<OFFICE97-2010>        Long
Private Const PTR_LENGTH As Long = 4
#End If                 '<WIN32>
'--------------------------------------------------------------------------------
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
''#if (_WIN32_WINNT >= 0x0500)
'    pvReserved As LongPtr
'    dwReserved As Long
'    FlagsEx As Long
''#endif // (_WIN32_WINNT >= 0x0500)
End Type
Private Type BROWSEINFO
    hOwner As LongPtr
    pidlRoot As LongPtr
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As LongPtr
    lParam As LongPtr
    iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1

Private Enum vbext_ProcKind
    vbext_pk_Proc = 0 'A Sub or Function procedure.
    vbext_pk_Let = 1 'A Property Let procedure.
    vbext_pk_Set = 2 'A Property Set procedure.
    vbext_pk_Get = 3 'A Property Get procedure.
End Enum
'!!! ������� ������ ��� ��� ObjectType
Private Enum vbext_ComponentType
    vbext_ct_Undef = 0
    vbext_ct_StdModule = 1
    vbext_ct_ClassModule = 2
    vbext_ct_MSForm = 3
    vbext_ct_ActiveXDesigner = 11
    vbext_ct_Document = 100
End Enum
Private Enum msys_ObjectType
    msys_ObjectUndef = 0
    msys_ObjectTable = 1
    msys_ObjectQuery = 5
    msys_ObjectLinked = 6
    msys_ObjectForm = &HFFFF8000
    msys_ObjectReport = &HFFFF8004
    msys_ObjectMacro = &HFFFF8002
    msys_ObjectModule = &HFFFF8007
End Enum

Public Enum appObjectType
' ��� ������� ����������
    appObjTypUndef = 0      ' &HFFFFFFFF
' ������ �������
    appObjTypMod = &H100        ' ������� (���������) ���� ������
    appObjTypBas = appObjTypMod * vbext_ct_StdModule        ' ����������� ������ (vbext_ct_StdModule, acObjectModule; msys_ObjectModule)
    appObjTypCls = appObjTypMod * vbext_ct_ClassModule      ' ������ ������ (vbext_ct_ClassModule, acObjectModule; msys_ObjectModule)
    appObjTypMsf = appObjTypMod * vbext_ct_MSForm           ' ������ MSForm (vbext_ct_MSForm)
    appObjTypAxd = appObjTypMod * vbext_ct_ActiveXDesigner  ' ������ ActiveXDesigner (vbext_ct_ActiveXDesigner)
    appObjTypDoc = appObjTypMod * vbext_ct_Document         ' ������ ��������� (vbext_ct_Document; (acForm; msys_ObjectForm)|(acReport; msys_ObjectReport))
' ��������� �������
#If APPTYPE = 0 Then        ' APPTYPE=Access
    appObjTypAccDoc = &H10000   ' ������� ���� ���������
    appObjTypAccTbl = acTable + appObjTypAccDoc             ' ������� (acTable; msys_ObjectTable)
    appObjTypAcclnk = acTable + &H80 + appObjTypAccDoc      ' ��������� ������� (acObjectTable & msys_ObjectLinked)
    appObjTypAccQry = acQuery + appObjTypAccDoc             ' ������ Access (acQuery; msys_ObjectQuery)
    appObjTypAccMac = acMacro + appObjTypAccDoc             ' ������ Access (acObjectMacro; msys_ObjectMacro)
    appObjTypAccFrm = acForm + appObjTypDoc + appObjTypAccDoc ' ����� Access (acForm; msys_ObjectForm; ModulePrefix="Form_")
    appObjTypAccRep = acReport + appObjTypDoc + appObjTypAccDoc ' ����� Access (acReport & msys_ObjectReport; ModulePrefix="Report_")
    appObjTypAccDap = acDataAccessPage + appObjTypAccDoc    '
    appObjTypAccSrv = acServerView + appObjTypAccDoc        '
    appObjTypAccDia = acDiagram + appObjTypAccDoc           '
    appObjTypAccPrc = acStoredProcedure + appObjTypAccDoc   '
    appObjTypAccFun = acFunction + appObjTypAccDoc          '
#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
    appObjTypXlsDoc = &H20000   ' ������� ���� ���������
'' ...
#Else                       ' APPTYPE=��������� ����� ���
'' ...
#End If                     ' APPTYPE
End Enum
Public Enum ObjectRwType
' ��� �������� ������/������ ��������
    orwUndef = 0
    orwSrcNewer = 1    ' ���������/��������� ���� �������� �����
    orwDestMiss = 2    ' ���������/��������� ���� ���������� �����������
    orwSrcOlder = 4    ' ���������/��������� ���� �������� ������
    orwSrcNewerOrDestMissing = orwSrcNewer + orwDestMiss
    orwAlways = orwSrcNewer + orwDestMiss + orwSrcOlder
End Enum

Private Enum m_CodeLineType
    m_CodeNone = -1     ' ������ �� ������ ��������� ������
    m_CodeHead = 0      ' ��������� ������. �������� �������� �� ������ ��� ������ ��������� ���������
    m_CodeName = 1      ' ��������� ������: ModName
    m_CodeDesc = 3      ' ��������� ������: ModDesc
    m_CodeVers = 4      ' ��������� ������: ModVers
    m_CodeDate = 5      ' ��������� ������: ModDate
    m_CodeAuth = 6      ' ��������� ������: ModAuth
    m_CodeComm = 7      ' ��������� ������: ModComm
    m_CodeHist = 8      ' ��������� ������: ModHist
    m_CodeProc = 100    ' ������� �������� ������ (������ ����� ���������)
End Enum

Private Enum m_ModErrors
' ������ ��� ��������� ������ ������
   m_errProcNameWrong = vbObjectError + 511:            ' ������������ ��� ���������
   m_errModuleNameWrong = vbObjectError + 512:          ' ������������ ��� ������
   m_errModuleIsActive = vbObjectError + 513:           ' ���������� �������� �������� ������
   m_errModuleDontFind = vbObjectError + 514:           ' ������ �� ������!
' ������ �������������� � ���������
   m_errObjectTypeUnknown = vbObjectError + 520         ' ���������� ��������� ������ ������� ����
   m_errObjectActionUndef = vbObjectError + 530         ' ��������� ��� ������ � ��������
   m_errObjectCantRemove = vbObjectError + 538          ' �� ������� ������� ������ �� �������
   m_errCantGetSrcVersion = vbObjectError + 541         ' �� ������� ��������� ������ ���������
   m_errCantGetDestVersion = vbObjectError + 542        ' �� ������� ��������� ������ ����������
' ������� �������� ��� ��������� � ���������
   m_errWrongVersion = vbObjectError + 1001             ' �������������� ������
   m_errDestMissing = vbObjectError + 1002              ' ������� ������ �����������
   m_errDestExists = vbObjectError + 1003               ' ������� ������ ��� ����������
   m_errSkippedByUser = vbObjectError + 1004            ' ������� ������������
   m_errSkippedByList = vbObjectError + 1005            ' ������� � ������ ��������
' ������� �������� ��� ��������� � ���������
   m_errZipError = vbObjectError + 2001                 ' ������ ��������
   m_errUnZipError = vbObjectError + 2002               ' ������ ����������
   m_errExportError = vbObjectError + 2101              ' ������ ��������
   m_errImportError = vbObjectError + 2102              ' ������ �������
End Enum

#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>
Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare PtrSafe Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Boolean
Private Declare PtrSafe Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As LongPtr
Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As LongPtr, ByVal pszPath As String) As Boolean

Private Declare PtrSafe Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal strSection As String, ByVal strRegKeyName As String, ByVal strDefault As String, ByVal strReturned As String, ByVal lngSize As Long, ByVal strFilename As String) As Long
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal strSection As String, ByVal strRegKeyNam As String, ByVal strValue As String, ByVal strFilename As String) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else                   '<OFFICE97-2010>
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (ByRef pOpenfilename As OPENFILENAME) As Boolean
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Boolean

Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal strSection As String, ByVal strRegKeyName As String, ByVal strDefault As String, ByVal strReturned As String, ByVal lngSize As Long, ByVal strFilename As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal strSection As String, ByVal strRegKeyNam As String, ByVal strValue As String, ByVal strFilename As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If                 '<WIN32>
'-------------------------
' ������� ��� ��������� ������� ���������� � ���������� � ������������
'-------------------------
Public Function UpdateAppVer(Optional strVersion As String)
' ���������� � ������������� � �������� ������ ����������
Const c_strProcedure = "UpdateAppVer"
Dim Result As Boolean ':Result = False
    On Error GoTo HandleError
Dim strTitle As String:     strTitle = "������ ����������"
Dim strMessage As String:   strMessage = "������� ������ ����������."
    If Len(strVersion) = 0 Then Call PropertyGet(c_strPrpKeyVer, strVersion)
    Dim strVer1 As String, strVer2 As String, datVerDate As Date
    VersionGet strVersion, VerShort:=strVer1, VerDate:=datVerDate   ' �������� �������� ������ �� ���������
    strVersion = VBA.Trim$(InputBox(strMessage, strTitle, strVer1)) ' ���������� ����� ������
    VersionGet strVersion, VerShort:=strVer2, VerDate:=datVerDate   ' �������� �������� ������ �� �����
    ' ���� ������ �� ��������� - ���������
    If strVer2 <> strVer1 Then datVerDate = Now: VersionSet strVersion, VerDate:=datVerDate: Call PropertySet(c_strPrpKeyVer, strVersion)
    Result = True
HandleExit:     UpdateAppVer = Result: Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function

Public Function AskAppData( _
    Optional strAppName As String, _
    Optional strVersion As String, _
    Optional strCodeName As String, _
    Optional strDescription As String _
    )
' ���������� ���/������ ���������� � ��������� ��������������� �������� �������
Const c_strProcedure = "AskAppData"
Dim strMessage As String, strTitle As String
Dim Result As Boolean

    Result = False
    On Error GoTo HandleError
    strTitle = "���������� ������� ����������!"
    strMessage = "����� ����������� ���������� ������� ����������," & vbCrLf _
        & "���������� �� �������� � ������� ������ ����������."
    MsgBox strMessage, vbOKOnly Or vbExclamation, strTitle
' ���������� � ������������� � �������� �������� ����������
    strTitle = "�������� ����������"
    strMessage = "������� �������� ����������."
    If Len(strAppName) = 0 Then Call PropertyGet(c_strPrpKeyApp, strAppName): If Len(strAppName) = 0 Then strVersion = "MyApp"
    strAppName = VBA.Trim$(InputBox(strMessage, strTitle, strAppName)): Call PropertySet(c_strPrpKeyApp, strAppName)
' ���������� � ������������� � �������� ������ ����������
    UpdateAppVer strVersion
' ���������� � ������������� ������� ��� ���������� (��� �������)
' VBE.ActiveVBProject.Name
    strTitle = "������� ��� ����������"
    strMessage = "������� ������� ��� ����������."
    If Len(strCodeName) = 0 Then strCodeName = VBE.ActiveVBProject.NAME
    strCodeName = VBA.Trim$(InputBox(strMessage, strTitle, strCodeName)): VBE.ActiveVBProject.NAME = strCodeName
' VBE.ActiveVBProject.Description
    strTitle = "�������� ����������"
    strMessage = "������� ������� �������� ����������."
    If Len(strDescription) = 0 Then strDescription = VBE.ActiveVBProject.Description
    strDescription = VBA.Trim$(InputBox(strMessage, strTitle, strDescription)): VBE.ActiveVBProject.Description = strDescription
    
    Result = True
HandleExit:     AskAppData = Result: Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function
Public Function AskAuthor( _
    Optional strAuthor As String, _
    Optional strSupport As String)
' ���������� ���/�������� ������������ � ��������� ��������������� �������� �������
Const c_strProcedure = "AskAuthor"
Dim strMessage As String, strTitle As String
    On Error GoTo HandleError
    strTitle = "���������� ������� ����������!"
    strMessage = "����� ����������� ���������� ������� ����������," & vbCrLf _
        & "���������� �� ��� � ���������� ������ ������������." & vbCrLf _
        & "��� ������ ������������ ��� �������� ������ ��������� " & vbCrLf _
        & "� ������������ ��� ������������� ������� UpdateModule � UpdateFunc."
    MsgBox strMessage, vbOKOnly Or vbExclamation, strTitle
' ���������� � ������������� � �������� ������� ��� ������������
    strTitle = "��� ������"
    strMessage = "������� ������ �� ����� ������������, " & vbCrLf & _
        "������� ����� ���������� � ������������ " & vbCrLf & _
        "� ���������� �������/�������."
    If Len(strAuthor) = 0 Then Call PropertyGet(c_strPrpKeyAuthor, strAuthor): If Len(strAuthor) = 0 Then strAuthor = "Unknown"
    strAuthor = VBA.Trim$(InputBox(strMessage, strTitle, strAuthor)): Call PropertySet(c_strPrpKeyAuthor, strAuthor)
' ���������� � ������������� � �������� ������� �������� ������������
    strTitle = "�������� ������"
    strMessage = "������� ���������� ������ ������������," & vbCrLf & _
        "������� ����� ���������� � ������������ " & vbCrLf & _
        "� ���������� �������/������� ������ � ������."
    If Len(strSupport) = 0 Then Call PropertyGet(c_strPrpKeySupport, strSupport): If Len(strSupport) = 0 Then strSupport = "dont@mail.me"
    strSupport = VBA.Trim$(InputBox(strMessage, strTitle, strSupport)): Call PropertySet(c_strPrpKeySupport, strSupport)
HandleExit:     Exit Function
HandleError:    Err.Clear: Resume HandleExit
End Function
Public Function Author() As String
Dim Result As String
    On Error GoTo HandleError
    Call PropertyGet(c_strPrpKeyAuthor, Result): If Len(Result) = 0 Then AskAuthor strAuthor:=Result
HandleExit:     Author = Result: Exit Function
HandleError:    Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function Support() As String
Dim Result As String
    On Error GoTo HandleError
    Call PropertyGet(c_strPrpKeySupport, Result): If Len(Result) = 0 Then AskAuthor strSupport:=Result
HandleExit:     Support = Result: Exit Function
HandleError:    Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function AppName() As String
Dim Result As String
    On Error GoTo HandleError
    Call PropertyGet(c_strPrpKeyApp, Result): If Len(Result) = 0 Then AskAppData strAppName:=Result
HandleExit:     AppName = Result: Exit Function
HandleError:    Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Public Function AppVersion() As String
Dim Result As String
    On Error GoTo HandleError
    Call PropertyGet(c_strPrpKeyVer, Result): If Len(Result) = 0 Then AskAppData strVersion:=Result
HandleExit:     AppVersion = Result: Exit Function
HandleError:    Result = vbNullString: Err.Clear: Resume HandleExit
End Function
'-------------------------
' ������� ����������/�������������� �������
'-------------------------
Public Function Init()
' ������������� �������
Dim Result As VbMsgBoxResult
    CloseAll
    Dim strMessage As String
    strMessage = "������������ ���� � �������� �����?"
    Result = MsgBox(strMessage, vbQuestion + vbYesNo)
    If Result = vbYes Then
        strMessage = "��� ���������� ��������� ����� �������������� ���� " & vbCrLf & _
            "��������� �������������� �� ������ ���� ����������� ������ � ��.��������." & vbCrLf & vbCrLf & _
            "������ ��� �������� ��������� ����� ����� �������������� �� ������ ������ Access." & vbCrLf & _
            "� ����� ������ ���� ������������� ���� � ��������� �������������� �������� ������ ""Setup"" �����."
        Call MsgBox(strMessage, vbExclamation + vbOKOnly)
        SourcesRestore
        'App.Setup ' !!! �������� ����� ��������������
    End If
'    ' ����������� ������ ������������
'    strMessage = "�������� ������ � ������������?"
'    Result = MsgBox(strMessage, vbQuestion + vbYesNo)
'    If Result = vbYes Then AskAuthor
End Function
Public Function SourcesBackup(Optional BackupPath As String, _
    Optional WithoutData As Boolean = True) As Boolean
' ��������� ��� ������� ������� � �������� ����� ��� ������������ ��������������
Const c_strProcedure = "SourcesBackup"
Dim Result As Boolean
On Error GoTo HandleError

'Dim WithoutData As Boolean:     WithoutData = True      ' ��� ������ ���������� ������� ������ (������� � �.�.)
Dim WriteType As ObjectRwType:  WriteType = orwAlways   ' ������ ����� ���� �������� ���������� ���������� �� ������
Dim AskBefore As Boolean:       AskBefore = False       ' �� ���������� ������������ ����� ����������� �������
Dim UseTypeFolders As Boolean:  UseTypeFolders = True   ' ����� � ������ ��������� �� ������ ��������
Dim DelAfterZip As Boolean:     DelAfterZip = True      ' ������� ��������� ����� ����� �������������

#If APPTYPE = 0 Then        ' APPTYPE=Access
    DoCmd.SetWarnings False ' ��������� ��������������
    Call SysCmd(504, 16484) ' ��������� ��� ������
#Else
#End If                     ' APPTYPE

Dim strCaption As String, strMessage As String
Dim ParentPath As String, FilePath As String, FileName As String
' ������� ���� ����������
    If Len(BackupPath) = 0 Then BackupPath = oFso.BuildPath(CurrentProject.path, c_strSrcPath)
    If Not oFso.FolderExists(BackupPath) Then Call oFso.CreateFolder(BackupPath) 'Then Err.Raise 76 ' Path not Found
' ������� ��� ������
    FileName = Split(CurrentProject.NAME, ".")(0)
    FileName = FileName & "_" & VBA.Format$(Now(), "YYYYMMDD_hhmmss")
    ParentPath = oFso.BuildPath(BackupPath, FileName)
    If Not oFso.FolderExists(ParentPath) Then Call oFso.CreateFolder(ParentPath) 'Then Err.Raise 76 ' Path not Found
' ��������� ��������� ���� �������� ���������� ����������
Dim colObjects As Collection:   Set colObjects = New Collection
    If Not p_ObjectsCollectionCreate(colObjects, WithoutData) Then Err.Raise m_errExportError
' �������������� ������������
Dim prg As clsProgress: Set prg = New clsProgress
    strCaption = "���������� �������� �������"
    strMessage = strCaption & " �: """ & BackupPath & """"
    prg.Init pCount:=1, pMin:=1, pMax:=colObjects.Count + 1, pCaption:=strCaption, pText:=strMessage ': prg.ProgressStep = 1
' ��������� �������� � ������������ ������ �������
    prg.Detail = "���������� ������� � ������������ ������ �������"
    FilePath = oFso.BuildPath(ParentPath, c_strAppNamPrj & "." & c_strObjExtPrj)
    p_PropertiesWrite FilePath
    p_ReferencesWrite FilePath
    prg.Update
' ��������� ��������� ��� �������
    Result = p_ObjectsBackup(colObjects, ParentPath, prg, WriteType, AskBefore, UseTypeFolders, strMessage) = 0:   If Not Result Then Err.Raise m_errExportError
    'If colObjects.Count > 0 Then ' �� ��� �������� ������� ���������
    Set colObjects = Nothing
    
' ����������� ��������
    FilePath = ParentPath & "." & c_strObjExtZip
    prg.Detail = "��������� ���������� ������� � ������������ ������ �������." & vbCrLf & _
                 "��� �������� ���������� �������� � �����: " & FilePath
#If USEZIPCLASS Then
    Result = oZip.AddFromFolder(ParentPath & "\*.*", True, , True)
    Result = oZip.CompressArchive(FilePath)
    If Result And DelAfterZip Then Result = oFso.DeleteFolder(ParentPath) = 0
#Else
    Result = ZipPack(FilePath:=ParentPath & "\*.*", ZipName:=FilePath, DelAfterZip:=True)
#End If ' USEZIPCLASS
    If Not Result Then Err.Raise m_errZipError
#If APPTYPE = 0 Then        ' APPTYPE=Access
    DoCmd.SetWarnings True  ' �������� ��������������
    Call SysCmd(504, 16484) ' ��������� ��� ������
#Else
#End If                     ' APPTYPE
' ����������� ������������ � �����������
    strCaption = "������� ��������"
    strMessage = strMessage & vbCrLf & "������� �������� ������� ���� ��������� �:" & vbCrLf & FilePath & "." & vbCrLf
    Call MsgBox(strMessage, vbInformation + vbOKOnly, strCaption)
HandleExit:     SourcesBackup = Result: Set prg = Nothing: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err.Number
    Case 75:    Message = "������ �������!" ' �������� ����� ��� �������
    If Len(BackupPath) > 0 Then Message = Message & " BackupPath=""" & BackupPath & """ "
    Case 76:    Message = "���� �� ������!" ' �������� ��� ������������ �����
    If Len(BackupPath) > 0 Then Message = Message & " BackupPath=""" & BackupPath & """ "
    Case 1004 '??? ' ������: ����������� ������ � ������� Visual Basic �� �������� ����������
                Message = "����������� ������ � ������� Visual Basic �� �������� ����������. ��� ����������� ������������ ����������/�������������� ������� ���������� ���������� ����������: ""������\������\������������\�������� ������ � Visual Basic Project"""
    Case m_errZipError:     Message = "������ �������� �������� �������"
    Case m_errExportError:  Message = "������ ��� �������� �������� �������"
    Case Else:  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    MsgBox Message & vbCrLf & "�������� �������� ����� �� ���������.", vbExclamation
    Err.Clear: Resume HandleExit
    Err.Clear: Resume 0
End Function
Public Function SourcesRestore(Optional SourcePath As String, _
    Optional WithoutData As Boolean = True) As Boolean
' ��������������� ��� ������� ������� �� �������� �����
Const c_strProcedure = "SourcesRestore"
Dim iTry As Integer
Dim Result As Boolean
On Error GoTo HandleError

'Dim WithoutData As Boolean:     WithoutData = True      ' ��� �������������� ���������� ������� ������ (������� � �.�.)
Dim ReadType As ObjectRwType:   ReadType = orwAlways    ' ��������������� ��� ������� ���������� ����������� � ������ ���������� �� ������
Dim AskBefore As Boolean:       AskBefore = True        ' ���������� ������������ ����� ��������������� ������������� �������
Dim UseTypeFolders As Boolean:  UseTypeFolders = True   ' ����� � ������ ������������� �� ������ ��������

#If APPTYPE = 0 Then        ' APPTYPE=Access
    DoCmd.SetWarnings False ' ��������� ��������������
    Call SysCmd(504, 16484) ' ��������� ��� ������
#Else
#End If                     ' APPTYPE
Dim strCaption As String, strMessage As String
Dim TempPath As String, FilePath As String

    If Len(SourcePath) = 0 Then SourcePath = oFso.BuildPath(CurrentProject.path, c_strSrcPath)
' ���������/������� ����
    Result = oFso.FileExists(SourcePath): If Result Then GoTo HandleCreateTempPath
    ' ����������� ��� ����� backup
    strCaption = "�������� �������� ����� ������� ��� ��������������"
    strMessage = "�������� ����� ������� " & VBA.Chr$(0) & "*.zip"
    SourcePath = p_SelectFile(SourcePath, strMessage, c_strObjExtZip, strCaption)
    Result = Len(SourcePath) > 0: If Not Result Then Err.Raise 76
HandleCreateTempPath:
    ' ������� ��������� �����
    TempPath = oFso.BuildPath(VBA.Environ$("Temp"), "~" & oFso.GetFileName(SourcePath))
    If Not oFso.FolderExists(TempPath) Then Call oFso.CreateFolder(TempPath) 'Then Err.Raise 76 ' Path not Found
' �������������� ������������

Dim prg As clsProgress: Set prg = New clsProgress
    strCaption = "�������������� �������� �������"
    strMessage = strCaption & " ��: """ & SourcePath & """"
    prg.Init pCount:=1, pMin:=1, pMax:=1, pCaption:=strCaption, pText:=strMessage ': prg.ProgressStep = 1
' ������������� �������� �� ��������� ����� ������
    prg.Detail = "���������� ������ ������ �� ��������� �����"
#If USEZIPCLASS Then
    Result = oZip.OpenArchive(SourcePath): If Not Result Then Err.Raise m_errUnZipError
    Result = oZip.Extract(TempPath): If Not Result Then Err.Raise m_errUnZipError
#Else
    Result = ZipUnPack(SourcePath, TempPath): If Not Result Then Err.Raise m_errUnZipError
#End If ' UseZipArchive

' ��������� ��������� ���� �������� ���������� ��������������
Dim colObjects As Collection:     Set colObjects = New Collection
    If UseTypeFolders Then  ' ������ ���������� ��������
Dim oItem As Object
        For Each oItem In oFso.GetFolder(TempPath).SubFolders
        ' ��������� ������������ ����� ����� ���������� ����� ��������
            Select Case oFso.GetFileName(oItem)
            Case c_strObjTypModule
#If APPTYPE = 0 Then        ' APPTYPE=Access
            Case c_strObjTypAccFrm, c_strObjTypAccRep
            Case c_strObjTypAccQry, c_strObjTypAccMac
            Case c_strObjTypAccTbl: If WithoutData Then GoTo HandleNextFolder
'#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
#Else                       '
#End If                     ' APPTYPE
            Case Else: GoTo HandleNextFolder
            End Select
'
            If Not p_ObjectFilesCollectionCreate(oItem.path, colObjects, WithoutData) Then Err.Raise m_errImportError
HandleNextFolder:
        Next oItem
        Set oItem = Nothing
    Else                     ' ������ ���������� ������ ������� �����
            If Not p_ObjectFilesCollectionCreate(TempPath, colObjects, WithoutData) Then Err.Raise m_errImportError
    End If

' �������������� ����� ������������ ���������
'    strCaption = "�������������� �������� �������"
'    strMessage = strCaption & " ��: """ & SourcePath & """"
    prg.ProgressMax = colObjects.Count + 1
' ��������������� �������� � ������������ ������ �������
    prg.Detail = "�������������� ������� � ������������ ������ �������"
    FilePath = oFso.BuildPath(TempPath, c_strAppNamPrj & "." & c_strObjExtPrj)
    p_PropertiesRead FilePath: p_ReferencesRead FilePath
    prg.Update
    'prg.Detail = "��������� �������������� ������� � ������������ ������ �������": prg.Detail = strMessage
' ���������� ��������� ��� �������
    Result = p_ObjectsRestore(colObjects, prg, ReadType, AskBefore, UseTypeFolders, strMessage) = 0:    If Not Result Then Err.Raise m_errExportError
    'If colObjects.Count > 0 Then ' �� ��� �������� ������� ������������
    Set colObjects = Nothing
' ������� ��������� ����� ' ������ ����� ������
    Result = oFso.DeleteFolder(TempPath) = 0
#If APPTYPE = 0 Then        ' APPTYPE=Access
    DoCmd.SetWarnings True  ' �������� ��������������
    Call SysCmd(504, 16484) ' ��������� ��� ������
#Else
#End If                     ' APPTYPE
' ����������� ������������ � �����������
    strCaption = "������ ��������"
    strMessage = strMessage & vbCrLf & "������� �������� ������� ���� ������������� ��:" & vbCrLf & SourcePath & "." & vbCrLf
    Call MsgBox(strMessage, vbInformation + vbOKOnly, strCaption)

HandleExit:     SourcesRestore = Result: Set prg = Nothing: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err.Number
    Case 58:    Err.Clear: Resume Next      ' �������� ����� ��� �������
    Case 70:    If iTry < 3 Then iTry = iTry + 1: Err.Clear: Sleep 333: Resume Next
                Message = "�� ������� ��������� ������� ��������� �����."
                If Len(TempPath) > 0 Then Message = Message & " TempPath=""" & TempPath & """ "
    Case 75:    Message = "������ �������!" ' �������� ����� ��� �������
    Case 76:    Message = "���� �� ������!" ' �������� ��� ������������ �����
    If Len(SourcePath) > 0 Then Message = Message & " SourcePath=""" & SourcePath & """ "
    If Len(TempPath) > 0 Then Message = Message & " TempPath=""" & TempPath & """ "
    Case 1004 '??? ' ������: ����������� ������ � ������� Visual Basic �� �������� ����������
                Message = "����������� ������ � ������� Visual Basic �� �������� ����������. ��� ����������� ������������ ����������/�������������� ������� ���������� ���������� ����������: ""������\������\������������\�������� ������ � Visual Basic Project"""
    Case m_errUnZipError:   Message = "������ ���������� �������� �������"
    Case m_errImportError:  Message = "������ ��� ������� �������� �������"
    Case Else:  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    MsgBox Message & vbCrLf & "�������������� �� �������� ����� �� ���������.", vbExclamation
    Err.Clear: Resume HandleExit
End Function
Public Function SourcesUpdateFromStorage(Optional SourcePath As String)
' ���������� �������� ���������� ������� �������������� � ��������� �������� ����� ������� ������
Const c_strProcedure = "SourcesUpdateFromStorage"

Dim Result As Boolean
On Error GoTo HandleError

'Dim WithoutData As Boolean:     WithoutData = True      ' ��� �������������� ���������� ������� ������ (������� � �.�.)
Dim ReadType As ObjectRwType:   ReadType = orwSrcNewer  ' ��������� ������ ���� ������ � ���������� �����
Dim AskBefore As Boolean:       AskBefore = True        ' ���������� ������������ ����� ����������� �������
Dim OnlyExisting As Boolean:    OnlyExisting = True     ' ��������� � ��������� �� ���������� ������ ������������ � ���������� �������
'Dim UseTypeFolders As Boolean:  UseTypeFolders = False  ' ����� � ���������� �� ������������� �� ������ �����

#If APPTYPE = 0 Then        ' APPTYPE=Access
    DoCmd.SetWarnings False ' ��������� ��������������
    Call SysCmd(504, 16484) ' ��������� ��� ������
#Else
#End If                     ' APPTYPE

Dim strCaption As String, strMessage As String
    
    If Len(SourcePath) = 0 Then SourcePath = c_strLibPath
' ���������/������� ����
    Result = oFso.FolderExists(SourcePath): If Result Then GoTo HandleUpdateLibPath
    ' ����������� ���� � ��������� �������
    strCaption = "������� ���� � ��������� ��������� ������ ������������ �������."
    SourcePath = p_SelectFolder(SourcePath, DialogTitle:=strCaption)
    Result = Len(SourcePath) > 0: If Not Result Then Err.Raise 75
HandleUpdateLibPath:
    If SourcePath = c_strLibPath Then GoTo HandleUpdateSources
' ��������� � ������ ������ �������� ��������� c_strLibPath �� SourcePath
Dim objModule As Object: If Not ModuleExists(c_strModule, objModule) Then Err.Raise m_errModuleDontFind
Dim CodeLine As Long: CodeLine = p_CodeLineFind(objModule, c_strLibPathLine, CodeLine) ' ������ �� ������ ������� �������
    With objModule
        If CodeLine > 0 Then .DeleteLines CodeLine, 1 Else CodeLine = .CountOfDeclarationLines
        .InsertLines CodeLine, c_strLibPathLine & " """ & SourcePath & """"
    End With

HandleUpdateSources:
' �������������� ������������
Dim prg As clsProgress: Set prg = New clsProgress
    strCaption = "���������� �������� �������"
    strMessage = strCaption & " �� ����������: """ & SourcePath & """"
    prg.Init pCount:=1, pMin:=1, pMax:=1, pCaption:=strCaption, pText:=strMessage ': prg.ProgressStep = 1
' ��������� ��������� ���� �������� ���������� ��������������
Dim colObjects As Collection:     Set colObjects = New Collection
' ������ ���������� ������ ������� �����
    If Not p_ObjectFilesCollectionCreate(SourcePath, colObjects, True, OnlyExisting) Then Err.Raise m_errImportError
' �������������� ����� ������������ ���������
    prg.ProgressMax = colObjects.Count ': prg.ProgressStep = 1
' ���������� ��������� ��� �������
    Result = p_ObjectsRestore(colObjects, prg, ReadType, AskBefore, False, strMessage) = 0:     If Not Result Then Err.Raise m_errExportError
    'If colObjects.Count > 0 Then ' �� ��� �������� ������� ������������
    Set colObjects = Nothing
' ����������� ������������ � �����������
    strCaption = "���������� ���������"
    strMessage = strMessage & vbCrLf & vbCrLf & "������� �������� ������� ���� ��������� ��:" & vbCrLf & SourcePath & "." & vbCrLf
    Call MsgBox(strMessage, vbInformation + vbOKOnly, strCaption)
    Set prg = Nothing
HandleExit:     SourcesUpdateFromStorage = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err.Number
    Case 75:    Message = "������ �������!" ' �������� ����� ��� �������
    If Len(SourcePath) > 0 Then Message = Message & " SourcePath=""" & SourcePath & """ "
    Case 76:    Message = "���� �� ������!" ' �������� ��� ������������ �����
    If Len(SourcePath) > 0 Then Message = Message & " SourcePath=""" & SourcePath & """ "
    Case 1004 '??? ' ������: ����������� ������ � ������� Visual Basic �� �������� ����������
                Message = "����������� ������ � ������� Visual Basic �� �������� ����������. ��� ����������� ������������ ����������/�������������� ������� ���������� ���������� ����������: ""������\������\������������\�������� ������ � Visual Basic Project"""
    Case m_errModuleNameWrong:  Message = "������� ������ ��� �������!"
    Case m_errObjectTypeUnknown: Message = "����������� ��� ������!"
    Case Else:                  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function SourcesUpdateStorage(Optional BackupPath As String)
' ���������� �������� ���������� ������� ������� ������� ����� ������� ������ � ��������� ������ ��������
Const c_strProcedure = "SourcesUpdateStorage"

Dim Result As Boolean
On Error GoTo HandleError

'Dim WithoutData As Boolean:     WithoutData = True      ' ��� �������������� ���������� ������� ������ (������� � �.�.)
Dim WriteType As ObjectRwType:  WriteType = orwSrcNewerOrDestMissing  ' ��������� ���� ������ � ������� ����� ��� ����������� � ����������
Dim AskBefore As Boolean:       AskBefore = True        ' ���������� ������������ ����� ����������� �������
Dim OnlyExisting As Boolean:    OnlyExisting = True     ' ��������� � ��������� �� ���������� ������ ������������ � ���������� �������
'Dim UseTypeFolders As Boolean:  UseTypeFolders = False  ' ����� � ���������� �� ������������� �� ������ �����

#If APPTYPE = 0 Then        ' APPTYPE=Access
    DoCmd.SetWarnings False ' ��������� ��������������
    Call SysCmd(504, 16484) ' ��������� ��� ������
#Else
#End If                     ' APPTYPE

Dim strCaption As String, strMessage As String
    
    If Len(BackupPath) = 0 Then BackupPath = c_strLibPath
' ���������/������� ����
    Result = oFso.FolderExists(BackupPath): If Result Then GoTo HandleUpdateLibPath
    ' ����������� ���� � ��������� �������
    strCaption = "������� ���� � ��������� ��������� ������ ������������ �������."
    BackupPath = p_SelectFolder(BackupPath, DialogTitle:=strCaption)
    Result = Len(BackupPath) > 0: If Not Result Then Err.Raise 75
HandleUpdateLibPath:
    If BackupPath = c_strLibPath Then GoTo HandleUpdateSources
' ��������� � ������ ������ �������� ��������� c_strLibPath �� SourcePath
Dim objModule As Object: If Not ModuleExists(c_strModule, objModule) Then Err.Raise m_errModuleDontFind
Dim CodeLine As Long: CodeLine = p_CodeLineFind(objModule, c_strLibPathLine, CodeLine) ' ������ �� ������ ������� �������
    With objModule
        If CodeLine > 0 Then .DeleteLines CodeLine, 1 Else CodeLine = .CountOfDeclarationLines
        .InsertLines CodeLine, c_strLibPathLine & " """ & BackupPath & """"
    End With

HandleUpdateSources:
' �������������� ������������
Dim prg As clsProgress: Set prg = New clsProgress
    strCaption = "���������� �������� �������"
    strMessage = strCaption & " � ����������: """ & BackupPath & """"
    prg.Init pCount:=1, pMin:=1, pMax:=1, pCaption:=strCaption, pText:=strMessage ': prg.ProgressStep = 1
' ��������� ��������� ���� �������� ���������� ��������������
Dim colObjects As Collection:     Set colObjects = New Collection
' ������ ���������� ������ ������� �����
    If Not p_ObjectsCollectionCreate(colObjects, True) Then Err.Raise m_errImportError
' �������������� ����� ������������ ���������
    prg.ProgressMax = colObjects.Count
' ��������� ��������� ��� �������
    Result = p_ObjectsBackup(colObjects, BackupPath, prg, WriteType, AskBefore, False, strMessage) = 0:      If Not Result Then Err.Raise m_errExportError
    'If colObjects.Count > 0 Then ' �� ��� �������� ������� ���������
    Set colObjects = Nothing
' ����������� ������������ � �����������
    strCaption = "���������� ���������"
    strMessage = strMessage & vbCrLf & vbCrLf & "������� �������� ������� ���� ��������� �:" & vbCrLf & BackupPath & "." & vbCrLf
    Call MsgBox(strMessage, vbInformation + vbOKOnly, strCaption)
    
HandleExit:     SourcesUpdateStorage = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err.Number
    Case 75:    Message = "������ �������!" ' �������� ����� ��� �������
    If Len(BackupPath) > 0 Then Message = Message & " BackupPath=""" & BackupPath & """ "
    Case 76:    Message = "���� �� ������!" ' �������� ��� ������������ �����
    If Len(BackupPath) > 0 Then Message = Message & " BackupPath=""" & BackupPath & """ "
    Case 1004 '??? ' ������: ����������� ������ � ������� Visual Basic �� �������� ����������
                Message = "����������� ������ � ������� Visual Basic �� �������� ����������. ��� ����������� ������������ ����������/�������������� ������� ���������� ���������� ����������: ""������\������\������������\�������� ������ � Visual Basic Project"""
    Case m_errModuleNameWrong:  Message = "������� ������ ��� �������!"
    Case m_errObjectTypeUnknown: Message = "����������� ��� ������!"
    Case Else:                  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_ObjectsCollectionCreate(colObjects As Collection, _
    Optional WithoutData As Boolean = True) As Boolean
' ��������� ��������� �������� ����������
Dim Result As Boolean: Result = False
' WithoutData - ���������� ����� �� �������� � ��������� ������� ������ (������� � ��.)
Dim BackupDocs As Boolean
#If APPTYPE = 0 Then        ' APPTYPE=Access
    BackupDocs = False      ' ������ ���� ������� ���� ����������� ������ � ������/�������
#Else
    BackupDocs = True       ' � Excel � ��. ��� ������ ���� �� �����
#End If                     ' APPTYPE
    On Error GoTo HandleError
Dim oItems As Object, oItem
' ������� ���������� ��� ������
    ' VBE.ActiveVBProject.VBComponents ���������� �� ���� ����������� Microsoft
    ' �� � ��� ����� �� ���� ������� - ����� ��������� ������ 1004
    ' ���� ����� ������� ���������: "�������� ������ � Visual Basic Project"
    Set oItems = Application.VBE.ActiveVBProject.VBComponents
    For Each oItem In oItems
        Select Case oItem.Type
        Case vbext_ct_StdModule, vbext_ct_ClassModule: colObjects.Add oItem, oItem.NAME
        Case vbext_ct_Document: If BackupDocs Then colObjects.Add oItem, oItem.NAME
        Case Else ' ???
        End Select
    Next oItem
' ����� ������������� ��� ���������� �������
#If APPTYPE = 0 Then        ' APPTYPE=Access
    ' � Access ��� ����� ����� ������� �������� �� ������� c_strMSysObjects
    '' ��� ����� �������� �� ���������� ��������
    'Set oItems = CurrentProject.AllModules:         For Each oItem In oItems: colObjects.Add oItem, oItem.NAME: Next oItem
    Set oItems = CurrentProject.AllForms:           For Each oItem In oItems: colObjects.Add oItem, oItem.NAME: Next oItem '
    Set oItems = CurrentProject.AllReports:         For Each oItem In oItems: colObjects.Add oItem, oItem.NAME: Next oItem
    Set oItems = CurrentProject.AllMacros:          For Each oItem In oItems: colObjects.Add oItem, oItem.NAME: Next oItem
    Set oItems = CurrentProject.AllDataAccessPages: For Each oItem In oItems: colObjects.Add oItem, oItem.NAME: Next oItem
    Set oItems = CurrentData.AllQueries:            For Each oItem In oItems: colObjects.Add oItem, oItem.NAME: Next oItem
    If Not WithoutData Then
    Set oItems = CurrentData.AllTables:             For Each oItem In oItems
                                                    ' ���������� ��������� �������
                                                        If Left(oItem.NAME, 4) <> "MSys" Then colObjects.Add oItem, oItem.NAME
                                                    Next oItem
    End If
#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
    ' � Excel �������� ���������� ��������� �������� � ThisWorkbook
#Else                       ' APPTYPE=��������� ����� ���
#End If                     ' APPTYPE
    Set oItems = Nothing: Set oItem = Nothing
    Result = True
HandleExit:  p_ObjectsCollectionCreate = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_ObjectFilesCollectionCreate(FilesPath As String, colObjects As Collection, _
    Optional WithoutData As Boolean = True, Optional OnlyExisting As Boolean = False) As Boolean
' ��������� ��������� ������ �������� ���������� ������������� �� ���������� ����
Dim Result As Boolean ': Result = False
' WithoutData - ���������� ����� �� �������� � ��������� ����� ������� ������ (������� � ��.)
' OnlyExisting - ���������� ����� �� �������� � ��������� ����� �������� ������������� � ����������
    On Error GoTo HandleError
Dim oItem As Object
    ' ���������� ��� ����� � ��������� ����� �������� � �������� �����
'' SubFolders ������ �� ����� ��������� �������
    For Each oItem In oFso.GetFolder(FilesPath).Files
        Select Case oFso.GetExtensionName(oItem.path)
        Case c_strObjExtBas, c_strObjExtCls
        'Case c_strObjExtDoc ' ������ ������ ���������
#If APPTYPE = 0 Then        ' APPTYPE=Access
        Case c_strObjExtFrm, c_strObjExtRep ': GoTo HandleNextFile
        'Case c_strObjExtDoc ' ������ ������ ��������� Access (Form ��� Report)
        Case c_strObjExtQry, c_strObjExtMac ': GoTo HandleNextFile
        ' ���������� ����� � �������
        Case c_strObjExtXml:    If WithoutData Then GoTo HandleNextFile  ' ��������� ������� � XML
        Case c_strObjExtTxt:    If WithoutData Then GoTo HandleNextFile  ' ��������� ������� � TXT
        Case c_strObjExtCsv:    If WithoutData Then GoTo HandleNextFile  ' ��������� ������� � CSV (����� � �������������)
        Case c_strObjExtLnk:    If WithoutData Then GoTo HandleNextFile  ' ��������� �������
'#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
#Else
#End If                     ' APPTYPE
        Case Else:              GoTo HandleNextFile
        End Select
' ��� ������������� ��������� ���������� �� ��������������� ������ � ����������
Dim ObjectName As String:    ObjectName = p_TextCode2Alpha(oFso.GetBaseName(oItem)) ' ��� ������� �� ����� �����
        If OnlyExisting Then If Not ObjectExists(ObjectName) Then GoTo HandleNextFile
' ��������� ���� � ��������� ������ ��� ��������������
        colObjects.Add oItem, ObjectName
HandleNextFile:
    Next oItem
    Set oItem = Nothing
    Result = True
HandleExit:  p_ObjectFilesCollectionCreate = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_ObjectsBackup(colObjects As Collection, BackupPath As String, prg As clsProgress, _
    Optional WriteType As ObjectRwType = orwAlways, Optional AskBefore As Boolean = True, Optional UseTypeFolders As Boolean = True, _
    Optional Message As String) As Long
' ���������� �������� ��������� ��������� � ����� �� ���������� ����
Const c_strProcedure = "p_ObjectsBackup"
Dim Result As Long
On Error GoTo HandleError
Dim oItem
Dim ParentPath As String, strFilePath As String, strFileExtn As String
Dim strObjName As String, strTypeName As String, strTypeDesc As String
Dim strSkip As String ', strDone As String
Dim strSkipByUser As String, strSkipByVers As String ', strSkipByList As String
' ������������� ������������
Dim strCaption As String, strMessage As String
'Dim i As Long, iMax As Long: i = 1: iMax = colObjects.Count
'Dim prg As clsProgress: Set prg = New clsProgress
'    strCaption = "���������� �������� �������"
'    prg.Init pCount:=1, pMin:=i, pMax:=iMax, pCaption:=strCaption, pText:=strCaption & " �: """ & BackupPath & """": prg.ProgressStep = 1

Dim iCount As Long
' ��������� ��������� ��� �������
    For Each oItem In colObjects
        strObjName = oItem.NAME
    ' �������� ���������� �� �������
        p_ObjectInfo strObjName, ObjectTypeName:=strTypeName, ObjectTypeDesc:=strTypeDesc
    ' ��������� �����������
        strMessage = "��� ���������� �������: " & strTypeDesc & " """ & strObjName & """":
        prg.Update: prg.Detail = strMessage
    ' ��������� ���� ���������� �������
        strFilePath = BackupPath: If UseTypeFolders Then strFilePath = oFso.BuildPath(strFilePath, strTypeName)
    ' ��������� ������ �� ���������� ����
        If Not oFso.FolderExists(strFilePath) Then Call oFso.CreateFolder(strFilePath) 'Then Err.Raise 76 ' Path not Found
    ' ��������� ��������� ���������� � ��������� ����� ��������� ������������
        Select Case p_ObjectWrite(strObjName, strFilePath, WriteType:=WriteType, AskBefore:=AskBefore)  ', Message:=strMessage)
        Case 0:     'strDone = strDone & c_strInDelim & """" & strObjName & """"  ' ������ ������� ���������
                    colObjects.Remove (strObjName): iCount = iCount + 1 ' ������� �� ��������� ������������ ������
'        ' ������� ��������:
        Case m_errSkippedByUser ' ��������� �� ������� ������������
            If Len(strObjName) > 0 Then strSkipByUser = strSkipByUser & c_strInDelim & """" & strObjName & """"   ' ������ ��������
        Case m_errWrongVersion  ' �������������� ������ ������� ����������
            If Len(strObjName) > 0 Then strSkipByVers = strSkipByVers & c_strInDelim & """" & strObjName & """"   ' ������ ��������
        Case m_errSkippedByList ' ������� � ������ ��������
            'If Len(strObjName) > 0 Then strSkipByList = strSkipByList & c_strInDelim & """" & strObjName & """"   ' ������ ��������
'        Case m_errDestMissing   ' ��������e�� ����������� ������
'        Case m_errDestExists    ' ����������� ������ ��� ����������
        Case Else:  If Len(strObjName) > 0 Then strSkip = strSkip & c_strInDelim & """" & strObjName & """"   ' ������ ��������
        End Select
    ' ��������� ������� ���������� ��������
        If prg.Canceled Then If MsgBox(c_strBreakProcessMessage, vbYesNo Or vbExclamation Or vbDefaultButton2) = vbYes Then GoTo HandleExit
'    ' ������� � ����������� ���������� � ���������� ��������
'        prg.Detail = strMessage
    Next oItem
' �������� ���������
    Message = "���������� �������� ������� ���������."
'    If Left(strDone, Len(c_strInDelim)) = c_strInDelim Then strDone = Mid(strDone, Len(c_strInDelim) + 1)
'    If Len(strDone) > 0 Then Message = Message & vbCrLf & "��������� ��������� �������: " & strDone
    Message = Message & vbCrLf & "����� ��������� " & iCount & " ��������." & vbCrLf '�: """ & BackupPath & """" & vbCrLf
    If colObjects.Count = 0 Then GoTo HandleExit
    Message = Message & vbCrLf & "��������� ��������� �������: "
    If Left(strSkipByUser, Len(c_strInDelim)) = c_strInDelim Then strSkipByUser = Mid(strSkipByUser, Len(c_strInDelim) + 1)
    If Len(strSkipByUser) > 0 Then Message = Message & vbCrLf & ChrW(&H2022) & " �� ������� ������������: " & vbCrLf & strSkipByUser
    If Left(strSkipByVers, Len(c_strInDelim)) = c_strInDelim Then strSkipByVers = Mid(strSkipByVers, Len(c_strInDelim) + 1)
    If Len(strSkipByVers) > 0 Then Message = Message & vbCrLf & ChrW(&H2022) & " � ����� � ��������������� ������ ������� ����������: " & vbCrLf & strSkipByVers
'    If Left(strSkipByList, Len(c_strInDelim)) = c_strInDelim Then strSkipByList = Mid(strSkipByList, Len(c_strInDelim) + 1)
'    If Len(strSkipByList) > 0 Then Message = Message & vbCrLf & ChrW(&h2022) & " ����� ���� ���������, �� ���� ��������� �.�. ������������ � �������� ����������: "  & vbCrLf & strSkipByList & vbcrLf & "��� ������������� �� ����� ��������� �������."
    If Left(strSkip, Len(c_strInDelim)) = c_strInDelim Then strSkip = Mid(strSkip, Len(c_strInDelim) + 1)
    If Len(strSkip) > 0 Then Message = Message & vbCrLf & ChrW(&H2022) & " �� ���� ��������: " & vbCrLf & strSkip
    Message = Message & vbCrLf & "����� ��������� " & colObjects.Count & " ��������."
HandleExit:     Set oItem = Nothing
                p_ObjectsBackup = Result: Exit Function
HandleError:
'    Dim Message As String
    Err.Clear: Resume 0
    Select Case Err.Number
    Case 75:    Message = "������ �������!" ' �������� ����� ��� �������
    If Len(strFilePath) > 0 Then Message = Message & " strFilePath=""" & strFilePath & """ "
    Case 76:    Message = "���� �� ������!" ' �������� ��� strFilePath �����
    If Len(strFilePath) > 0 Then Message = Message & " BackupPath=""" & strFilePath & """ "
    Case 1004 '??? ' ������: ����������� ������ � ������� Visual Basic �� �������� ����������
                Message = "����������� ������ � ������� Visual Basic �� �������� ����������. ��� ����������� ������������ ����������/�������������� ������� ���������� ���������� ����������: ""������\������\������������\�������� ������ � Visual Basic Project"""
    Case Else:  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Result = Err: Err.Clear: Resume HandleExit
End Function
Private Function p_ObjectsRestore(colObjects As Collection, prg As clsProgress, _
    Optional ReadType As ObjectRwType = orwAlways, Optional AskBefore As Boolean = True, Optional UseTypeFolders As Boolean = True, _
    Optional Message As String) As Long
' �������������� �������� �� ������ ��������� ��������� � ������
Const c_strProcedure = "p_ObjectsRestore"
Dim Result As Long
On Error GoTo HandleError
Dim oItem
Dim strCaption As String, strMessage As String
Dim ParentPath As String, strFilePath As String, strFileExtn As String
Dim strObjName As String, strTypeName As String, strTypeDesc As String
Dim strSkip As String ', strDone As String
Dim strSkipByUser As String, strSkipByVers As String, strSkipByList As String
'' ������������� ������������
'Dim i As Long, iMax As Long: i = 1: iMax = colObjects.Count
'Dim prg As clsProgress: Set prg = New clsProgress
'    strCaption = "�������������� �������� �������"
'    prg.Init pCount:=1, pMin:=i, pMax:=iMax, pCaption:=strCaption, pText:=strCaption & " �: """ & SourcePath & """": prg.ProgressStep = 1

Dim iCount As Long
' ��������� ��������� ��� �������
    For Each oItem In colObjects
        strFilePath = oItem.path: strObjName = strFilePath
    ' �������� ���������� �� �������
        p_ObjectInfo strObjName, ObjectTypeName:=strTypeName, ObjectTypeDesc:=strTypeDesc
    ' ��������� �����������
        strMessage = "��� �������������� �������: " & strTypeDesc & " """ & strObjName & """": prg.Update: prg.Detail = strMessage
    ' ��������� ��������� �������������� � ��������� ����� ��������� ������������
        Select Case p_ObjectRead(strFilePath, ReadType:=ReadType, AskBefore:=AskBefore)   ', Message:=strMessage)
        Case 0:     'strDone = strDone & c_strInDelim & """" & strObjName & """"  ' ������ ������� ���������
                    colObjects.Remove (strObjName): iCount = iCount + 1  ' ������� �� ��������� ������������ ������
'        ' ������� ��������:
        Case m_errSkippedByUser ' ��������� �� ������� ������������
            If Len(strObjName) > 0 Then strSkipByUser = strSkipByUser & c_strInDelim & """" & strObjName & """"
        Case m_errWrongVersion  ' �������������� ������ ������� ����������
            If Len(strObjName) > 0 Then strSkipByVers = strSkipByVers & c_strInDelim & """" & strObjName & """"
        Case m_errSkippedByList ' ������� � ������ ��������
            If Len(strObjName) > 0 Then strSkipByList = strSkipByList & c_strInDelim & """" & strObjName & """"
'        Case m_errDestMissing   ' ����������� ����������� ������
'        Case m_errDestExists    ' ����������� ������ ��� ����������
        Case Else:  If Len(strObjName) > 0 Then strSkip = strSkip & c_strInDelim & """" & strObjName & """"   ' ������ ��������
        End Select
    ' ��������� ������� ���������� ��������
        If prg.Canceled Then If MsgBox(c_strBreakProcessMessage, vbYesNo Or vbExclamation Or vbDefaultButton2) = vbYes Then GoTo HandleExit
    ' ������� � ����������� ���������� � ���������� ��������
        prg.Detail = strMessage
    Next oItem
' �������� ���������
    Message = "���������� �������� ������� ���������."
    Message = Message & vbCrLf & "����� ��������� " & iCount & " ��������." & vbCrLf ' ��: """ & RestorePath & """" & vbCrLf
'    If Left(strDone, Len(c_strInDelim)) = c_strInDelim Then strDone = Mid(strDone, Len(c_strInDelim) + 1)
'    If Len(strDone) > 0 Then Message = Message & vbCrLf & "��������� ��������� �������: " & strDone
    If colObjects.Count = 0 Then GoTo HandleExit
    Message = Message & vbCrLf & "��������� ��������� �������: "
    If Left(strSkipByUser, Len(c_strInDelim)) = c_strInDelim Then strSkipByUser = Mid(strSkipByUser, Len(c_strInDelim) + 1)
    If Len(strSkipByUser) > 0 Then Message = Message & vbCrLf & ChrW(&H2022) & " �� ������� ������������: " & vbCrLf & strSkipByUser
    If Left(strSkipByVers, Len(c_strInDelim)) = c_strInDelim Then strSkipByVers = Mid(strSkipByVers, Len(c_strInDelim) + 1)
    If Len(strSkipByVers) > 0 Then Message = Message & vbCrLf & ChrW(&H2022) & " � ����� � ��������������� ������ ������� ����������: " & vbCrLf & strSkipByVers
    If Left(strSkipByList, Len(c_strInDelim)) = c_strInDelim Then strSkipByList = Mid(strSkipByList, Len(c_strInDelim) + 1)
    If Len(strSkipByList) > 0 Then Message = Message & vbCrLf & ChrW(&H2022) & " ����� ���� ���������, �� ���� ��������� �.�. ������������ � �������� ����������: " & vbCrLf & strSkipByList & vbCrLf & "��� ������������� �� ����� �������� �������."
    If Left(strSkip, Len(c_strInDelim)) = c_strInDelim Then strSkip = Mid(strSkip, Len(c_strInDelim) + 1)
    If Len(strSkip) > 0 Then Message = Message & vbCrLf & ChrW(&H2022) & " �� ���� ��������: " & vbCrLf & strSkip
    Message = Message & vbCrLf & "����� ��������� " & colObjects.Count & " ��������."
HandleExit:     Set oItem = Nothing
                p_ObjectsRestore = Result: Exit Function
HandleError:
'    Dim Message As String
    Err.Clear: Resume 0
    Select Case Err.Number
    Case 75:    Message = "������ �������!" ' �������� ����� ��� �������
    If Len(strFilePath) > 0 Then Message = Message & " strFilePath=""" & strFilePath & """ "
    Case 76:    Message = "���� �� ������!" ' �������� ��� ������������ �����
    If Len(strFilePath) > 0 Then Message = Message & " strFilePath=""" & strFilePath & """ "
    Case 1004 '??? ' ������: ����������� ������ � ������� Visual Basic �� �������� ����������
                Message = "����������� ������ � ������� Visual Basic �� �������� ����������. ��� ����������� ������������ ����������/�������������� ������� ���������� ���������� ����������: ""������\������\������������\�������� ������ � Visual Basic Project"""
    Case Else:  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Result = Err: Err.Clear: Resume HandleExit
End Function

Public Function ReferencesRestore(Optional FilePath As String)
' �������������� ������ �� ����������
Const c_strProcedure = "ReferencesRestore"
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
'    p_ReferencesDropBroken
'    p_ReferencesRestore

    Call SysCmd(504, 16484) ' ��������� ��� ������
    ' ��������� ������� ���������� ����/�����
    Result = oFso.FileExists(FilePath): If Result Then GoTo HandleUpdate
    FilePath = oFso.BuildPath(CurrentProject.path, c_strSrcPath) '& "\" & strText
    ' ����������� ��� ����� backup
Dim strTitle As String: strTitle = "�������� ���� ������� �������"
Dim strText As String: strText = "�������� ����� ������� ������� " & VBA.Chr$(0) & strText & "*." & c_strObjExtZip & ";" & "*." & c_strObjExtPrj
    FilePath = p_SelectFile(FilePath, strText, c_strObjExtZip & ";" & c_strObjExtPrj, strTitle)
    Result = Len(FilePath) > 0: If Result Then GoTo HandleUpdate
    MsgBox Prompt:="�� ������� ��� ����� �������� �����." & vbCrLf & _
        "������ �� ���� �������������.", Buttons:=vbOKOnly + vbInformation
    GoTo HandleExit
HandleUpdate:
    ' ��������� ���������� �����
    Select Case oFso.GetExtensionName(FilePath)
    Case c_strObjExtPrj ' ���� ������� �������
    ' ������ ������ �� ����� ������� �������
        Result = p_ReferencesRead(FilePath)
    Case c_strObjExtZip ' ���� ������ �������
    ' ��������� ������ ���� ������� �������
Dim TempPath As String, FileName As String
    ' ������� ��������� ����� ��� ����������
        FileName = "~" & oFso.GetFileName(FilePath) 'oFso.GetBaseName(FilePath) & "_" & VBA.Format$(Now(), "YYYYMMDD_hhmmss")
        TempPath = oFso.BuildPath(VBA.Environ$("Temp"), FileName)
        If Not oFso.FolderExists(TempPath) Then Call oFso.CreateFolder(TempPath) 'Then Err.Raise 76 ' Path not Found
    ' ��������� ������� ���� � �������� ����
        FileName = c_strAppNamPrj & "." & c_strObjExtPrj
        oApp.Namespace((TempPath)).CopyHere ((oFso.BuildPath(FilePath, FileName))) ': DoEvents: DoEvents
    ' ������ ������ �� ����� ������� �������
        Result = p_ReferencesRead(oFso.BuildPath(TempPath, FileName))
    ' ������� ��������� ����� ����� ����������
        oFso.DeleteFolder (TempPath)
    Case Else: Err.Raise 75
    End Select
HandleExit:     ReferencesRestore = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err.Number
    Case 75:                    Message = "������ �������! ����: " & FilePath ' Path/File access error
    Case 76:                    Message = "���� �� ������! ����: " & FilePath ' Path/File access error
    Case Else:                  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function ReferencesBackup(Optional FilePath As String)
' ���������� ������ ������� ��������� � ������ ������� �������������� ������ �� ����������
Const c_strProcedure = "ReferencesBackup"
Dim Result As Boolean: Result = False
On Error GoTo HandleError
    With oFso
    If Len(FilePath) = 0 Then
        FilePath = .BuildPath(CurrentProject.path, c_strSrcPath)
        FilePath = .BuildPath(FilePath, Split(CurrentProject.NAME, ".")(0) & "_" & VBA.Format$(Now(), "YYYYMMDD_hhmmss") & "." & c_strObjExtPrj)
    End If
    If Not .FolderExists(.GetParentFolderName(FilePath)) Then Call .CreateFolder(.GetParentFolderName(FilePath)) 'Then Err.Raise 76 ' Path not Found
    End With
' ��������� � ����
    'Result = p_PropertiesWrite(FilePath)
    Result = p_ReferencesWrite(FilePath)
HandleExit:     ReferencesBackup = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err.Number
    Case 75:                    Message = "������ �������! ����: " & FilePath ' Path/File access error
    Case 76:                    Message = "���� �� ������! ����: " & FilePath ' Path/File access error
    Case Else:                  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Sub ReferencesPrint()
' ���������� - ������� � Immediate ��� ������������ ������ �������� �������
    Dim Itm As Object
    Debug.Print "Project """ & VBE.ActiveVBProject.NAME & """ References:"
    For Each Itm In References
        With Itm
            If .BuiltIn Then GoTo HandleNext
            Debug.Print " " & IIf(.IsBroken, "X", " ") & .NAME, .GUID & " " & " " & .FullPath ' & .Major & " " & .Minor
        End With
HandleNext:
    Next Itm
End Sub
Public Function PropertiesBackup(Optional FilePath As String)
' ���������� ������� �������
Const c_strProcedure = "PropertiesBackup"
Dim Result As Boolean
' ������� ��� �����
    With oFso
    If Len(FilePath) = 0 Then
        FilePath = .BuildPath(CurrentProject.path, c_strSrcPath)
        FilePath = .BuildPath(FilePath, Split(CurrentProject.NAME, ".")(0) & "_" & VBA.Format$(Now(), "YYYYMMDD_hhmmss") & "." & c_strObjExtPrj)
    End If
    If Not .FolderExists(.GetParentFolderName(FilePath)) Then Call .CreateFolder(.GetParentFolderName(FilePath)) 'Then Err.Raise 76 ' Path not Found
    End With
' ��������� ��������
    Result = p_PropertiesWrite(FilePath)
    'Result = p_ReferencesWrite(FilePath)
HandleExit:     PropertiesBackup = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err.Number
    Case 75:                    Message = "������ �������! ����: " & FilePath ' Path/File access error
    Case 76:                    Message = "���� �� ������! ����: " & FilePath ' Path/File access error
    Case Else:                  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function PropertiesRestore(Optional FilePath As String)
' �������������� ������� �������
Const c_strProcedure = "PropertiesRestore"
Dim Result As Boolean ': Result = False
    On Error GoTo HandleError
    
    Call SysCmd(504, 16484) ' ��������� ��� ������
    ' ��������� ������� ���������� ����/�����
    Result = oFso.FileExists(FilePath): If Result Then GoTo HandleUpdate
    FilePath = oFso.BuildPath(CurrentProject.path, c_strSrcPath) '& "\" & strText
    ' ����������� ��� ����� backup
Dim strTitle As String: strTitle = "�������� ���� ������� �������"
Dim strText As String: strText = "�������� ����� ������� ������� " & VBA.Chr$(0) & strText & "*." & c_strObjExtZip & ";" & "*." & c_strObjExtPrj
    FilePath = p_SelectFile(FilePath, strText, c_strObjExtZip & ";" & c_strObjExtPrj, strTitle)
    Result = Len(FilePath) > 0: If Result Then GoTo HandleUpdate
    MsgBox Prompt:="�� ������� ��� ����� �������� �����." & vbCrLf & _
        "������ �� ���� �������������.", Buttons:=vbOKOnly + vbInformation
    GoTo HandleExit
HandleUpdate:
    ' ��������� ���������� �����
    Select Case oFso.GetExtensionName(FilePath)
    Case c_strObjExtPrj ' ���� ������� �������
    ' ������ �������� �� ����� ������� �������
        Result = p_PropertiesRead(FilePath)
    Case c_strObjExtZip ' ���� ������ �������
    ' ��������� ������ ���� ������� �������
Dim TempPath As String, FileName As String
    ' ������� ��������� ����� ��� ����������
        FileName = "~" & oFso.GetFileName(FilePath) 'oFso.GetBaseName(FilePath) & "_" & VBA.Format$(Now(), "YYYYMMDD_hhmmss")
        TempPath = oFso.BuildPath(VBA.Environ$("Temp"), FileName)
        If Not oFso.FolderExists(TempPath) Then Call oFso.CreateFolder(TempPath) 'Then Err.Raise 76 ' Path not Found
    ' ��������� ������� ���� � �������� ����
        FileName = c_strAppNamPrj & "." & c_strObjExtPrj
        oApp.Namespace((TempPath)).CopyHere ((oFso.BuildPath(FilePath, FileName))) ': DoEvents: DoEvents
    ' ������ �������� �� ����� ������� �������
        Result = p_PropertiesRead(oFso.BuildPath(TempPath, FileName))
    ' ������� ��������� ����� ����� ����������
        oFso.DeleteFolder (TempPath)
    Case Else: Err.Raise 75
    End Select
    
HandleExit:     PropertiesRestore = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err.Number
    Case 75:                    Message = "������ �������! ����: " & FilePath ' Path/File access error
    Case 76:                    Message = "���� �� ������! ����: " & FilePath ' Path/File access error
    Case Else:                  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Sub PropertiesClear()
' ������� ��� ���������������� ��������
    With CurrentProject.Properties: Do While .Count > 0: .Remove .Item(0).NAME: Loop: End With
End Sub
Public Sub PropertiesPrint()
' ���������� - ������� � Immediate ��� �������� �������� �������
    Dim Itm As Object
    Debug.Print "Project """ & VBE.ActiveVBProject.NAME & """ Properties:"
    For Each Itm In CurrentProject.Properties
        Debug.Print Itm.NAME & "=" & Itm.Value
    Next Itm
End Sub
Public Function PropertyGet(PropName As String, PropValue As Variant, Optional PropObject As Object) As Boolean
' ������ �������� ������������� �������
Const c_strProcedure = "PropertyGet"
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
HandleExit:  PropertyGet = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function PropertySet(PropName As String, PropValue As Variant, Optional PropObject As Object, Optional PropType As eDataType = dbText) As Boolean
' ��������� ���������������� �������� � ������� DAO ��� AccessObject
Const c_strProcedure = "PropertySet"
' PropName      - ��� ��������
' PropValue     - �������� ��������
' PropObject    - ������ � �������� ����������� ��������
' PropType      - ��� ������ ��������
Dim Result As Boolean
    On Error Resume Next
    If PropObject Is Nothing Then Set PropObject = CurrentProject ' ��-���������
    ' �������� �������� ��������
    PropObject.Properties(PropName) = PropValue
    Select Case Err.Number
    Case 0: Result = True: GoTo HandleExit
    Case 3270, 2455: ' �������� �� �������
    Case Else: On Error GoTo HandleExit: Err.Raise Err.Number
    End Select
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
        Err.Raise 438 ' Object doesn't support this property or method ' vbObjectError + 512
    End If
    Result = True
HandleExit:     PropertySet = Result: Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_IsSkippedObject(ObjectName) As Boolean
' ��������� ����������� �� ������ � ������������ ��� ��������
Dim Result As Boolean:   ' Result = False
On Error GoTo HandleError
Dim Arr, i As Long ', iMax As Long
    Arr = Split(c_strObjIgnore, ";")
    For i = LBound(Arr) To UBound(Arr)
        Result = (VBA.UCase$(Arr(i)) = VBA.UCase$(ObjectName))
        If Result Then Exit For
    Next i
HandleExit:  p_IsSkippedObject = Result: Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
'-------------------------
' ������� ������������ �������
'-------------------------
Public Function CompileAll()
' ����� ������� ������� SysCmd ��� �������������� ����������/���������� �������.
'Enum accSaveVbaCode
'    accSaveVBAUnkn01 = 16481
'    accSaveVBAUnkn02 = 16482
'    accSaveVBAwCode = 16483 'Save VBA with compiled code
'    accSaveVBAwoCode = 16484 'Save VBA without compiled code
'End Enum
'    For i = 0 To Modules.Count - 1
'        UpdateModule Modules(i).Name
'    Next i
    If Not IsCompiled Then Call SysCmd(504, 16483)
End Function
Public Function CloseAll( _
    Optional OnlyInDesignMode As Boolean = False)
' ��������� �������� ������� Access
Const c_strProcedure = "CloseAll"
#If APPTYPE = 0 Then        ' APPTYPE=Access
Dim i As Byte
Dim oColl, oItem, eObjType As AcObjectType
    On Error GoTo HandleError
'    If SysCmd(714) Then ' �� ������ ��������� ��������
    ' ��������� ���� �� ������� �������� � ������ ������������
    
    ' ��������� ��� �������� �������
    For eObjType = acTable To acModule
        Select Case eObjType
        Case acTable:   Set oColl = CurrentData.AllTables
        Case acQuery:   Set oColl = CurrentData.AllQueries
        Case acForm:    Set oColl = CurrentProject.AllForms
        Case acReport:  Set oColl = CurrentProject.AllReports
        Case acMacro:   Set oColl = CurrentProject.AllMacros
        Case acModule:  Set oColl = CurrentProject.AllModules
        Case Else:      GoTo HandleNextType
        End Select
        For Each oItem In oColl
            With oItem
                If .IsLoaded Then
                    If OnlyInDesignMode And .CurrentView <> 0 Then GoTo HandleNextItem
                    DoCmd.Close eObjType, .NAME, acSaveYes
'Debug.Print IIf(.IsLoaded, "Can't close object", "Object was closed") & ": """ & oItem.Name & """."
                End If
            End With
HandleNextItem:
        Next oItem
HandleNextType:
    Next eObjType
#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
#Else                       ' APPTYPE=��������� ����� ���
#End If                     ' APPTYPE
'    End If
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
#If APPTYPE = 0 Then        ' APPTYPE=Access
Public Function TablesReLink(DatabasePath As String)
' ���������� ����� � ��������� ��������
Const c_strProcedure = "TablesReLink"
Const c_ConnString = ";DATABASE="
Dim tdf As Object 'TableDef
Dim i As Long, iMax As Long
Dim lngBrokenLinks As Long
Dim strCaption As String
Dim prg As clsProgress

    On Error GoTo HandleError
    i = 0:    iMax = CurrentDb.TableDefs.Count
    strCaption = "���������� ������ ������"
'    SysCmd acSysCmdInitMeter, "���������� ������ ������ " & DatabasePath, _
    Set prg = New clsProgress
    prg.Init pCount:=1, pMin:=i, pMax:=iMax, pCaption:=strCaption, pText:=strCaption & " �: """ & DatabasePath & """"
    lngBrokenLinks = 0
    On Error Resume Next
    For Each tdf In CurrentDb.TableDefs
        If Len(tdf.Connect) > 0 Then
            If VBA.Mid$(tdf.Connect, Len(c_ConnString) + 1) <> DatabasePath Then
                With prg
                    If .Canceled Then If MsgBox(c_strBreakProcessMessage, vbYesNo Or vbExclamation Or vbDefaultButton2) = vbYes Then GoTo HandleExit
                    .Canceled = False
                    .Detail = strCaption & "�: " & tdf.NAME
                    .Progress = i
                End With
                tdf.Connect = c_ConnString & DatabasePath
                tdf.RefreshLink
                lngBrokenLinks = lngBrokenLinks + 1
            End If
        End If
        i = i + 1
'        SysCmd acSysCmdUpdateMeter, i
    Next tdf
HandleExit:
    Set prg = Nothing
    If lngBrokenLinks > 0 Then MsgBox strCaption & " c:" & vbCrLf & DatabasePath & vbCrLf & "���������.", vbOKOnly, "�������� ���������!"
'    SysCmd (acSysCmdClearStatus)
    CurrentDb.Close
    Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
#End If                     ' APPTYPE

'-------------------------
' ������� ���������� ������ �������/�������
'-------------------------
Public Function UpdateModule( _
    ObjectName As String, _
    Optional COMMENT As String, _
    Optional VerType As appVerType = appVerBuild, _
    Optional SkipDialog As Boolean = False)
' ��������� ���������� � ������ ���������� ������
Const c_strProcedure = "UpdateModule"
Dim strBegLineMarker As String
Dim BegLine As Long, EndLine As Long
Dim VerDate As Date
Dim strVersion As String, strVerShort As String
Dim strValue As String, strComment As String
Dim IsLoaded As Boolean
Dim Result As Boolean
    Result = False
    On Error GoTo HandleError
DoCmd.Echo False
    If ObjectName = vbNullString Then ObjectName = InputBox("������� ��� ������������ �������:", , ObjectName)
Dim ModuleName As String: ModuleName = ObjectName
Dim ObjectType As AcObjectType
' ��������� ��������� ���
    If ModuleName = vbNullString Then
        Err.Raise vbObjectError + 512
'    ElseIf ModuleName = VBE.ActiveCodePane.CodeModule Then
'        Err.Raise vbObjectError + 513
    ElseIf Not IsModuleExists(ModuleName, ObjectName, ObjectType) Then
        Err.Raise vbObjectError + 514
    'ElseIf Not IsFuncExists(ModuleName) Then
    '    Result = Update Func ... : Goto HandleExit
    End If
' ��������� ������ � ������ �������
    Select Case ObjectType
    Case acModule:  'Do Nothing
    Case acForm:    DoCmd.OpenForm ObjectName, acDesign
    Case acReport:  DoCmd.OpenReport ObjectName, acDesign
    Case Else:      Err.Raise vbObjectError + 514
    End Select
    ' ���� ������ ������ - ���������
    'DoCmd.Save acModule, ModuleName ': DoCmd.Close acModule, ModuleName, acSaveYes
    ' �������� ������� ������ ������
    strVersion = ModuleVersGet(ModuleName)
    ' ����������� ������ ������
    VersionInc strVersion, VerShort:=strVerShort, VerDate:=VerDate, IncType:=VerType
    If SkipDialog Then
        strComment = COMMENT
    Else
        strValue = InputBox("������� ����� ����� ������ " & vbCrLf & "������ " & ModuleName & ":", , strVerShort)
        ' ���� ������ ������
        If strValue <> vbNullString Then strVersion = strValue: VersionSet strVersion, VerShort:=strVerShort, VerDate:=VerDate
        strComment = InputBox("�������� �������� ��������� � ����� ������" & vbCrLf & "������ " & ModuleName & ":", , COMMENT)
    End If
    ModuleVersSet ModuleName, strVersion
    ModuleDateSet ModuleName, CStr(VerDate)
    'ModuleAuthSet ModuleName, Author '& " (" & Support & ")"
' ��������� ����������� � ������
    If Len(strComment) > 0 Then
Dim tmpString As String: tmpString = ModuleAuthGet(ModuleName)
Dim strAuthor As String, strSupport As String: strAuthor = Author: strSupport = Support
    ' ���� ��� ������ ��������� �� ��������� � ������ ������ ������ ��������� ������ � ���� ���������
        Select Case VBA.UCase$(tmpString)
        Case VBA.UCase$(strAuthor), VBA.UCase$(strAuthor) & " (" & VBA.UCase$(strSupport) & ")": tmpString = vbNullString
        Case Else: tmpString = vbNullString
            If Len(strAuthor) > 0 Then tmpString = tmpString & strAuthor
            If Len(strSupport) > 0 Then tmpString = tmpString & " (" & strSupport & ")"
            If Len(tmpString) > 0 Then tmpString = " {" & tmpString & "}"
        End Select
    ' ������� � ����������� ����� ������ � ������� �������� ���������
        tmpString = c_strPrefVerComm & strVerShort & VBA.String$(IIf(c_strPrefLen - Len(strVerShort) - 5 > 0, c_strPrefLen - Len(strVerShort) - 5, 1), " ") & ": " & VBA.Format$(Now, "dd.mm.yyyy") & " - " & strComment & tmpString
        ModuleCommSet ModuleName, tmpString
        'DoCmd.Save acModule, ModuleName
    End If
    'ModuleDebugSet ModuleName, c_bolDebugMode
' ��������� ������ � ������ �������
    DoCmd.Close ObjectType, ObjectName, acSaveYes
' ��������� ������ ����������
    ' �������� ������� ������ ����������
    Call PropertyGet(c_strPrpKeyVer, strVersion)
    ' ��������� ������� ������ ����������
    VersionInc strVersion, VerShort:=strVerShort, VerDate:=VerDate, IncType:=appVerRevis
    ' ��������� ������� ������ ���������� � ��������
    'Call PropertySet(c_strPrpKeyVer, strVersion)
    CurrentProject.Properties(c_strPrpKeyVer) = strVersion
    Result = True
HandleExit:  DoCmd.Echo True: UpdateModule = Result: Exit Function
HandleError:
    Result = False
    Select Case Err.Number
    Case vbObjectError + 512: Debug.Print "�� ������ ��� ������!"
    Case vbObjectError + 513: Debug.Print "�� �������� �������� �������� ������: """ & ModuleName & """!"
    Case vbObjectError + 514: Debug.Print "������: """ & ModuleName & """ �� ������!"
    Case Else: Debug.Print Err.Description
    End Select
    Err.Clear: Resume HandleExit
End Function
Private Function IsModuleExists( _
    ByRef ModuleName As String, _
    Optional ByRef ObjectName As String, _
    Optional ByRef ObjectType As AcObjectType _
    ) As Boolean
' ��������� ������� ���������� ������
Dim Result As Boolean
Dim strObjName As String
' ���������� True, ���� ���� ������ � ����� ������.

    Result = False
    On Error Resume Next
' Application.Modules ����� ������ ����������� ������.
'   c������������� ������������ ������ ���� �� Not IsLoaded ����� �� ������
' CurrentProject.AllModules ����� ������ ������ � �� ����� ������� ���� � �������
    With CurrentProject
' ��������� ��������� ������� �������
        If (.AllModules(ModuleName).NAME = ModuleName) Then ObjectName = ModuleName
        Result = (Err = 0): If Result Then ObjectType = acModule: GoTo HandleExit
        Err.Clear
' ��������� ��������� ���� �������
        If Left$(ObjectName, Len(c_strFrmModPref)) = c_strFrmModPref Then ObjectName = Mid$(ModuleName, Len(c_strFrmModPref) + 1)
        DoCmd.OpenForm ObjectName, acDesign '1 = acDesign
        ModuleName = Forms(ObjectName).Module.NAME
        Result = (Err = 0): If Result Then ObjectType = acForm: GoTo HandleExit
        Err.Clear
' ��������� ��������� ������� �������
        If Left$(ObjectName, Len(c_strFrmModPref)) = c_strRepModPref Then ObjectName = Mid$(ModuleName, Len(c_strFrmModPref) + 1)
        DoCmd.OpenReport ObjectName, acDesign '1 = acDesign
        ModuleName = Forms(ObjectName).Module.NAME
        Result = (Err = 0): If Result Then ObjectType = acReport: GoTo HandleExit
        Err.Clear
    End With
    DoCmd.Close ObjectType, ObjectName
HandleExit:     IsModuleExists = Result:    Exit Function
HandleError:    Result = False: Err.Clear:  Resume HandleExit
End Function
Public Sub UpdateFunc( _
    funcName As String, _
    Optional ModuleName As String, _
    Optional VerType As appVerType = appVerBuild, _
    Optional SkipDialog As Boolean = False)
' ��������� ���������� � ������ �������
Const c_strProcedure = "UpdateFunc"
' v.1.0.0       : 02.04.2019 - �������� ������
Dim BegLine As Long, EndLine As Long
Dim i As Long
Dim Result As Boolean

    Result = False
    On Error GoTo HandleError
    If funcName = vbNullString Then Err.Raise m_errProcNameWrong
    If ModuleName = vbNullString Then
' �������� ������� ������ � ���� clsForm.Form(Set), ����� �������� ��� ������
        ModuleName = Split(funcName, ".")(0)
        funcName = VBA.Mid$(funcName, Len(ModuleName) + 2)
    End If
' ��������� ��� �������
Dim ProcKind As vbext_ProcKind: ProcKind = vbext_pk_Proc
    If VBA.Right$(funcName, 1) = ")" Then
        Select Case VBA.Mid$(funcName, Len(funcName) - 3, 3)
        Case "Let": ProcKind = vbext_pk_Let
        Case "Set": ProcKind = vbext_pk_Set
        Case "Get": ProcKind = vbext_pk_Get
        End Select
        If VBA.Left$(VBA.Right$(funcName, 5), 1) = "(" Then
            funcName = VBA.Left$(funcName, Len(funcName) - 5)
        ElseIf VBA.Left$(VBA.Right$(funcName, 14), 10) = "(Property " Then
            funcName = VBA.Left$(funcName, Len(funcName) - 14)
        Else
            Err.Raise m_errProcNameWrong
        End If
    End If
' ��������� ����������� ������� � �� �������
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
    With objModule
        BegLine = .ProcBodyLine(funcName, ProcKind) ' ������ �� ������ �������
        EndLine = BegLine + .ProcCountLines(funcName, ProcKind) - 3
        BegLine = CodeLineNext(ModuleName, BegLine)
' ��������� ���������� � ������ �������
Dim strVersion As String, strVerShort As String, VerDate As Date
Dim strComment As String
Dim tmpString As String: tmpString = "'" & c_strPrefVerComm & "*.*.*:"
Dim NxtLine As Long
'        tmpString = c_strPrefVerComm & strVerShort & String$(IIf(c_strPrefLen - Len(strVerShort) - 5 > 0, c_strPrefLen - Len(strVerShort) - 5, 1), " ") & ": " & VBA.format$(Now, "dd.mm.yyyy") & " - " & strComment & tmpString
'        Do
            NxtLine = EndLine
            Result = .Find(tmpString, _
                StartLine:=BegLine, EndLine:=NxtLine, _
                StartColumn:=0, EndColumn:=0, _
                PatternSearch:=True)
            If Result Then
    ' ���� ������� - ��������� ����� ��������� ������ �� �������
                strVersion = VBA.Trim(VBA.Mid$(.Lines(BegLine, 1), Len(c_strPrefVerComm) + 2))
                strVersion = VBA.Trim$(VBA.Left$(strVersion, InStr(strVersion, ":") - 1))
'                Exit Do
'                BegLine = NxtLine + 1
'            Else
'                Exit Do
            End If
'        Loop
    ' ����������� ������ �������
        VersionInc strVersion, VerShort:=strVerShort, VerDate:=VerDate, IncType:=VerType
        If Not SkipDialog Then
            tmpString = InputBox("������� ����� ����� ������ " & vbCrLf & "������� " & ModuleName & "." & funcName & ":", , strVerShort)
            ' ���� ������ ������
            If tmpString <> vbNullString Then strVersion = tmpString: VersionSet strVersion, VerShort:=strVerShort, VerDate:=VerDate
            strComment = InputBox("�������� �������� ��������� � ����� ������" & vbCrLf & "������� " & ModuleName & "." & funcName & ":")
        End If
    ' ��������� ����������� � ������
    tmpString = ModuleAuthGet(ModuleName)
Dim strAuthor As String, strSupport As String: strAuthor = Author: strSupport = Support
    ' ���� ��� ������ ��������� �� ��������� � ������ ������ ������ ��������� ������ � ���� ���������
        Select Case VBA.UCase$(tmpString)
        Case VBA.UCase$(strAuthor), VBA.UCase$(strAuthor) & " (" & VBA.UCase$(strSupport) & ")": tmpString = vbNullString
        Case Else:  tmpString = " (" & strAuthor & " (" & strSupport & ")"
        End Select
    ' ������� � ����������� ����� ������ � ������� �������� ���������
        tmpString = "'" & c_strPrefVerComm & strVerShort & VBA.String$(IIf(c_strPrefLen - Len(strVerShort) - 5 > 0, c_strPrefLen - Len(strVerShort) - 5, 1), " ") & ": " & VBA.Format$(Now, "dd.mm.yyyy") & " - " & strComment & tmpString
        .InsertLines BegLine, tmpString
    ' ������� � ����������� ����� ������ � ������� �������� ���������
        tmpString = "��������� � " & funcName & " - " & strComment
        UpdateModule ModuleName, tmpString, SkipDialog:=True
    End With
HandleExit:     Exit Sub
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "������� ������ ��� �������!"
    Case m_errModuleIsActive: Debug.Print "���������� �������� �������� ������: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "������: """ & ModuleName & """ �� ������!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Sub
'-------------------------
' ������� ��� ������ � ��������
'-------------------------
Public Function VersionSet( _
    ByRef VerText As String, _
    Optional ByRef VerDate As Date, _
    Optional ByRef VerShort As String _
    ) As Long
' ��������� �� ������ ������ ������ ������, ���������
Const c_strProcedure = "VersionSet"
' VerText - �������� ������ ������, �� ������ �������� ������ ������
' VerDate - �������� ���� ������ ��� ���������� � Revision
' VerShort- �� ������ �������� ������� ������
' ������� ���������� �������� ��� ������ � ���� AABBB
Dim VerType As typVersion
Dim Result As Long

    Result = False
    On Error GoTo HandleError
    Call p_VersionFromString(VerText, VerType)
    With VerType
        ' VerText - �������� ������ ������, �� ������ �������� ������ ������ � ������ ���������� ������
        ' VerShort- �� ������ �������� ������� ������ � ������ ���������� ������
    ' �������� ���������� � ������
        Result = .VerCode
        .Revision = VersionDate2Rev(VerDate): .VerDate = VerDate
        VerShort = .Major & c_strVerDelim & .Minor & c_strVerDelim & .Build
        VerText = VerShort & c_strVerDelim & .Revision
    End With
HandleExit:     VersionSet = Result: Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function
Public Function VersionGet( _
    ByRef VerText As String, _
    Optional ByRef VerShort As String, _
    Optional ByRef VerDate As Date, _
    Optional RelType As Byte = 1 _
    ) As Long
' �������� �� ������ ������ �� ����������
Const c_strProcedure = "VersionGet"
' VerText - �������� ������ ������, �� ������ �������� ������ ������
' VerShort- �� ������ �������� ������� ������
' VerDate - �� ������ �������� ���� ������ ���������� �� ����������� Revision
' RelType - ��� ������������ ���������� � ������ 0 - ���, 1 - ������� ��������, 3 - ������ ��������
' ������� ���������� �������� ��� ������ � ���� AABBB
Dim VerType As typVersion
Dim Result As Long

    Result = False
    On Error GoTo HandleError
    Call p_VersionFromString(VerText, VerType)
    With VerType
        ' VerText - �������� ������ ������, �� ������ �������� ������ ������ � ������ ���������� ������
        ' VerShort- �� ������ �������� ������� ������ � ������ ���������� ������
    ' �������� ���������� � ������
        Result = .VerCode: VerDate = .VerDate
        VerShort = .Major & c_strVerDelim & .Minor & c_strVerDelim & .Build
        VerText = VerShort & c_strVerDelim & .Revision
    ' ��������� ���������� � ������
        Select Case RelType
        Case 0: VerText = VerText & c_strVerDelim & .Release: If .RelSubNum > 0 Then VerText = VerText & " " & .RelSubNum
        Case 1: If .Release > 0 Then VerText = VerText & " " & .RelShort: If .RelSubNum > 0 Then VerText = VerText & .RelSubNum
        Case 2: If .Release > 0 Then VerText = VerText & " " & .RelFull: If .RelSubNum > 0 Then VerText = VerText & " " & .RelSubNum
        End Select
    End With
HandleExit:     VersionGet = Result: Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function
Public Function VersionInc( _
    ByRef VerText As String, _
    Optional ByRef VerShort As String, _
    Optional ByRef VerDate As Date, _
    Optional IncStep As Integer = 1, _
    Optional IncType As appVerType = appVerBuild, _
    Optional RelType As Byte = 1 _
    ) As Long
' �������� ����� ������ �������� ������� � ���������� �� �������� ��������
Const c_strProcedure = "VersionInc"
' VerText - �������� ������ ������, �� ������ �������� ������ ������ � ������ ���������� ������
' VerShort- �� ������ �������� ������� ������ � ������ ���������� ������
' VerDate - �� ������ �������� ���� ������ ���������� �� ����������� Revision
' IncStep - ���������� ������ ��������� ������� ������ (+-)
' IncType - ����� ����������� ������� ������
' RelType - ��� ������������ ���������� � ������ 0 - ���, 1 - ������� ��������, 3 - ������ ��������
' ������� ���������� �������� ��� ������ � ���� AABBB
Dim VerType As typVersion
Dim Result As Long

    Result = False
    On Error GoTo HandleError
' ��������� ������� ������
    Call p_VersionFromString(VerText, VerType): If IncStep = 0 Then GoTo HandleExit
    With VerType
' ��������� ������� ������
    ' ������� ��������� ��� ������ ���������� ���������� �� ��������� IncType
        .Revision = VersionDate2Rev(Now): VerDate = VersionRev2Date(.Revision)
        Select Case IncType
        Case appVerMajor: .Major = .Major + IncStep
        Case appVerMinor: .Minor = .Minor + IncStep
        Case appVerBuild: .Build = .Build + IncStep
        Case appRelease:  If .Release > appReleaseNotDefine Then .Release = .Release + IncStep
        Case appRelSubNum:  If .Release > appReleaseNotDefine Then .RelSubNum = .RelSubNum + IncStep
        End Select
    End With
' ������������ ����������
    p_VersionCheck VerType
    With VerType
        ' VerText - �������� ������ ������, �� ������ �������� ������ ������ � ������ ���������� ������
        ' VerShort- �� ������ �������� ������� ������ � ������ ���������� ������
    ' �������� ���������� � ������
        Result = .VerCode: VerDate = .VerDate
        VerShort = .Major & c_strVerDelim & .Minor & c_strVerDelim & .Build
        VerText = VerShort & c_strVerDelim & .Revision
    ' ��������� ���������� � ������
        Select Case RelType
        Case 0: VerText = VerText & c_strVerDelim & .Release: If .RelSubNum > 0 Then VerText = VerText & " " & .RelSubNum
        Case 1: If .Release > 0 Then VerText = VerText & " " & .RelShort: If .RelSubNum > 0 Then VerText = VerText & .RelSubNum
        Case 2: If .Release > 0 Then VerText = VerText & " " & .RelFull: If .RelSubNum > 0 Then VerText = VerText & " " & .RelSubNum
        End Select
    End With
HandleExit:     VersionInc = Result: Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function
Public Function VersionCmp(Ver1 As String, Ver2 As String) As Integer
' ���������� ������
Const c_strProcedure = "VersionCmp"
' ����������:
' 0 - ���� Ver1=Ver2
' 1 - ���� Ver1>Ver2
'-1 - ���� Ver1<Ver2
'-32768 - ���� ������
Dim Result As Integer
    On Error GoTo HandleError
Dim v1 As typVersion: If Not p_VersionFromString(Ver1, v1) Then Err.Raise vbObjectError + 512
Dim v2 As typVersion: If Not p_VersionFromString(Ver2, v2) Then Err.Raise vbObjectError + 512
    If v1.Revision > 0 And v2.Revision > 0 Then Result = Sgn(v1.Revision - v2.Revision): GoTo HandleExit
    Result = Sgn(v1.Major - v2.Major): If Result <> 0 Then GoTo HandleExit
    Result = Sgn(v1.Minor - v2.Minor): If Result <> 0 Then GoTo HandleExit
    Result = Sgn(v1.Build - v2.Build) ': If Result <> 0 Then GoTo HandleExit
HandleExit:     VersionCmp = Result: Exit Function
HandleError:    Result = -32768: Err.Clear: Resume HandleExit
End Function
Public Function VersionDate2Rev(DateTime As Date) As Long: VersionDate2Rev = CCur(DateTime) * 10 ^ c_bytTimeDig: End Function
Public Function VersionRev2Date(RevString As Long) As String: VersionRev2Date = VBA.Format$(CDate(RevString / 10 ^ c_bytTimeDig), "dd.mm.yyyy hh:nn:ss"): End Function
Private Function p_VersionCheck(ByRef VerType As typVersion) As Boolean
' ��������� ������ � ������� ������
Const c_strProcedure = "p_VersionCheck"
Dim Result As Boolean

    Result = False
    On Error GoTo HandleError
    With VerType
    ' ��������� ���������
        If .Major < 0 Then .Major = 0 Else If .Major > 10 ^ c_bytMajorDig - 1 Then .Major = 10 ^ c_bytMajorDig - 1: Err.Raise 6 'Overflow
        If .Minor < 0 Then .Minor = 0 Else If .Minor > 10 ^ c_bytMinorDig - 1 Then .Minor = 10 ^ c_bytMinorDig - 1: Err.Raise 6 'Overflow
        If .Build < 0 Then .Build = 0
        If .Release < 0 Then .Release = 0: If .Release > appReleaseEOL Then .Release = appReleaseNotDefine: Err.Raise 6 'Overflow
        If .RelSubNum < 0 Then .RelSubNum = 0
        If .Revision > 0 Then
    ' �������� �� ������� ���� �����
        Dim tmp As Long: tmp = 10 ^ (Len(CStr(.Revision)) - c_bytDateDig): If tmp = 0 Then tmp = 1
            .VerDate = CDate(.Revision / tmp)
        End If
    ' ��������� �������� ��� ������ (AABBB)
        .VerCode = .Major * 10 ^ c_bytMinorDig + .Minor
    End With
    Result = True
HandleExit:     p_VersionCheck = Result: Exit Function
HandleError:    Result = False
    Select Case Err.Number
    Case 6:     Err.Clear: Resume Next ' Overflow
    Case Else:  Err.Clear: Resume HandleExit
    End Select
End Function
Private Function p_VersionFromString( _
    ByRef VerText As String, _
    ByRef VerType As typVersion _
    ) As Boolean
' ��������� ������� ������ �� ������
Const c_strProcedure = "p_VersionFromString"
' VerText - ������ ������ ������ ���� A.B.C.D[r][n]
' VerType - ���������� � ������ ����������� �� ����������
Dim strRel As String
Dim Result As Boolean

    Result = False
    On Error GoTo HandleError
    VerText = VBA.Trim$(VerText)
Dim arrVer: arrVer = Split(VerText, c_strVerDelim)
Dim i As Long, iMax As Long: i = 0: iMax = UBound(arrVer) - LBound(arrVer)
' ��������� ������ ������
    With VerType
        Do While i <= iMax
            Select Case i
            Case appVerMajor: .Major = arrVer(i) ' A � ������� ����� ������
            Case appVerMinor: .Minor = arrVer(i) ' B � ��������������� ����� ������
            Case appVerBuild: .Build = arrVer(i) ' C � ����� ����� ������
            Case appVerRevis                     ' D - ����� ������� (�����-���� �����)
                .Revision = Fix(Val(arrVer(i)))
                ' �������� ��� ����� ������� + ��� ������, - ��������� � ���������
                If Not IsNumeric(arrVer(i)) Then strRel = VBA.Mid$(arrVer(i), Len(.Revision) + 1): Exit Do
            Case Else:  strRel = arrVer(i)       '[r]- ��� ������
            End Select
            i = i + 1
        Loop
        If Len(strRel) > 0 Then .Release = p_GetReleaseType(strRel, .RelSubNum, .RelShort, .RelFull)
    End With
' ������������ ������ ������
    p_VersionCheck VerType
    Result = True
HandleExit:     p_VersionFromString = Result: Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function
Private Function p_GetReleaseType( _
    ReleaseData As String, _
    ReleaseSubNum As Integer, _
    Optional ReleaseShortDesc As String, _
    Optional ReleaseFullDesc As String _
    ) As appRelType
' ���������� ��� � �������� ���� ������
Const c_strProcedure = "p_GetReleaseType"
' ReleaseData -      �������� ������
' ReleaseSubNum -    �������� ������ ��� ���� ������ (rc1)
' ReleaseShortDesc - ������� �������� ���� ������
' ReleaseFullDesc -  ������ �������� ���� ������
Dim strData As String
Dim Result As appRelType
    On Error GoTo HandleError
    strData = VBA.Trim$(ReleaseData)
    Result = False: ReleaseShortDesc = vbNullString: ReleaseFullDesc = vbNullString
' �� ����� � ������ ���� �����
    Dim c As String * 1, strNum As String
    Dim i As Long: i = Len(strData)
    Do While i > 0
        c = VBA.Right$(strData, 1)
        Select Case c
        Case 0 To 9: strNum = c & strNum
        Case Else:  Exit Do
        End Select
        i = i - 1: strData = VBA.Left$(strData, i)
    Loop
    strData = VBA.Trim$(strData)
' �������� �������������� ����� ���� ������
    If Len(strNum) > 0 Then ReleaseSubNum = CByte(strNum)
' �������� �������� ���� ������
    Select Case strData
    Case 1, "pa", "Pre-alpha":                  Result = 1: ReleaseShortDesc = "pa":  ReleaseFullDesc = "Pre-alpha"
    Case 2, "a", "Alpha":                       Result = 2: ReleaseShortDesc = "a":   ReleaseFullDesc = "Alpha"
    Case 3, "b", "Beta":                        Result = 3: ReleaseShortDesc = "b":   ReleaseFullDesc = "Beta"
    Case 4, "rc", "Release Candidate":          Result = 4: ReleaseShortDesc = "rc":  ReleaseFullDesc = "Release Candidate"
    Case 5, "rtm", "Release to manufacturing":  Result = 5: ReleaseShortDesc = "rtm": ReleaseFullDesc = "Release to manufacturing"
    Case 6, "ga", "General availability":       Result = 6: ReleaseShortDesc = "ga":  ReleaseFullDesc = "General availability"
    Case 7, "eol", "End of life":               Result = 7: ReleaseShortDesc = "eol": ReleaseFullDesc = "End of life"
    End Select
HandleExit:     p_GetReleaseType = Result: Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function
'----------------------
' ������� ������ � ������� �������
'----------------------
Private Function ModuleExists(ByVal ModuleName As String, Optional objModule As Object) As Boolean
' ��������� ������� ���������� ������
Dim Result As Boolean
' ���������� True, ���� ���� ������ � ����� ������.
    On Error Resume Next
    Set objModule = Application.VBE.ActiveVBProject.VBComponents(ModuleName).CodeModule
    Result = (Err = 0): Err.Clear
' Application.Modules ����� ������ ����������� ������.
'   c������������� ������������ ������ ���� �� Not IsLoaded ����� �� ������
HandleExit:     ModuleExists = Result:    Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function
Public Function ObjectExists(ObjectName As String) As Boolean
' ��������� ������� ���������� ������� � ����������
    ObjectExists = p_ObjectInfo(ObjectName)
End Function
Private Function p_ObjectInfo(ObjectName As String, _
    Optional ObjectType As appObjectType, Optional ObjectTypeName, Optional ObjectTypeDesc, _
    Optional ObjectFileName, Optional ObjectFileExt, Optional ObjectModuleName) As Boolean
' ���������� ���������� �� ������� ����������
Const c_strProcedure = "p_ObjectInfo"
' ObjectName - ��� ������������ �������
' ObjectType - ������������ ��� �������
' ObjectTypeName - �������� ���� ������� (= ��� ����� � ����� ������)
' ObjectTypeDesc - �������� ���� �������
' ObjectFileName - ��� �����
' ObjectFileExt - ���������� ��� �����
' ObjectModuleName - ��� ������ ������� (��� �������)
' ���������� True, ���� ������ � ����� ������ ���������� � ����������.
'-------------------------
Dim Result As Long ': Result = False
    On Error Resume Next
' ��-���������
    ObjectType = appObjTypUndef:    ObjectTypeDesc = "������"
    ObjectFileExt = vbNullString:   ObjectModuleName = vbNullString
' ��������� ������ �������
    Result = p_GetModuleType(ObjectName)
    If Result Then
    Select Case Result
    Case vbext_ct_StdModule:        ObjectType = appObjTypBas: ObjectTypeName = c_strObjTypModule: ObjectTypeDesc = "����������� ������": ObjectModuleName = ObjectName: ObjectFileExt = c_strObjExtBas
    Case vbext_ct_ClassModule:      ObjectType = appObjTypCls: ObjectTypeName = c_strObjTypModule: ObjectTypeDesc = "������ ������": ObjectModuleName = ObjectName: ObjectFileExt = c_strObjExtCls
    'Case vbext_ct_Document:
    'Case vbext_ct_MSForm
    'Case vbext_ct_ActiveXDesigner
    End Select
    End If
    If ObjectType <> appObjTypUndef Then ObjectFileName = p_TextAlpha2Code(ObjectName) & "." & ObjectFileExt: GoTo HandleExit
' ��������� ������� ����������� ��� ���������� ����������
#If APPTYPE = 0 Then        ' APPTYPE=Access
' ��������� ��������� ������� Access �� MSysObjects
    Result = Nz(DLookup("Type", c_strMSysObjects, "Name=""" & ObjectName & """"), msys_ObjectUndef) '(ObjectName)
    If Result Then
    Select Case Result
    ' ��������� ������� ������� � ����� ������ � ���� �� �������
    Case msys_ObjectForm:   ObjectType = appObjTypAccFrm: ObjectTypeName = c_strObjTypAccFrm: ObjectTypeDesc = "����� Access": ObjectFileExt = c_strObjExtFrm
        If p_GetModuleType(c_strFrmModPref & ObjectName) = vbext_ct_Document Then ObjectModuleName = c_strFrmModPref & ObjectName
    Case msys_ObjectReport: ObjectType = appObjTypAccRep: ObjectTypeName = c_strObjTypAccRep: ObjectTypeDesc = "����� Access": ObjectFileExt = c_strObjExtRep
        If p_GetModuleType(c_strRepModPref & ObjectName) = vbext_ct_Document Then ObjectModuleName = c_strRepModPref & ObjectName
    Case msys_ObjectMacro:  ObjectType = appObjTypAccMac: ObjectTypeName = c_strObjTypAccMac: ObjectTypeDesc = "������ Access": ObjectFileExt = c_strObjExtMac
    Case msys_ObjectQuery:  ObjectType = appObjTypAccQry: ObjectTypeName = c_strObjTypAccQry: ObjectTypeDesc = "������": ObjectFileExt = c_strObjExtQry
    Case msys_ObjectTable:  ObjectType = appObjTypAccTbl: ObjectTypeName = c_strObjTypAccTbl: ObjectTypeDesc = "������� Access": ObjectFileExt = c_strObjExtXml
    Case msys_ObjectLinked: ObjectType = appObjTypAcclnk: ObjectTypeName = c_strObjTypAccTbl: ObjectTypeDesc = "��������� �������": ObjectFileExt = c_strObjExtLnk
    Case Else: Err.Raise m_errModuleNameWrong
    End Select
    End If
'#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
#Else                       '
'Stop    ' ???
#End If                     ' APPTYPE
' ���� ����� ������ � ������� �������� ��� ����� �� ����� ������� � �������
' ���� �� ����� - �������� ��� ������� ��� ���� � �����
    If ObjectType <> appObjTypUndef Then ObjectFileName = p_TextAlpha2Code(ObjectName) & "." & ObjectFileExt: GoTo HandleExit
    Result = oFso.FileExists(ObjectName): If Not Result Then GoTo HandleExit
    ObjectFileExt = oFso.GetExtensionName(ObjectName)           ' ����������
    ObjectName = p_TextCode2Alpha(oFso.GetBaseName(ObjectName)) ' ��� ������� �� ����� �����
' ���� ������ ���� � ����� �������� - �������� ��� �� ����������
    Select Case ObjectFileExt
    Case c_strObjExtBas:    ObjectType = appObjTypBas: ObjectTypeName = c_strObjTypModule: ObjectTypeDesc = "����������� ������":   ObjectModuleName = ObjectName
    Case c_strObjExtCls:    ObjectType = appObjTypCls: ObjectTypeName = c_strObjTypModule: ObjectTypeDesc = "������ ������": ObjectModuleName = ObjectName
#If APPTYPE = 0 Then        ' APPTYPE=Access
    Case c_strObjExtFrm:    ObjectType = appObjTypAccFrm: ObjectTypeName = c_strObjTypAccFrm: ObjectTypeDesc = "����� Access":      ObjectModuleName = c_strFrmModPref & ObjectName
    Case c_strObjExtRep:    ObjectType = appObjTypAccRep: ObjectTypeName = c_strObjTypAccRep: ObjectTypeDesc = "����� Access":      ObjectModuleName = c_strRepModPref & ObjectName
    Case c_strObjExtMac:    ObjectType = appObjTypAccMac: ObjectTypeName = c_strObjTypAccMac: ObjectTypeDesc = "������ Access"
    Case c_strObjExtQry:    ObjectType = appObjTypAccQry: ObjectTypeName = c_strObjTypAccQry: ObjectTypeDesc = "������"
    Case c_strObjExtTxt:    ObjectType = appObjTypAccTbl: ObjectTypeName = c_strObjTypAccTbl: ObjectTypeDesc = "������� Access (TXT)"
    Case c_strObjExtCsv:    ObjectType = appObjTypAccTbl: ObjectTypeName = c_strObjTypAccTbl: ObjectTypeDesc = "������� Access (CSV)"
    Case c_strObjExtXml:    ObjectType = appObjTypAccTbl: ObjectTypeName = c_strObjTypAccTbl: ObjectTypeDesc = "������� Access (XML)"
    Case c_strObjExtLnk:    ObjectType = appObjTypAcclnk: ObjectTypeName = c_strObjTypAccTbl: ObjectTypeDesc = "��������� �������"
    Case Else: Err.Raise m_errModuleNameWrong
'#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
#Else                       '
'Stop    ' ???
#End If                     ' APPTYPE
    End Select
HandleExit:     p_ObjectInfo = ObjectType <> appObjTypUndef:    Exit Function
HandleError:    Dim Message As String
    Select Case Err.Number
    Case m_errModuleNameWrong:  Message = "������� ������ ��� �������!"
    Case Else:                  Message = Err.Description
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_GetModuleType(ModuleName As String) As vbext_ComponentType
    On Error GoTo HandleError
    p_GetModuleType = Application.VBE.ActiveVBProject.VBComponents(ModuleName).Type
HandleExit:  Exit Function
HandleError: Err.Clear: Resume HandleExit
End Function
Public Function ObjectDelete(ObjectName As String, Optional ObjectType As appObjectType) As Boolean
' ������� ������ �� �������
Dim Result As Boolean
' �������� � ��������� ������� �� ��������, ����� �����-������ ������� ��������� ������ ����������
    On Error Resume Next
Dim oTmp As Object
    If ObjectType = appObjTypUndef Then Result = p_ObjectInfo(ObjectName, ObjectType): If Not Result Then Err.Raise m_errObjectTypeUnknown
    Select Case ObjectType
    Case appObjTypBas, appObjTypCls:  With Application.VBE.ActiveVBProject: Set oTmp = .VBComponents(ObjectName): .VBComponents.Remove oTmp: End With
    'Case appObjTypDoc: :  With Application.VBE.ActiveVBProject: Set oTmp = .VBComponents(ObjectName): .VBComponents.Remove oTmp: End With
    'Case appObjTypMsf, appObjTypAxd: :  With Application.VBE.ActiveVBProject: Set oTmp = .VBComponents(ObjectName): .VBComponents.Remove oTmp: End With
#If APPTYPE = 0 Then        ' ���� Access
    Case appObjTypAccTbl, appObjTypAcclnk: DoCmd.DeleteObject ObjectType And &HFF&, ObjectName
    Case appObjTypAccQry, appObjTypAccMac: DoCmd.DeleteObject ObjectType And &HFF&, ObjectName
    Case appObjTypAccFrm, appObjTypAccRep ', appObjTypAccDoc: DoCmd.DeleteObject ObjectType And &HFF&, ObjectName
    'Case appObjTypAccDap, appObjTypAccSrv: DoCmd.DeleteObject ObjectType And &HFF&, ObjectName
    'Case appObjTypAccDia, appObjTypAccPrc, appObjTypAccFun: DoCmd.DeleteObject ObjectType And &HFF&, ObjectName
'#ElseIf APPTYPE = 1 Then    ' ���� Excel
    'Case appObjTypXlsDoc
#Else                       '
#End If
    Case Else: Err.Raise vbObjectError + 512
    End Select
    Result = (Err = 0): Err.Clear
HandleExit:     ObjectDelete = Result:    Exit Function
HandleError:    Result = False: Err.Clear: Resume HandleExit
End Function
Public Function ObjectRead(SourceFile As String) As Boolean
Dim Result As Long
    On Error GoTo HandleError
Dim strMessage As String: strMessage = "������ ������� ��������� �������."
    Result = p_ObjectRead(SourceFile, ReadType:=orwAlways, AskBefore:=True, Message:=strMessage)
    MsgBox strMessage
HandleExit:  ObjectRead = Result = 0: Exit Function
HandleError: Result = Err: Err.Clear: Resume HandleExit
End Function
Public Function ObjectWrite(ObjectName As String, _
    Optional ByVal FilePath As String) As Boolean
Dim Result As Long
    On Error GoTo HandleError
Dim strMessage As String: strMessage = "������� ������� ��������� �������."
    Result = p_ObjectWrite(ObjectName, FilePath, WriteType:=orwAlways, AskBefore:=True, Message:=strMessage)
    MsgBox strMessage
HandleExit:  ObjectWrite = Result = 0: Exit Function
HandleError: Result = Err: Err.Clear: Resume HandleExit
End Function

Private Function p_ObjectRead(SourceFile As String, _
    Optional TargetName As String, _
    Optional ObjectType As appObjectType = appObjTypUndef, _
    Optional ReadType As ObjectRwType = orwDestMiss, _
    Optional AskBefore As Boolean = True, _
    Optional Message As String) As Long
' ������ ������ �� ����������� ������ � ������� ������ � ��������� ������ (��� ��������)
Const c_strProcedure = "p_ObjectRead"
' SourceFile    - ���� � ����� ��������� �������
' TargetName    - ��� ������� � ������� ������� ������ ���� ��������/������
' ObjectType    - ��� ��������� �������
' ReadType      - ��� �������� ������
' AskBefore     - ����� ����������� ������ ���������� ������������� ������������/��������� ��� �������������
' Message       - ���������� ��������� �� ������ �������� (��� ������ ���������� ��������)
' ���������� ��������� ��������:   0 - ��������, ����� ��� ������
Const cstrMsgYesNo = """��""  - ���� ������ ����� �������� ���� ���������, " & vbCrLf & _
                     """���"" - ���� ������ ���������� ��������"
Const cstrMsgCancel = """������"" - ���� ������ ����� ������ � ��� ����������� �������� ����������� ��� ��������������."
Dim Result As Long: Result = 1
Dim bolClear As Boolean:  bolClear = False ' � ������ ������ ������ �� �������� �������� "���������"
Dim bolRestored As Boolean ' ������� ���� ��� ���� ��� ������ ��� ����������� �������
Dim errCount As Integer    ' ������� ������
    On Error GoTo HandleError
'    DoCmd.SetWarnings False
'Dim SourceFile As String, strObjName As String ', strFileExtn As String
Dim strObjName As String, lngObjType As appObjectType
Dim strObjType As String, strObjFile As String, strObjExtn As String, strModName As String
Dim bolProceed As Boolean       ' ���� ������������� ������ ������� �� �����
Dim intCmp As Integer
' ��������� ������� �����
    strObjName = SourceFile
Dim bolSrcExist As Boolean:  bolSrcExist = p_ObjectInfo(strObjName, ObjectType, , strObjType, strObjFile, strObjExtn): If Not bolSrcExist Then Err.Raise 76
    'strFileExtn = VBA.LCase$(.GetExtensionName(SourceFile))    ' ���������� �����
    If Len(TargetName) = 0 Then TargetName = strObjName
' ��������� ������� ������� � �������
    strObjName = TargetName
Dim bolDestExist As Boolean: bolDestExist = p_ObjectInfo(strObjName, lngObjType, ObjectModuleName:=strModName)
'!!! ���� ��� ����������� �� ����� � ���������� �� ��������� - ���������� �����-�� ������
    If bolDestExist And (ObjectType <> lngObjType) Then ObjectType = lngObjType ': Stop
' ��������� ����������� ������ ������ � ������ ��
    If (ObjectType And &HFFFF&) > 0 Then
Dim strSrcVer As String, datSrcDate As Date ', strFilDesc As String
Dim strDestVer As String, datDestDate As Date ', strModDesc As String
' ������ ����� ������������� ������ � ��������� �� ���� ������ � ������
' ToDo: ��������� �� ������ �������� ��� ������� TargetName
        If bolDestExist Then If Len(strModName) > 0 Then Call ModuleInfo(strModName, strDestVer, datDestDate)
' ������ ���� ��� ����� � ��������� �� ���� ������ � ������, ���� � ��.
        If bolSrcExist Then Call ModuleInfoFromFile(SourceFile, strSrcVer, datSrcDate)
    ' ���������� ������
        If bolSrcExist And bolDestExist Then intCmp = VersionCmp(strSrcVer, strDestVer)
    End If
' ��������� ������������� ���������� �������� ����� �������� ������
    Message = strObjType & " """ & TargetName & """"
    Select Case ReadType
    Case orwAlways                  ' ������ ������ (��������������)
        Message = Message & " ������ �������������� �� ��������� �����."
    Case orwSrcNewerOrDestMissing   ' ������ ���� ���� ����� ��� ������ �����������
        Message = Message & " �������������� �� ��������� �����."
        If (intCmp <> 1) And bolDestExist Then Err.Raise m_errWrongVersion
    Case orwSrcNewer                ' ������ ���� ���� ����� � ������ ����������
        Message = Message & " ���������� �� �����."
        If Not (intCmp = 1) Then Err.Raise m_errWrongVersion
        If Not bolDestExist Then Err.Raise m_errDestMissing
    Case orwDestMiss                ' ������ ���� ������ �����������
        Message = Message & " ���������� �� �����."
        If bolDestExist Then Err.Raise m_errDestExists
    Case orwSrcOlder                ' ������ ���� ���� ������ � ������ ����������
        Message = Message & " �������������� ���������� ������ �� �����."
        If Not (intCmp = -1) Then Err.Raise m_errWrongVersion
        If Not bolDestExist Then Err.Raise m_errDestMissing
    Case Else: Err.Raise m_errObjectActionUndef ' �������������� ������� - ��������� ��� ������ � ������
    End Select
' ��������� � ����� ���������� � �������
    If Len(strDestVer) > 0 Or datDestDate > 0 Then
        Message = Message & vbCrLf & "������� ������"
        If Len(strDestVer) > 0 Then Message = Message & ": " & strDestVer
        If datDestDate > 0 Then Message = Message & " �� " & datDestDate
    End If
    If Len(strSrcVer) > 0 Or datSrcDate > 0 Then
        Message = Message & vbCrLf & "����� ������"
        If Len(strSrcVer) > 0 Then Message = Message & ": " & strSrcVer
        If datSrcDate > 0 Then Message = Message & " �� " & datSrcDate
    End If
' !!! ���������� ������� ����������� ��� ������ �������� ������� ��� ��������� ���������� ������������ ����
    If p_IsSkippedObject(TargetName) Then Err.Raise m_errSkippedByList

' �������� ���������� ������������ �� ���������� ������
    If AskBefore Then ' intAsk <> vbCancel Then
        Select Case MsgBox(Message & vbCrLf & vbCrLf & cstrMsgYesNo & ", " & vbCrLf & cstrMsgCancel, vbYesNoCancel Or vbInformation, "���������� �������")
        Case vbYes                                      ' �������� ������ �������
        Case vbCancel:  AskBefore = False               ' �������� ������� � �����������
        Case Else:      Err.Raise m_errSkippedByUser    ' ���������� �������
        End Select
    End If
' ���� ������ ���������� ������� ��� �� �������
    If bolDestExist Then If Not ObjectDelete(TargetName, ObjectType) Then Err.Raise m_errObjectCantRemove
' � ����������� �� ���� ��������� �������
' !!! ��� �������� �������� ������������ ������ �� ������������ !!!
    Select Case ObjectType
    Case appObjTypMod, appObjTypBas, appObjTypCls
                                            VBE.ActiveVBProject.VBComponents.Import SourceFile
'    Case appObjTypMsf
'    Case appObjTypAxd
'    Case appObjTypDoc
#If APPTYPE = 0 Then        ' APPTYPE=Access
    Case appObjTypAccTbl
        Select Case strObjExtn    ' ���������� �� ���������� �����
        Case c_strObjExtXml:                ImportXML DataSource:=SourceFile, ImportOptions:=acStructureAndData
        Case c_strObjExtTxt:                DoCmd.TransferText TransferType:=acImportDelim, TableName:=TargetName, FileName:=SourceFile, HasFieldNames:=True
        Case c_strObjExtCsv:                DoCmd.TransferText TransferType:=acImportDelim, TableName:=TargetName, FileName:=SourceFile, HasFieldNames:=True
        Case c_strObjExtUndef:              ImportXML DataSource:=SourceFile, ImportOptions:=acStructureAndData
        Case Else:                          Err.Raise m_errObjectTypeUnknown
        End Select
    Case appObjTypAcclnk:                   p_LinkedRead TargetName, SourceFile    ' ������ �� ����� SourceFile
    Case appObjTypAccMac, appObjTypAccQry:  LoadFromText (ObjectType And &HFF&), TargetName, SourceFile
    Case appObjTypAccFrm, appObjTypAccRep:  LoadFromText (ObjectType And &HFF&), TargetName, SourceFile
'    Case appObjTypAccDap
'    Case appObjTypAccSrv
'    Case appObjTypAccDia
'    Case appObjTypAccPrc
'    Case appObjTypAccFun
#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
'    Case appObjTypXlsDoc
#Else                       ' APPTYPE=��������� ����� ���
#End If                     ' APPTYPE
    Case Else:                              Err.Raise m_errObjectTypeUnknown
    End Select
' ���� ��������������� ���� ��� ������ ��� �������� � ������ �� �������������
' ���� ������� �� ��������� ����� ��� ���������������� ������
    If bolRestored And Not bolClear Then Kill SourceFile
    Result = 0
HandleExit:  p_ObjectRead = Result: Exit Function
HandleError:    'Dim Message As String
    Select Case Err 'Err.Number
    Case 75: Err.Clear: Resume Next     ' Path/File access error
    Case 76: Message = Err.Description ' Path not found
    Case m_errModuleNameWrong: Message = "������� ������ ��� �������!": If Len(TargetName) > 0 Then Message = Message & ": """ & TargetName & """"
    Case m_errObjectActionUndef: Message = "��������� ��� ������ � ������": If Len(TargetName) > 0 Then Message = Message & ": """ & TargetName & """"
    Case m_errObjectCantRemove: Message = "�� ������� ������� ������": If Len(TargetName) > 0 Then Message = Message & ": """ & TargetName & """"
    Case m_errObjectTypeUnknown: Message = "���������� ��������� ������ ������� ����": If Len(TargetName) > 0 Then Message = Message & ": """ & TargetName & """"
' ������� ��������:
    Case m_errWrongVersion: Message = "������ ��������� ������ ����� � �������": If Len(TargetName) > 0 Then Message = Message & ": """ & TargetName & """"
        Result = Err: Err.Clear: Resume HandleExit
    Case m_errDestMissing: Message = "����������� ������ ����������� � ����������": If Len(TargetName) > 0 Then Message = Message & ": """ & TargetName & """"
        Result = Err: Err.Clear: Resume HandleExit
    Case m_errDestExists: Message = "����������� ������ ��� ���������� � ����������": If Len(TargetName) > 0 Then Message = Message & ": """ & TargetName & """"
        Result = Err: Err.Clear: Resume HandleExit
    Case m_errSkippedByUser: Message = "��������� �� ������� ������������": If Len(TargetName) > 0 Then Message = Message & ": """ & TargetName & """"
        Result = Err: Err.Clear: Resume HandleExit
    Case m_errSkippedByList: Message = "������� � ������ ��������": If Len(TargetName) > 0 Then Message = Message & ": """ & TargetName & """"
        Result = Err: Err.Clear: Resume HandleExit
    Case 2285 '���������� "Microsoft Office Access" �� ������� ������� �������� ����.
        Message = Err.Description
        Message = Message & vbCrLf & "�������� ������: " & strObjType & " """ & TargetName & """" & vbCrLf & _
                    "������� � ������� ������ ������ ����������."
        If Not bolRestored Then
            If MsgBox(Message & vbCrLf & vbCrLf & _
                    "�������� ������� ��������� ���� � ������ � ��������� ������� ����������." & vbCrLf & _
                    "�������� ���� �� ����������, ������� ����� ������ ��������?", _
                    vbYesNo Or vbQuestion, c_strProcedure) _
                    = vbYes Then
            ' ������� ������� ����� �� ���������� ������ ����������
                bolRestored = p_ObjectFileClear(SourceFile, TargetName, ObjectType, bolClear)
                If bolRestored Then errCount = errCount + 1: Err.Clear: Resume 0
            End If
        End If: Message = Replace(Message, vbCrLf, " ")
    Case 1004 '??? ' ������: ����������� ������ � ������� Visual Basic �� �������� ����������
            Message = "����������� ������ � ������� Visual Basic �� �������� ����������. ��� ����������� ������������ ����������/�������������� ������� ���������� ���������� ����������: ""������\������\������������\�������� ������ � Visual Basic Project"""
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Result = Err: Err.Clear: Resume HandleExit
End Function
Private Function p_ObjectWrite(SourceName As String, _
    Optional ByVal TargetPath As String, _
    Optional ObjectType As appObjectType = appObjTypUndef, _
    Optional WriteType As ObjectRwType = orwDestMiss, _
    Optional AskBefore As Boolean = True, _
    Optional Message As String) As Long
' ������ ������ �� ����������� ������ � ������� ������ � ��������� ������ (��� ��������)
Const c_strProcedure = "p_ObjectWrite"
' SourceName    - ��� ������� � ������� ������� ������ ���� �������
' TargetPath    - ���� ��� ���������� �������
' ObjectType    - ��� ������������ �������
' WriteType     - ��� �������� ����������
' AskBefore     - ����� ����������� ������ ���������� ������������� ������������/��������� ��� �������������
' Message       - ���������� ��������� �� ������ �������� (��� ������ ���������� ��������)
' ���������� ��������� ��������:   0 - ��������, ����� ��� ������
Const cstrMsgYesNo = """��""  - ���� ������ ����� �������� ���� ���������, " & vbCrLf & _
                     """���"" - ���� ������ ���������� ��������"
Const cstrMsgCancel = """������"" - ���� ������ ����� ������ � ��� ����������� �������� ����������� ��� ��������������."
Dim Result As Long: Result = 1
Dim bolClear As Boolean:  bolClear = False ' ���� ������������� ������� ������ ����� ����������� �� ���������� ������� �������������� �� ��������� ��������
Dim errCount As Integer    ' ������� ������
    On Error GoTo HandleError
'    DoCmd.SetWarnings False
'Dim SourceName As String, strObjName As String ', strFileExtn As String
Dim strObjName As String, lngObjType As appObjectType
Dim strObjType As String, strObjFile As String, strModName As String
Dim bolProceed As Boolean       ' ���� ������������� ������ ������� �� �����
Dim intCmp As Integer
' ��������� ������� ������� � �������
Dim bolSrcExist As Boolean:  bolSrcExist = p_ObjectInfo(SourceName, ObjectType, , strObjType, strObjFile, , strModName): If Not bolSrcExist Then Err.Raise 76
' ��������� - ������� ���� � ����� ��� ��� �����
' � ������� ����� �� ���������� ����
Dim bolDestExist As Boolean
    With oFso
        If Len(TargetPath) = 0 Then
    ' �� ������ - ����� ���� �������, ���� �������� ��� �����
            TargetPath = .BuildPath(CurrentProject.path, c_strSrcPath)
            TargetPath = .BuildPath(TargetPath, strObjFile)
            bolDestExist = .FileExists(TargetPath)
        ElseIf .FolderExists(TargetPath) Then
    ' ����� ����, ���� �������� ��� �����
            TargetPath = .BuildPath(TargetPath, strObjFile)
            bolDestExist = .FileExists(TargetPath)
        ElseIf Right(TargetPath, 1) = "\" Then
    ' ����� �������������� ����, ���� �������� ��� �����
            If Not .FolderExists(TargetPath) Then Call .CreateFolder(TargetPath) 'Then Err.Raise 76 ' Path not Found
            TargetPath = .BuildPath(TargetPath, strObjFile)
        ElseIf .GetExtensionName(TargetPath) > 0 Then
    ' ������ ��� ����� ' !!! ����� - ��������� ��� ��� ���� �� ������� ���������� � �����
        Else: Err.Raise 76
        End If
    End With
' ������ ��� �� �����
    strObjName = TargetPath: bolDestExist = p_ObjectInfo(strObjName, lngObjType) ': If Not bolDestExist Then Err.Raise 76
'!!! ���� ��� ����������� �� ����� � ���������� �� ��������� - ���������� �����-�� ������
    If bolDestExist And (ObjectType <> lngObjType) Then Err.Raise m_errObjectActionUndef
' ��������� ����������� ������ ������ � ������ ��
    If (ObjectType And &HFFFF&) > 0 Then
Dim strSrcVer As String, datSrcDate As Date ', strFilDesc As String
Dim strDestVer As String, datDestDate As Date ', strModDesc As String
' ������ ���� ��� ����� � ��������� �� ���� ������ � ������, ���� � ��.
' ToDo: ��������� �� ����� �������� ��� ������� TargetPath
'Stop
        If bolDestExist Then Call ModuleInfoFromFile(TargetPath, strDestVer, datDestDate)
' ������ ����� ������������� ������ � ��������� �� ���� ������ � ������
        If bolSrcExist And Len(strModName) > 0 Then Call ModuleInfo(strModName, strSrcVer, datSrcDate)
    ' ���������� ������
        If bolSrcExist And bolDestExist Then intCmp = VersionCmp(strSrcVer, strDestVer)
    End If
' ��������� ������������� ���������� �������� ����� �������� ������
    Message = strObjType & " """ & SourceName & """"
    Select Case WriteType
    Case orwAlways                  ' ��������� ������ (�����)
        Message = Message & " ������ ���������� � ��������� �����."
    Case orwSrcNewerOrDestMissing   ' ��������� ���� ������ ����� ��� ���� �����������
        Message = Message & " ���������� ��������� �����."
        If (intCmp <> 1) And bolDestExist Then Err.Raise m_errWrongVersion
    Case orwSrcNewer                ' ��������� ���� ������ ����� � ���� ����������
        Message = Message & " ���������� ���������� ������ �� �������."
        If Not (intCmp = 1) Then Err.Raise m_errWrongVersion
        If Not bolDestExist Then Err.Raise m_errDestMissing
    Case orwDestMiss                ' ��������� ���� ���� �����������
        Message = Message & " ���������� � ����� ����������� �� �������."
        If bolDestExist Then Err.Raise m_errDestExists
    Case orwSrcOlder                ' ��������� ���� ������ ������ � ���� ����������
        Message = Message & " ������ ���������� ������� �� �������."
        If Not (intCmp = -1) Then Err.Raise m_errWrongVersion
        If Not bolDestExist Then Err.Raise m_errDestMissing
    Case Else: Err.Raise m_errObjectActionUndef ' �������������� ������� - ��������� ��� ������ � ������
    End Select
' ��������� � ����� ���������� � �������
    If Len(strDestVer) > 0 Or datDestDate > 0 Then
        Message = Message & vbCrLf & "���������� ������"
        If Len(strDestVer) > 0 Then Message = Message & ": " & strDestVer
        If datDestDate > 0 Then Message = Message & " �� " & datDestDate
    End If
    If Len(strSrcVer) > 0 Or datSrcDate > 0 Then
        Message = Message & vbCrLf & "����������� ������"
        If Len(strSrcVer) > 0 Then Message = Message & ": " & strSrcVer
        If datSrcDate > 0 Then Message = Message & " �� " & datSrcDate
    End If
'' !!! ���������� ������� ����������� ��� ������ �������� ������� ��� ��������� ���������� ������������ ����
'    If p_IsSkippedObject(SourceObject) Then Err.Raise m_errSkippedByList
' �������� ���������� ������������ �� ���������� ������
    If AskBefore Then ' intAsk <> vbCancel Then
        Select Case MsgBox(Message & vbCrLf & vbCrLf & cstrMsgYesNo & ", " & vbCrLf & cstrMsgCancel, vbYesNoCancel Or vbInformation, "���������� �������")
        Case vbYes                                      ' �������� ������ �������
        Case vbCancel:  AskBefore = False               ' �������� ������� � �����������
        Case Else:      Err.Raise m_errSkippedByUser    ' ���������� �������
        End Select
    End If
' ���� ���� ���������� ������� ���
    If bolDestExist Then Kill TargetPath
' � ����������� �� ���� ��������� �������
    Select Case ObjectType
    Case appObjTypMod, appObjTypBas, appObjTypCls
                                            VBE.ActiveVBProject.VBComponents(SourceName).Export TargetPath
'    Case appObjTypMsf
'    Case appObjTypAxd
'    Case appObjTypDoc
#If APPTYPE = 0 Then        ' APPTYPE=Access
    Case appObjTypAccTbl:                   ExportXML DataSource:=SourceName, DataTarget:=TargetPath, ObjectType:=acExportTable, OtherFlags:=acEmbedSchema
    Case appObjTypAcclnk:                   p_LinkedWrite SourceName, TargetPath
    Case appObjTypAccMac, appObjTypAccQry:  SaveAsText (ObjectType And &HFF&), SourceName, TargetPath
    Case appObjTypAccFrm, appObjTypAccRep:  SaveAsText (ObjectType And &HFF&), SourceName, TargetPath
                                            If bolClear Then Call p_ObjectFileClear(TargetPath, ObjectType:=ObjectType, ReplaceOriginal:=True)
'    Case appObjTypAccDap
'    Case appObjTypAccSrv
'    Case appObjTypAccDia
'    Case appObjTypAccPrc
'    Case appObjTypAccFun
#ElseIf APPTYPE = 1 Then    ' APPTYPE=Excel
'    Case appObjTypXlsDoc
#Else                       ' APPTYPE=��������� ����� ���
#End If                     ' APPTYPE
    Case Else:                              Err.Raise m_errObjectTypeUnknown
    End Select
    Result = 0
HandleExit:  p_ObjectWrite = Result: Exit Function
HandleError:    'Dim Message As String
    Select Case Err 'Err.Number
    Case 75: Err.Clear: Resume Next     ' Path/File access error
    Case 76: Message = Err.Description ' Path not found
    Case m_errModuleNameWrong: Message = "������� ������ ��� �������!": If Len(SourceName) > 0 Then Message = Message & ": """ & SourceName & """"
    Case m_errObjectActionUndef: Message = "��������� ��� ������ � ������": If Len(SourceName) > 0 Then Message = Message & ": """ & SourceName & """"
    Case m_errObjectCantRemove: Message = "�� ������� ������� ������": If Len(SourceName) > 0 Then Message = Message & ": """ & SourceName & """"
    Case m_errObjectTypeUnknown: Message = "���������� ��������� ������ ������� ����": If Len(SourceName) > 0 Then Message = Message & ": """ & SourceName & """"
' ������� ��������:
    Case m_errWrongVersion: Message = "������ ��������� ������ ����� � �������": If Len(SourceName) > 0 Then Message = Message & ": """ & SourceName & """"
        Result = Err: Err.Clear: Resume HandleExit
    Case m_errDestMissing: Message = "���� ������������ ������� �����������": If Len(SourceName) > 0 Then Message = Message & ": """ & SourceName & """"
        Result = Err: Err.Clear: Resume HandleExit
    Case m_errDestExists: Message = "���� ������������ ������� ��� ����������": If Len(SourceName) > 0 Then Message = Message & ": """ & SourceName & """"
        Result = Err: Err.Clear: Resume HandleExit
    Case m_errSkippedByUser: Message = "��������� �� ������� ������������": If Len(SourceName) > 0 Then Message = Message & ": """ & SourceName & """"
        Result = Err: Err.Clear: Resume HandleExit
    Case m_errSkippedByList: Message = "������� � ������ ��������": If Len(SourceName) > 0 Then Message = Message & ": """ & SourceName & """"
        Result = Err: Err.Clear: Resume HandleExit
 '
    Case 1004 '??? ' ������: ����������� ������ � ������� Visual Basic �� �������� ����������
                Message = "����������� ������ � ������� Visual Basic �� �������� ����������. ��� ����������� ������������ ����������/�������������� ������� ���������� ���������� ����������: ""������\������\������������\�������� ������ � Visual Basic Project"""
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Result = Err: Err.Clear: Resume HandleExit
End Function
Private Function p_ObjectFileClear( _
    SourceFile As String, _
    Optional TargetName As String, _
    Optional ObjectType As appObjectType = appObjTypUndef, _
    Optional ReplaceOriginal As Boolean = False) As Boolean
' ������� ����� ���������� �� ���������� �������������� ��� ��������
Dim Result As Boolean: Result = False
' based on: https://github.com/bkidwell/msaccess-vcs-integration/issues/12?utm_source=pocket_mylist#issuecomment-29564498
    On Error GoTo HandleError
    'strReadLine = oFso.ReadLine
    If Not oFso.FileExists(SourceFile) Then Err.Raise 76 ' Path not Found
Dim TargetFile As String: TargetFile = VBA.Environ$("Temp") & "\~@#" & Format$(Now, "yyyymmddhhnnss")

Dim aSkipStrings(), aSkipBlocks() 'As String
Dim sBlockBeg As String, sBlockEnd As String
    Select Case ObjectType
    Case appObjTypAccFrm, appObjTypAccRep
' ��� ���� � ������� ��� ������ ����������
'' ������:
    aSkipStrings = Array( _
        "Checksum", _
        "FilterOnLoad", _
        "AllowLayoutView", _
        "NoSaveCTIWhenDisabled", _
        "Overlaps", "BorderLineStyle", _
        "WebImagePaddingLeft", "WebImagePaddingTop", "WebImagePaddingRight", "WebImagePaddingBottom")
'' �����:
    sBlockBeg = "Begin": sBlockEnd = "End"
    aSkipBlocks = Array( _
        "PrtDevNames", "PrtDevNamesW", _
        "PrtDevMode", "PrtDevModeW", _
        "NameMap", "NameMapW") ', _
        "GUID") ', _
        "PrtMip", "PrtMipW") ', _
        "dbLongBinary ""DOL""", _

    Case Else: Err.Raise vbObjectError + 2048 ' ���������������� ������
    End Select
Dim strReadLine As String

'1. ������ ���� ���������
Dim iFileIn As Integer:     iFileIn = FreeFile:  Open SourceFile For Input As iFileIn
Dim iFileOut As Integer:    iFileOut = FreeFile: Open TargetFile For Output Access Write As #iFileOut

Dim strLine As String, strTemp As String, varTemp '
Dim bolEndOfBlock As Boolean
Dim LineType As m_CodeLineType
    Do Until EOF(iFileIn) ' Or LineType = m_CodeProc
    ' ������ ��������� ���� �� ��������� �����
        Line Input #iFileIn, strLine: If LineType = m_CodeProc Then GoTo HandleOutput
'Stop
' ��� Access !!!
    ' ��������� ������ �� ����� ��������� ������� (���������� ������ ���������)
        If InStrRegEx(1, strLine, c_strCodeProcBeg) > 0 Then LineType = m_CodeProc: GoTo HandleOutput
'2. ���������� ������������ ���������
    ' ��������� �� ������ �������� ����� � ���������� ���� �����
        For Each varTemp In aSkipStrings
            strTemp = "^\s*" & varTemp & "\s*=\s*"
            If InStrRegEx(1, strLine, strTemp) > 0 Then GoTo HandleNextLine
        Next varTemp
    ' ��������� �� ������ �������� ������
        For Each varTemp In aSkipBlocks
        ' ���� ����������� ���� ������
            strTemp = "^\s*" & varTemp & "\s*=\s*" & sBlockBeg
            If InStrRegEx(1, strLine, strTemp$) > 0 Then
                strTemp = "^\s*" & sBlockEnd
                Do
        ' ������ ��������� ���� �� ������ �� ����������� ���� ������ ��� �� ��������� �����
                    Line Input #iFileIn, strLine
                    If InStrRegEx(1, strLine, strTemp) > 0 Then Exit Do ' �������� ����
                    If EOF(iFileIn) Then Err.Raise vbObjectError + 2049 ' ���������� ����
                    If InStrRegEx(1, strLine, c_strCodeProcBeg) > 0 Then Err.Raise vbObjectError + 2049 ' ���������� ����
                Loop
        ' � ���������� ���� ���� ���� �����
                GoTo HandleNextLine
            End If
        Next varTemp
'3. ��������� ��������� � Temp
HandleOutput:
        Print #iFileOut, strLine
HandleNextLine:
    Loop
    Close #iFileIn: Close #iFileOut
' ��������� ���������
    If ReplaceOriginal Then
' �������� �������� ���� ��������������� � ���������� ���� ��������� �����
        Kill PathName:=SourceFile
        FileCopy Source:=TargetFile, Destination:=SourceFile
        Kill PathName:=TargetFile
    Else
' ���������� ���� � ���������������� �����
        SourceFile = TargetFile
    End If
    Result = True
HandleExit:  p_ObjectFileClear = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Private Sub ModuleHeaderCreate( _
    ModuleName As String)
Const c_strProcedure = "ModuleHeaderCreate"
Dim Result As Boolean ': Result = False
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
' ������ �����������
Dim CodeLine As Long, PrevLine As Long
    CodeLine = CodeLineFind(ModuleName, c_strPrefModLine)
    If CodeLine = 0 Then objModule.InsertLines 1, c_strPrefModLine: PrevLine = CodeLine + 1
' ��� ������
    CodeLine = CodeLineFind(ModuleName, c_strPrefModName, PrevLine)
    If CodeLine = 0 Then CodeLine = PrevLine + 1 Else objModule.DeleteLines CodeLine, 1
    objModule.InsertLines CodeLine, c_strPrefModName & VBA.Chr(34) & ModuleName & VBA.Chr(34): PrevLine = CodeLine
' ������ �����������
    objModule.InsertLines CodeLine, c_strPrefModLine: PrevLine = CodeLine
' ��������
    CodeLine = CodeLineFind(ModuleName, c_strPrefModDesc, PrevLine)
    If CodeLine = 0 Then CodeLine = PrevLine + 1 Else objModule.DeleteLines CodeLine, 1
    objModule.InsertLines CodeLine, c_strPrefModDesc: PrevLine = CodeLine
' ������
    CodeLine = CodeLineFind(ModuleName, c_strPrefModVers, PrevLine)
    If CodeLine = 0 Then CodeLine = PrevLine + 1 Else objModule.DeleteLines CodeLine, 1
    objModule.InsertLines CodeLine, c_strPrefModVers: PrevLine = CodeLine
' ����
    CodeLine = CodeLineFind(ModuleName, c_strPrefModDate, PrevLine)
    If CodeLine = 0 Then CodeLine = PrevLine + 1 Else objModule.DeleteLines CodeLine, 1
    objModule.InsertLines CodeLine, c_strPrefModDate: PrevLine = CodeLine
' �����
    CodeLine = CodeLineFind(ModuleName, c_strPrefModAuth, PrevLine)
    If CodeLine = 0 Then CodeLine = PrevLine + 1 Else objModule.DeleteLines CodeLine, 1
    objModule.InsertLines CodeLine, c_strPrefModAuth: PrevLine = CodeLine
' ����������
    CodeLine = CodeLineFind(ModuleName, c_strPrefModComm, PrevLine)
    If CodeLine = 0 Then CodeLine = PrevLine + 1 Else objModule.DeleteLines CodeLine, 1
    objModule.InsertLines CodeLine, c_strPrefModComm: PrevLine = CodeLine + 1
' ������ �����������
    objModule.InsertLines CodeLine, c_strPrefModLine
    
    Result = True
HandleExit:     Exit Sub
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "������� ������ ��� �������!"
    Case m_errModuleIsActive: Debug.Print "���������� �������� �������� ������: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "������: """ & ModuleName & """ �� ������!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Sub
Public Function ModuleInfo( _
    ModuleName As String, _
    Optional ModVers, Optional ModDate, _
    Optional ModDesc, Optional ModAuth, _
    Optional ModComm, Optional ModHist) As Boolean
' ������ ���������� � ������ �� ������ ������
Const c_strProcedure = "ModuleInfo"
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
Dim CodeLine As Long
Dim LineType As m_CodeLineType, LineBreak As Boolean
' LineType - ���� ����������� ������ ��������
' LineBreak - ������� ����������� ������ ������ (������ ������������� �������� ��������)
    LineType = m_CodeHead ' ������ ����� ���������� � ����

Dim strLine As String, strResult As String
    For CodeLine = 1 To objModule.CountOfLines '.CountOfDeclarationLines - ���������� ����������� � ����� ��������� ������ ������ �� �������� ��� ��� ����������� ����� ���������
        If LineType = m_CodeProc Then Exit For
    ' ������ ��������� ���� �� ��������� ����� ��� �� ������ �� ���������� ������ ���������
        strLine = objModule.Lines(CodeLine, 1)
    ' ��������� � ����� ����� ������ ���������
    ' � � ����������� �� ����������� ������ ���������� ��������
        ' � ��������� ���� � ������� ������� ���� �������� ������������ ��������� �����
        ' ������� ���������� ��� ������ �� ������ ������ ������
        ' � ������� ������ �������� � ������ ������
        If LineType = m_CodeNone Then
            If VBA.Trim$(strLine) = c_strCodeHeadBeg Then LineType = m_CodeHead
            GoTo HandleNextLine
        End If
        ' ���������� ��� ������
        strLine = p_CodeLineGet(strLine, LineType, LineBreak, vbCrLf)
        ' ��������� ������ ����������
        strResult = strResult & strLine
        ' ������� ������ ���������� - ������� �����
        strResult = Replace(strResult, c_strBrokenQuotes, vbNullString)   ' ���������� ����������� ��������� ������
HandleRead:
        If LineType <> m_CodeHead Then
        ' ���������� ���������
            If LineBreak Then GoTo HandleNextLine
            If Not IsMissing(ModVers) And LineType = m_CodeVers And Len(strResult) > 0 Then ModVers = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModDate) And LineType = m_CodeDate And Len(strResult) > 0 Then ModDate = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModDesc) And LineType = m_CodeDesc And Len(strResult) > 0 Then ModDesc = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModAuth) And LineType = m_CodeAuth And Len(strResult) > 0 Then ModAuth = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModComm) And LineType = m_CodeComm And Len(strResult) > 0 Then ModComm = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModHist) And LineType = m_CodeHist And Len(strResult) > 0 Then ModHist = ModHist & vbCrLf & strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
        End If
        LineType = m_CodeHead
        strResult = vbNullString
HandleNextLine:
    Next
' ��������� ��������� ����������
    If Not IsMissing(ModVers) Then If Len(ModVers) = 0 Then ModVers = cEmptyVers
    If Not IsMissing(ModDate) Then If Len(ModDate) = 0 Then ModDate = cEmptyDate
    If Not IsMissing(ModHist) Then If Left(ModHist, Len(vbCrLf)) = vbCrLf Then ModHist = Mid(ModHist, Len(vbCrLf) + 1)
HandleExit:  ModuleInfo = Result: Exit Function
HandleError: Result = False
Dim Message As String
    Select Case Err.Number
    Case m_errModuleNameWrong: Message = "�� ������ ��� ������!"
    Case m_errModuleIsActive: Message = "���������� �������� �������� ������!"
    Case m_errModuleDontFind: Message = "������ �� ������!"
    Case Else: Message = Err.Description
    End Select
    If Len(ModuleName) > 0 Then Message = Message & "; ModuleName=""" & ModuleName & """ "
    'If Len(strResult) > 0 Then Message = Message & "; strResult=""" & strResult & """ "
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
End Function
Public Function ModuleVersGet(ModuleName As String) As String: Call ModuleInfo(ModuleName, ModVers:=ModuleVersGet): End Function
Public Function ModuleDateGet(ModuleName As String) As String: Call ModuleInfo(ModuleName, ModDate:=ModuleDateGet): End Function
Public Function ModuleAuthGet(ModuleName As String) As String: Call ModuleInfo(ModuleName, ModAuth:=ModuleAuthGet): End Function
Public Function ModuleDescGet(ModuleName As String) As String: Call ModuleInfo(ModuleName, ModDesc:=ModuleDescGet): End Function
Public Function ModuleCommGet(ModuleName As String) As String: Call ModuleInfo(ModuleName, ModComm:=ModuleCommGet): End Function
Public Function ModuleHistGet(ModuleName As String): Call ModuleInfo(ModuleName, ModHist:=ModuleHistGet): End Function

Public Function ModuleInfoFromFile( _
    FilePath As String, _
    Optional ModVers, Optional ModDate, _
    Optional ModDesc, Optional ModAuth, _
    Optional ModComm, Optional ModHist) As Boolean
' ������ ���������� � ������ �� ������ ����� ������
Const c_strProcedure = "ModuleInfoFromFile"
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
Dim CodeLine As Long
Dim strPath As String:  strPath = VBA.Trim$(FilePath)
' ToDo: ��������� �� ����� �������� ��� ������� TargetPath
    ' ��������� ������� ���������� ����/�����
    If Not oFso.FileExists(strPath) Then Err.Raise 76 ' Path not Found
' ��������� ��� ����� (�� ����������)
Dim oType As msys_ObjectType, cType As vbext_ComponentType
Dim strExtn As String: strExtn = VBA.LCase$(oFso.GetExtensionName(strPath))
    Select Case strExtn
    Case c_strObjExtBas: oType = msys_ObjectModule  ' ����������� ������
    Case c_strObjExtCls: oType = msys_ObjectModule  ' ������ ������
    Case c_strObjExtDoc: oType = msys_ObjectModule  ' ������ ������ ��������� Access (Form ��� Report)
    Case c_strObjExtFrm: oType = msys_ObjectForm    ' ����� Access (������� ������)
    Case c_strObjExtRep: oType = msys_ObjectReport  ' ����� Access (������� ������)
    Case c_strObjExtUndef: oType = msys_ObjectModule  ' ������ ���������� ����� ������ ������� Access
    'Case c_strObjExtXml: oType = msys_ObjectTable   ' ��������� ������� � XML
    'Case c_strObjExtTxt: oType = msys_ObjectTable   ' ��������� ������� � TXT
    'Case c_strObjExtCsv: oType = msys_ObjectTable   ' ��������� ������� � CSV (����� � �������������)
    'Case c_strObjExtLnk: oType = msys_ObjectLinked  ' ��������� �������
    'Case c_strObjExtQry: oType = msys_ObjectQuery   ' ������ Access
    'Case c_strObjExtMac: oType = msys_ObjectMacro   ' ������ Access
    Case Else: Err.Raise vbObjectError + 512
    End Select
Dim LineType As m_CodeLineType, LineBreak As Boolean
' LineType - ���� ����������� ������ ��������
' LineBreak - ������� ����������� ������ ������ (������ ������������� �������� ��������)
    Select Case oType
    Case msys_ObjectModule:                     LineType = m_CodeHead ' ������ ����� ���������� � ����
    Case msys_ObjectForm, msys_ObjectReport:    LineType = m_CodeNone ' �����/����� ������� �������� ���������� � ������������ ���������
    Case Else: Err.Raise vbObjectError + 512              ' ������ �������� �� �������������� �.�. �� �������� ����
    End Select

Dim iFile As Integer:   iFile = FreeFile
    Open strPath For Input As #iFile: CodeLine = 0

Dim strLine As String, strResult As String
    Do Until EOF(iFile) Or LineType = m_CodeProc
    ' ������ ��������� ���� �� ��������� ����� ��� �� ������ �� ���������� ������ ���������
        Line Input #iFile, strLine
        CodeLine = CodeLine + 1
    ' ��������� � ����� ����� ������ ���������
    ' � � ����������� �� ����������� ������ ���������� ��������
        ' � ��������� ���� � ������� ������� ���� �������� ������������ ��������� �����
        ' ������� ���������� ��� ������ �� ������ ������ ������
        ' � ������� ������ �������� � ������ ������
        If LineType = m_CodeNone Then
            If VBA.Trim$(strLine) = c_strCodeHeadBeg Then LineType = m_CodeHead
            GoTo HandleNextLine
        End If
        ' ���������� ��� ������
        strLine = p_CodeLineGet(strLine, LineType, LineBreak, vbCrLf)
        ' ��������� ������ ����������
        strResult = strResult & strLine
        ' ������� ������ ���������� - ������� �����
        strResult = Replace(strResult, c_strBrokenQuotes, vbNullString)   ' ���������� ����������� ��������� ������
HandleRead:
        If LineType <> m_CodeHead Then
        ' ���������� ���������
            If LineBreak Then GoTo HandleNextLine
            If Not IsMissing(ModVers) And LineType = m_CodeVers Then ModVers = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModDate) And LineType = m_CodeDate Then ModDate = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModDesc) And LineType = m_CodeDesc Then ModDesc = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModAuth) And LineType = m_CodeAuth Then ModAuth = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModComm) And LineType = m_CodeComm Then ModComm = strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
            If Not IsMissing(ModHist) And LineType = m_CodeHist Then ModHist = ModHist & vbCrLf & strResult: Result = True ': LineType = m_CodeHead: strResult = vbNullString: GoTo HandleNextLine
        End If
        LineType = m_CodeHead
        strResult = vbNullString
HandleNextLine:
    Loop
    Close #iFile
' ��������� ��������� ����������
    If Not IsMissing(ModVers) Then If Len(ModVers) = 0 Then ModVers = cEmptyVers
    If Not IsMissing(ModDate) Then If Len(ModDate) = 0 Then ModDate = cEmptyDate
    If Not IsMissing(ModHist) Then If Left(ModHist, Len(vbCrLf)) = vbCrLf Then ModHist = Mid(ModHist, Len(vbCrLf) + 1)
HandleExit:   ModuleInfoFromFile = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
Public Function ModuleVersGetFromFile(ModulePath As String) As String: Call ModuleInfoFromFile(ModulePath, ModVers:=ModuleVersGetFromFile): End Function
Public Function ModuleDateGetFromFile(ModulePath As String) As String: Call ModuleInfoFromFile(ModulePath, ModDate:=ModuleDateGetFromFile): End Function
Public Function ModuleAuthGetFromFile(ModulePath As String) As String: Call ModuleInfoFromFile(ModulePath, ModAuth:=ModuleAuthGetFromFile): End Function
Public Function ModuleDescGetFromFile(ModulePath As String) As String: Call ModuleInfoFromFile(ModulePath, ModDesc:=ModuleDescGetFromFile): End Function
Public Function ModuleCommGetFromFile(ModulePath As String) As String: Call ModuleInfoFromFile(ModulePath, ModComm:=ModuleCommGetFromFile): End Function
Public Function ModuleHistGetFromFile(ModulePath As String): Call ModuleInfoFromFile(ModulePath, ModHist:=ModuleHistGetFromFile): End Function

Public Function ModuleVersSet( _
    ModuleName As String, _
    VersionString As String, _
    Optional CodeLine As Long)
Const c_strProcedure = "ModuleVersSet"
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
    CodeLine = CodeLineFind(ModuleName, c_strPrefModVers, CodeLine) ' ������ �� ������ ������� �������
    If CodeLine > 0 Then
        objModule.DeleteLines CodeLine, 1:
    Else
        CodeLine = 1 ' ��� ������ ������� � ������ ������
    End If
    objModule.InsertLines CodeLine, c_strPrefModVers & " " & VersionString
    Result = True
HandleExit:     ModuleVersSet = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "������� ������ ��� �������!"
    Case m_errModuleIsActive: Debug.Print "���������� �������� �������� ������: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "������: """ & ModuleName & """ �� ������!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function ModuleDateSet( _
    ModuleName As String, _
    DateString As String, _
    Optional CodeLine As Long)
Const c_strProcedure = "ModuleDateSet"
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
    CodeLine = CodeLineFind(ModuleName, c_strPrefModDate, CodeLine) ' ������ �� ������ ������� �������
    If CodeLine > 0 Then
        objModule.DeleteLines CodeLine, 1:
    Else
        CodeLine = 1 ' ��� ���� ������� � ������ ������
    End If
    objModule.InsertLines CodeLine, c_strPrefModDate & " " & DateString
    Result = True
HandleExit:     ModuleDateSet = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "������� ������ ��� �������!"
    Case m_errModuleIsActive: Debug.Print "���������� �������� �������� ������: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "������: """ & ModuleName & """ �� ������!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function ModuleDescSet( _
    ModuleName As String, _
    DescString As String, _
    Optional CodeLine As Long)
Const c_strProcedure = "ModuleDescSet"
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
    CodeLine = CodeLineFind(ModuleName, c_strPrefModDesc, CodeLine) ' ������ �� ������ ������� �������
    If CodeLine > 0 Then
        objModule.DeleteLines CodeLine, 1:
    Else
        CodeLine = 1 ' ��� �������� ������� � ������ ������
    End If
    objModule.InsertLines CodeLine, c_strPrefModDesc & " " & DescString
    Result = True
HandleExit:     ModuleDescSet = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "������� ������ ��� �������!"
    Case m_errModuleIsActive: Debug.Print "���������� �������� �������� ������: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "������: """ & ModuleName & """ �� ������!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function ModuleCommSet( _
    ModuleName As String, _
    CommString As String, _
    Optional Replace As Boolean = False, _
    Optional CodeLine As Long)
' ���������� � ��������� ������ �����������
Const c_strProcedure = "ModuleCommSet"
' ���� Replace = True - �������� �������, ����� ��������� ��� ���������� �����
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
Dim NumLines As Long: NumLines = 1
    CodeLine = CodeLineFind(ModuleName, c_strPrefModComm, CodeLine) ' ������ �� ������ ������� �������
    If CodeLine > 0 Then
        NumLines = CodeLineNext(ModuleName, CodeLine) - CodeLine
        If Replace Then objModule.DeleteLines CodeLine, NumLines: NumLines = 1
    Else
        CodeLine = 1: Replace = True ' ��� ����������� ������� � ������ ������
    End If
    With objModule
        If Replace Then .InsertLines CodeLine, c_strPrefModComm
        CodeLine = CodeLine + NumLines
        .InsertLines CodeLine, "'" & CommString
    End With
    Result = True
HandleExit:     ModuleCommSet = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "������� ������ ��� �������!"
    Case m_errModuleIsActive: Debug.Print "���������� �������� �������� ������: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "������: """ & ModuleName & """ �� ������!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function ModuleAuthSet(ModuleName As String, AuthString As String, Optional CodeLine As Long)
' ������������� ������ ������ ������
Const c_strProcedure = "ModuleAuthSet"
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
    CodeLine = CodeLineFind(ModuleName, c_strPrefModAuth, CodeLine) ' ������ �� ������ ������� �������
    If CodeLine > 0 Then
        objModule.DeleteLines CodeLine, 1:
    Else
        CodeLine = 1 ' ��� ����������� ������� � ������ ������
    End If
    objModule.InsertLines CodeLine, c_strPrefModAuth & " " & AuthString
    Result = True
HandleExit:     ModuleAuthSet = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "������� ������ ��� �������!"
    Case m_errModuleIsActive: Debug.Print "���������� �������� �������� ������: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "������: """ & ModuleName & """ �� ������!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function ModuleDebugSet( _
    ModuleName As String, _
    DEBUGGING As Boolean, _
    Optional CodeLine As Long)
'#Const DEBUGGING = False
Const c_strProcedure = "ModuleDebugSet"
Dim Result As Boolean: Result = False
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
    CodeLine = CodeLineFind(ModuleName, c_strPrefModDebg, CodeLine) ' ������ �� ������ ������� �������
    If CodeLine > 0 Then
        objModule.DeleteLines CodeLine, 1:
    Else
        CodeLine = 1 ' ��� ����������� ������� � ������ ������
    End If
    objModule.InsertLines CodeLine, c_strPrefModDebg & "=" & DEBUGGING
    Result = True
HandleExit:     ModuleDebugSet = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "������� ������ ��� �������!"
    Case m_errModuleIsActive: Debug.Print "���������� �������� �������� ������: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "������: """ & ModuleName & """ �� ������!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function

'======================
Public Function CodeLineNext(ModuleName As String, Optional BegLine As Long = 1) As Long
' ���������� ����� ��������� ������ � ������ � ������ ������ ��������
Const c_strProcedure = "CodeLineNext"
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
    CodeLineNext = p_CodeLineNext(objModule, BegLine)
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "������� ������ ��� �������!"
    Case m_errModuleIsActive: Debug.Print "���������� �������� �������� ������: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "������: """ & ModuleName & """ �� ������!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function CodeLineGet(ModuleName As String, Optional BegLine As Long = 1, Optional NumOfLines As Long) As String
' ���������� ������ ������ � ������ �� ������� ������ ��������
Const c_strProcedure = "CodeLineGet"
' ����� NumOfLines �������� ���������� ����� ������� ���� ����������
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
    CodeLineGet = p_CodeLineFull(objModule, BegLine, NumOfLines)
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "������� ������ ��� �������!"
    Case m_errModuleIsActive: Debug.Print "���������� �������� �������� ������: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "������: """ & ModuleName & """ �� ������!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Public Function CodeLineFind(ModuleName As String, FindString As String, Optional BegLine As Long = 1, Optional FindNum As Integer = 0) As Long
' ���� ������ ���� � ������ ModuleName ������������ � FindString ���������� ����� ������
Const c_strProcedure = "CodeLineFind"
    On Error GoTo HandleError
    If ModuleName = vbNullString Then Err.Raise m_errModuleNameWrong
'    If ModuleName = VBE.ActiveCodePane.CodeModule Then Err.Raise m_errModuleIsActive
Dim objModule As Object: If Not ModuleExists(ModuleName, objModule) Then Err.Raise m_errModuleDontFind
    CodeLineFind = p_CodeLineFind(objModule, FindString, BegLine, FindNum)
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case m_errModuleNameWrong: Message = "������� ������ ��� �������!"
    Case m_errModuleIsActive: Debug.Print "���������� �������� �������� ������: """ & ModuleName & """!"
    Case m_errModuleDontFind: Debug.Print "������: """ & ModuleName & """ �� ������!"
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_CodeLineNext(objModule As Object, Optional BegLine As Long = 1) As Long
' ���������� ����� ��������� ������ � ������ � ������ ������ ��������
Const c_strProcedure = "p_CodeLineNext"
Dim Result As Long: Result = False
    On Error GoTo HandleError
Dim i As Long: i = BegLine
    With objModule
        Do While VBA.Right$(.Lines(i, 1), Len(c_strHyphen)) = c_strHyphen
            i = i + 1
        Loop
    End With
    Result = i + 1
HandleExit:     p_CodeLineNext = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_CodeLineFull(objModule As Object, _
    Optional BegLine As Long = 1, Optional NumOfLines As Long, _
    Optional TrimSpaces As Boolean, _
    Optional ReplaceHyphensWith As String = c_strSpace, _
    Optional LinePrefToWrap As String) As String
' ���������� ������ ������ � ������, ��������� ����� ������ ����������� �������� ��������
Const c_strProcedure = "p_CodeLineFull"
' ����� NumOfLines �������� ���������� ����� ������� ���� ����������
Dim Result As String
    On Error GoTo HandleError
    NumOfLines = 0
Dim tmpString As String
Dim LineBreak As Boolean, LineType As m_CodeLineType
    With objModule
        Do
            tmpString = VBA.Trim$(.Lines(BegLine + NumOfLines, 1))
            Result = Result & p_CodeLineGet(tmpString, LineType, LineBreak, ReplaceHyphensWith)
            NumOfLines = NumOfLines + 1
        Loop While LineBreak
    End With
' ����� ���������� ������� ������ - ������� �����
    Result = Replace(Result, c_strBrokenQuotes, vbNullString)   ' ���������� ����������� ��������� ������
HandleExit:  p_CodeLineFull = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_CodeLineGet(CodeLine As String, _
    Optional LineType As m_CodeLineType, _
    Optional LineBreak As Boolean, _
    Optional ReplaceHyphensWith As String = c_strSpace, _
    Optional TrimPrefix As Boolean = True, _
    Optional TrimSpaces As Boolean) As String
' ���������� ��������� ������ ������ ����, ������� �������� � ����� ��������, ���������� ����� ���� ������
Const c_strProcedure = "p_CodeLineGet"
' CodeLine - �������������� ������
' LineType - ��� ������ ���� (������������� �� �����������)
' LineBreak - ������� ����������� ������ ������ (������ ������������� �������� ��������)
' ReplaceHyphensWith - ������ ��� ������ �������� �������� ������
' TrimPrefix - ������� ������������� �������� ��������� �����
' TrimSpaces - ������� ������������� ������� ������� �������/����� ������
Dim Result As String: Result = CodeLine
    On Error GoTo HandleError
Static arrPrefs(): arrPrefs = Array(m_CodeVers, c_strPrefModVers, _
                                    m_CodeDate, c_strPrefModDate, _
                                    m_CodeDesc, c_strPrefModDesc, _
                                    m_CodeAuth, c_strPrefModAuth, _
                                    m_CodeComm, c_strPrefModComm, _
                                    m_CodeHist, c_strPrefModHist, _
                                    m_CodeProc, c_strCodeProcBeg)
Dim i As Long
Dim strPref As String, lngLineType As Long
' ���� ��� ����������� ������ - ������� ������ ������� ������� ������
    If LineBreak Then
        Select Case LineType
        Case m_CodeDesc, m_CodeComm, m_CodeHist: strPref = c_strPrefModNone
        End Select
        GoTo HandleProceed
    End If
' ����� �������� ������ ��� ��� �� ������
    Select Case LineType
    Case m_CodeHead
' ���� �������� ���������� � ��������� ������
        For i = LBound(arrPrefs) To UBound(arrPrefs) Step 2
            lngLineType = arrPrefs(i): strPref = arrPrefs(i + 1)
            Select Case lngLineType
            Case m_CodeProc: If InStrRegEx(1, Result, strPref) > 0 Then LineType = lngLineType: strPref = vbNullString: Exit For
            Case m_CodeHist: If InStrRegEx(1, Result, strPref) > 0 Then LineType = lngLineType: strPref = "'": Exit For
            Case Else:       If VBA.Left$(Result, Len(strPref)) = strPref Then LineType = lngLineType: Exit For
            End Select
            strPref = vbNullString ' ���� �� ����� �������� �������
        Next i
    Case m_CodeProc
' ���� �������� ���������� � ���������
'Stop
    Case Else
    End Select
'Stop
HandleProceed:
' ��������� ������� �������� ������
    LineBreak = (VBA.Right$(Result, Len(c_strHyphen)) = c_strHyphen)
' ���� ������� �������� � ������ ���������� � ���������� �������� - ������� ���
    If TrimPrefix And Len(strPref) > 0 Then Result = VBA.Trim$(VBA.Mid$(Result, Len(strPref) + 1))
' ���� ������� ������� � ������ ����������/������������� �� ������� - ��������
    If TrimSpaces Then Result = Trim$(Result)
' ���� ������ ������������� ��  _ �������� �� ������ ����������� � ���������� ������
    If LineBreak Then Result = VBA.Left$(Result, Len(Result) - Len(c_strHyphen)) & ReplaceHyphensWith
HandleExit:     p_CodeLineGet = Result: Exit Function
HandleError:    Result = vbNullString
    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_CodeLineFind(objModule As Object, FindString As String, Optional BegLine As Long = 1, Optional FindNum As Integer = 0) As Long
' ���� ������ ���� � ������ ModuleName ������������ � FindString ���������� ����� ������
Const c_strProcedure = "p_CodeLineFind"
Dim Result As Long ':Result = 0
    On Error GoTo HandleError
Dim lngBegLine As Long, lngEndLine As Long: lngBegLine = BegLine
Dim lngBegCol As Long, lngEndCol As Long
Dim FindCount As Long ': FindCount = 0
Dim bolIsLoaded As Boolean
    With objModule
        Do ' ������ �� ������
            If .Find(FindString, lngBegLine, lngBegCol, 0, -1) Then
                If lngBegCol = 1 Then   ' ���� � ������ ������
                    Result = lngBegLine
                    FindCount = FindCount + 1
                    Select Case FindNum
                    Case 0:  Exit Do    ' ������ ���������
                    Case -1:            ' ��������� ���������
                    Case Else: If FindCount = FindNum Then Exit Do ' n-��� ���������
                    End Select
                End If
                lngBegLine = lngBegLine + 1 ' ���� �� ����� �� ����� - ��������� �� ��������� ������ � ���������� �����
            Else
                Exit Do
            End If
        Loop
    End With
HandleExit:     p_CodeLineFind = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function

'-------------------------
' ������� ��� ������ � zip ��������
'-------------------------
Public Function ZipCreate( _
    ZipName As String _
    ) As Boolean
' ������� ZIP ����
' ���������� True ���� ������
Dim strZipFileHeader As String
Dim Result As Boolean
    On Error GoTo HandleError
    ' �������� ������� ���������� zip � ������ ����-����� �����
    If VBA.UCase$(oFso.GetExtensionName(ZipName)) <> c_strObjExtZip Then Exit Function
    ' �������� ������� zip ������
    strZipFileHeader = "PK" & VBA.Chr(5) & VBA.Chr(6) & VBA.String$(18, 0)
    oFso.OpenTextFile(ZipName, 2, True).Write strZipFileHeader
    Dim oArch As Object: Set oArch = oApp.Namespace((ZipName))
    ' �������� �������� ������
    Result = Not (oArch Is Nothing)
HandleExit:  ZipCreate = Result: Exit Function
HandleError: Result = False: Resume HandleExit
End Function
Public Function ZipPack( _
    FilePath As String, _
    Optional ZipName As String, _
    Optional DelAfterZip As Boolean = False _
    ) As Boolean
' ��������� ����� � ������
Const c_strProcedure = "ZipPack"
' FileNames - ����� ������������ ������/����� ����� ����� � �������
' FilePath - ���� � ����� ��� ��������� ������������ �����
' ZipName - ��� ����� ������ (������ ���� � ������)
' v.1.0.1       : 24.10.2022 - ������ ������ �������� ���������� ������������� ��������
Const DelayAfterZip = 333
Const iTryMax = 3
Dim iTry As Integer
Dim i As Long, iMax As Long
Dim strFilePath As String, strFilename As String
Dim strZipPath As String, strZipName As String ', strFileName As String
Dim Result As Boolean
    On Error GoTo HandleError
' ���������� ����� � �������� ������
    If Len(FilePath) = 0 Then Err.Raise vbObjectError + 512
    
    With oFso
        strFilePath = .GetParentFolderName(FilePath)
        strFilename = .GetFileName(FilePath)
        If Len(ZipName) > 0 Then
            strZipPath = .GetParentFolderName(ZipName)
            strZipName = .GetFileName(ZipName)
        Else
    ' ���� ���� � ����� ������ �� ����� ����� ��� ������������ ����� ���� FilePath
            strZipPath = .GetParentFolderName(strFilePath)
            strZipName = .GetBaseName(strFilePath) & "." & c_strObjExtZip
            ZipName = .BuildPath(strZipPath, strZipName)
        End If
    ' ���� ��������� ���� ������ ���������� ��������� � ��������
        Result = .FileExists(ZipName): If Result Then GoTo HandlePack
    ' ��������� ���� ������ �� ����������
        ' ������� ��������� ����
        If Not oFso.FolderExists(strZipPath) Then Call oFso.CreateFolder(strZipPath)  ' Then Err.Raise 76 ' Path not Found
        ZipName = .BuildPath(strZipPath, strZipName)
        ' ������� ���� ������
        Result = ZipCreate(ZipName): If Not Result Then GoTo HandleExit ' ���� �� ������� ������� - ������� �� ������
    End With
HandlePack:
' ���������� �������������
    Dim oItm As Object, oZip As Object, lItm As Long, sItm As String
    
    Set oZip = oApp.Namespace((ZipName))
    For Each oItm In oApp.Namespace((strFilePath)).Items
        If oItm.IsFolder Then
        ' ���� ��� �����
            ' �������� ���������� ������ � �����
            sItm = oItm.NAME: lItm = oItm.GetFolder.Items.Count
            ' ���� ����� ����� - ��������� � ��������� �������
            If lItm = 0 Then GoTo HandleNext
        End If
        ' ���������� ����� � �����
        oZip.MoveHere (oItm.path), 4 + 8 + 16 + 1024
        ' ������� ��������� ������ ������
        Do
            Sleep DelayAfterZip: DoEvents: DoEvents
            If oItm.IsFolder Then Result = oApp.Namespace((ZipName & "\" & sItm)).Items.Count = lItm
        Loop Until Result
HandleNext:
    Next oItm
    Sleep DelayAfterZip: DoEvents: DoEvents
HandleDelAfterZip:
    If DelAfterZip Then Result = oFso.DeleteFolder(strFilePath) = 0
    Set oItm = Nothing: Set oZip = Nothing
HandleExit:     ZipPack = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case 91:    If iTry < iTryMax Then iTry = iTry + 1: Err.Clear: Sleep 333: Resume 0
                Message = "������ ��� ������� �������� �������"
                If Len(sItm) > 0 Then Message = Message & " sItm=""" & sItm & """ "
                'Stop: Err.Clear: Resume 0
    Case 70, 76: If iTry < iTryMax Then iTry = iTry + 1: Err.Clear: Sleep 333: Resume Next
                Message = "�� ������� ��������� ������� �����."
                If Len(strFilePath) > 0 Then Message = Message & " FilePath=""" & strFilePath & """ "
                'Stop: Err.Clear: Resume 0
    Case 10094: Err.Clear: Resume 0: '�������� �� �� ��� ����� ���� ������� ��������� � �����, ���������� ������� ZIP-�����" �� ������� ��������� ������� ��������� (���������, ��� ����� �� ����� ������ �� ������, � ��� �� ������ ���������� �� �� ��������)
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
    Err.Clear: Resume 0
End Function
Public Function ZipUnPack( _
    ZipFile As String, _
    Optional FilePath As String = vbNullString, _
    Optional Overwrite As Boolean = True _
    ) As Boolean
' ��������� ����� � �������� �����
Const c_strProcedure = "ZipUnPack"
Dim strPath As String, strZipName As String
Dim Result As Boolean
On Error GoTo HandleError
' ���������� ������� ������ ��� ����������
    If oFso.FileExists(ZipFile) = False Then
    ' ���� ���� ����������� - �������
        MsgBox "System could not find " & ZipFile & " Unpack cancelled.", vbInformation, "Error Unziping File"
        GoTo HandleExit 'Exit Function
    End If
    strZipName = oFso.GetFileName(ZipFile)
' ���������� ���� ����������
    If Len(FilePath) = 0 Then
    ' ���� ����� ������ ���� - ��������� �� ��������� �����
        strPath = CurrentProject.path & "\" & c_strSrcPath & "\"
    Else
        If oFso.FolderExists(FilePath) Then
            strPath = oFso.GetFolder(FilePath).path 'oFso.BuildPath(FilePath)
        Else
            Err.Raise 53, , "Folder not found"
        End If
    End If
' ��������� ����� � �������� ����������
    Dim oZip As Object:  Set oZip = oApp.Namespace((ZipFile & "\"))
    Dim oItm As Object
  
    With oZip
        If Overwrite Then
            For Each oItm In .Items
            ' ���� ���� ��� ����� ��� ���������� - ������� ����� �����������
                If oFso.FileExists(strPath & oItm.NAME) Then
                    Kill strPath & oItm.NAME
                ElseIf oFso.FolderExists(strPath & oItm.NAME) Then
                    Kill strPath & oItm.NAME & "\*.*"
                    RmDir strPath & oItm.NAME
                End If
            Next
        End If
        oApp.Namespace(CVar(strPath)).CopyHere .Items, 4 + 16
    End With
'    '���� ������� ������� ����� ����� ���������� - �������
'    If DelAfterUnZip Then Kill ZipFile
    Result = True
HandleExit:     On Error Resume Next
    ZipUnPack = Result: Exit Function
HandleError:    Result = False
    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit: Resume
End Function
Public Function ZipItemsCount(ZipName As String) As Integer
' ���������� ���������� ������ � ������
' ZipName - ������ ���� � ������
Dim oItms As Object
Dim Result As Integer
    On Error GoTo HandleError
Dim oFld As Object: Set oFld = oApp.Namespace((ZipName))
    Result = oFld.Items().Count
HandleExit:  ZipItemsCount = Result: Exit Function
HandleError: Result = -1: Resume HandleExit
End Function
Public Function ZipItemName( _
    ZipName As String, _
    Optional i As Integer = 0, _
    Optional fExt As Boolean = True) As String
' ���������� ��� i-���� ����� � ������
' ZipName - ��� ������
' i - ����� ����� � ������ (������ � 0), �� ��������� - 0
' fExt - �������� ���������� � ��� �����, �� ��������� - true
Dim Result As String
    On Error GoTo HandleError
Dim oFld As Object: Set oFld = oApp.Namespace((ZipName))
    With oFld.Items().Item((i)):  Result = IIf(fExt, .path, .NAME): End With
HandleExit:  ZipItemName = Result: Exit Function
HandleError: Result = vbNullString: Resume HandleExit
End Function
'-------------------------
' ������� ���������� ��������� ���������� �������
'-------------------------
Private Function p_LinkedRead(LocalName As String, FilePath As String) ', Optional TableName)
' ������ ��������� � ������� ��������� ������� �� ����� FilePath
Const c_strProcedure = "p_LinkedRead"
Dim strName As String, Connect As String, TableName As String, Attributes  As Long
    On Error GoTo HandleError
    LocalName = VBA.Trim$(LocalName): If Len(LocalName) = 0 Then LocalName = VBA.Trim$(p_SettingKeyRead(FilePath, c_strLnkSecParam, c_strLnkKeyLocal))
    TableName = VBA.Trim$(p_SettingKeyRead(FilePath, c_strLnkSecParam, c_strLnkKeyTable))
    Connect = VBA.Trim$(p_SettingKeyRead(FilePath, c_strLnkSecParam, c_strLnkKeyConnect))
    Attributes = VBA.Trim$(p_SettingKeyRead(FilePath, c_strLnkSecParam, c_strLnkKeyAttribute))
' ������� ������������� ������� � ��������� �����������
    If Len(TableName) = 0 Or Len(Connect) = 0 Then GoTo HandleExit
    If Len(LocalName) = 0 Then LocalName = TableName
    Dim tdf As Object 'dao.TableDef
    Set tdf = CurrentDb.CreateTableDef(LocalName) '��� ��������� �������
    With tdf
        .Connect = Connect ' ������ ����������� � ���� �� �������,����� �� ������ ������ ��� ������
        .SourceTableName = TableName ' ��� ������� ���������
        '.Attributes = Attributes
    End With
' ��������� ���� �� �������
    With CurrentDb.TableDefs: .Append tdf: .Refresh: End With: Set tdf = Nothing
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_LinkedWrite(TableName As String, FilePath As String)
' �������� � ���� FilePath ��������� �������������� ������� TableName
Const c_strProcedure = "p_LinkedWrite"
    On Error GoTo HandleError
    p_SettingKeyWrite FilePath, c_strLnkSecParam, c_strLnkKeyLocal, CurrentDb.TableDefs(TableName).NAME
    p_SettingKeyWrite FilePath, c_strLnkSecParam, c_strLnkKeyConnect, CurrentDb.TableDefs(TableName).Connect
    p_SettingKeyWrite FilePath, c_strLnkSecParam, c_strLnkKeyTable, CurrentDb.TableDefs(TableName).SourceTableName
    p_SettingKeyWrite FilePath, c_strLnkSecParam, c_strLnkKeyAttribute, CurrentDb.TableDefs(TableName).Attributes
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function

Private Function p_ReferencesDropBroken() As Boolean
' �������� �������
On Error Resume Next
Dim refs As Access.References, ref As Access.Reference, i As Integer, fBroken As Boolean
Dim wshShell As Object, strGUID As String, strKey As String
  
    Set refs = Access.References
    '������ �� ������� � �������� �������
    For i = refs.Count To 0 Step -1
        Set ref = refs(i)
        fBroken = ref.IsBroken: If Err.Number Then fBroken = True: Err.Clear ' Err.Number=48
        If Not fBroken Then GoTo HandleNext
        If ref.BuiltIn Then fBroken = True: Exit For
        '������ ������� �������� ����� ������
        refs.Remove ref
        If Err.Number = 0 Then GoTo HandleNext
        Err.Clear
        If ref.Kind = 1 Then fBroken = True: GoTo HandleNext
        '���� �������� ������ ��� �������� ������ �� ����������, �������� WSH ��������
        '�������� � ������ ����� � GUID � ������� �� ����� ������, ������� ������,
        '� ����� ������� �����.
        If wshShell Is Nothing Then Set wshShell = CreateObject("WScript.Shell")
        If Err.Number <> 0 Then fBroken = True: GoTo HandleNext
        strGUID = ref.GUID
        strKey = "HKCR\TypeLib\" & strGUID & "\" & ref.Major & "." & ref.Minor & "\"
        wshShell.RegWrite strKey & "0\win32\", ""
        '������ ������� �������� ������ (���, ����, ����������������)
        refs.Remove ref
        If Err.Number <> 0 Then Err.Clear: fBroken = True
        
        wshShell.RegDelete strKey & "0\win32\"
        wshShell.RegDelete strKey & "0\"
        wshShell.RegDelete strKey
        
        '�������� ������� ������ �� ����� ������ ������������������ ����������.
        refs.AddFromGuid strGUID, 0, 0
        Err.Clear
HandleNext:
    Next i
    '  Next ref
    p_ReferencesDropBroken = fBroken
HandleExit:
    Set ref = Nothing
    Set refs = Nothing
    Set wshShell = Nothing
    Exit Function
End Function
Private Function p_ReferencesRead(FilePath As String) ', Optional TableName)
' ������ � ��������������� ������ (References) �� ����� FilePath
Const c_strProcedure = "p_ReferencesRead"
    On Error GoTo HandleError
    Dim strText As String: strText = p_SettingSecRead(FilePath, c_strRefSecName)
    If Len(strText) = 0 Then Err.Raise 76 ' Path not Found
    Dim aRefs() As String: aRefs = Split(strText, ";")
    Dim ref As Reference
' ��������� ��� ��������� ������ � ������� ���� ��� ��������
    On Error Resume Next
    Dim Broken As Boolean: Broken = False
    For Each ref In References
' �� ����� ������� ��������� ������ �� MSComctlLib ����� ������� Win64
        Err.Clear
        If ref.BuiltIn Then GoTo HandleRemoveNext
        If Not ref.IsBroken Then GoTo HandleRemoveNext
        strText = ref.GUID      ' ���������� GUID
        References.Remove ref
    ' ������� ������
        DoEvents
    ' ���� ������� ������� - �������� ������������ �� GUID
        If Err.Number = 0 Then References.AddFromGuid strText, 0, 0: If Err.Number = 0 Then GoTo HandleRemoveNext
''        '���� �������� ������ ��� �������� ������ �� ����������, �������� WSH ��������
''        '�������� � ������ ����� � GUID � ������� �� ����� ������, ������� ������,
''        '� ����� ������� �����.
''        Err.Clear
''        If oWsh Is Nothing Then Set oWsh = CreateObject("WScript.Shell")
''        If Err.Number <> 0 Then fBroken = True: GoTo HandleNext
''        strRegKey = c_strRegKey & strGUID & "\" & lngMajor & "." & lngMinor & "\0\win32\"
''        oWsh.RegWrite strRegKey, ""
''        '������ ������� �������� ������ (���, ����, ����������������)
''        refs.Remove ref
''        If Err.Number <> 0 Then Err.Clear: fBroken = True
''
''        oWsh.RegDelete c_strRegKey 'c_strRegKey & strGUID & "\" & iMajor & "." & iMinor & "\0\win32\"
''        oWsh.RegDelete c_strRegKey & strGUID & "\" & lngMajor & "." & lngMinor & "\0\"
''        oWsh.RegDelete c_strRegKey & strGUID & "\" & lngMajor & "." & lngMinor & "\"
'' ���� ��� ����� �� ������� ������������ - �������� ������ ��� ������
        Broken = Broken Or Err.Number: Err.Clear
HandleRemoveNext:
    Next ref
' ��������������� ������ �� �����
    Dim Itm, aRef() As String ', strName As String, strDesc As String
    On Error Resume Next
    For Each Itm In aRefs
        Err.Clear
    ' �������� ������� ��� ������������ ������
        aRef = Split(Itm, "=")
    ' ��������� �� �������
        Set ref = References(aRef(0)): If Err.Number = 0 Then GoTo HandleNext
        Err.Clear
    ' �������� �� ��������� �� �������� ����������� ����������
        aRef = Split(aRef(1), "|") ' GUID|Major|Minor|FullPath
        If Err.Number = 0 Then
            Set ref = References.AddFromGuid(aRef(0), aRef(1), aRef(2))  ' GUID|Major|Minor
'            Set ref = References.AddFromFile(aRef(3))' FullPath
        End If
HandleNext:
    Next Itm
    If Not Broken Then GoTo HandleExit
' �� ��� ������ ������� ��������� - ����������� ������� ������������
    Dim strTitle As String, strMessage As String
    strTitle = "��������!"
    strMessage = "��� �������������� ����� ��������� ������ ������������� ������������ �� �������." & vbCrLf & _
        "��� �������������� ������ ������� ���������� ����� � �������� VBA (Alt+F11)," & vbCrLf & _
        "������� ���� Tools\References, ����� ������� �� ������ ���������� MISSING," & vbCrLf & _
        "� ������������� ������ ������."
    MsgBox Title:=strTitle, Prompt:=strMessage, Buttons:=vbOKOnly + vbInformation
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_ReferencesWrite(FilePath As String)
' �������� � ���� FilePath ������ (References)
Const c_strProcedure = "p_ReferencesWrite"
' ��������� ���:
' [References]
'   Name=GUID|Major|Minor|FullPath
    On Error GoTo HandleError
    Dim Itm As Object
    For Each Itm In References
' ���������� ���������� � ��������� ������
        If Itm.BuiltIn Then GoTo HandleNext
        'If Itm.IsBroken Or Err.Number <> 0 Then Else GoTo HandleNext
' ��������� � ����: Name=GUID|Major|Minor|FullPath
        With Itm: p_SettingKeyWrite FilePath, c_strRefSecName, .NAME, Join(Array(.GUID, .Major, .Minor, .FullPath), "|"): End With
HandleNext:
    Next Itm
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_PropertiesRead(FilePath As String)
' ������ � ��������������� �������� ������� �� ����� FilePath
Const c_strProcedure = "p_PropertiesRead"
    On Error GoTo HandleError
' ��������������� �������� �������
    With VBE.ActiveVBProject
        .NAME = p_SettingKeyRead(FilePath, c_strPrjSecName, c_strPrjKeyName)
        .Description = p_SettingKeyRead(FilePath, c_strPrjSecName, c_strPrjKeyDesc)
        .HelpFile = p_SettingKeyRead(FilePath, c_strPrjSecName, c_strPrjKeyHelp)
    End With
Dim strText As String
Dim aPrps() As String, Itm
Dim strName As String, strValue As String, intType As eDataType
' ��������������� ���������������� ��������
    strText = p_SettingSecRead(FilePath, c_strPrpSecName): If Len(strText) = 0 Then Err.Raise 76 ' Path not Found
    ' ������� ��� ������ ��������
    With CurrentProject.Properties
        Do While .Count > 0: .Remove .Item(0).NAME: Loop
    ' ��������� ����������� ��������
    aPrps = Split(strText, ";")
    For Each Itm In aPrps
        Itm = VBA.Trim$(Itm)
        If Len(Itm) > 0 Then
            strValue = p_PropertyStringRead(CStr(Itm), PropName:=strName)
            If Len(strName) > 0 Then .Add strName, strValue
        End If
    Next Itm
    End With
' ��������������� �������� ���� ������
    strText = p_SettingSecRead(FilePath, c_strDbsSecName): If Len(strText) = 0 Then Err.Raise 76 ' Path not Found
    ' ��������� ����������� ��������
    aPrps = Split(strText, ";")
On Error Resume Next
Dim prp As DAO.Property
    With CurrentDb
    For Each Itm In aPrps
    ' ������ ������ ����������
        Itm = VBA.Trim$(Itm)
        If Len(Itm) > 0 Then
            strValue = p_PropertyStringRead(CStr(Itm), PropName:=strName, PropType:=intType)
            If Len(strName) > 0 Then PropertySet strName, strValue, CurrentDb, intType
        End If
    Next Itm
    End With
Err.Clear
On Error GoTo HandleError
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_PropertiesWrite(FilePath As String)
' �������� � ���� FilePath �������� �������
Const c_strProcedure = "p_PropertiesWrite"
    On Error GoTo HandleError
' ��������� �������� �������
    With VBE.ActiveVBProject
        p_SettingKeyWrite FilePath, c_strPrjSecName, c_strPrjKeyName, p_PropertyStringCreate(.NAME, PropType:=dbText)
        p_SettingKeyWrite FilePath, c_strPrjSecName, c_strPrjKeyDesc, p_PropertyStringCreate(.Description, PropType:=dbText)
        p_SettingKeyWrite FilePath, c_strPrjSecName, c_strPrjKeyHelp, p_PropertyStringCreate(.HelpFile, PropType:=dbText)
    End With
Dim Itm As Object, strName As String, varValue, intType As eDataType: intType = dbText
' ��������� ���������������� ��������
    With CurrentProject
    For Each Itm In .Properties
        With Itm:  strName = .NAME: varValue = .Value: End With
        varValue = p_PropertyStringCreate(varValue, PropType:=intType)
        p_SettingKeyWrite FilePath, c_strPrpSecName, strName, CStr(varValue)
    Next Itm
    End With
' ��������� �������� ���� ������
On Error Resume Next
    Dim i As Long
    With CurrentDb
    For Each Itm In .Properties
        Err.Clear
        With Itm
            strName = .NAME
            Select Case strName
            ' ���������� �������� �������� �� �����
            Case "DesignMasterID", "Name", "Transactions", "Updatable", "CollatingOrder", _
                 "Version", "RecordsAffected", "ReplicaID", "Connection", "AccessVersion", _
                 "Build", "ProjVer"
                GoTo HandleNext
            End Select
            varValue = .Value: intType = .Type
        ' ��� �������� ����������� ���� ��� ������ �������� �������� � ���� �������� � ����� ������
            .Value = varValue
        End With
        ' ���� ������ - �������� ���������� ����� �������� - ��� ������������� ��� ���������
        If Err.Number Then Debug.Print strName & " - " & Err.Number & " " & Err.Description: Err.Clear: GoTo HandleNext
        ' ���� ��� ������ - ��������� �������� ��� ������������ ��������������
        varValue = p_PropertyStringCreate(varValue, PropType:=intType)
        p_SettingKeyWrite FilePath, c_strDbsSecName, strName, CStr(varValue)
HandleNext:
    Next Itm
    End With
    Err.Clear
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_PropertyStringRead(PropString As String, _
    Optional PropName, Optional PropType As eDataType _
    ) As Variant
' ������ �������� ��������� �������� �� ������
Const c_strProcedure = "p_PropertyStringRead"
Dim Result

    Result = PropString: If Len(Result) = 0 Then GoTo HandleExit
    On Error GoTo HandleError
    If Not IsMissing(PropName) Then
        PropName = Left$(PropString, InStr(1, PropString, "=") - 1)
        Result = Mid$(PropString, InStr(1, PropString, "=") + 1): If Len(Result) = 0 Then GoTo HandleExit
    End If
    If (Len(Result) > 1) And ((Left$(Result, 1) = """") And (Right$(Result, 1) = """")) Then
        PropType = dbText: Result = Mid$(Result, 2, Len(Result) - 2)
    ElseIf IsNumeric(Result) Then
        PropType = dbNumeric: Result = Val(Result)
    ElseIf (Len(Result) > 1) And ((Left$(Result, 1) = "#") And (Right$(Result, 1) = "#")) Then
        PropType = dbDate: Result = CDate(Mid$(Result, 2, Len(Result) - 2))
    ElseIf UCase(Result) = "TRUE" Then
        PropType = dbBoolean: Result = CBool(True)
    ElseIf UCase(Result) = "FALSE" Then
        PropType = dbBoolean: Result = CBool(False)
    Else
    Dim strType As String: strType = UCase(Left$(Result, InStr(1, Result, ":") - 1))
    If Len(strType) > 0 Then Result = Mid$(Result, InStr(1, Result, ":") + 1)
        Select Case strType
        Case "DECIMAL": PropType = dbDecimal:   Result = CDec(Result)
        Case "BYTE":    PropType = dbByte:      Result = CByte(Result)
        Case "SINGLE":  PropType = dbSingle:    Result = CSng(Result)
        Case "DOUBLE":  PropType = dbDouble:    Result = CDbl(Result)
        Case "INTEGER": PropType = dbInteger:   Result = CInt(Result)
        Case "LONG":    PropType = dbLong:      Result = CLng(Result)
        Case "CURRENCY": PropType = dbCurrency: Result = CCur(Result)
        Case "FLOAT":   PropType = dbFloat:     Result = Val(Result)
        Case Else:      Err.Raise vbObjectError + 512
        End Select
    End If
HandleExit:  p_PropertyStringRead = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Private Function p_PropertyStringCreate(PropValue As Variant, _
    Optional PropName As String, Optional PropType As eDataType = dbText _
    ) As String
' ��������� ������ �������� ��� ��������
Const c_strProcedure = "p_PropertyStringCreate"
Dim Result 'As String

    On Error GoTo HandleError
    Select Case PropType
    Case dbText:    Result = """" & CStr(PropValue) & """"
'    Case dbChar
'    Case dbMemo
    Case dbBoolean: Result = IIf(CBool(PropValue), "True", "False")
    Case dbNumeric: Result = Val(PropValue)
    Case dbDecimal: Result = "Decimal:" & CDec(PropValue)
    Case dbByte:    Result = "Byte:" & CByte(PropValue)
    Case dbSingle:  Result = "Single:" & CSng(PropValue)
    Case dbDouble:  Result = "Double:" & CDbl(PropValue)
    Case dbInteger: Result = "Integer:" & CInt(PropValue)
    Case dbLong:    Result = "Long:" & CLng(PropValue)
    Case dbCurrency: Result = "Currency:" & CCur(PropValue)
    Case dbFloat:   Result = "Float:" & CDec(PropValue)
'    Case dbBigInt
'    Case dbBinary
'    Case dbLongBinary
'    Case dbVarBinary
'    Case dbGUID
    Case dbDate:    Result = "#" & CDate(PropValue) & "#"
'    Case dbTime:    Result = "#" & CDate(PropValue) & "#"
'    Case dbTimeStamp
    Case Else: Err.Raise vbObjectError + 512
    End Select
    If Len(PropName) > 0 Then Result = PropName & "=" & Result
HandleExit:  p_PropertyStringCreate = Result:  Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
'-------------------------
' ������� ��� ������ � ini
'-------------------------
Private Function p_SettingKeyRead(strFile As String, strSection As String, strRegKeyName As String) As String
' ���������� ��������� �������� �� INI �����
Const c_strProcedure = "p_SettingKeyRead"
' strSection - ��� ������
' strRegKeyName - ��� �������� ���������
' strFile - ���� � ini �����
Dim strBuffer As String * 4096
Dim intSize As Integer

    On Error GoTo HandleError
    intSize = GetPrivateProfileString(strSection, strRegKeyName, "", strBuffer, 4096, strFile)
    p_SettingKeyRead = VBA.Left$(strBuffer, intSize)
HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_SettingKeyWrite(strFile As String, strSection As String, strRegKeyName As String, strValue As String) As Integer
' ��������� �������� ��������� � ����
Const c_strProcedure = "p_SettingKeyWrite"
' strSection - ��� ������
' strRegKeyName - ��� �������� ���������
' strValue - �������� ���������
' strFile - ���� � ini �����
' ���������� True ���� �������, ����� - False
Dim intStatus As Integer

    On Error GoTo HandleError
    intStatus = WritePrivateProfileString(strSection, strRegKeyName, strValue, strFile)
    p_SettingKeyWrite = (intStatus <> 0)

HandleExit:     Exit Function
HandleError:    Dim Message As String
    Select Case Err 'Err.Number
    Case Else:  Message = Err.Description ': Resume 0
    End Select
    Debug.Print c_strModule & "." & c_strProcedure, "Error# " & Err.Number & ": " & Message
    Err.Clear: Resume HandleExit
End Function
Private Function p_SettingSecRead(strFile As String, strSection As String) As String
' ������ ����� � �������� ������ � �������� ������ �� ����� .INI
Const c_strProcedure = "p_SettingSecRead"
'����������: Param1=Val1;Param2=Val2...
' strSection - ��� ������
' strFile - ���� � ini �����
Dim strBuffer As String * 4096
Dim intSize As Integer
    intSize = GetPrivateProfileSection(strSection, strBuffer, 4096, strFile)
    p_SettingSecRead = Replace(VBA.Left$(strBuffer, intSize), VBA.Chr$(0), ";")
End Function
'-------------------------
' ��������������� �������
'-------------------------
Private Function p_SelectFile(Optional InitPath As String, _
    Optional FileMask As String = "*.*", Optional Extention As String = "*", _
    Optional DialogTitle As String) As Variant
' ������ ������ �����
Const c_strProcedure = "p_SelectFile"
Dim of As OPENFILENAME
Dim Result: Result = vbNullString
    On Error GoTo HandleError
    With of
        .lStructSize = Len(of)
        .hwndOwner = hWndAccessApp
        .lpstrInitialDir = InitPath
        .lpstrFilter = FileMask
        .nFilterIndex = 1
        .lpstrFile = VBA.String$(512, 0)
        .nMaxFile = 511
        .lpstrDefExt = Extention
        .lpstrTitle = DialogTitle
    End With
    If GetOpenFileName(of) Then Result = StrZ(of.lpstrFile)
HandleExit:  p_SelectFile = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Private Function p_SelectFolder( _
    Optional InitPath As String = vbNullString, _
    Optional InitView As Integer = msoFileDialogViewList, _
    Optional DialogTitle As String _
    ) As Variant
' ������ ������ �����
Const c_strProcedure = "p_SelectFolder"
Dim Result As Variant: Result = vbNullString
    On Error GoTo HandleError
'
'Dim fRet As Long, bi As BROWSEINFO, dwIList As Long
'Dim szPath As String, wPos As Integer
'
'    With bi
'        .hOwner = hWndAccessApp
'        .lpszTitle = DialogTitle
'        .ulFlags = BIF_RETURNONLYFSDIRS
'    End With
'
'    dwIList = SHBrowseForFolder(bi)
'    szPath = Space$(512)
'    fRet = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
'
'    If fRet Then Result = StrZ(szPath)
Dim InitFolder As String
' SHBrowseForFolder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = DialogTitle
        .InitialView = InitView
        If Len(InitPath) > 0 Then
            If dir(InitPath, vbDirectory) <> vbNullString Then
                InitFolder = InitPath
                If Right(InitFolder, 1) <> "\" Then
                    InitFolder = InitFolder & "\"
                End If
                .InitialFileName = InitFolder
            End If
        End If
        .Show
        On Error Resume Next
        Err.Clear
        Result = .SelectedItems(1)
        If Err.Number <> 0 Then Result = vbNullString
    End With
HandleExit:  p_SelectFolder = Result: Exit Function
HandleError: Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Private Function p_TextAlpha2Code(Source As String)
' �������� ������� �� �������� � ������ ���������� �� ���������������� Asc ����� ���� %XX
Dim c As Long, cMax As Long
Dim PermissedSymb As String
Dim Char As String
Dim Result As String

    Result = vbNullString
' ������ ����������� �������
    PermissedSymb = VBA.UCase$(c_strSymbDigits & c_strSymbRusAll & c_strSymbEngAll & c_strOthers)
    c = 1: cMax = Len(Source)
    Do Until c > cMax
        Char = VBA.Mid$(Source, c, 1)
        If InStr(1, PermissedSymb, VBA.UCase$(Char)) = 0 Then
            Char = VBA.Hex$(Asc(Char))
            If Len(Char) < c_strCodeLen Then Char = VBA.String$(c_strCodeLen - Len(Char), "0") & Char
            Char = c_strCodeSym & Char
        End If
        Result = Result & Char
        c = c + 1
    Loop
HandleExit: p_TextAlpha2Code = Result
End Function
Private Function p_TextCode2Alpha(Source As String)
' �������� ���������������� Asc ��� ���� %XX, �������� �� �������� � ������ ����������, �� ���������
Dim c As Long, cMax As Long
Dim i As Byte, tmpChar As String
Dim Char As String
Dim Result As String
    Result = vbNullString
    c = 1: cMax = Len(Source)
    Do Until c > cMax
        Char = VBA.Mid$(Source, c, 1)
        If Char = c_strCodeSym Then
            tmpChar = VBA.UCase$(VBA.Mid$(Source, c + 1, c_strCodeLen))
            i = 1
            Do Until i > c_strCodeLen
                Select Case VBA.Mid$(tmpChar, i, 1)
                Case "0" To "9", "A" To "F"
                Case Else: GoTo HandleSkip
                End Select
                i = i + 1
            Loop
            c = c + c_strCodeLen
            Char = VBA.Chr$(Val(c_strHexPref & tmpChar))
HandleSkip:
        End If
        Result = Result & Char
        c = c + 1
    Loop
HandleExit: p_TextCode2Alpha = Result
End Function
Private Function InStrRegEx( _
    Start As Long, _
    String1 As String, String2 As String, _
    Optional Found As String, _
    Optional Compare As VbCompareMethod = vbTextCompare) As Long
' InStr ����������� ������ ���������� �� ����� ��������� RegEx
Const c_strProcedure = "p_InstrRegEx"
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
'Static oRegEx As Object: If oRegEx Is Nothing Then Set oRegEx = CreateObject("VBScript.RegExp")
    ' ����� RegExp � �������� ��� �����
    With oRegEx: .IgnoreCase = (Compare = vbTextCompare): .Pattern = String2: Set oMatches = .Execute(S1): End With
    If oMatches.Count = 0 Then GoTo HandleError
    For Each oMatch In oMatches
        Result = oMatch.FirstIndex + Start: Found = oMatch
        Exit For
    Next
HandleExit:  InStrRegEx = Result: Exit Function
HandleError: Result = False: Err.Clear: Resume HandleExit
End Function
'-------------------------
' ������� ������ ����������� ��������
'-------------------------
Private Function oFso() As Object
Static soFSO As Object: If soFSO Is Nothing Then Set soFSO = CreateObject("Scripting.FileSystemObject")
    Set oFso = soFSO
End Function
Private Function oApp() As Object
Static soApp As Object: If soApp Is Nothing Then Set soApp = CreateObject("Shell.Application")
    Set oApp = soApp
End Function
Private Function oRegEx() As Object
Static soRegEx As Object: If soRegEx Is Nothing Then Set soRegEx = CreateObject("VBScript.RegExp")
    Set oRegEx = soRegEx
End Function
'-------------------------
' ��������������� �������
'-------------------------
Private Function StrZ(par As String) As String
Dim nSize As Long, i As Long ', Rez As String
   nSize = Len(par)
   i = InStr(1, par, VBA.Chr(0)) - 1
   If i > nSize Then i = nSize
   If i < 0 Then i = nSize
   StrZ = VBA.Mid$(par, 1, i)
End Function
#If USEZIPCLASS Then
Private Function oZip() As clzZipArchive
Static soZip As clzZipArchive: If soZip Is Nothing Then Set soZip = New clzZipArchive
    Set oZip = soZip
End Function
#End If
'-------------------------
' �������������� �������
'-------------------------
Private Function p_WinUserName() As String
' ��� ������������ Windows
Const c_strProcedure = "p_WinUserName"
Dim sBuffer As String, lSize As Long
Dim Result As String
    On Error GoTo HandleError
    sBuffer = VBA.Space$(255): lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    Result = vbNullString
    If lSize > 0 Then Result = VBA.Left$(sBuffer, InStr(sBuffer, VBA.Chr(0)) - 1)
    'If lSize > 0 Then Result = VBA.Left$(sBuffer, lSize)
HandleExit:     p_WinUserName = Result: Exit Function
HandleError:    Result = vbNullString: Err.Clear: Resume HandleExit
End Function
Private Function p_GetSysDiskSerial() As String
' ������� �������� ��������� ������ �����
Dim Result As String
Dim VolLabel As String, VolSize As Long, Serial As Long, MaxLen As Long
Dim strName As String, NameSize As Long, Flags As Long
Const DiskName = "C:\"

    Result = VBA.String(8, "0")
    On Error GoTo HandleError
    If GetVolumeInformation(DiskName, VolLabel, VolSize, Serial, MaxLen, Flags, strName, NameSize) _
        Then Result = VBA.Format$(VBA.Hex$(Serial), VBA.String$(8, "0"))
HandleExit:     p_GetSysDiskSerial = Result: Exit Function
HandleError:    Result = VBA.String$(8, "0"): Err.Clear: Resume HandleExit
End Function
