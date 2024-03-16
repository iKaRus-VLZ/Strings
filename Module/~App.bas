Attribute VB_Name = "~App"
Option Compare Database
Option Base 1
'=========================
Private Const c_strModule As String = "~App"
'=========================
' ��������      :
' ������        : 1.0.0.0
' ����          : 01.07.2016 15:00:00
' �����         : ������ �.�. (KashRus@gmail.com)
' ����������    :
'=========================
' ��������� �����������
'=========================
'-------------------------
' �������� �����
'-------------------------
Public Enum AppColorScheme
    appColorGrey = &H969696
    appColorLightGrey = &H767676
' �������� �����
    appColorDark = &H993333 '&H732A0A             ' ������ ����
    appColorBright = &HDFA000           ' ����� ����
    appColorLight = &HE5D1C0            ' ������� ����
' ������ 1
    appColorDark2 = &H730A1F
    appColorBright2 = &HDF3000
    appColorLight2 = &HE5C0C1
' ������ 2
    appColorDark3 = &H730A53
    appColorBright3 = &HDF003F
    appColorLight3 = &HE5C0D4
End Enum
'-------------------------
' �����
'-------------------------
Public Const appFontNameDef = "Arial"
Public Const appFontSizeDef = 10
Public Const appFontSize1 = 8
Public Const appFontSize2 = 12
Public Const appFontSize3 = 18
'-------------------------

'=========================
' ��������� �����������
'=========================
'-------------------------
Private Const c_strApplication = "��������"
Private Const c_strAuthor As String = "������ �.�."
Private Const c_strVersion As String = "2.0.0"
Private Const c_strSupport As String = "KashRus@gmail.com"
Private Const c_strShowClock As Boolean = True
'-------------------------
'Public Const cpHash = "SHA256", HashType = eHashSHA256 '
'-------------------------
' �������� �������� ������
'SysBookmarks, SysFields, SysLog, SysMenu, SysObjData, SysObjTypes, SysOrderTypes, SysUsers, SysVersion
Public Const c_strSysObjects = "MSysObjects"    ' ��������� �������
Public Const c_strTableVers = "SysVersion"      ' �������� ������ ����������
Public Const c_strTableUser = "SysUsers"        ' �������� ������ �������������
Public Const c_strTableMenu = "SysMenu"         ' �������� ������� ����
Public Const c_strTableData = "SysObjData"      ' �������� ��������� ��������
Public Const c_strTableLogs = "SysLog"          ' �������� ��������� ������
Public Const c_strTableDate = "SysCalendar"     ' ���������� ���
'Public Const c_strTableBook = "SysBookmarks"    ' ���������� �������� �������
'Public Const c_strTableFlds = "SysFields"       ' ���������� ������������ �����
'Public Const c_strTableTObj = "SysObjTypes"     ' ���������� ����� ��������
'Public Const c_strTableTOrd = "SysOrderTypes"   ' ���������� ����� ��������

'-------------------------
' ����� ������� �������
'-------------------------
' �������� ����������
Public Const c_strPropAppName = "Application" '= c_strApplication ' ="��������"
Public Const c_strPropAuthor = "Author"
Public Const c_strPropSupport = "Support"
Public Const c_strPropVersion = "Version"
Public Const c_strPropVerDate = "VersionDate"
Public Const c_strPropLastDate = "LastDate"
Public Const c_strPropLastUser = "LastUserName"
Public Const c_strPropFirstRun = "FirstRun"
Public Const c_strPropShowClock = "ShowClock"
' �������� ����� ����������
Public Const c_strPropSrvPath = "SrvPath"
Public Const c_strPropDatPath = "DatPath"
Public Const c_strPropSecPath = "SecPath"
Public Const c_strPropLogPath = "LogPath"
Public Const c_strPropDllPath = "DllPath"
Public Const c_strPropDocPath = "DocPath"
Public Const c_strPropTmpPath = "TmpPath"
' �������� ���������� ������ ��� ���������� ����������
Public Const c_strDesignRes = "DesignRes" ' 1280x1024
Public Const c_strDesignDpi = "DesignDpi" '
Public Const c_strResDelim = "x" '
'-------------------------
' �������� ������� �������
'-------------------------
Public Const c_strLastUserName As String = "�������������" ' ������������� �� SysUsers
Public Const c_strSrvPath As String = "\"
Public Const c_strDatPath As String = "DAT"
Public Const c_strDllPath As String = "LIB"
Public Const c_strLogPath As String = "LOG"
Public Const c_strTmpPath As String = "DOT"
Public Const c_strDocPath As String = "DOC"
Public Const c_strSrcPath As String = "SRC"
Public Const c_strDbfPath As String = "DBF"
Public Const c_bolShowClock As Boolean = True
'-------------------------
' ������
'-------------------------
Public Const c_strAppIco = "App"
Public Const c_strMenuIco = "ContextMenu"
'-------------------------
' ��� �������� ����������� ��� ����������
'-------------------------
' ������� ������ � ��������� ������� �������
Private Const strBegLineMarker = "'=== BEGIN INSERT ==="
Private Const strEndLineMarker = "'==== END INSERT ===="
' ���� �� ��� ���� �� ������ � DoEvents �� ��������� = 333
Public Const appDoEventsPause = 100

Public Const c_strTagDelim = "_"
Public Const c_strDelim = ";"
Public Const c_strInDelim = ","
' �������� ���������� SQL
Public Const sqlSelect = "SELECT ", sqlAll = "*"
Public Const sqlUpdate = "UPDATE ", sqlSet = " SET "
Public Const sqlInsert = "INSERT ", sqlInto = " INTO "
Public Const sqlTransform = "TRANSFORM ", sqlPivot = " PIVOT "
Public Const sqlDelete = "DELETE ", sqlUnion = "UNION "
Public Const sqlDrop = "DROP ", sqlTable = " TABLE ", sqlIndex = " INDEX "
Public Const sqlAs = " AS "
Public Const sqlDistinct = "DISTINCT ", sqlDistinctRow = "DISTINCTROW "
Public Const sqlFrom = " FROM ", sqlWhere = " WHERE "
Public Const sqlOrder = " ORDER BY ", sqlGroup = " GROUP BY "
Public Const sqlHaving = " HAVING ", sqlTop = " TOP ", sqlTop1 = "TOP 1 ", sqlPercent = " PERCENT "
Public Const sqlJoin = " JOIN ", sqlInner = " INNER", sqlLeft = " LEFT", sqlRight = " RIGHT", sqlOn = " ON "
Public Const sqlIdentity = "@@Identity"
Public Const sqlSelectAll = sqlSelect & sqlAll & sqlFrom
Public Const sqlSelect1st = sqlSelect & sqlTop1 & sqlAll & sqlFrom
Public Const sqlDeleteAll = sqlDelete & sqlAll & sqlFrom
Public Const sqlDropTable = sqlDrop & sqlTable, sqlDropIndex = sqlDrop & sqlIndex
Public Const sqlOR = " OR ", sqlAnd = " AND ", sqlNot = " NOT "
Public Const sqlEqual = "=", sqlGreater = ">", sqlLess = "<"
Public Const sqlGreaterOrEqual = ">=", sqlLessOrEqual = "<=", sqlNotEqual = "<>"
Public Const sqlIn = " IN ", sqlLike = " LIKE ", sqlBetween = " BETWEEN "
Public Const sqlAsc = " ASC", sqlDesc = " DESC"
Public Const sqlSimilar = "SIMILAR"  ' ������������� - �������� �����
Public Const sqlIs = " IS ", sqlNull = "NULL", sqlTrue = "True", sqlFalse = "False"
Public Const sqlIsNull = sqlIs & sqlNull, sqlIsNotNull = sqlIs & sqlNot & sqlNull
' �������� �������� ������
'SysBookmarks, SysFields, SysLog, SysMenu, SysObjData, SysObjTypes, SysOrderTypes, SysUsers, SysVersion
'Private Const c_strSysObjects = "MSysObjects"    ' ��������� �������

' �������� �������������� ����������
Public Const c_strParamType = "Type"
Public Const c_strParamMode = "Mode"
Public Const c_strParamKey = "Key"
'-------------------------
' �������� �������� Access
'-------------------------
' AccessObjectType
Public Const c_strTmpTypePref = "tmp" ' ��������������� ������
Public Const c_strTmpTablPref = "@&%" ' ��������� �������

' �������� ��� ����������� � �������� "On[...]" ������� ��� ��������� ��� �������
Public Const c_strCustomProc = "[Event Procedure]"
Public Const c_strCmdMnuProc = "ContextMenu_Click"
' ��������� �������� ��� ��������� ��������� �������
Public frmDROP_Date_Controls As Collection ' ��������� ��������� ����� frmDROP_Date
' FormType
Public Const c_strMenuType = "MENU" ' ����
' ���
Public Const c_strServType = "SERV" ' ���������
Public Const c_strDropType = "DROP" ' ���������� �����
' FormMode
Public Const c_strMainMode = "MAIN" ' ��� ���� - ��������
' ��� ���������� ���� FormType=c_strDropType
Public Const c_strRealMode = "Real" ' ���������� ���.��������
Public Const c_strCalcMode = "Calc" ' ���������� �����������
Public Const c_strDateMode = "Date" ' ���������� ���������
' ��� ��������� ���� FormType=c_strServType
Public Const c_strUserMode = "User"
Public Const c_strUChgMode = "UserChg"
Public Const c_strFloat = "Float"   ' ��������� ������
Public Const c_strNavBar = "NavBar" ' ������ ��������� �� �������
Public Const c_strPrtBar = "PrtBar" ' ������ �� ������������ �������
' ��� �������� ����������� �� ����� �����������

'==============================
Public Enum appErrors
' ���������������� ������ ����������
    errAuthGrant = vbObjectError + 1000     ' ����� �����
    errAuthError = vbObjectError + 1001     ' ������ �����������
    errAuthFailed = vbObjectError + 1002    ' ����� �����������
    errAuthEnd = vbObjectError + 1009       ' ����� ��������
    errAppNoConn = vbObjectError + 1010     ' ����������� ����������� � ������� ������
    errAppNoData = vbObjectError + 1011     ' ����������� ������
    errAppClose = vbObjectError + 1109      ' ������ �������� ����������
    errAppPathWrong = vbObjectError + 1110  ' ������ �������� ���� ����������
End Enum
Public Enum enmPathType
' ���� ����� ���������� ������������
    enmPathUndef = 0 ' �� ���������
    enmPathAll = 255 ' ��� ����
    enmPathSrv = 1  ' ���� � ������� ������ ����������
    enmPathDll = 2  ' ���� � ����� ������� ��������� ����������
    enmPathTmp = 3  ' ���� � ����� �������� ������� ����������
    enmPathDoc = 4  ' ���� � ����� ������� ����������
    enmPathDat = 5  ' ���� � ����� ������
    enmPathSec = 6  ' ���� � ����� ������� ������
    enmPathLoc = 7  ' ���� � ��������� ���� (����������)
    enmPathLog = 8  ' ���� � ����� ���������� ����������
    enmPathSrc = 9  ' ���� � ����� ����� ����������
    enmPathDbf = 10 ' ���� � ����� ��������
End Enum
Public Enum appUserType
' ���� ���� ������������
    appUserTypeAdmin = 100
    appUserTypeUser = 200
End Enum
Public Enum appModeType
' ������������� ������� ����������
    appModeDebug = -1                   ' ����� �������
    appModeNormal = 0                   ' ������� �����
End Enum
Public Enum appRecState
' �������� ��������� ������ (SPReal)
    appRecStateTemp = -1 '��������� - �� ����������
    appRecStateReal = 0  '�������� - ����������
    appRecStateOld = 10  '������ - ���� ��������, ���� ����� ����������
    appRecStateArc = 11  '�������� - ���� ��������, ����� ���������� ���
    appRecStateDen = 91  '���������
    appRecStateDel = 99  '�������� ���������
End Enum
'======================
Private bolFirstRun As Boolean
'----------------------
' POINTER
'----------------------
#If VBA7 = 0 Then       'LongPtr trick by @Greedo (https://github.com/Greedquest)
Public Enum LongPtr
    [_]
End Enum
#End If
#If VBA7 And Win64 Then '<WIN64 & OFFICE2010+>  PtrSafe, LongPtr and LongLong
Private Const PTR_LENGTH As Long = 8
Private Const VARIANT_SIZE As Long = 24
#Else                   '<OFFICE97-2010>        Long
Private Const PTR_LENGTH As Long = 4
Private Const VARIANT_SIZE As Long = 16
#End If                 '<WIN32>
'======================
'Public Function App(): Static myApp As New clsApp: Set App = myApp: End Function                ' ���������� ������ �� ����� �������� ����������
'Public Function Cmd(): Static myCmd As New clsCommands: Set Cmd = myCmd: End Function           ' ���������� ������ �� ����� �������� �������� ����������
'Public Function Dbg(): Static myDebug As New clsDebug: Set Dbg = myDebug: End Function          ' ���������� ������ �� ����� ������ �������
'Public Function Crypto(): Static myCrypto As New clsCrypto: Set Crypto = myCrypto: End Function ' ���������� ������ �� ����� ������ ����������
'Public Function Fso(): Static myFso As Object: Set Fso = CreateObject("Scripting.FileSystemObject"): End Function ' ���������� ������ �� ����� ������ ��������� �������/������
'Public Function Wdr(): Static myWord As New clsWordReport: Set Wdr = myWord: End Function ' ���������� ������ �� ����� ������� Word
'======================
Public Function StartApp(): Call App.AppStart(True): End Function
Public Function StopApp(): App.AppStop (False): End Function
Public Function UpdateAppPath(): Call App.UpdatePath: End Function
Public Function UpdateAppMode(): App.ModeSwitch: End Function
Public Function UpdateLocalRefs(): App.UpdateRefs: End Function
Public Function Setup()
' ��������� ������� ��� �������������� ���������
' ��������� ����������
'    RestoreRefs     ' �������������� ������ �� ����������
    RestoreProp     ' ��������� ������� ����������
    CloseAll        ' ��������� ��������� ��� �������
    CompileAll      ' ����������
End Function
Public Sub RestoreProp()
Const c_strProcedure = "RestoreProp"
'������������� ������� ����������
On Error GoTo HandleError
    SetOption ("Auto Compact"), True            ' ������� ��� ������
    SetOption ("ShowWindowsInTaskbar"), True    ' ��������� ���� � ������ �����
'    SetOption ("Show Status Bar"), False        ' ��������� ������ ���������
    With CurrentProject.Properties
'=== BEGIN INSERT ===
        .Add "Application", c_strApplication ' ="��������"
        .Add "Version", c_strVersion ' =""
        .Add "Author", c_strAuthor   ' ="������ �.�."
        .Add "Support", c_strSupport ' ="KashRus@gmail.com"
        .Add "FirstRun", c_strFirstRun ' ="0"
'==== END INSERT ====
    End With
HandleExit:
    Exit Sub
HandleError:
    Dbg.Error Err.Number, Err.Description, Err.Source, c_strModule & "." & c_strProcedure, Erl()
    Resume HandleExit
End Sub
'======================

Public Sub ContextMenu_Click()
' ���������� ������� ������������ ����
    With Application.CommandBars.ActionControl
Stop
Debug.Print .Tag, .Caption
    End With
End Sub



'=========================
' ������������ �������
'=========================
Sub Test() 'Optional WordCase As DeclineCase = 0, Optional NumbType As NumeralType = 1)
Dim strName As String, strComment As String
Dim Result As String
'Call 1  Expression = [comma-separated 12,345 elements list, 1 char per element]
'Call 2  Expression = [comma-separated 1,234 elements list, 1 char per element]
'Call 3  Expression = [comma-separated 123 elements list, 10 chars per element]
'Call 4  Expression = [comma-separated 12 elements list, 100 chars per element]
Dim i As Long, iMax As Long
Dim sLen As Integer, dLen As Integer
Dim Arr() As String, Delim As String: Delim = ";"
Dim Source As String, Pos As Long, Data As String
    dLen = 3 ' ���� ��������� ����� ������ ���-�� sLen (Len(String)=sLen-dLen+Int(d2*Len*rnd))
'    iMax = 100
'    Pos = 90
'    sLen = 12
'    Source = "4FPkEodTl;T2fden5NWBomJ;Ugl6Riavfc9cwF;LoG6clotBYH8q1M;Zocpn9NBUZgywyg;7pROhQH2p;TUPnryXPu99sl5;YHYmM3y82v;h9fSeFWT1FU30L6;kBwuy7jWm25;OAYbsS7ii2nf;6Np3nTg6lZqAP;IRe61Y03ngaY;IWixDoT51y2qh;" & _
'        "pyysAK8Ifky;WM3jJ8Vck9C;DBOEVuwBMEkX;EXCQAhL7RU6;ETYlvS38BZdUmjx;HNpEYB5dMbEMiVu;qbQAPyVOauT;8YHu4YA5SUhln;oguy88AStYWF;GOjAyH882hc7U;gt7eNbof3;G0iQjVeOf;q64a9dYECvw;Ei4ekSatJM4V4QI;1U82pT6v1Ze;" & _
'        "3YRTc7cdQUlS;OPEljsO7xA;fDm0bxUifMRuR6;HjU4lo8UtgyY;Th4yDGq1EYj;fgn3LIO2QYeQe;IOt8QwH80s2jUr;lYgKd8qLm76;0WUOFDEn5I8Hm;tPK65KuWOk6C;BntiaZJNU9;7dUH5lIvsQ;ygLAK71LuyMUK5b;f7ZfK5YHtN;fbyS9RTWUf5;" & _
'        "Ja3sTFy1XD;N33nihOcLdI7;5U3AgEiXoi4xOZ;tohkjoRnEDG7;X1qubBLuCvU;7IQPeM3AApWG3;nAHAVUDUxAb;Gng5b7qQE29uh;YWsQEc4X7;CAfOLTxErrh;KCLuMP9eyaw;p5iiEVKL5Tdk3hF;0KF3AwSn6h;fpg42zU3IROA;ODDbAwOaWEhakqM;" & _
'        "sdaQjahyAjXX3;rtTedpriRWNydt;mxxuqpZEc2;otsV6tOPYOHGr;W9NIG8kSVaTY;Gl2W2Amp7rr2cU;sqKitC8n4fQ3OLS;dvg9XX3Wh0;XcV8tQUHJ7eDP;fasUax0JUU4;11Z5ODDas;fCwUDwg2ZMVUy;8hvBEW4k5RiN;WtIVkEezTpyHI;uljsEMjr1ghS;" & _
'        "Obs2Dkm2Id6DI;OiMdg0yrcC9NZ;VVnOrCP40B;iioCYnzJ0EISf1;SJL9JPRYOXUgTS;XnQojEdN07;woyyjd3nWmOF;Su7Xfs6yhP8;yl0ZicVg8CdZu0;d9qJl1N7snO5YR;bjwULh3U1UnNl1;NctiE4PERa;MVef6t5BkEXobl1;uEGKgEexnH3;QAxFq4CdK;" & _
'        "tG5LHmiiCaB6ti;lx3iUPwF6Uei;iUaq95h65bpL4V;jNbeBbLa7QK;R9GUDOtBvU3ZGR;3joyyVpcO6NM;QwUC7L3fYdEOG;ery6X68kR;AH6U6JpbeJayV;F5kAPgpAE2R;eG38LV4eZF"
'    arr() = Split(Source, Delim)
    iMax = InputBox(Prompt:="������� ���������� ���������:", Title:="���������� ���������", Default:=100)
    sLen = InputBox(Prompt:="������� ����� �������� (" & Chr$(&HB1) & dLen & "): ", Title:="����� ���������", Default:=12)
    Data = "TestTest"
    Pos = (iMax - 1) * Rnd + 1
    ReDim Arr(0 To iMax - 1)
    For i = 0 To iMax - 1: Arr(i) = GenPassword(sLen + CInt(2 * dLen * Rnd()), NewSeed:=False): Next i
    
    sLen = sLen - dLen: If sLen < 1 Then sLen = 1: dLen = 1 / 2 * dLen
    strComment = "(" & iMax & " elements list, [" & sLen & " to " & sLen + 2 * dLen & "] chars per element)"
' ��������� ������
    Source = Join(Arr, Delim)
    iMax = UBound(Split(Source, Delim)) + 1
    Dbg.Message "iMax=" & iMax '& ", " & "Source=""" & Source & """"
' ��������� �� �������
    Source = DelimStringShrink(Source, Delim)
    iMax = UBound(Split(Source, Delim)) + 1
    Dbg.Message "����� �������: iMax=" & iMax '& ", " & "Source=""" & Source & """"
    
strName = "DelimStringGet (����� Split)"
    Dbg.Counter CounterStart, strName, COMMENT:=strComment & ", Pos=" & Pos
    Result = Split(Source, Delim)(Pos - 1)
    Dbg.Counter CounterStop, strName ', Comment:="Result=""" & Result & """"
strName = "DelimStringGet (����� InStr)"
    Dbg.Counter CounterStart, strName, COMMENT:=strComment & ", Pos=" & Pos
    Result = DelimStringGet(Source, Pos, Delim)
    Dbg.Counter CounterStop, strName ', Comment:="Result=""" & Result & """"
strName = "DelimStringGet (����� InStr) ������������� ������"
    Dbg.Counter CounterStart, strName, COMMENT:=strComment & ", Pos=" & -Pos
    Result = DelimStringGet(Source, -Pos, Delim)
    Dbg.Counter CounterStop, strName ', Comment:="Result=""" & Result & """"

strName = "DelimStringDel (����� Split)"
    Dbg.Counter CounterStart, strName, COMMENT:=strComment & ", Pos=" & Pos
    Arr = Split(Source, Delim): Arr(Pos - 1) = vbNullString
    Result = Replace(Join(Arr, Delim), Delim & Delim, Delim): Erase Arr()
    Dbg.Counter CounterStop, strName ', Comment:="Result=""" & Result & """"
strName = "DelimStringDel (����� InStr)"
    Dbg.Counter CounterStart, strName, COMMENT:=strComment & ", Pos=" & Pos
    Result = DelimStringDel(Source, Pos, Delim)
    Dbg.Counter CounterStop, strName ', Comment:="Result=""" & Result & """"

strName = "DelimStringSet (����� Split)"
    Dbg.Counter CounterStart, strName, COMMENT:=strComment & ", Pos=" & Pos & ", Data=" & Data
    Arr = Split(Source, Delim)
    'arr(Pos - 1) = Data  ' ������� � �������
    Arr(Pos - 1) = Data & Delim & Arr(Pos - 1) ' ������� �� �������
    Result = Join(Arr, Delim): Erase Arr()
    Dbg.Counter CounterStop, strName ', Comment:="Result=""" & Result & """"
strName = "DelimStringSet (����� InStr)"
    Dbg.Counter CounterStart, strName, COMMENT:=strComment & ", Pos=" & Pos & ", Data=" & Data
    Result = DelimStringSet(Source, Pos, Data, Delim)
    Dbg.Counter CounterStop, strName ', Comment:="Result=""" & Result & """"

strName = "Split"
    Erase Arr
    Dbg.Counter CounterStart, strName, COMMENT:=strComment
    Arr = Split(Result, Delim)
    Dbg.Counter CounterStop, strName ', Comment:="Result=""" & Result & """"
    strName = "xSplit"
    Erase Arr
    Dbg.Counter CounterStart, strName, COMMENT:=strComment
    Call xSplit(Result, Arr, Delim)
    Dbg.Counter CounterStop, strName ', Comment:="Result=""" & Result & """"

strName = "Join"
    Dbg.Counter CounterStart, strName, COMMENT:=strComment
    Result = Join(Arr, Delim)
    Dbg.Counter CounterStop, strName ', Comment:="Result=""" & Result & """"
    strName = "xJoin"
    Dbg.Counter CounterStart, strName, COMMENT:=strComment & "(not compiled)"
    Result = xJoin(Arr, Delim)
    Dbg.Counter CounterStop, strName ', Comment:="Result=""" & Result & """"

' Replace
Dim sText As String: sText = Replicate(iMax, "ha")
Dim sFind As String: sFind = "a"
Dim sReplace As String: sReplace = "1!"
strComment = "(Len=" & sLen & ", Find=""" & sFind & """, Replace=""" & sReplace & """)"
strName = "Replace"
    Dbg.Counter CounterStart, strName, COMMENT:=strComment
    Result = Replace(sText, sFind, sReplace)
    Dbg.Counter CounterStop, strName ', Comment:="Result=""" & Result & """"
    strName = "xReplace"
    Dbg.Counter CounterStart, strName, COMMENT:=strComment
    Result = xReplace(sText, sFind, sReplace)
    Dbg.Counter CounterStop, strName ', Comment:="Result=""" & Result & """"
End Sub
Public Sub Test2()
Dim Col As New Collection
'Dim i As Long: For i = 1 To 10: col.Add "Val" & i, "Key" & i: Next i
Dim i As Long, strTag As String
Dim strTest As String, strResult As String
Dim strName As String, strComment As String
Dim Keys
'' �������� ������
'    strTest = "����� [%Key1%] �������� �� [%Key2%] ������� [%Key1%] ��������� [%Key4%] �������� [%Key5%] �������� � ������� [%Key6%] ��������� [%Key7%] ������� ������ [%Key8%] ������ [%Key9%] �������� [%Key10%]"
'    Dbg.Message "Source:", strTest
'strName = "PlaceHoldersSet"
'' ������ 1
'    Set col = New Collection
'    For i = 1 To 10: strTag = strTag & ";Key" & i & "=Val" & i: Next i: strTag = Mid$(strTag, 2)
'    Call TaggedString2Collection(strTag, col) ': strTag = vbNullString
'    strResult = PlaceHoldersSet(strTest, col)
'    Dbg.Message "1." & strName, Comment:=strResult
'' ������ 2
'    Set col = New Collection
'    For i = 1 To 5: strTag = strTag & ";Key" & i & "=����" & i: Next i: strTag = Mid$(strTag, 2)
'    Call TaggedString2Collection(strTag, col): strTag = vbNullString
'    strResult = PlaceHoldersSet(strTest, col)
'    Dbg.Message "2." & strName, Comment:=strResult
'' ������ 3
'    Set col = New Collection
'    strTest = "2*[%Par3%]"
'    strTag = "Par2=23;Par3=[%Par1%]/[%Par2%];Par1=7"
'    Call TaggedString2Collection(strTag, col)
'    strResult = PlaceHoldersSet(strTest, col, True)
'    Dbg.Message "3." & strName, Comment:=strTest & "=" & Eval(strResult) & ", ���: " & strTag
'' ������ 4
'    Set col = New Collection
'    strTest = "� [%Word01{��������:���,����}%] ��� [%Word03{��������:���,����}%], � � [%Word02{��������:���,����}%] ��� [%Mumb01{�����������:���}%] [%Word03{��������:���,��,����}%]."
'    strTag = "Word01=�;Word02=��;Word03=����;Mumb01=10"
'    Call TaggedString2Collection(strTag, col)
'    strResult = PlaceHoldersSet(strTest, col, True)
'    Dbg.Message "4." & strName, Comment:=strResult & ", ���: " & strTag
strName = "PlaceHoldersGet"
strComment = "Method=0"
    Set Col = New Collection
    strResult = "(107/23)*(71)+1"
    strTest = "([%Val1%]/[%Val2%])*([%Val3%])+1"
    Dbg.Counter CounterStart, strName, COMMENT:=strComment
    Call PlaceHoldersGet(strResult, strTest, Col, Keys, ReplaceExisting:=1, Method:=0)
    Dbg.Counter CounterStop, strName
strComment = "Method=1"
    Set Col = New Collection
    strResult = "(107/23)*(71)+X"
    strTest = "([%Val1%][/*+-][%Val2%])[/*+-]([%Val3%])[/*+-]*"
    Dbg.Counter CounterStart, strName, COMMENT:=strComment
    Call PlaceHoldersGet(strResult, strTest, Col, Keys, ReplaceExisting:=1, Method:=1)
    Dbg.Counter CounterStop, strName
strComment = "Method=2"
    Set Col = New Collection
    strResult = "����� �������� �������� ������ �� ���������"
'    Debug.Print InStrRegEx(10, strResult, strTest, strTag)
    strTest = "\s[%Val1{����:��,��,��,��,��,��}%]\s"
    Dbg.Counter CounterStart, strName, COMMENT:=strComment
    Call PlaceHoldersGet(strResult, strTest, Col, Keys, ReplaceExisting:=1, Method:=2)
    Dbg.Counter CounterStop, strName
Stop
End Sub
Public Sub Test3()
Dim REx As Object, oMatch As Object
Dim strSource As String, strTest As String
Dim strName  As String, strDelim As String, strTagDelim As String
    strDelim = ";": strTagDelim = "="
    
'    strSource = "��������� ������������� ��������� ����������� ����� �������"
'    strTest = "(?:\s?)([�-�]+?)(��|����|����){0,2}(��|��|��|��|��)?(?:\s?)"
    strSource = "disabled;hide=""a1, a2, a3"";active"
    strTest = strDelim & "+(?=(?:[^\""]*\""[^\""]*\"")*[^\""]*$)"
    Set REx = RegEx
    With REx
        .Pattern = strTest
        Set oMatch = .Execute(strSource)
Stop
    End With
End Sub
