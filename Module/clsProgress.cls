VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
#Const APPTYPE = 0          '0|1        '0=ACCESS,1=EXCEL ' not yet
'=========================
Private Const c_strModule As String = "clsProgress"
'=========================
' ��������      : ����� ��� ������ � �������������
' ������        : 1.0.0.448466087
' ����          : 12.10.2022 14:36:32
' �����         :
' ����������    : ���������� ������� ���� frmSERV_Progress
'=========================
' �������� �������� � ������ ������:
'-------------------------
' Init - ������������� ������������
' Show - ����� �� �����
' Update - ���������� ��������� ��������
' Progress - ������� ��������� ��������
' Canceled - ���� ���������� ��������
'=========================
' ������� �������������:
''------------------------------------------
''*** ������ �������������� ������������ ***
''------------------------------------------
'Dim prg As clsProgress
'    Set prg = New clsProgress
'    With prg
'        .Init pMin:=-10220.12, pMax:=1750250.56, _
'            pCaption:="������������ ���������", _
'            pText:="����������..."
'        .Show
'        Do Until .Progress = .ProgressMax 'And Not .Canceled
'            .Detail = "��������� �������������� �����������" & vbCrLf & _
'                "������� �������� " & Format$(.Progress, "Standard") & ", � ���������" & " �� " & .ProgressMin & " �� " & .ProgressMax
'            Sleep 50 ' <����� ����������� �����-�� ��������
'            If .Canceled Then
'            ' ���������� ��������
'                MsgBox "������� ������� �� ��������: " & .Progress
'                Exit Do
'            End If
'            .Update
'        Loop
'    End With
'    Set prg = Nothing
''------------------------------------------
''*** ������ �������������� ������������ ***
''------------------------------------------
'Dim prg As clsProgress
'    Set prg = New clsProgress
'    With prg
'        .Init pCount:=2
'        .Caption = "������������ �������� ���������"
'        .FloodColor(eMajor) = vbRed: .BarTextColor(eMajor) = vbBlue
'        .FloodColor(eMinor) = vbCyan: .BarTextColor(eMinor) = vbWhite
'        .ProgressMax(eMajor) = 10: .ProgressStep(eMajor) = 1
'        Do Until .Progress(eMajor) = .ProgressMax(eMajor) 'And Not .Canceled
'            .ProgressMax(eMinor) = 100 * (.Progress(eMajor) + 1)
'            Do Until .Progress(eMinor) = .ProgressMax(eMinor) 'And Not .Canceled
'            ' ���������� ��������
'                If .Canceled Then
'                    Select Case MsgBox("�������� ����������?", vbYesNo Or vbExclamation Or vbDefaultButton2)
'                     Case vbYes: Exit Do
'                     Case Else: .Canceled = False
'                    End Select
'                End If
'                Sleep 50 ' <����� ����������� �����-�� �������� ������� ������
'                .Update (eMinor)
'            Loop
'            If .Canceled Then Exit Do
'            Sleep 50 ' <����� ����������� �����-�� �������� ������� ������
'            .Update (eMajor)
'        Loop
'    End With
'    Set prg = Nothing
'=========================

#If APPTYPE = 0 Then        ' APPTYPE=Access
Const cstrHostForm = "frmSERV_Progress"     ' ��� �������� ����� ��������
Const cstrCaptionLabel = "lblCaptionText"   ' ��� �������� ��������� ����� �������
Const cstrMajorTextLabel = "lblMajorText"   ' ��� �������� ��������� ��������� �����
Const cstrMinorTextLabel = "lblMinorText"   ' ��� �������� ��������� ��������� �����
Const cstrDetailLabel = "lblDetailText"     ' ��� �������� ��������� ����� �������

Private mProgressForm As Access.Form  ' ������ �� ����� ���������
Private Const c_strCustomProc = "[Event Procedure]"
#Else                       '

#End If                     ' APPTYPE

Private Enum eType
    eCaption = 0
    eMajor = 1
    eMinor = 2
    eDetail = 3
End Enum

Private mProgressCount As Byte
Private mCaptionText As String, mDetailText As String

Private mMajorValue As Double, mMajorStep As Single
Private mMajorMin As Double, mMajorMax As Double
Private mMajorText As String, mMajorBarText As String
Private mMajorFloodColor As Long, mMajorBarTextColor As Long

Private mMinorValue As Double, mMinorStep As Single
Private mMinorMin As Double, mMinorMax As Double
Private mMinorText As String, mMinorBarText As String
Private mMinorFloodColor As Long, mMinorBarTextColor As Long


Private mCancel As Boolean
'Public Event Cancel(ExitProgress As Long)
'========================
Public Sub Init( _
    Optional pCount As Byte = 1, _
    Optional pCaption As String = "", _
    Optional pMin = 0, Optional pMax = 100, Optional pText As String = "", _
    Optional pMinorMin = 0, Optional pMinorMax = 100, Optional pMinorText As String = "", _
    Optional pFloodColor As Long = 8402176, Optional pBarTextColor As Long = vbWhite _
    )
' ������������� ��������
' pCount - ���������� ������������ ������������� (1 ��� 2)
' pCaption - ��������� �����
' pMin/pMinorMin - ������������ �������� ��� ��������/������� ������������
' pMax/pMinorMax - ������������ �������� ��� ��������/������� ������������
' pText/pMinorText - ������� ��� ��������/������� ������������
' pFloodColor - ���� ���������� �������������
' pBarTextColor As Long - ���� ������ �������������

    If Not IsNumeric(pMin) Or Not IsNumeric(pMax) Then GoTo HandleExit
#If APPTYPE = 0 Then        ' APPTYPE=Access
'    Set mProgressForm = New Form_frmSERV_Progress
    DoCmd.OpenForm cstrHostForm, acNormal, , , , acHidden
    Set mProgressForm = Application.Forms(cstrHostForm)
#Else                       '
#End If                     ' APPTYPE
' ������������� �������
    mCancel = False
    Caption = pCaption
' 1 �����������
    mMajorMin = pMin: mMajorMax = pMax
    'mMajorStep = (mMajorMax - mMajorMin) / 100: If mMajorStep <= 0 Then mMajorStep = 1
    mMajorStep = 1
    FloodColor = pFloodColor: BarTextColor = pBarTextColor
    Progress = pMin: Text = pText
    Select Case pCount
     Case 1
        mProgressCount = 1
     Case 2
' 2 �����������
        mProgressCount = 2
    If Not IsNumeric(pMinorMin) Or Not IsNumeric(pMinorMax) Then GoTo HandleExit
        mMinorMin = pMinorMin: mMinorMax = pMinorMax
        'mMinorStep = (mMinorMax - mMinorMin) / 100:If mMinorStep <= 0 Then mMinorStep = 1
        mMinorStep = 1
        FloodColor(eMinor) = pFloodColor: BarTextColor(eMinor) = pBarTextColor
        Progress(eMinor) = pMinorMin: Text(eMinor) = pMinorText
     Case Else
        mCancel = True
        GoTo HandleExit
    End Select
HandleInit:
#If APPTYPE = 0 Then        ' APPTYPE=Access
' ������������� �����
    mProgressForm.KeyPreview = True
    mProgressForm.OnKeyDown = c_strCustomProc
    mProgressForm.Init mProgressCount
#Else                       '

#End If                     ' APPTYPE

HandleExit:
End Sub
Public Sub Show()
    mProgressForm.Visible = True
    mProgressForm.Repaint
End Sub
Public Sub Clear(Optional ProgressNum As Byte = eMajor)
' ������� ��������� ������������
    mCancel = False
' ������� ���������� ���� ������
    If ProgressNum = eMajor Then Progress = mMajorMin
' ������ ���������� ������
    Progress(eMinor) = mMajorMin
End Sub
Public Sub Update(Optional ProgressNum As Byte = eMajor, Optional ProgressStep)
' ��������� ��������� ��������
Dim pStep As Single, pValue As Double
    DoEvents
    mCancel = mProgressForm.Canceled
    Select Case ProgressNum
    Case eMajor
        If IsMissing(ProgressStep) Then
            pStep = mMajorStep
        ElseIf Not IsNumeric(ProgressStep) Then
            pStep = mMajorStep
        Else
            pStep = ProgressStep
        End If
        ' ���� ����� ������� �������� �������� - �������� ������ (�.�. � �� ����)
        If mProgressCount > 1 Then Progress(eMinor) = mMinorMin
        ' ������������� �������� �������� ���������
        Progress(eMajor) = mMajorValue + pStep
    Case eMinor
        If IsMissing(ProgressStep) Then
            pStep = mMinorStep
        ElseIf Not IsNumeric(ProgressStep) Then
            pStep = mMinorStep
        Else
            pStep = ProgressStep
        End If
        ' ������������� �������� ������� ���������
        Me.Progress(eMinor) = mMinorValue + pStep
    End Select
End Sub
'=============================
Private Sub Class_Initialize()
    mProgressCount = 1
#If APPTYPE = 0 Then        ' APPTYPE=Access
    If p_IsFormOpen(cstrHostForm) Then DoCmd.Close acForm, cstrHostForm, acSaveNo
#Else                       '
#End If                     ' APPTYPE
End Sub
Private Sub Class_Terminate()
#If APPTYPE = 0 Then        ' APPTYPE=Access
    Set mProgressForm = Nothing
    If p_IsFormOpen(cstrHostForm) Then DoCmd.Close acForm, cstrHostForm, acSaveNo
#Else                       '
#End If                     ' APPTYPE
End Sub
'=============================
' ���������� �������������
Public Property Let ProgressCount(rData As Byte)
#If APPTYPE = 0 Then        ' APPTYPE=Access
    mProgressForm.Init rData
#Else                       '
#End If                     ' APPTYPE
    mProgressCount = rData
End Property
Public Property Get ProgressCount() As Byte: ProgressCount = mProgressCount: End Property
' ��������� ����� ���������
Public Property Let Caption(rData As String)
#If APPTYPE = 0 Then        ' APPTYPE=Access
    mProgressForm.lblCaptionText.Caption = rData
#Else                       '
#End If                     ' APPTYPE
End Property
Public Property Get Caption() As String:
#If APPTYPE = 0 Then        ' APPTYPE=Access
    Caption = mProgressForm.lblCaptionText.Caption
#Else                       '
#End If                     ' APPTYPE
End Property
' ��������� ���������� �� ����� ��������
Public Property Let Detail(rData As String)
#If APPTYPE = 0 Then        ' APPTYPE=Access
    mProgressForm.lblDetailText.Caption = rData
#Else                       '
#End If                     ' APPTYPE
End Property
Public Property Get Detail() As String
#If APPTYPE = 0 Then        ' APPTYPE=Access
    Detail = mProgressForm.lblDetailText.Caption
#Else                       '
#End If                     ' APPTYPE
End Property
' ���������� � ��������� ������������
Public Property Let Text(Optional ProgressNum As Byte = eMajor, rData As String)
    Select Case ProgressNum
    Case eMajor:    mMajorText = rData
#If APPTYPE = 0 Then        ' APPTYPE=Access
                    mProgressForm.lblMajorText.Caption = rData
#Else                       '
#End If                     ' APPTYPE
    Case eMinor:    mMinorText = rData
#If APPTYPE = 0 Then        ' APPTYPE=Access
            mProgressForm.lblMinorText.Caption = rData
#Else                       '
#End If                     ' APPTYPE
    End Select
#If APPTYPE = 0 Then        ' APPTYPE=Access
    mProgressForm.Repaint
#Else                       '
#End If                     ' APPTYPE
End Property
' ����� ��������� � �����������
Public Property Let BarText(Optional ProgressNum As Byte = eMajor, rData As String)
    Select Case ProgressNum
    Case eMajor:    mMajorBarText = rData
#If APPTYPE = 0 Then        ' APPTYPE=Access
                    mProgressForm.lblMajorPercent.Caption = rData
#Else                       '
#End If                     ' APPTYPE
    Case eMinor:    mMinorBarText = rData
#If APPTYPE = 0 Then        ' APPTYPE=Access
                    mProgressForm.lblMinorPercent.Caption = rData
#Else                       '
#End If                     ' APPTYPE
    End Select
#If APPTYPE = 0 Then        ' APPTYPE=Access
    mProgressForm.Repaint
#Else                       '
#End If                     ' APPTYPE
End Property
' ���� ������ ������������
Public Property Let BarTextColor(Optional ProgressNum As Byte = eMajor, rData As Long)
    Select Case ProgressNum
    Case eMajor:
#If APPTYPE = 0 Then        ' APPTYPE=Access
                    mProgressForm.lblMajorPercent.ForeColor = rData
#Else                       '
#End If                     ' APPTYPE
    Case eMinor:
#If APPTYPE = 0 Then        ' APPTYPE=Access
                    mProgressForm.lblMinorPercent.ForeColor = rData
#Else                       '
#End If                     ' APPTYPE
    End Select
#If APPTYPE = 0 Then        ' APPTYPE=Access
    mProgressForm.Repaint
#Else                       '
#End If                     ' APPTYPE
End Property
' ���� ������������
Public Property Let FloodColor(Optional ProgressNum As Byte = eMajor, rData As Long)
    Select Case ProgressNum
    Case eMajor:
#If APPTYPE = 0 Then        ' APPTYPE=Access
                    mProgressForm.lblMajorFlood.BackColor = rData
#Else                       '
#End If                     ' APPTYPE
    Case eMinor:
#If APPTYPE = 0 Then        ' APPTYPE=Access
                    mProgressForm.lblMinorFlood.BackColor = rData
#Else                       '
#End If                     ' APPTYPE
    End Select
#If APPTYPE = 0 Then        ' APPTYPE=Access
    mProgressForm.Repaint
#Else                       '
#End If                     ' APPTYPE
End Property
' ������� �������� ���������
Public Property Get Progress(Optional ProgressNum As Byte = eMajor) As Double
    Select Case ProgressNum
    Case eMajor: Progress = mMajorValue
    Case eMinor: Progress = mMinorValue
    End Select
End Property
Public Property Let Progress(Optional ProgressNum As Byte = eMajor, rData As Double)
Dim pPercent As Single
Dim sText As String, sBarText As String
' ����� ������ ������� �������� �������� ��������
    DoEvents
#If APPTYPE = 0 Then        ' APPTYPE=Access
    mCancel = mProgressForm.Canceled
#Else                       '
#End If                     ' APPTYPE

    Select Case ProgressNum
    Case eMajor
        mMajorValue = rData
        If mMajorValue > mMajorMax Then mMajorValue = mMajorMax
        If mMajorValue < mMajorMin Then mMajorValue = mMajorMin
        If mMajorMax <> mMajorMin Then pPercent = (mMajorValue - mMajorMin) / (mMajorMax - mMajorMin) Else pPercent = 1
        sText = "������� ��������: " & p_FormatedNum(mMajorValue) & ", � ��������� �� " & mMajorMin & " �� " & mMajorMax
        sBarText = Format$(pPercent, "Percent")
#If APPTYPE = 0 Then        ' APPTYPE=Access
        With mProgressForm
            .lblMajorFlood.Width = pPercent * .FloodWidth
            If Len(mMajorText) = 0 Then .lblMajorText.Caption = sText
            If Len(mMajorBarText) = 0 Then .lblMajorPercent.Caption = sBarText
        End With
#Else                       '
#End If                     ' APPTYPE
    Case eMinor
        mMinorValue = rData
        If mMinorValue > mMinorMax Then mMinorValue = mMinorMax
        If mMinorValue < mMinorMin Then mMinorValue = mMinorMin
        If mMinorMax <> mMinorMin Then pPercent = (mMinorValue - mMinorMin) / (mMinorMax - mMinorMin) Else pPercent = 1
        sText = "������� ��������: " & p_FormatedNum(mMinorValue) & ", � ��������� �� " & mMinorMin & " �� " & mMinorMax
        sBarText = Format$(pPercent, "Percent")
#If APPTYPE = 0 Then        ' APPTYPE=Access
        With mProgressForm
        ' ������������ ������� ��������
            If Len(mMajorText) = 0 Or Len(mMajorBarText) = 0 Then Me.Progress(eMajor) = mMajorValue
        ' ������������ ������� ��������
            .lblMinorFlood.Width = pPercent * .FloodWidth
            If Len(mMinorText) = 0 Then .lblMinorText.Caption = sText
            If Len(mMinorBarText) = 0 Then .lblMinorPercent.Caption = sBarText
#Else                       '
#End If                     ' APPTYPE
        End With
    End Select
#If APPTYPE = 0 Then        ' APPTYPE=Access
    If Not mProgressForm.Visible Then Show
    mProgressForm.Repaint
#Else                       '
#End If                     ' APPTYPE
End Property
Public Property Let ProgressMin(Optional ProgressNum As Byte = eMajor, rData As Double)
' ����������� �������� ���������
' ���� ���������� ����������� - ������������� ���
    Select Case ProgressNum
    Case eMajor: mMajorMin = rData: If mMajorStep <= 0 Then mMajorStep = (mMajorMax - mMajorMin) / 100
    Case eMinor: mMinorMin = rData: If mMinorStep <= 0 Then mMinorStep = (mMinorMax - mMinorMin) / 100
    End Select
End Property
Public Property Get ProgressMin(Optional ProgressNum As Byte = eMajor) As Double
    Select Case ProgressNum
    Case eMajor: ProgressMin = mMajorMin
    Case eMinor: ProgressMin = mMinorMin
    End Select
End Property
Public Property Let ProgressMax(Optional ProgressNum As Byte = eMajor, rData As Double)
' ������������ �������� ���������
' ���� ���������� ������������ - ������������� ���
    Select Case ProgressNum
    Case eMajor: mMajorMax = rData: If mMajorStep <= 0 Then mMajorStep = (mMajorMax - mMajorMin) / 100
    Case eMinor: mMinorMax = rData: If mMinorStep <= 0 Then mMinorStep = (mMinorMax - mMinorMin) / 100
    End Select
End Property
Public Property Get ProgressMax(Optional ProgressNum As Byte = eMajor) As Double
    Select Case ProgressNum
    Case eMajor: ProgressMax = mMajorMax
    Case eMinor: ProgressMax = mMinorMax
    End Select
End Property
Public Property Get ProgressStep(Optional ProgressNum As Byte = eMajor) As Single
' ������������ �������� ���������
    Select Case ProgressNum
    Case eMajor: ProgressStep = mMajorStep
    Case eMinor: ProgressStep = mMinorStep
    End Select
End Property
Public Property Let ProgressStep(Optional ProgressNum As Byte = eMajor, rData As Single)
    If rData <= 0 Then rData = 1
    Select Case ProgressNum
    Case eMajor: mMajorStep = rData 's = (mMajorMax-mMajorMin)/100
    Case eMinor: mMinorStep = rData 's = (mMinorMax-mMinorMin)/100
    End Select
End Property
' ���� ���������� ��������
Public Property Get Canceled() As Boolean: Canceled = mCancel: End Property
Public Property Let Canceled(rData As Boolean)
#If APPTYPE = 0 Then        ' APPTYPE=Access
    mProgressForm.Canceled = rData
#Else                       '
#End If                     ' APPTYPE
    mCancel = rData
End Property
'=============================
Private Function p_FormatedNum(rNumber As Double) As String
Dim fPart As Double, dig As Byte, tmp As Byte
Dim msk As String
' ��� ����� �����
    msk = "#,##0"
    fPart = Abs(rNumber - Fix(rNumber))
    If fPart > 0 Then
' ���� ������� �����
        dig = 1 'dig = 2 ' ��� ������� ����� �������
        Do While fPart * 10 ^ dig Mod 10 ^ dig = 0: dig = dig + 1: Loop
        msk = msk & "." & String(dig - 1, "#") & "0"
    End If
    p_FormatedNum = Format$(rNumber, msk)
End Function
#If APPTYPE = 0 Then        ' APPTYPE=Access
Private Function p_IsFormOpen(sFormName As String) As Boolean
Dim tmp As String
    On Error Resume Next
    tmp = Application.Forms(sFormName).Caption
    p_IsFormOpen = Not CBool(Err.Number)
End Function
#Else                       '
#End If                     ' APPTYPE

