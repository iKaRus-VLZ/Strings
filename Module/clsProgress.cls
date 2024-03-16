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
' Описание      : Класс для работы с прогрессбаром
' Версия        : 1.0.0.448466087
' Дата          : 12.10.2022 14:36:32
' Автор         :
' Примечание    : необходимо наличие форм frmSERV_Progress
'=========================
' Основные свойства и методы модуля:
'-------------------------
' Init - инициализация прогрессбара
' Show - вывод на экран
' Update - обновление состояния процесса
' Progress - текущее состояние счетчика
' Canceled - флаг прерывания процесса
'=========================
' Примеры использования:
''------------------------------------------
''*** пример одноуровневого прогрессбара ***
''------------------------------------------
'Dim prg As clsProgress
'    Set prg = New clsProgress
'    With prg
'        .Init pMin:=-10220.12, pMax:=1750250.56, _
'            pCaption:="Демонстрация прогресса", _
'            pText:="Выполнение..."
'        .Show
'        Do Until .Progress = .ProgressMax 'And Not .Canceled
'            .Detail = "проверяем однопроцессный прогрессбар" & vbCrLf & _
'                "Текущее значение " & Format$(.Progress, "Standard") & ", в диапазоне" & " от " & .ProgressMin & " до " & .ProgressMax
'            Sleep 50 ' <здесь выполняются какие-то действия
'            If .Canceled Then
'            ' прерывание процесса
'                MsgBox "Процесс прерван на значении: " & .Progress
'                Exit Do
'            End If
'            .Update
'        Loop
'    End With
'    Set prg = Nothing
''------------------------------------------
''*** пример двухуровневого прогрессбара ***
''------------------------------------------
'Dim prg As clsProgress
'    Set prg = New clsProgress
'    With prg
'        .Init pCount:=2
'        .Caption = "Демонстрация двойного прогресса"
'        .FloodColor(eMajor) = vbRed: .BarTextColor(eMajor) = vbBlue
'        .FloodColor(eMinor) = vbCyan: .BarTextColor(eMinor) = vbWhite
'        .ProgressMax(eMajor) = 10: .ProgressStep(eMajor) = 1
'        Do Until .Progress(eMajor) = .ProgressMax(eMajor) 'And Not .Canceled
'            .ProgressMax(eMinor) = 100 * (.Progress(eMajor) + 1)
'            Do Until .Progress(eMinor) = .ProgressMax(eMinor) 'And Not .Canceled
'            ' прерывание процесса
'                If .Canceled Then
'                    Select Case MsgBox("Прервать выполнение?", vbYesNo Or vbExclamation Or vbDefaultButton2)
'                     Case vbYes: Exit Do
'                     Case Else: .Canceled = False
'                    End Select
'                End If
'                Sleep 50 ' <здесь выполняются какие-то действия второго уровня
'                .Update (eMinor)
'            Loop
'            If .Canceled Then Exit Do
'            Sleep 50 ' <здесь выполняются какие-то действия первого уровня
'            .Update (eMajor)
'        Loop
'    End With
'    Set prg = Nothing
'=========================

#If APPTYPE = 0 Then        ' APPTYPE=Access
Const cstrHostForm = "frmSERV_Progress"     ' имя основной формы прогреса
Const cstrCaptionLabel = "lblCaptionText"   ' имя контрола текстовой метки подписи
Const cstrMajorTextLabel = "lblMajorText"   ' имя контрола первичной текстовой метки
Const cstrMinorTextLabel = "lblMinorText"   ' имя контрола вторичной текстовой метки
Const cstrDetailLabel = "lblDetailText"     ' имя контрола текстовой метки деталей

Private mProgressForm As Access.Form  ' ссылка на форму прогресса
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
' инициализация объектов
' pCount - количество отображаемых прогрессбаров (1 или 2)
' pCaption - заголовок формы
' pMin/pMinorMin - миниммальное значение для верхнего/нижнего прогрессбара
' pMax/pMinorMax - максимальное значение для верхнего/нижнего прогрессбара
' pText/pMinorText - подпись для верхнего/нижнего прогрессбара
' pFloodColor - цвет заполнения прогрессбаров
' pBarTextColor As Long - цвет текста прогрессбаров

    If Not IsNumeric(pMin) Or Not IsNumeric(pMax) Then GoTo HandleExit
#If APPTYPE = 0 Then        ' APPTYPE=Access
'    Set mProgressForm = New Form_frmSERV_Progress
    DoCmd.OpenForm cstrHostForm, acNormal, , , , acHidden
    Set mProgressForm = Application.Forms(cstrHostForm)
#Else                       '
#End If                     ' APPTYPE
' инициализация свойств
    mCancel = False
    Caption = pCaption
' 1 прогрессбар
    mMajorMin = pMin: mMajorMax = pMax
    'mMajorStep = (mMajorMax - mMajorMin) / 100: If mMajorStep <= 0 Then mMajorStep = 1
    mMajorStep = 1
    FloodColor = pFloodColor: BarTextColor = pBarTextColor
    Progress = pMin: Text = pText
    Select Case pCount
     Case 1
        mProgressCount = 1
     Case 2
' 2 прогрессбар
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
' инициализация формы
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
' очищает состояние прогрессбара
    mCancel = False
' верхнюю сбрасываем если задано
    If ProgressNum = eMajor Then Progress = mMajorMin
' нижнюю сбрасываем всегда
    Progress(eMinor) = mMajorMin
End Sub
Public Sub Update(Optional ProgressNum As Byte = eMajor, Optional ProgressStep)
' обновляет состояние процесса
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
        ' если новый уровень верхнего счетчика - обнуляем нижний (м.б. и не надо)
        If mProgressCount > 1 Then Progress(eMinor) = mMinorMin
        ' устанавливаем значение верхнего прогресса
        Progress(eMajor) = mMajorValue + pStep
    Case eMinor
        If IsMissing(ProgressStep) Then
            pStep = mMinorStep
        ElseIf Not IsNumeric(ProgressStep) Then
            pStep = mMinorStep
        Else
            pStep = ProgressStep
        End If
        ' устанавливаем значение нижнего прогресса
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
' количество прогрессбаров
Public Property Let ProgressCount(rData As Byte)
#If APPTYPE = 0 Then        ' APPTYPE=Access
    mProgressForm.Init rData
#Else                       '
#End If                     ' APPTYPE
    mProgressCount = rData
End Property
Public Property Get ProgressCount() As Byte: ProgressCount = mProgressCount: End Property
' заголовок формы процессов
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
' детальная информация об этапе процесса
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
' информация в заголовке прогрессбара
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
' текст выводимый в прогрессбар
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
' цвет текста прогрессбара
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
' цвет прогрессбара
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
' текущее значение прогресса
Public Property Get Progress(Optional ProgressNum As Byte = eMajor) As Double
    Select Case ProgressNum
    Case eMajor: Progress = mMajorValue
    Case eMinor: Progress = mMinorValue
    End Select
End Property
Public Property Let Progress(Optional ProgressNum As Byte = eMajor, rData As Double)
Dim pPercent As Single
Dim sText As String, sBarText As String
' прямо задаем текущее значение счетчика итераций
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
        sText = "Текущее значение: " & p_FormatedNum(mMajorValue) & ", в диапазоне от " & mMajorMin & " до " & mMajorMax
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
        sText = "Текущее значение: " & p_FormatedNum(mMinorValue) & ", в диапазоне от " & mMinorMin & " до " & mMinorMax
        sBarText = Format$(pPercent, "Percent")
#If APPTYPE = 0 Then        ' APPTYPE=Access
        With mProgressForm
        ' обрабатываем старший прогресс
            If Len(mMajorText) = 0 Or Len(mMajorBarText) = 0 Then Me.Progress(eMajor) = mMajorValue
        ' обрабатываем младший прогресс
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
' минимальное значение прогресса
' если изменилось минимальное - пересчитываем шаг
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
' максимальное значение прогресса
' если изменилось максимальное - пересчитываем шаг
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
' максимальное значение прогресса
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
' флаг прерывания процесса
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
' для целых чисел
    msk = "#,##0"
    fPart = Abs(rNumber - Fix(rNumber))
    If fPart > 0 Then
' есть дробная часть
        dig = 1 'dig = 2 ' два разряда после запятой
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

