VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GetFuncNe"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim strFn() As String
Dim X As Double
Dim intCount As Integer
Dim intKor As Long
Dim strPromKor As String
Dim strFKor As String
Dim strSKor As String
Dim intPosKor As Long
Dim dblKor As Double
Dim FFunc As Double
Dim SFunc As Double
Dim FFuncSt As Double
Dim SFuncSt As Double
Dim dblStop As Double
Dim dblExct As Double
Dim dblMax As Double
Dim dblMin As Double
Dim dblWork As Double
Dim GetF As New GetFunc

Public Function GetFun(strFunc As String, _
Min As Double, _
Max As Double, _
Optional Step = 0.1, _
Optional Exact = 8) As String

Dim dblWork As Double
dblExct = Exact
intKor = 0
dblStop = Min
dblMax = Max
dblMin = Min
frmWork.Show

'>=
If InStr(strFunc, ">=") <> 0 Then
Fdr:
dblWork = Step
    dblKor = GetFunMaxBR(strFunc, ">=", dblStop, Max, dblWork)
    If dblKor <> 1000000.3 Then intKor = intKor + 1: GetFun = GetFun & dblKor & " "
        If frmWork.blnWork = False Then GetFun = "Прервано_" & frmWork.lblPer & "(" & dblStop & ")_:" & GetFun: GoTo EndFn
    If dblStop < Max + 0.001 Then GoTo Fdr
GoTo EndFn
End If

'<=
If InStr(strFunc, "<=") <> 0 Then
Fmr:
dblWork = Step
    dblKor = GetFunMaxMR(strFunc, "<=", dblStop, Max, dblWork)
    If dblKor <> 1000000.3 Then intKor = intKor + 1: GetFun = GetFun & dblKor & " "
        If frmWork.blnWork = False Then GetFun = "Прервано_" & frmWork.lblPer & "(" & dblStop & ")_:" & GetFun: GoTo EndFn
    If dblStop < Max + 0.001 Then GoTo Fmr
GoTo EndFn
End If

'<>
If InStr(strFunc, "<>") <> 0 Then
Fnr:
dblWork = Step
    dblKor = GetFunMaxNR(strFunc, "<>", dblStop, Max, dblWork)
    If dblKor <> 1000000.3 Then intKor = intKor + 1: GetFun = GetFun & dblKor & " "
            If frmWork.blnWork = False Then GetFun = "Прервано_" & frmWork.lblPer & "(" & dblStop & ")_:" & GetFun: GoTo EndFn
    If dblStop < Max + 0.001 Then GoTo Fnr
GoTo EndFn
End If

'>
If InStr(strFunc, ">") <> 0 Then
Fb:
dblWork = Step
    dblKor = GetFunMaxB(strFunc, ">", dblStop, Max, dblWork)
    If dblKor <> 1000000.3 Then intKor = intKor + 1: GetFun = GetFun & dblKor & " "
            If frmWork.blnWork = False Then GetFun = "Прервано_" & frmWork.lblPer & "(" & dblStop & ")_:" & GetFun: GoTo EndFn
    If dblStop < Max + 0.001 Then GoTo Fb
GoTo EndFn
End If

'<
If InStr(strFunc, "<") <> 0 Then
Fm:
dblWork = Step
    dblKor = GetFunMaxM(strFunc, "<", dblStop, Max, dblWork)
    If dblKor <> 1000000.3 Then intKor = intKor + 1: GetFun = GetFun & dblKor & " "
            If frmWork.blnWork = False Then GetFun = "Прервано_" & frmWork.lblPer & "(" & dblStop & ")_:" & GetFun: GoTo EndFn
    If dblStop < Max + 0.001 Then GoTo Fm
GoTo EndFn
End If

'уравнения
If InStr(strFunc, "=") <> 0 Then
Fu:
dblWork = Step
    dblKor = GetFunMax(strFunc, "=", dblStop, Max, dblWork)
    If dblKor <> 1000000.3 Then intKor = intKor + 1: GetFun = GetFun & dblKor & " "
            If frmWork.blnWork = False Then GetFun = "Прервано_" & frmWork.lblPer & "(" & dblStop & ")_:" & GetFun: GoTo EndFn
    If dblStop < Max + 0.001 Then GoTo Fu
End If
    
  
EndFn:
'создание промежутков
If GetFun = "" Then GoTo Fo
Dim strWork() As String
GetFun = Replace(GetFun, ",", ".")
GetFun = Left(GetFun, Len(GetFun) - 1)
If InStr(GetFun, " ") = 0 Then GoTo Fo
strWork = Split(GetFun, " ")
GetFun = ""
SFunc = Val(strWork(0))
For intCount = 1 To UBound(strWork)
     FFunc = SFunc
     SFunc = Val(strWork(intCount))
     If Abs(FFunc - SFunc) <= 0.0011 Then
        If intPosKor = 0 Then strFKor = FFunc
        intPosKor = intPosKor + 1
        If intPosKor < 3 Then
            strPromKor = strPromKor & SFunc
        Else
            strPromKor = "[" & strFKor & ";" & SFunc & "]"
        End If
    Else
        If intPosKor > 0 Then GetFun = GetFun & strPromKor & " "
        If intPosKor = 0 Then GetFun = GetFun & SFunc & " "
        If intCount = 1 Then GetFun = FFunc & " " & SFunc & " "
        strFKor = ""
        strPromKor = ""
        intPosKor = 0
    End If
Next intCount
If intPosKor > 0 Then GetFun = GetFun & strPromKor

Fo:
'вывод результата
If GetFun = "" Then GetFun = "нет корней"
GetFun = strFunc & vbCrLf & "на_промежутке_[" & Str(Min) & ";" & Str(Max) & "]_с_шагом_" & Str(Step) & "_с_точностью_" & Exact & "_найдено_" & Str(intKor) & "_корней:" & vbCrLf & GetFun
Unload frmWork
End Function

Public Function GetFunCheck(strFun As String) As Boolean
Dim intOper As Integer
If UBound(Split(strFun, ">=")) = 1 Then intOper = intOper + 1: strFun = Replace(strFun, ">=", "")
If UBound(Split(strFun, "<=")) = 1 Then intOper = intOper + 1: strFun = Replace(strFun, "<=", "")
If UBound(Split(strFun, "<>")) = 1 Then intOper = intOper + 1: strFun = Replace(strFun, "<>", "")
If UBound(Split(strFun, "<")) = 1 Then intOper = intOper + 1: strFun = Replace(strFun, "<", "")
If UBound(Split(strFun, ">")) = 1 Then intOper = intOper + 1: strFun = Replace(strFun, ">", "")
If UBound(Split(strFun, "=")) = 1 Then intOper = intOper + 1: strFun = Replace(strFun, "=", "")
If Not (GetF.GetFunCheck(strFun)) Then GetFunCheck = False: Exit Function
If intOper > 1 Or intOper = 0 Then GetFunCheck = False: Exit Function
GetFunCheck = True
End Function

Private Function GetFunMaxBR(strFunc As String, Splt As String, Min As Double, Max As Double, StepM As Double) As Double
strFn = Split(strFunc, Splt)
FFuncSt = GetF.GetFun(strFn(0), Min)
SFuncSt = GetF.GetFun(strFn(1), Min)
If StepM > Max - Min Then StepM = Max - Min
For X = Min To Max Step StepM
        dblWork = Round((X - dblMin) / (dblMax - dblMin) * 100, 2)
        If dblWork >= 100 Then Unload frmWork
        dblWork = Len(Str(StepM)) - 2
       If dblWork < 3 Then dblWork = 3
        X = Round(X, dblWork)
        frmWork.lblPer = dblWork & "%"
                frmWork.prbPer.Value = dblWork: DoEvents
        strFn = Split(strFunc, Splt)
        FFunc = FFuncSt
        FFuncSt = Round(GetF.GetFun(strFn(0), X + StepM), dblExct)
        SFunc = SFuncSt
        SFuncSt = Round(GetF.GetFun(strFn(1), X + StepM), dblExct)
        If FFunc >= SFunc And FFunc <> 1000000.3 And SFunc <> 1000000.3 Then GetFunMaxBR = X: dblStop = X + 0.001: Exit Function
        If FFunc > SFunc And FFuncSt < SFuncSt Or FFunc < SFunc And FFuncSt > SFuncSt Then
            If SFuncSt = 1000000.3 Or FFuncSt = 1000000.3 Or SFunc = 1000000.3 Or FFunc = 1000000.3 Or dblWork > dblExct Then GoTo 1
            dblKor = GetFunMaxBR(strFunc, Splt, X, X + StepM, StepM / 10)
            If dblKor <> 1000000.3 Then GetFunMaxBR = dblKor: dblStop = X + 0.001: Exit Function
1:
        End If
Next X
dblStop = X + 0.001
GetFunMaxBR = 1000000.3
End Function

Private Function GetFunMaxMR(strFunc As String, Splt As String, Min As Double, Max As Double, StepM As Double) As Double
strFn = Split(strFunc, Splt)
FFuncSt = GetF.GetFun(strFn(0), Min)
SFuncSt = GetF.GetFun(strFn(1), Min)
If StepM > Max - Min Then StepM = Max - Min
For X = Min To Max Step StepM
         dblWork = Round((X - dblMin) / (dblMax - dblMin) * 100, 2)
        If dblWork >= 100 Then Unload frmWork
        dblWork = Len(Str(StepM)) - 2
       If dblWork < 3 Then dblWork = 3
        X = Round(X, dblWork)
        frmWork.lblPer = dblWork & "%"
                 frmWork.prbPer.Value = dblWork: DoEvents
        strFn = Split(strFunc, Splt)
        FFunc = FFuncSt
        FFuncSt = Round(GetF.GetFun(strFn(0), X + StepM), dblExct)
        SFunc = SFuncSt
        SFuncSt = Round(GetF.GetFun(strFn(1), X + StepM), dblExct)
        If FFunc <= SFunc And FFunc <> 1000000.3 And SFunc <> 1000000.3 Then GetFunMaxMR = X: dblStop = X + 0.001: Exit Function
        If FFunc > SFunc And FFuncSt < SFuncSt Or FFunc < SFunc And FFuncSt > SFuncSt Then
            If SFuncSt = 1000000.3 Or FFuncSt = 1000000.3 Or SFunc = 1000000.3 Or FFunc = 1000000.3 Or dblWork > dblExct Then GoTo 1
            dblKor = GetFunMaxMR(strFunc, Splt, X, X + StepM, StepM / 10)
            If dblKor <> 1000000.3 Then GetFunMaxMR = dblKor: dblStop = X + 0.001: Exit Function
1:
            End If
Next X
dblStop = X + 0.001
GetFunMaxMR = 1000000.3
End Function

Private Function GetFunMaxNR(strFunc As String, Splt As String, Min As Double, Max As Double, StepM As Double) As Double
strFn = Split(strFunc, Splt)
FFuncSt = GetF.GetFun(strFn(0), Min)
SFuncSt = GetF.GetFun(strFn(1), Min)
If StepM > Max - Min Then StepM = Max - Min
For X = Min To Max Step StepM
         dblWork = Round((X - dblMin) / (dblMax - dblMin) * 100, 2)
        If dblWork >= 100 Then Unload frmWork
        dblWork = Len(Str(StepM)) - 2
       If dblWork < 3 Then dblWork = 3
        X = Round(X, dblWork)
        frmWork.lblPer = dblWork & "%"
                 frmWork.prbPer.Value = dblWork: DoEvents
        strFn = Split(strFunc, Splt)
        FFunc = FFuncSt
        FFuncSt = Round(GetF.GetFun(strFn(0), X + StepM), dblExct)
        SFunc = SFuncSt
        SFuncSt = Round(GetF.GetFun(strFn(1), X + StepM), dblExct)
        If FFunc <> SFunc And FFunc <> 1000000.3 And SFunc <> 1000000.3 Then GetFunMaxNR = X: dblStop = X + 0.001: Exit Function
        If FFunc > SFunc And FFuncSt < SFuncSt Or FFunc < SFunc And FFuncSt > SFuncSt Then
            If SFuncSt = 1000000.3 Or FFuncSt = 1000000.3 Or SFunc = 1000000.3 Or FFunc = 1000000.3 Or dblWork > dblExct Then GoTo 1
            dblKor = GetFunMaxNR(strFunc, Splt, X, X + StepM, StepM / 10)
            If dblKor <> 1000000.3 Then GetFunMaxNR = dblKor: dblStop = X + 0.001: Exit Function
1:
            End If
Next X
dblStop = X + 0.001
GetFunMaxNR = 1000000.3
End Function

Private Function GetFunMaxB(strFunc As String, Splt As String, Min As Double, Max As Double, StepM As Double) As Double
strFn = Split(strFunc, Splt)
FFuncSt = GetF.GetFun(strFn(0), Min)
SFuncSt = GetF.GetFun(strFn(1), Min)
If StepM > Max - Min Then StepM = Max - Min
For X = Min To Max Step StepM
         dblWork = Round((X - dblMin) / (dblMax - dblMin) * 100, 2)
        If dblWork >= 100 Then Unload frmWork
        dblWork = Len(Str(StepM)) - 2
       If dblWork < 3 Then dblWork = 3
        X = Round(X, dblWork)
        frmWork.lblPer = dblWork & "%"
                frmWork.prbPer.Value = dblWork: DoEvents
        strFn = Split(strFunc, Splt)
        FFunc = FFuncSt
        FFuncSt = Round(GetF.GetFun(strFn(0), X + StepM), dblExct)
        SFunc = SFuncSt
        SFuncSt = Round(GetF.GetFun(strFn(1), X + StepM), dblExct)
        If FFunc > SFunc And FFunc <> 1000000.3 And SFunc <> 1000000.3 Then GetFunMaxB = X: dblStop = X + 0.001: Exit Function
        If FFunc > SFunc And FFuncSt < SFuncSt Or FFunc < SFunc And FFuncSt > SFuncSt Then
            If SFuncSt = 1000000.3 Or FFuncSt = 1000000.3 Or SFunc = 1000000.3 Or FFunc = 1000000.3 Or dblWork > dblExct Then GoTo 1
            dblKor = GetFunMaxB(strFunc, Splt, X, X + StepM, StepM / 10)
            If dblKor <> 1000000.3 Then GetFunMaxB = dblKor: dblStop = X + 0.001: Exit Function
1:
        End If
Next X
dblStop = X + 0.001
GetFunMaxB = 1000000.3
End Function

Private Function GetFunMaxM(strFunc As String, Splt As String, Min As Double, Max As Double, StepM As Double) As Double
strFn = Split(strFunc, Splt)
FFuncSt = GetF.GetFun(strFn(0), Min)
SFuncSt = GetF.GetFun(strFn(1), Min)
If StepM > Max - Min Then StepM = Max - Min
For X = Min To Max Step StepM
        dblWork = Round((X - dblMin) / (dblMax - dblMin) * 100, 2)
        If dblWork >= 100 Then Unload frmWork
        frmWork.lblPer = dblWork & "%"
        frmWork.prbPer.Value = dblWork: DoEvents
        dblWork = Len(Str(StepM)) - 2
       If dblWork < 3 Then dblWork = 3
        X = Round(X, dblWork)
        strFn = Split(strFunc, Splt)
        FFunc = FFuncSt
        FFuncSt = Round(GetF.GetFun(strFn(0), X + StepM), dblExct)
        SFunc = SFuncSt
        SFuncSt = Round(GetF.GetFun(strFn(1), X + StepM), dblExct)
        If FFunc < SFunc And FFunc <> 1000000.3 And SFunc <> 1000000.3 Then GetFunMaxM = X: dblStop = X + 0.001: Exit Function
        If FFunc > SFunc And FFuncSt < SFuncSt Or FFunc < SFunc And FFuncSt > SFuncSt Then
            If SFuncSt = 1000000.3 Or FFuncSt = 1000000.3 Or SFunc = 1000000.3 Or FFunc = 1000000.3 Or dblWork > dblExct Then GoTo 1
            dblKor = GetFunMaxM(strFunc, Splt, X, X + StepM, StepM / 10)
            If dblKor <> 1000000.3 Then GetFunMaxM = dblKor: dblStop = X + 0.001: Exit Function
1:
        End If
Next X
dblStop = X + 0.001
GetFunMaxM = 1000000.3
End Function


Private Function GetFunMax(strFunc As String, Splt As String, Min As Double, Max As Double, StepM As Double) As Double
strFn = Split(strFunc, Splt)
FFuncSt = GetF.GetFun(strFn(0), Min)
SFuncSt = GetF.GetFun(strFn(1), Min)
If StepM > Max - Min Then StepM = Max - Min
For X = Min To Max Step StepM
        dblWork = Round((X - dblMin) / (dblMax - dblMin) * 100, 2)
        If dblWork >= 100 Then Unload frmWork
        frmWork.lblPer = dblWork & "%"
        frmWork.prbPer.Value = dblWork: DoEvents
        dblWork = Len(Str(StepM)) - 2
        If dblWork < 3 Then dblWork = 3
        X = Round(X, dblWork)
        strFn = Split(strFunc, Splt)
        FFunc = FFuncSt
        FFuncSt = GetF.GetFun(strFn(0), X + StepM)
        SFunc = SFuncSt
        SFuncSt = GetF.GetFun(strFn(1), X + StepM)
        If Abs(FFunc - SFunc) <= 10 ^ (-dblExct) And FFunc <> 1000000.3 Then GetFunMax = X: dblStop = X + 0.001: Exit Function
        If FFunc > SFunc And FFuncSt < SFuncSt Or FFunc < SFunc And FFuncSt > SFuncSt Then
            If SFuncSt = 1000000.3 Or FFuncSt = 1000000.3 Or SFunc = 1000000.3 Or FFunc = 1000000.3 Or dblWork > dblExct Then GoTo 1
            dblKor = GetFunMax(strFunc, Splt, X, X + StepM, StepM / 10)
            If dblKor <> 1000000.3 Then GetFunMax = dblKor: dblStop = X + 0.001: Exit Function
1:
        End If
Next X
dblStop = X + 0.001
GetFunMax = 1000000.3
End Function
