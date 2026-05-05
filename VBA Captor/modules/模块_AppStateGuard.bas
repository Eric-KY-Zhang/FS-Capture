Attribute VB_Name = "模块_AppStateGuard"
Option Explicit

' Phase 4k Step 2: 全 Excel 状态守护,任何入口宏报错时一次性恢复
Public Type TAppState
    ScreenUpdating As Boolean
    DisplayAlerts As Boolean
    DisplayStatusBar As Boolean
    EnableEvents As Boolean
    Calculation As XlCalculation
    StatusBarValue As Variant
    DisplayPageBreaks As Boolean
    HasActiveSheet As Boolean
End Type


Public Function BeginAppState(Optional ByVal statusText As String = "") As TAppState
    Dim st As TAppState
    st.ScreenUpdating = Application.ScreenUpdating
    st.DisplayAlerts = Application.DisplayAlerts
    st.DisplayStatusBar = Application.DisplayStatusBar
    st.EnableEvents = Application.EnableEvents
    st.Calculation = Application.Calculation
    st.StatusBarValue = Application.StatusBar
    st.HasActiveSheet = Not (ActiveSheet Is Nothing)
    If st.HasActiveSheet Then st.DisplayPageBreaks = ActiveSheet.DisplayPageBreaks

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = True
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    If st.HasActiveSheet Then ActiveSheet.DisplayPageBreaks = False
    If Len(statusText) > 0 Then Application.StatusBar = statusText

    BeginAppState = st
End Function


Public Sub EndAppState(ByRef st As TAppState)
    On Error Resume Next
    Application.Calculation = st.Calculation
    Application.EnableEvents = st.EnableEvents
    Application.DisplayStatusBar = st.DisplayStatusBar
    Application.DisplayAlerts = st.DisplayAlerts
    Application.ScreenUpdating = st.ScreenUpdating
    Application.StatusBar = st.StatusBarValue
    If st.HasActiveSheet Then ActiveSheet.DisplayPageBreaks = st.DisplayPageBreaks
    Err.Clear
    On Error GoTo 0
End Sub
