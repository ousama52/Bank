Attribute VB_Name = "modPerformance"
Option Explicit

Private m_isOff As Boolean
Private m_calc As Long
Private m_events As Boolean
Private m_screen As Boolean

Public Sub TurnOff()
    ' Remember the user's settings only once, so a second call cannot overwrite them with "off".
    If Not m_isOff Then
        m_calc = Application.Calculation
        m_events = Application.EnableEvents
        m_screen = Application.ScreenUpdating
        m_isOff = True
    End If
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Public Sub TurnOn()
    If m_isOff Then
        Application.Calculation = m_calc
        Application.EnableEvents = m_events
        Application.ScreenUpdating = m_screen
        m_isOff = False
    Else
        Application.ScreenUpdating = True
        Application.EnableEvents = True
    End If
    Application.StatusBar = False
End Sub
