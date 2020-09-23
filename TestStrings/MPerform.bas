Attribute VB_Name = "MPerform"
Option Explicit

Public secIn As Currency, secOut As Currency

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'From Hardcore VB
Private secFreq As Currency
Public Sub ProfileStart(secStart As Currency)
    If secFreq = 0 Then QueryPerformanceFrequency secFreq
    QueryPerformanceCounter secStart
End Sub

Public Sub ProfileStop(secStart As Currency, secTiming As Currency)
    QueryPerformanceCounter secTiming
    If secFreq = 0 Then
        secTiming = 0 ' Handle no high-resolution timer
    Else
        secTiming = (secTiming - secStart) / secFreq
    End If
End Sub

