VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Green, Bullen, Bovey, Alexander
' Excel 2007 VBA Programmer’s Reference
' (более эффективный, чем пример из Гетц, Гилберт
' Программирование на Visual Basic 6 и VBA)

Option Explicit

#If Win64 Then
  #If VBA7 Then
    Private Declare PtrSafe Function QueryFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (lpFrequency As Currency) As LongPtr
    Private Declare PtrSafe Function QueryCounter Lib "kernel32" Alias "QueryPerformanceCounter" (lpPerformanceCount As Currency) As LongPtr
  #Else
    Private Declare Function QueryFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (lpFrequency As Currency) As LongPtr
    Private Declare Function QueryCounter Lib "kernel32" Alias "QueryPerformanceCounter" (lpPerformanceCount As Currency) As LongPtr
  #End If
#Else
  #If VBA7 Then
    Private Declare PtrSafe Function QueryFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (lpFrequency As Currency) As Long
    Private Declare PtrSafe Function QueryCounter Lib "kernel32" Alias "QueryPerformanceCounter" (lpPerformanceCount As Currency) As Long
  #Else
    Private Declare Function QueryFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (lpFrequency As Currency) As Long
    Private Declare Function QueryCounter Lib "kernel32" Alias "QueryPerformanceCounter" (lpPerformanceCount As Currency) As Long
  #End If
#End If


Private mcyFrequency As Currency
Private mcyOverhead As Currency
Private mcyStarted As Currency
Private mcyStopped As Currency

'When first initialized, determine the overhead incurred when retrieving the
'high-performance counter value
Private Sub Class_Initialize()
    Dim cyCount1 As Currency, cyCount2 As Currency
    'Get the counter frequency
    QueryFrequency mcyFrequency
    'Call the hi-res counter twice, to check how long it takes
    QueryCounter cyCount1
    QueryCounter cyCount2
    'Store the call overhead
    mcyOverhead = cyCount2 - cyCount1
End Sub

Public Sub StartTimer()
    'Get the time that you started
    QueryCounter mcyStarted
End Sub

Public Sub StopTimer()
    'Get the time that you stopped
    QueryCounter mcyStopped
    Debug.Print "time of execution: " & Round(Elapsed, 6) & " сек."
End Sub

Public Property Get Elapsed() As Double
    Dim cyTimer As Currency
    'Have you stopped or not?
    If mcyStopped = 0 Then
        QueryCounter cyTimer
    Else
        cyTimer = mcyStopped
    End If
    
    'If you have a frequency, return the duration, in seconds
    If mcyFrequency > 0 Then
        Elapsed = (cyTimer - mcyStarted - mcyOverhead) / mcyFrequency
    End If
End Property
