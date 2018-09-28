Attribute VB_Name = "modWaitForFileClose"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modWaitForFileClose
' By Chip Pearson, www.cpearson.com chip@cpearson.com
'
' This module contains the WaitForFileClose and IsFileOpen functions.
' See http://www.cpearson.com/excel/WaitForFileClose.htm for more documentation.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''
' Windows API Declares
''''''''''''''''''''''
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long


Public Function WaitForFileClose(FileName As String, ByVal TestIntervalMilliseconds As Long, _
    ByVal TimeOutMilliseconds As Long) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WaitForFileClose
' This function tests to see if a specified file is open. If the file is not
' open, it returns a value of True and exists immediately. If FileName is
' open, the code goes into a wait loop, testing whether the is still open
' every TestIntervalMilliSeconds. If the is closed while the function is
' waiting, the function exists with a result of True. If TimeOutMilliSeconds
' is reached and file remains open, the function exits with a result of
' False. The function will return True is FileName does not exist.
' If TimeOutMilliSeconds is reached and the file remains open, the function
' returns False.
' If FileName refers to a workbook that is open Shared, the function returns
' True and exits immediately.
' This function requires the IsFileOpen function and the Sleep and GetTickCount
' API functions.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim StartTickCount As Long
Dim EndTickCount As Long
Dim TickCountNow As Long
Dim FileIsOpen As Boolean
Dim Done As Boolean
Dim CancelKeyState As Long

'''''''''''''''''''''''''''''''''''''''''''''''
' Before we do anything, first test if the file
' is open. If it is not, get out immediately.
'''''''''''''''''''''''''''''''''''''''''''''''
FileIsOpen = IsFileOpen(FileName:=FileName)
If FileIsOpen = False Then
    WaitForFileClose = True
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' If TestIntervalMilliseconds <= 0, use a default value of 500.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If TestIntervalMilliseconds <= 0 Then
    TestIntervalMilliseconds = 500
End If


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Here, we save the state of EnableCancelKey, and set it to
' xlErrorHandler. This will cause an error 18 to raised if the
' user press CTLR+BREAK. In this case, we'll abort the wait
' procedure and return False.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CancelKeyState = Application.EnableCancelKey
Application.EnableCancelKey = xlErrorHandler
On Error GoTo ErrHandler:

'''''''''''''''''''''''''''''''
' Get the current tick count.
'''''''''''''''''''''''''''''''
StartTickCount = GetTickCount()
If TimeOutMilliseconds <= 0 Then
    ''''''''''''''''''''''''''''''''''''''''
    ' If TimeOutMilliSeconds is negative,
    ' we'll wait forever.
    ''''''''''''''''''''''''''''''''''''''''
    EndTickCount = -1
Else
    ''''''''''''''''''''''''''''''''''''''''
    ' If TimeOutMilliseconds > 0, get the
    ' tick count value at which we will
    ' give up on the wait and return
    ' false.
    ''''''''''''''''''''''''''''''''''''''''
    EndTickCount = StartTickCount + TimeOutMilliseconds
End If

Done = False
Do Until Done
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Test if the file is open. If it is closed,
    ' exit with a result of True.
    ''''''''''''''''''''''''''''''''''''''''''''''''
    If IsFileOpen(FileName:=FileName) = False Then
        WaitForFileClose = True
        Application.EnableCancelKey = CancelKeyState
        Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''''''''
    ' Go to sleep for TestIntervalMilliSeconds
    ' milliseconds.
    '''''''''''''''''''''''''''''''''''''''''
    Sleep dwMilliseconds:=TestIntervalMilliseconds
    TickCountNow = GetTickCount()
    If EndTickCount > 0 Then
        '''''''''''''''''''''''''''''''''''''''''''''
        ' If EndTickCount > 0, a specified timeout
        ' value was provided. Test if we have
        ' exceeded the time. Do one last test for
        ' FileOpen, and exit.
        '''''''''''''''''''''''''''''''''''''''''''
        If TickCountNow >= EndTickCount Then
            WaitForFileClose = Not (IsFileOpen(FileName))
            Application.EnableCancelKey = CancelKeyState
            Exit Function
        Else
            '''''''''''''''''''''''''''''''''''''''''
            ' TickCountNow is less than EndTickCount,
            ' so continue to wait.
            '''''''''''''''''''''''''''''''''''''''''
        End If
    Else
        ''''''''''''''''''''''''''''''''
        ' EndTickCount < 0, meaning wait
        ' forever. Test if the file
        ' is open. If the file is not
        ' open, exit with a TRUE result.
        ''''''''''''''''''''''''''''''''
        If IsFileOpen(FileName:=FileName) = False Then
            WaitForFileClose = True
            Application.EnableCancelKey = CancelKeyState
            Exit Function
        End If
        
    End If
    DoEvents
Loop
Exit Function
ErrHandler:
'''''''''''''''''''''''''''''''''''
' This is the error handler block.
' For any error, return False.
'''''''''''''''''''''''''''''''''''
Application.EnableCancelKey = CancelKeyState
WaitForFileClose = False

End Function


Private Function IsFileOpen(FileName As String) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsFileOpen
' By Chip Pearson www.cpearson.com/excel chip@cpearson.com
' This function determines whether a file is open by any program. Returns TRUE or FALSE
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim FileNum As Integer
Dim ErrNum As Integer

On Error Resume Next   ' Turn error checking off.

'''''''''''''''''''''''''''''''''''''''''''
' If we were passed in an empty string,
' there is no file to test so return FALSE.
'''''''''''''''''''''''''''''''''''''''''''
If FileName = vbNullString Then
    IsFileOpen = False
    Exit Function
End If

'''''''''''''''''''''''''''''''
' If the file doesn't exist,
' it isn't open so get out now.
'''''''''''''''''''''''''''''''
If Dir(FileName) = vbNullString Then
    IsFileOpen = False
    Exit Function
End If
''''''''''''''''''''''''''
' Get a free file number.
''''''''''''''''''''''''''
FileNum = FreeFile()
'''''''''''''''''''''''''''
' Attempt to open the file
' and lock it.
'''''''''''''''''''''''''''
Err.Clear
Open FileName For Input Lock Read As #FileNum
''''''''''''''''''''''''''''''''''''''
' Save the error number that occurred.
''''''''''''''''''''''''''''''''''''''
ErrNum = Err.Number
On Error GoTo 0        ' Turn error checking back on.
Close #FileNum       ' Close the file.
''''''''''''''''''''''''''''''''''''
' Check to see which error occurred.
''''''''''''''''''''''''''''''''''''
Select Case ErrNum
    Case 0
    '''''''''''''''''''''''''''''''''''''''''''
    ' No error occurred.
    ' File is NOT already open by another user.
    '''''''''''''''''''''''''''''''''''''''''''
        IsFileOpen = False

    Case 70
    '''''''''''''''''''''''''''''''''''''''''''
    ' Error number for "Permission Denied."
    ' File is already opened by another user.
    '''''''''''''''''''''''''''''''''''''''''''
        IsFileOpen = True

    '''''''''''''''''''''''''''''''''''''''''''
    ' Another error occurred. Assume the file
    ' cannot be accessed.
    '''''''''''''''''''''''''''''''''''''''''''
    Case Else
        IsFileOpen = True
        
End Select

End Function




