VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ADODictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If Development Then
  'for early binding
  Private dic As Scripting.Dictionary
#Else
  'for late binding
  Private dic As Object
#End If

Private Sub Class_Initialize()
  #If Development Then
    'early binding
    Set dic = New Scripting.Dictionary
  #Else
    'late binding
    Set dic = CreateObject("Scripting.Dictionary")
  #End If
  dic.CompareMode = vbTextCompare
End Sub

Private Sub Class_Terminate()
  If Not dic Is Nothing Then: Set dic = Nothing
  If Not tempRs Is Nothing Then: Set tempRs = Nothing
End Sub

Public Property Set rs(inputrs As Object)
  Set myrecordset = inputrs
End Property

Public Property Get mydictionary() As Scripting.Dictionary
  Call UserDL
  Set mydictionary = dic
End Property

Private Sub UserDL()
  '  Dim USList As Range
  '  Dim USArr(0 To 11) As Variant
  Dim sKey As String
  Dim item As String
  Dim UserList As Scripting.Dictionary

  Set UserList = New Scripting.Dictionary
  '  Call ConnecttoDB

  '  Set cmd = New ADODB.Command: Set rs = New ADODB.RecordSet

  '  With cmd
  '    .CommandTimeout = 30
  '    .ActiveConnection = CN
  '    .CommandText = "CSLL.DLUsers"
  '
  '    Set rs = .Execute
  '  End With

  With myrecordset
    If Not .BOF And Not .EOF Then
      .MoveLast
      .MoveFirst
      While (Not .EOF)
        'For i = 1 To 11
        '  USArr(i - 1) = rs(i)
        'Next i
        With UserList
          sKey = myrecordset("Key")
          item = myrecordset("Item")
          'If Not .Exists(sKey) Then
          .Add sKey, item
          'End If
        End With
        .MoveNext
      Wend
    End If
  End With


  Set dic = UserList
  '  IA = UserList.Items
  '  Debug.Print UserList.Count & " Items in the dictionary"
  '  For Each element In IA
  '    For i = 0 To 10
  '      Debug.Print element(i)
  '    Next i
  '  Next element
  '
  '  Set cmd = Nothing: Set rs = Nothing            ': Set UserList = Nothing
End Sub

Private Property Set IADictionary_rs(ByRef inputrs As Object)
  Err.Raise 5                                    'TODO implement interface member
End Property

Private Property Get IADictionary_mydictionary() As Scripting.Dictionary
  Err.Raise 5                                    'TODO implement interface member
End Property

