VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Custom extention class for Scripting.Dictionary"
Option Explicit

'-------------------------------------------------------------
'Ver.2.0 7/1/2018
'without values as Object
'-------------------------------------------------------------
Private currentDelimiter As String
Private Const STANDARDDELIMITER = "|"

#If Development Then
  'for early binding
  Private dic As Scripting.Dictionary
#Else
  'for late binding
  Private dic As Object
#End If

Private Sub Class_Initialize()
  On Error GoTo ErrHandler
  #If Development Then
    'for early binding
    Set dic = New Scripting.Dictionary
  #Else
    'for late binding
    Set dic = CreateObject("Scripting.Dictionary")
  #End If
  dic.CompareMode = vbBinaryCompare 'vbTextCompare
  Exit Sub
ErrHandler:
End Sub

Private Sub Class_Terminate()
  If Not dic Is Nothing Then Set dic = Nothing
End Sub

Public Property Get Self() As Dictionary
  Set Self = Me
End Property

Property Get size() As Long
  size = dic.count
End Property

Property Let delimiter(ByVal char As String)
Attribute delimiter.VB_Description = "Get or set delimiter for complex item. Delimiter by default is ""|"""
On Error GoTo ErrHandler
If size > 0 Then Exit Property
If Len(char) <> 1 Then GoTo ErrHandler
If char = vbNullString Then
    currentDelimiter = STANDARDDELIMITER
  Else: currentDelimiter = char
  End If
  Exit Property
ErrHandler:
End Property

Property Get delimiter() As String
Attribute delimiter.VB_Description = "Get or set delimiter for complex item. Delimiter by default is ""|"""
If currentDelimiter <> vbNullString Then
delimiter = currentDelimiter
Else: delimiter = STANDARDDELIMITER
End If
End Property

Property Get dimension() As Integer
Attribute dimension.VB_Description = "Get dimension of complex item. Depends on delimiter in use"
  Dim firstEntry As String
  firstEntry = item(keys()(0))
  dimension = CountDelimiters(firstEntry, delimiter) + 1
End Property

Property Get keys() As Variant()
Attribute keys.VB_Description = "Get array of all keys in the dictionary"
  keys = dic.keys
End Property

Property Get item(ByVal key As Variant) As Variant
Attribute item.VB_Description = "Get item for a given key"
Attribute item.VB_UserMemId = 0
  item = dic.item(key)
End Property

Property Get ObjectItem(ByVal key As Variant) As Object
Attribute ObjectItem.VB_Description = "Get object item for a given key"
  Set ObjectItem = dic.item(key)
End Property

Property Get Items() As Variant()
  Items = dic.Items
End Property

Property Get SubItem(ByVal key As Variant, ByVal position As Integer) As Variant
Dim foundItem As Variant
If position > dimension Then GoTo ErrHandler
foundItem = RetrieveSubItem(Me.item(key), position)
If Not IsEmpty(foundItem) Then SubItem = foundItem
Exit Property
ErrHandler:
End Property

Property Get Exists(ByVal key As Variant) As Boolean
Attribute Exists.VB_Description = "Check if the key exist in the dictionary"
  Exists = dic.Exists(key)
End Property

Public Sub AddObject(ByVal key As Variant, obj As Object)
  On Error Resume Next
  If Not Exists(key) Then dic.Add key, obj
End Sub

Public Sub Add(ByVal value As Variant, Optional ByVal key As Variant)
  On Error GoTo ErrHandler
  If IsMissing(key) Then key = value
  dic.item(key) = value
  Exit Sub
ErrHandler:
End Sub

'removing only by key
Public Sub Remove(ByVal key As Variant)
  On Error GoTo ErrHandler
  If Exists(key) Then dic.Remove (key)
  Exit Sub
ErrHandler:
End Sub

Public Sub RemoveAll()
  On Error GoTo ErrHandler
  dic.RemoveAll
  Exit Sub
ErrHandler:
End Sub

Private Function CountDelimiters(ByVal entry As String, usedDelimiter As String) As Integer
  Dim count As Integer
  Dim Place As Integer: Place = 1
  Do
    Place = InStr(Place, entry, usedDelimiter, vbBinaryCompare)
    If Place > 0 Then
      count = count + 1
      Place = Place + 1
    End If
  Loop While Place > 0
  CountDelimiters = count
End Function

Private Function RetrieveSubItem(ByVal item As Variant, ByVal subItemPos As Integer) As Variant
  Dim pos As Integer
  On Error GoTo ErrHandler
  'SubItem first position is 1
  pos = subItemPos - 1
  RetrieveSubItem = Split(item, delimiter)(pos)
  Exit Function
ErrHandler:
End Function
