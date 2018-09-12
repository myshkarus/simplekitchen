VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LabelBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ClassName As String = "LabelBuilder"

Private Type TAttribute
  TName As String
  tRange As Range
  tVisible As Boolean
  tCaption As String
  tStyle As ControlStyle
  'tID As String
  tTabIndex As Long
  tIcon As String
  tMark As String
  tAlignment As MsoParagraphAlignment
  tFontSize As Long
  tFieldStartOffset As Integer
  tFieldColumns As Long
  tFieldRows As Long
  tFieldRangeName As String
  tFieldList As Boolean
  tFieldFlag As Boolean
  tFieldBorder As Boolean
End Type

Private this As TAttribute
Private thisReset As TAttribute
Private mSheet As Worksheet

Public Function Self() As LabelBuilder
  Set Self = Me
End Function

Private Function myLabel() As IControl
  Dim lbl As IControl
  With New CLabel
    .Caption = this.tCaption
    .Alignment = this.tAlignment
    .Mark = this.tMark
    .FontSize = this.tFontSize
    .Icon = this.tIcon
    .FieldBorder = this.tFieldBorder
    .FieldColumns = this.tFieldColumns
    .FieldFlag = this.tFieldFlag
    .FieldList = this.tFieldList
    .FieldName = this.tFieldRangeName
    .FieldOffset = this.tFieldStartOffset
    .FieldRows = this.tFieldRows
    Set .CRange = this.tRange
    Set lbl = .Self
  End With
  With lbl
    .Name = this.TName
    .Style = this.tStyle
    .TabIndex = this.tTabIndex
    .Visible = this.tVisible
    .Create mSheet
  End With
  Set myLabel = lbl
  this = thisReset
  Set lbl = Nothing
End Function

Public Function Build() As IControl
  Set Build = myLabel
End Function

Public Function OfPlaceAndSize(ByVal rng As Range) As LabelBuilder
  If Not rng Is Nothing Then Set this.tRange = rng
  Set OfPlaceAndSize = Self
End Function

Public Function Field(Optional ByVal startOffset As Integer, Optional ByVal columns As Long, Optional ByVal rows As Long, Optional ByVal rangeName As String, _
                      Optional ByVal list As Boolean, Optional ByVal flag As Boolean, Optional ByVal border As Boolean) As LabelBuilder
  With this
    If startOffset <> 0 Then
      .tFieldStartOffset = startOffset
    Else: .tFieldStartOffset = 1
    End If
    If columns <> 0 Then
      .tFieldColumns = columns
    Else: .tFieldColumns = 1
    End If
    If rows <> 0 Then
      .tFieldRows = rows
    Else: .tFieldRows = 1
    End If
    If rangeName <> vbNullString Then
      .tFieldRangeName = rangeName
    End If
    .tFieldBorder = border
    .tFieldFlag = flag
    .tFieldList = list
  End With
  Set Field = Self
End Function

Public Function Name(ByVal text As String) As LabelBuilder
  If text = vbNullString Then Exit Function
  this.TName = text
  Set Name = Self
End Function

Public Function SetIcon(ByVal Icon As String) As LabelBuilder
  If Icon = vbNullString Then Exit Function
  this.tIcon = Icon
  Set SetIcon = Self
End Function

'Public Function Visible(ByVal value As Boolean) As LabelBuilder
'  this.tVisible = value
'  Set Visible = Self
'End Function

Public Function Style(ByVal value As ControlStyle) As LabelBuilder
  this.tStyle = value
  Set Style = Self
End Function

Public Function Caption(ByVal text As String) As LabelBuilder
  If text = vbNullString Then Exit Function
  this.tCaption = text
  Set Caption = Self
End Function

Public Function FontSize(ByVal value As Variant) As LabelBuilder
  this.tFontSize = value
  Set FontSize = Self
End Function

Public Function Alignment(ByVal val As MsoParagraphAlignment) As LabelBuilder
  this.tAlignment = val
  Set Alignment = Self
End Function

Public Function ToString() As String
  ToString = ClassName & " " & "to be build"
End Function

Public Property Get sheet() As Worksheet
  Set sheet = mSheet
End Property

Friend Property Set sheet(ByVal wsh As Worksheet)
  If Not wsh Is Nothing Then Set mSheet = wsh
End Property

Public Function Mark(ByVal val As String) As LabelBuilder
  this.tMark = val
  Set Mark = Self
End Function

Private Sub Class_Terminate()
  Set mSheet = Nothing
  Set this.tRange = Nothing
End Sub

