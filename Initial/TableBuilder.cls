VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ClassName As String = "TableBuilder"

Private Type TAttribute
  TName As String '+
  tRange As Range '+
  tHeader As Variant '+
  tStyle As ControlStyle '+
  tTabIndex As Long
  tMark As String
  tAlignment As MsoParagraphAlignment 'variant ?
  tFontSize As Long
End Type

Private this As TAttribute
Private thisReset As TAttribute
Private mSheet As Worksheet

Public Function Self() As TableBuilder
  Set Self = Me
End Function

Private Function myTable() As IControl
  Dim tbl As IControl
  With New CTable
    .Alignment = this.tAlignment
    .Mark = this.tMark
    .FontSize = this.tFontSize
    .Header = this.tHeader
    Set .CRange = this.tRange
    Set tbl = .Self
  End With
  With tbl
    .Name = this.TName
    .Style = this.tStyle
    .TabIndex = this.tTabIndex
    '.Visible = this.tVisible
    .Create mSheet
  End With
  Set myTable = tbl
  this = thisReset
  Set tbl = Nothing
End Function

Public Function Build() As IControl
  Set Build = myTable
End Function

Public Function OfPlaceAndSize(ByVal rng As Range) As TableBuilder
  If Not rng Is Nothing Then Set this.tRange = rng
  Set OfPlaceAndSize = Self
End Function

Public Function Name(ByVal text As String) As TableBuilder
  If text = vbNullString Then Exit Function
  this.TName = text
  Set Name = Self
End Function

Public Function Style(Optional ByVal value As ControlStyle) As TableBuilder
  this.tStyle = value
  Set Style = Self
End Function

Public Function Header(ByVal text As Variant) As TableBuilder
  If IsEmpty(text) Then Exit Function
  this.tHeader = text
  Set Header = Self
End Function

Public Function FontSize(ByVal value As Variant) As TableBuilder
  this.tFontSize = value
  Set FontSize = Self
End Function

'variant?
Public Function Alignment(ByVal val As MsoParagraphAlignment) As TableBuilder
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

