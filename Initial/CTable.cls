VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ClassName As String = "CTable"
Private tbl As ListObject

Private Type tControl
  tID As String
  tTabIndex As Long                              'last Characters of AlternativeText
  tStyle As ControlStyle                         'do not restore!
  tSheet As Worksheet
  tMark As String                                'first Character of AlternativeText
  TName As String
  tFontSize As Long
  tAlignment As Variant
  tHeader As Variant
  tRange As Range
End Type

Private this As tControl
Implements IControl

Public Property Set sheet(ByVal wsh As Worksheet)
  If Not wsh Is Nothing Then Set this.tSheet = wsh
End Property

Private Property Get IControl_ID() As String
  IControl_ID = this.tID
End Property

Private Property Get IControl_ControlType() As String
  IControl_ControlType = TypeName(Self)
End Property

Private Property Get IControl_Name() As String
  If Not tbl Is Nothing Then IControl_Name = tbl.Name
End Property

Private Property Let IControl_Name(ByVal text As String)
  this.TName = text
  If text <> vbNullString And Not tbl Is Nothing Then tbl.Name = text
End Property

Private Property Let IControl_Caption(ByVal text As String)
  Caption = text
End Property

Private Property Get IControl_Mark() As String
  If Not tbl Is Nothing Then
    With tbl
      IControl_Mark = left$(.AlternativeText, IIf(Len(.AlternativeText) > 1, 1, 0))
    End With
  End If
End Property

Private Property Get IControl_TabIndex() As Long
  If Not tbl Is Nothing Then
    With tbl
      IControl_TabIndex = Right$(.AlternativeText, IIf(Len(.AlternativeText) > 1, Len(.AlternativeText) - 1, Len(.AlternativeText)))
    End With
  End If
End Property

Private Property Let IControl_TabIndex(ByVal value As Long)
  this.tTabIndex = value
End Property

'для CTable установлено условно
Private Property Get IControl_Visible() As Boolean
  IControl_Visible = this.tVisible
End Property

'для CTable установлено условно
Private Property Let IControl_Visible(ByVal value As Boolean)
  this.tVisible = value
End Property

Private Property Get IControl_Style() As ControlStyle
  IControl_Style = this.tStyle
End Property

Private Property Let IControl_Style(ByVal value As ControlStyle)
  this.tStyle = value
  If Not tbl Is Nothing Then Call SetupTableStyle(value)
End Property

Private Property Let IControl_Movable(ByVal status As Boolean)
End Property

Private Property Get IControl_Height() As Single
  If Not tbl Is Nothing Then IControl_Height = tbl.Range.Height
End Property

Private Property Get IControl_Width() As Single
  If Not tbl Is Nothing Then IControl_Width = tbl.Range.Width
End Property

Private Property Get IControl_Top() As Single
  If Not tbl Is Nothing Then IControl_Top = tbl.Range.top
End Property

Private Property Get IControl_Left() As Single
  If Not tbl Is Nothing Then IControl_Left = tbl.Range.left
End Property

Private Property Get IControl_Bottom() As Single
  If Not tbl Is Nothing Then IControl_Bottom = tbl.Range.top + tbl.Range.Height
End Property

Private Property Get IControl_Right() As Single
  If Not tbl Is Nothing Then IControl_Right = tbl.Range.left + tbl.Range.Width
End Property

Private Property Get IControl_BottomRightRange() As Range
  If Not tbl Is Nothing Then Set IControl_BottomRightRange = tbl.Range.Offset(1, 0).Resize(1, 1)
End Property

Private Property Get IControl_BottomLeftRange() As Range
  If Not tbl Is Nothing Then Set IControl_BottomLeftRange = tbl.Range.Offset(-1, 0)
End Property

Public Property Get FontSize() As Long
  If Not tbl Is Nothing Then
    FontSize = tbl.DataBodyRange.Font.size
  Else: FontSize = this.tFontSize
  End If
End Property

Public Property Let FontSize(ByVal value As Long)
  this.tFontSize = value
  If Not tbl Is Nothing Then
    If value = 0 Then
      tbl.DataBodyRange.Font.size = FontSizeEnum.Standard
    Else: tbl.DataBodyRange.Font.size = value
    End If
  End If
End Property

Public Property Get Alignment() As Variant
  If Not tbl Is Nothing Then
    Alignment = tbl.DataBodyRange.HorizontalAlignment
  Else: Alignment = this.tAlignment
  End If
End Property

Public Property Let Alignment(ByVal val As Variant)
  If Not tbl Is Nothing Then
    If IsEmpty(val) Then
      tbl.DataBodyRange.HorizontalAlignment = xlLeft
    Else
      tbl.DataBodyRange.HorizontalAlignment = val
      If val = xlRight Then tbl.DataBodyRange.IndentLevel = 1
    End If
  Else: this.tAlignment = val
  End If
End Property

Public Property Let Mark(ByVal text As String)
  this.tMark = text
End Property

Public Property Let Header(ByVal title As Variant)
  If Not IsEmpty(title) Then this.tHeader = title
End Property

Public Property Get CRange() As Range
  If Not tbl Is Nothing Then Set CRange = tbl.Range
End Property

Public Property Set CRange(ByVal rng As Range)
  If Not rng Is Nothing Then
    Set this.tRange = rng
  End If
End Property

Private Function ToString() As String
  ToString = ClassName & " [id=" & ID & "] " & " to be build"
End Function

Public Property Get Self() As IControl
  Set Self = Me
End Property

Private Sub Class_Terminate()
  Set tbl = Nothing
  Set this.tSheet = Nothing
End Sub

Private Sub SetupTableStyle(Optional ByVal Style As ControlStyle)
  If Not tbl Is Nothing Then
    With tbl
      this.tSheet.Range(.Range.address).Interior.Color = xlNone
      .TableStyle = "TableStyleMedium23"
      .HeaderRowRange.Font.Name = "Calibri Light"
      .HeaderRowRange.Font.size = FontSizeEnum.Medium
      .HeaderRowRange.Font.Color = Color.TableTitleBackground
      .HeaderRowRange.HorizontalAlignment = xlCenter
      .DataBodyRange.Locked = False
      .DataBodyRange.Font.Color = Color.OffBackground
    End With
  End If
End Sub

Private Sub IControl_Create(ByVal wsh As Worksheet)
  Dim i As Integer
  Dim rngcol() As Single
  On Error Resume Next
  If Not wsh Is Nothing Then Set this.tSheet = wsh
  If this.tRange Is Nothing Then Exit Sub
  
  ReDim rngcol(this.tRange.columns.count)
  For i = 0 To this.tRange.columns.count - 1
    rngcol(i) = this.tRange.columns(i + 1).ColumnWidth
  Next i
  Set tbl = this.tSheet.ListObjects.Add(xlSrcRange, this.tRange)
    With tbl
    If this.TName <> vbNullString Then tbl.Name = this.TName
        For i = 0 To UBound(this.tHeader)
            .ListColumns(i + 1).Name = this.tHeader(i)
            .ListColumns(i + 1).Range.ColumnWidth = rngcol(i)
        Next i
    End With
  With tbl

    If Not IsEmpty(this.tAlignment) Then .DataBodyRange.HorizontalAlignment = this.tAlignment
    
    Call SetupTableStyle(this.tStyle)
    
    If this.tFontSize = 0 Then
      .DataBodyRange.Font.size = FontSizeEnum.Standard
    Else: .DataBodyRange.Font.size = this.tFontSize
    End If
    
    .AlternativeText = left$(this.tMark, 1) & CStr(this.tTabIndex)

  End With
End Sub

Private Sub Class_Initialize()
  this.tID = Hash(TypeName(Me) & CLng(Rnd * (2 ^ 31)))
End Sub
