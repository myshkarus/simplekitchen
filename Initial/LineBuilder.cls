VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LineBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private line As Shape

Public Function DrawLine(ByVal sheet As Worksheet, ByVal begX As Single, ByVal begY As Single, ByVal endX As Single, ByVal endY As Single) As Shape
  Set line = sheet.Shapes.AddLine(begX, begY, endX, endY)
  If Not line Is Nothing Then
    Call LineStyle(csGlowed)
    Set DrawLine = line
  End If
End Function

Private Sub LineStyle(ByVal Style As ControlStyle)
  With line
    Select Case Style
    Case ControlStyle.csGlowed
      .line.ForeColor.RGB = Color.OffBackground
      .Glow.Color.RGB = Color.OffBackground
      .Glow.Transparency = 0.85
      .Glow.Radius = 7
      .Visible = msoTrue
      .line.Weight = 1
      .Placement = xlFreeFloating
    End Select
  End With
End Sub
