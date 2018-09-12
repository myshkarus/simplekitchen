VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Controls As Collection

Public Property Get Self() As ControlCollection
  Set Self = Me
End Property

Public Sub Add(ByVal ctrl As IControl, Optional ByVal varBefore As Variant)
On Error GoTo FailExit
If Not ctrl Is Nothing Then
    Controls.Add ctrl, ctrl.ID, varBefore
End If
Exit Sub
FailExit:
End Sub

Public Sub Remove(ByVal varID As Variant)
On Error GoTo FailExit
Controls.Remove varID
Exit Sub
FailExit:
End Sub

Public Function ItemByName(ByVal Name As String) As IControl
Dim control As IControl
On Error GoTo FailExit
  For Each control In Controls
    If control.Name = Name Then
      Set ItemByName = Controls.item(control.ID)
      'Debug.Print TypeName(control)
      Exit Function
    End If
  Next control
Exit Function
FailExit:
End Function

Public Function item(ByVal varID As Variant) As IControl
Set item = Controls(varID)
End Function

Property Get count() As Long
count = Controls.count
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Set NewEnum = Controls.[_NewEnum]
End Function

Private Sub Class_Initialize()
Set Controls = New Collection
End Sub

Private Sub Class_Terminate()
Set Controls = Nothing
End Sub
