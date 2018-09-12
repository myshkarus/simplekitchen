Attribute VB_Name = "Constructors"
Option Explicit
Option Private Module

Private ptCollection As ControlCollection
Private rtCollection As ControlCollection

Public Enum table
  product = 1
  Dish = 2
  References = 3
'  TC
'  Cost
'  MenuCard
'  Puchase
'  WriteOff
'  WOutput
'  MainPage
'  WInput
'  Stock
'  DishesBalance
'  Test
End Enum

Public Property Get pCollection() As ControlCollection
  If Not ptCollection Is Nothing Then Set pCollection = ptCollection
End Property

Public Property Set pCollection(ByRef customCollection As ControlCollection)
  Set ptCollection = customCollection
End Property

Public Property Get rCollection() As ControlCollection
  If Not rtCollection Is Nothing Then Set rCollection = rtCollection
End Property

Public Property Set rCollection(ByRef customCollection As ControlCollection)
  Set rtCollection = customCollection
End Property

Public Function NewButtonBuilder(ByVal sheet As Worksheet) As ButtonBuilder
  On Error GoTo FailExit
  If sheet Is Nothing Then Exit Function
  With New ButtonBuilder
    Set .sheet = sheet
    Set NewButtonBuilder = .Self
  End With
  Exit Function
FailExit:
End Function

Public Function NewCheckBoxBuilder(ByVal sheet As Worksheet) As CheckBoxBuilder
  On Error GoTo FailExit
  If sheet Is Nothing Then Exit Function
  With New CheckBoxBuilder
    Set .sheet = sheet
    Set NewCheckBoxBuilder = .Self
  End With
  Exit Function
FailExit:
End Function

Public Function NewLabelBuilder(ByVal sheet As Worksheet) As LabelBuilder
  On Error GoTo FailExit
  If sheet Is Nothing Then Exit Function
  With New LabelBuilder
    Set .sheet = sheet
    Set NewLabelBuilder = .Self
  End With
  Exit Function
FailExit:
End Function

Public Function NewTableBuilder(ByVal sheet As Worksheet) As TableBuilder
  On Error GoTo FailExit
  If sheet Is Nothing Then Exit Function
  With New TableBuilder
    Set .sheet = sheet
    Set NewTableBuilder = .Self
  End With
  Exit Function
FailExit:
End Function
