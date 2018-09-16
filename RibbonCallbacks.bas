Attribute VB_Name = "RibbonCallbacks"
Option Explicit

Public myRibbon As IRibbonUI
Const shortCut As String = "�"
Private dicOptions As Dictionary

Public Sub OnLoad(ribbon As IRibbonUI)
  Set myRibbon = ribbon
  'Call LayoutsList
  'Call AddNameForRibbonPointer
  'ThisWorkbook.Names("RibbonPointer").value = ObjPtr(myRibbon)
  
  'Call StartOptions
  'Set dicOptions = RetrieveOptions.Read
  RestoreRibbon
  'SendKeys "%" & shortCut & "{F6}" '"%" & shortCut & "{F6}"
  'SendKeys "%" & shortCut1 & "{F6}"
End Sub

Public Sub UpdateRibbon()
  RestoreOptions
  RestoreRibbon
  myRibbon.Invalidate
  
End Sub

'Private Sub RestoreOptions()
'  Set dicOptions = RetrieveOptions.Read
'End Sub

Private Sub GetVisible(control As IRibbonControl, ByRef visible)
'  Dim ctl As String
'  Dim i As Integer
'  Dim count As Integer
'  ctl = OnlyLiterals(control.id)
'  If dicOptions Is Nothing Then RestoreOptions
'  count = FileBtnCount
'  Select Case ctl
'  Case "btnFile"
'    For i = 1 To count
'      If CInt(Numbers(control.id)) = i Then visible = True
'    Next i
'  End Select
End Sub

Sub GetImage(control As IRibbonControl, ByRef image)
  Dim ctl As String
'  Dim i As Integer
  ctl = ElementID(control.id)
'  If dicOptions Is Nothing Then RestoreOptions
  Select Case ctl
  Case "gpImport" '???
'  Case "mImport"
'  Case "btnImportOptions"
'  Case "gpFiles"
'  Case "btnImport":      image = "ImportExcel" '"XmlImport" '"FileOpen" '"NewFolder" '"DatabaseQueryNew" '"GetExternalDataImportClassic" '"ImportExcel"
'  Case "btnFile":
'    For i = 1 To FileBtnCount
'      If CInt(Numbers(control.id)) = i Then
'        image = Split(dicOptions.Item("file" & i), "|")(4)
'      End If
'    Next i
'  Case "btnFilesOptions"
'  Case "gpFormat"
'  Case "mFormat":        image = "pencil"
'  Case "btnFormat"
'
'  Case "btnFormatOptions"
  End Select
End Sub

Sub GetKeyTip(control As IRibbonControl, ByRef keytip)
  Dim ctl As String
'  Dim i As Integer
'  Dim tempString As String
'
  ctl = ElementID(control.id)
'  'If dicOptions Is Nothing Then RestoreOptions
  Select Case ctl
  Case "tabCB"
    label = "���"
  End Select
End Sub

Sub GetEnabled(control As IRibbonControl, ByRef enabled)
  Dim ctl As String
'  Dim i As Integer
'  Dim tempString As String
'
  ctl = ElementID(control.id)
'  'If dicOptions Is Nothing Then RestoreOptions
  Select Case ctl
  Case "sbImport": enabled = True
  Case "btnImportOptions": enabled = True
  End Select
End Sub

Sub GetLabel(control As IRibbonControl, ByRef label)
  Dim ctl As String
'  Dim i As Integer
'  Dim tempString As String
'
  ctl = ElementID(control.id)
'  'If dicOptions Is Nothing Then RestoreOptions
  Select Case ctl
  Case "tabCB"
    label = "������ ����"
  Case "gpImport": label = "������"
  Case "gpAuthorization": label = "�����������"
  Case "btnLogin": label = "�����"
'  Case "mImport"
'  Case "btnImportOptions"
'  Case "gpFiles"
'    label = "������ � ������"
'  Case "btnImport"
'    For i = 1 To ImportBtnCount
'      If CInt(Numbers(control.id)) = i Then
'        label = Split(dicOptions.Item("import" & i), "|")(0)
'      End If
'    Next i
'  Case "btnFile"
'    For i = 1 To FileBtnCount
'      If CInt(Numbers(control.id)) = i Then
'        label = Split(dicOptions.Item("file" & i), "|")(0)
'      End If
'    Next i
'  Case "btnFilesOptions"
'    label = "������ � ������"
'  Case "gpFormat"
'  Case "mFormat"
'  Case "btnFormat"
'  Case "btnFormatOptions"
'  Case "btnPublicationOptions"
'    label = "������ � ������"
'    Debug.Print "publications"
'
  End Select
End Sub

Sub ButtonOnAction(control As IRibbonControl)
'    Dim ctl As String
'    Dim path As String
'    Dim i As Integer
'    ctl = OnlyLiterals(control.id)
'    If dicOptions Is Nothing Then RestoreOptions
'    Select Case ctl
'    Case "btnDownloadFolder":
'      Call ImportData(Split(dicOptions.Item("UserDownloadFolder"), "|")(0))
'    Case "btnImport":
'      For i = 1 To ImportBtnCount
'        If CInt(Numbers(control.id)) = i Then
'          If Split(dicOptions.Item("import" & i), "|")(1) = "" Then
'              frmHelp.Mode = 2: frmHelp.Show (vbModal)
'          Else: Call ImportData(Split(dicOptions.Item("import" & i), "|")(1))
'          End If
'        End If
'      Next i
'    Case "btnImportOptions": OpenForm (control.id)
'    Case "btnFile"
'      For i = 1 To FileBtnCount
'        If CInt(Numbers(control.id)) = i Then
'          If (i = 1 And Split(dicOptions.Item("file" & i), "|")(0) = "������������� ������") Then
'            frmHelp.Mode = 1
'            frmHelp.Show (vbModal)
'          Else:
'            Call OpenFile(Split(dicOptions.Item("file" & i), "|")(1), _
'                          Split(dicOptions.Item("file" & i), "|")(2), _
'                          Split(dicOptions.Item("file" & i), "|")(3))
'          End If
'        End If
'      Next i
'    Case "btnFilesOptions"
'      Call OpenForm(control.id)
'    Case "btnFormat1"
'    Case "btnFormat2"
'    Case "btnFormat3"
'
'
'    Case "btnFormatOptions"
'    Case "btnPublicationOptions"
'      MsgBox "� ����������"
'    End Select
End Sub

Sub GetScreenTip(control As IRibbonControl, ByRef screentip)
'    Dim ctl As String
'    Dim i As Integer
'    ctl = OnlyLiterals(control.id)
'    If dicOptions Is Nothing Then RestoreOptions
'    Select Case ctl
'    Case "gpImport"
'    Case "mImport"
'    Case "gpFiles"
'
'    Case "btnImportOptions"
'      screentip = "��������� ������� � ����� �������� ������"
'    Case "btnImport"
'      screentip = "������ ������� ������"
'      For i = 1 To ImportBtnCount
'        If CInt(Numbers(control.id)) = i Then
'          If Split(dicOptions.Item("import" & i), "|")(1) = "" Then
'            screentip = "������ ������� ������ (������ ������ ����� ��������� � ���� ��������)"
'          Else:
'            screentip = Split(dicOptions.Item("import" & i), "|")(5)
'          End If
'        End If
'      Next i
'
'    Case "btnFile"
'      For i = 1 To FileBtnCount
'        If CInt(Numbers(control.id)) = i Then
'          If (i = 1 And Split(dicOptions.Item("file" & i), "|")(0) = "������������� ������") Then
'            screentip = "������������� ������ ��� �������� ������"
'          Else:
'            screentip = "������� ���� " & Split(dicOptions.Item("file" & i), "|")(0)
'          End If
'        End If
'      Next i
'    Case "btnFilesOptions"
'      screentip = "��������� ����� � �������� ���������� ������"
'    Case "gpFormat"
'    Case "mFormat"
'
'    Case "btnFormatOptions"
'    Case "btnPublicationOptions"
'      screentip = "��������� ����������� ������ � �������� ����������"
'
'    End Select
'
End Sub

Sub GetSize(control As IRibbonControl, ByRef size)
  Dim ctl As String
'    Dim i As Integer
  ctl = ElementID(control.id)
'    If dicOptions Is Nothing Then RestoreOptions
  Select Case ctl
  Case "sbImport": size = 1
'    'Case "btnImport"
'    Case "btnFile"
'      For i = 1 To FileBtnCount
'        If FileBtnCount < 3 Then
'          size = 1
'        Else: size = 0
'        End If
'      Next i
'    Case "mFormat"
'      size = 1
  End Select
End Sub

Sub GetSuperTip(control As IRibbonControl, ByRef supertip)
'    Dim ctl As String
'    'Dim keys As Variant
'    Dim i As Integer
'    ctl = OnlyLiterals(control.id)
'    If dicOptions Is Nothing Then RestoreOptions
'    Select Case ctl
'    Case "gpImport"
'    Case "mImport"
'    Case "gpFiles"
'    Case "btnImportOptions"
'    Case "btnImport"
'      If control.id = "btnImport1" Then
'        supertip = "������ ������ �� ����� (*.dbf, *.xls*, *.csv)" & vbCrLf & _
'                   "" & vbCrLf & vbCrLf & _
'                   "��������� ����� ����: " & vbCrLf & _
'                   "------------------------------------------------" & vbCrLf & _
'                   "  (1) ����������� ���� ��� ���������� ������;" & vbCrLf & _
'                   "  (2) ������� '�������� ������' � �������� ����" & vbCrLf & _
'                   "-------------------------------------------------" & vbCrLf
'      ElseIf control.id = "btnImport3" Then
'        supertip = "� ������ ����� ������������� " & vbCrLf & _
'        "����������� ��� �����, �������" & vbCrLf & _
'        "���� ������������� � ����"
'      End If
'    Case "btnFile"
'      For i = 1 To FileBtnCount
'        If CInt(Numbers(control.id)) = i Then
'          If (i = 1 And Split(dicOptions.Item("file" & i), "|")(0) = "������������� ������") Then
'            supertip = "�������� ��������� �� 6 ������ � " & vbCrLf & _
'                  "������� ����������� ���� ��������"
'          Else:
'            supertip = "������ ����:" & vbCrLf & Split(dicOptions.Item("file" & Numbers(control.id)), "|")(1)
'          End If
'        End If
'      Next i
'    Case "btnFilesOptions"
'      supertip = "� ������� �������� ����� ��������� �������" & vbCrLf & _
'                 "������ � ���������� ������." & vbCrLf & _
'                 "������������ ���������� ������ - 6." & vbCrLf & _
'                 "��������� ���������:" & vbCrLf & _
'                 "------------------------------------------" & vbCrLf & _
'                 " (1) ��������� ���� � �����;" & vbCrLf & _
'                 " (2) �������� ������ ������;" & vbCrLf & _
'                 " (3) ���������� ���������; " & vbCrLf & _
'                 " (4) ��������� ������" & vbCrLf & _
'                 "------------------------------------------" & vbCrLf
'    Case "gpFormat"
'    Case "mFormat"
'    Case "btnFormat"
'    Case "btnFormatOptions"
'    Case "btnPublicationOptions"
'      supertip = "� ������� �������� �����: " & vbCrLf & _
'                   " - ��������� ��������������� ����������� ������; " & vbCrLf & _
'                   " - ��������� ������ ��������; " & vbCrLf & _
'                   " - ��������� �������� ��������� � " & vbCrLf & _
'                   "   ������ ������..."
'
'
'    End Select
End Sub

Sub OpenForm(ByVal id As String)
  Dim ctl As String
  ctl = OnlyLiterals(id)
  
  Select Case ctl
  Case "btnImportOptions":  frmImportOptions.Show (vbModal)
  Case "btnFilesOptions":  frmFilesOptions.Show (0)
  End Select
  
End Sub



Private Function ElementID(ByVal name As String) As String
  If name <> vbNullString Then ElementID = OnlyLiterals(name)
End Function



