Attribute VB_Name = "RibbonCallbacks"
Option Explicit

Public myRibbon As IRibbonUI
Const shortCut As String = "Б"
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
    label = "ЁПТ"
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
    label = "Клиент банк"
  Case "gpImport": label = "импорт"
  Case "gpAuthorization": label = "авторизация"
  Case "btnLogin": label = "логин"
'  Case "mImport"
'  Case "btnImportOptions"
'  Case "gpFiles"
'    label = "доступ к файлам"
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
'    label = "доступ к файлам"
'  Case "gpFormat"
'  Case "mFormat"
'  Case "btnFormat"
'  Case "btnFormatOptions"
'  Case "btnPublicationOptions"
'    label = "доступ к файлам"
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
'          If (i = 1 And Split(dicOptions.Item("file" & i), "|")(0) = "Настраиваемая кнопка") Then
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
'      MsgBox "в разработке"
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
'      screentip = "настройка доступа к самым полезным папкам"
'    Case "btnImport"
'      screentip = "Импорт внешних данных"
'      For i = 1 To ImportBtnCount
'        If CInt(Numbers(control.id)) = i Then
'          If Split(dicOptions.Item("import" & i), "|")(1) = "" Then
'            screentip = "Импорт внешних данных (данную кнопку можно настроить в меню настроек)"
'          Else:
'            screentip = Split(dicOptions.Item("import" & i), "|")(5)
'          End If
'        End If
'      Next i
'
'    Case "btnFile"
'      For i = 1 To FileBtnCount
'        If CInt(Numbers(control.id)) = i Then
'          If (i = 1 And Split(dicOptions.Item("file" & i), "|")(0) = "Настраиваемая кнопка") Then
'            screentip = "Настраиваемая кнопка для открытия файлов"
'          Else:
'            screentip = "Открыть файл " & Split(dicOptions.Item("file" & i), "|")(0)
'          End If
'        End If
'      Next i
'    Case "btnFilesOptions"
'      screentip = "настройка путей к наиболее актуальным файлам"
'    Case "gpFormat"
'    Case "mFormat"
'
'    Case "btnFormatOptions"
'    Case "btnPublicationOptions"
'      screentip = "настройка публикуемых файлов и рассылки информации"
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
'        supertip = "Импорт данных из файла (*.dbf, *.xls*, *.csv)" & vbCrLf & _
'                   "" & vbCrLf & vbCrLf & _
'                   "Выполните такие шаги: " & vbCrLf & _
'                   "------------------------------------------------" & vbCrLf & _
'                   "  (1) активируйте лист для добавления данных;" & vbCrLf & _
'                   "  (2) нажмите 'Получить данные' и выберите файл" & vbCrLf & _
'                   "-------------------------------------------------" & vbCrLf
'      ElseIf control.id = "btnImport3" Then
'        supertip = "В данную папку автоматически " & vbCrLf & _
'        "переносятся все файлы, которые" & vbCrLf & _
'        "были импортированы в базу"
'      End If
'    Case "btnFile"
'      For i = 1 To FileBtnCount
'        If CInt(Numbers(control.id)) = i Then
'          If (i = 1 And Split(dicOptions.Item("file" & i), "|")(0) = "Настраиваемая кнопка") Then
'            supertip = "Возможно настроить до 6 кнопок с " & vbCrLf & _
'                  "помощью диалогового окна настроек"
'          Else:
'            supertip = "Полный путь:" & vbCrLf & Split(dicOptions.Item("file" & Numbers(control.id)), "|")(1)
'          End If
'        End If
'      Next i
'    Case "btnFilesOptions"
'      supertip = "С помощью настроек можно назначить кнопкам" & vbCrLf & _
'                 "доступ к актуальным файлам." & vbCrLf & _
'                 "Максимальное количество кнопок - 6." & vbCrLf & _
'                 "Настройки позволяют:" & vbCrLf & _
'                 "------------------------------------------" & vbCrLf & _
'                 " (1) запомнить путь к файлу;" & vbCrLf & _
'                 " (2) поменять значок кнопки;" & vbCrLf & _
'                 " (3) определить подсказку; " & vbCrLf & _
'                 " (4) запомнить пароли" & vbCrLf & _
'                 "------------------------------------------" & vbCrLf
'    Case "gpFormat"
'    Case "mFormat"
'    Case "btnFormat"
'    Case "btnFormatOptions"
'    Case "btnPublicationOptions"
'      supertip = "С помощью настроек можно: " & vbCrLf & _
'                   " - назначить местонахождение публикуемых файлов; " & vbCrLf & _
'                   " - создавать списки рассылки; " & vbCrLf & _
'                   " - настроить почтовую программу и " & vbCrLf & _
'                   "   многое другое..."
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



