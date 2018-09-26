VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Настройки"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6885
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private WithEvents Listener As Listener
Attribute Listener.VB_VarHelpID = -1
Private listenerCollection As New Collection



Private lRecords As Records
Private gRecords As Records
Private opt As IOptions
Private langStorage As IStorage
Private mainStorage As IStorage

Private Sub cbLanguage_AfterUpdate()
  With Range("language")
    .value = cbLanguage.ListIndex + 1
    .Calculate
  End With
  Debug.Print Range("language").value
  
  Call ChangeFormLanguage
End Sub

Private Sub FormLayout()
  With Me
    .Width = 356: .Height = 288
    .MultiPage.Width = 336: .MultiPage.Height = 222: .MultiPage.Top = 2: .MultiPage.Left = (.InsideWidth - .MultiPage.Width) / 2
    With .MultiPage.Pages("MainPage")
      With .frUser
       .Left = 7: .Top = 7: .Width = 153: .Height = 110
        .lblLogin.Left = 9: .lblLogin.Top = 11: .lblLogin.Height = 9.75: .lblLogin.Width = 36
        .lblLoginValue.Left = .lblLogin.Left + .lblLogin.Width + HINDENT: .lblLoginValue.Top = .lblLogin.Top: .lblLoginValue.Height = .lblLogin.Height: .lblLoginValue.Width = 95
        .lblLoginPwd.Left = .lblLogin.Left: .lblLoginPwd.Top = .lblLogin.Top + .lblLogin.Height + VSHORTINDENT: .lblLoginPwd.Height = .lblLogin.Height: .lblLoginPwd.Width = .lblLogin.Width
        .lblPwdValue.Left = .lblLoginValue.Left: .lblPwdValue.Top = .lblLoginPwd.Top: .lblPwdValue.Height = .lblLogin.Height: .lblPwdValue.Width = .lblLoginValue.Width
        .chkLoginPwdRemember.Left = .lblLogin.Left: .chkLoginPwdRemember.Top = .lblLoginPwd.Top + .lblLoginPwd.Height + VSHORTINDENT - 2: .chkLoginPwdRemember.Height = 16: .chkLoginPwdRemember.Width = 130
        .lblStatus.Left = .lblLogin.Left: .lblStatus.Top = .chkLoginPwdRemember.Top + .chkLoginPwdRemember.Height + VSHORTINDENT - 2: .lblStatus.Height = .lblLogin.Height: .lblStatus.Width = .lblLogin.Width
        .lblStatusValue.Left = .lblLoginValue.Left: .lblStatusValue.Top = .lblStatus.Top: .lblStatusValue.Height = .lblLogin.Height: .lblStatusValue.Width = .lblLoginValue.Width
        .cmdChangeUser.Height = DEFAULTHEIGHT: .cmdChangeUser.Top = .lblStatus.Top + .lblStatus.Height + VSHORTINDENT: .cmdChangeUser.Width = 135: .cmdChangeUser.Left = (.Width - .cmdChangeUser.Width) / 2
      End With
      With .frMenu
        .Left = Me.frUser.Left + Me.frUser.Width + HINDENT: .Top = Me.frUser.Top: .Height = 77: .Width = 163
        .lblSession.Left = 9: .lblSession.Top = Me.frUser.lblLogin.Top: .lblSession.Height = Me.frUser.lblLogin.Height: .lblSession.Width = 45
        .txtSession.Left = .lblSession.Left + .lblSession.Width + HINDENT: .txtSession.Width = 32:  .txtSession.Height = DEFAULTHEIGHT - 3: .txtSession.Top = .lblSession.Top - BUTTONSHIFT * 2
        .lblMin.Left = .txtSession.Left + .txtSession.Width + HINDENT: .lblMin.Top = .lblSession.Top: lblMin.Height = Me.frUser.lblLogin.Height: lblMin.Width = 22
        .chkMenuAuto.Left = .lblSession.Left: .chkMenuAuto.Top = .lblSession.Top + .lblSession.Height + VSHORTINDENT: .chkMenuAuto.Height = .txtSession.Height: .chkMenuAuto.Width = 85
        .lblShortCut.Left = .chkMenuAuto.Left: .lblShortCut.Top = .chkMenuAuto.Top + .chkMenuAuto.Height + VSHORTINDENT - BUTTONSHIFT: .lblShortCut.Height = Me.frUser.lblLogin.Height: .lblShortCut.Width = 92
        .txtShortCut.Left = .lblShortCut.Left + .lblShortCut.Width + HINDENT: .txtShortCut.Top = .lblShortCut.Top - BUTTONSHIFT * 3: .txtShortCut.Height = DEFAULTHEIGHT - 3: .txtShortCut.Width = .txtSession.Width
      End With
      .lblSharedWork.Left = .frMenu.Left + BUTTONSHIFT * 3: .lblSharedWork.Top = .frMenu.Top + .frMenu.Height + HINDENT:  .lblSharedWork.Height = .lblLogin.Height: .lblSharedWork.Width = 85
      .optAutonomous.Left = .lblSharedWork.Left: .optAutonomous.Top = .lblSharedWork.Top + .lblSharedWork.Height + BUTTONSHIFT * 3: .optAutonomous.Height = Me.frMenu.chkMenuAuto.Height: .optAutonomous.Width = 59
      .optWebSync.Left = .optAutonomous.Left + .optAutonomous.Width: .optWebSync.Top = .optAutonomous.Top: .optWebSync.Height = .optAutonomous.Height: .optWebSync.Width = 97
      With .frPrivateKey
        .Left = Me.frUser.Left: .Top = Me.frUser.Top + Me.frUser.Height + VSHORTINDENT: .Height = 73: .Width = Me.frMenu.Left + Me.frMenu.Width - Me.frUser.Left
        .lblPrivateKeyPath.Left = Me.frUser.lblLogin.Left: .lblPrivateKeyPath.Top = 11: .lblPrivateKeyPath.Height = Me.frUser.lblLogin.Height: .lblPrivateKeyPath.Width = 115
        .txtPrivateKeyPath.Left = .lblPrivateKeyPath.Left: .txtPrivateKeyPath.Top = .lblPrivateKeyPath.Top + .lblPrivateKeyPath.Height + VSHORTINDENT \ 2: .txtPrivateKeyPath.Height = DEFAULTHEIGHT: .txtPrivateKeyPath.Width = 180
        .cmdPrivateKeyPath.Left = .txtPrivateKeyPath.Left + .txtPrivateKeyPath.Width + BUTTONSHIFT: .cmdPrivateKeyPath.Top = .txtPrivateKeyPath.Top: .cmdPrivateKeyPath.Height = DEFAULTHEIGHT: .cmdPrivateKeyPath.Width = .cmdPrivateKeyPath.Height - BUTTONSHIFT
        .cmdRequestKey.Left = .cmdPrivateKeyPath.Left + .cmdPrivateKeyPath.Width + HINDENT: .cmdRequestKey.Top = .cmdPrivateKeyPath.Top: .cmdRequestKey.Height = DEFAULTHEIGHT: .cmdRequestKey.Width = 100
        .lblPrivateKeyLife.Left = .lblPrivateKeyPath.Left: .lblPrivateKeyLife.Top = .cmdRequestKey.Top + .cmdRequestKey.Height + VSHORTINDENT: .lblPrivateKeyLife.Height = Me.frUser.lblLogin.Height: .lblPrivateKeyLife.Width = 200
      End With
    End With
    With .MultiPage.Pages("FormatPage")
      .cbLanguage.Height = DEFAULTHEIGHT
    End With
    .cmdClose.Width = 66: .cmdClose.Left = .InsideWidth - .cmdClose.Width - HINDENT: .cmdClose.Height = 22.5: .cmdClose.Top = .InsideHeight - .cmdClose.Height - VSHORTINDENT
    .cmdOK.Width = .cmdClose.Width: .cmdOK.Left = .cmdClose.Left - .cmdOK.Width - BUTTONSHIFT * 3: .cmdOK.Top = .cmdClose.Top: .cmdOK.Height = .cmdClose.Height
  End With
End Sub
'здесь должна быть процедура по подключению/отключению памяти пароля
'Private Sub IOptions_AddNewUser(ByVal newUser As String, ByVal password As String, ByVal passwordRemember As Boolean)
'  Dim initialSettingsAdded As Boolean
'  Dim certified As Boolean
'  Dim someSetting As EntryUser
'  Set userStorage = NewStorage(newUser, False)
'  If Registered(newUser, userStorage) Then
'    initialSettingsAdded = NewUserOptions(newUser, password)
'  End If
'  If passwordRemember Then
'    Set someSetting = myRecordsUserSetting.Item("RememberPassword")
'    certified = someSetting.ChangeEditMode(True)
'    If certified Then someSetting.value = True
'  End If
'  CatchLogger InfoLevel, , "Добавлен новый пользователь (логин: " & newUser & ")", "IOptions_AddNewUser"
'End Sub
'
'Private Function newUserSetting() As EntryUser
'  Set newUserSetting = NewEntry(userStorage)
'  Set myRecordsUserSetting = NewRecords(newUserSetting)
'End Function

Private Sub CommandButton14_Click()

End Sub

Private Sub cmdFactoryReset_Click()

End Sub

Private Sub cmdChangeUser_Click()
  frmPasswordChange.Show (vbModal)
End Sub

Private Sub cmdResetToDefault_Click()
'Dim d As Date
Dim opt As New IOptions
Set opt = New options
opt.ResetToDefault
Set opt = Nothing
End Sub

Private Sub cmdRestoreLanguage_Click()
Dim d As Date
Dim opt As New IOptions
Set opt = New options
opt.setupLanguage
Set opt = Nothing
End Sub

Private Sub frPrivateKey_Click()

End Sub

Private Sub lblPrivateKeyPath_Click()

End Sub

Private Sub MultiPage_Change()

End Sub

Private Sub UserForm_Initialize()
  Dim user As IEntry
  On Error GoTo FailExit
  Application.ScreenUpdating = False
  Set langStorage = NewStorage(LANG_SETTINGS, True)
  Set mainStorage = NewStorage(MAIN_SETTINGS, True)
  Set user = NewEntry(mainStorage)
  Set gRecords = NewRecords(user)
  Call ChangeFormLanguage
  CatchLogger InfoLevel, , "Инициализация формы...", , "UserForm_Initialize", className
  Set user = Nothing
  Exit Sub
FailExit:
  CatchLogger ErrorLevel, UnInitialized, "Обнаружена проблема с инициализацией формы", , "UserForm_Initialize", className
  End
End Sub

Private Sub UserForm_Terminate()
  On Error GoTo FailExit
  Listener.StopListening
  Set Listener = Nothing
  Set listenerCollection = Nothing
  Set langStorage = Nothing
  Set mainStorage = Nothing
  Set opt = Nothing
  Set gRecords = Nothing
  Set lRecords = Nothing
  Application.ScreenUpdating = True
  CatchLogger InfoLevel, , "Завершена работа формы", , "UserForm_Terminate", className
  Exit Sub
FailExit:
  CatchLogger ErrorLevel, SubRuntimeError, "Обнаружена проблема при завершении работы формы", "UserForm_Terminate", className
End Sub

Private Sub UserForm_Activate()
Dim ctItem As MSForms.control, ctMultiPageItem As MSForms.control
'Debug.Print "form activate"
'запускаем прослушку на commandbutton
'Set gRecords = gRecords(genSetting)
'Set uRecords = uRecords(userSetting)
'Set lRecords = lRecords(lang)
Me.cmdClose.Cancel = True
Call FormLayout

  With cbLanguage
    .AddItem "українська"
    .AddItem "русский"
    .AddItem "English"
    .ListIndex = Range("language") - 1 'заменить на опции и сделать проверку чтобы не было <1 >3
  End With


For Each ctItem In Me.Controls
  If TypeOf ctItem Is MSForms.CommandButton Then
    'Set Listener = New Listener
    'Set Listener.btn = ctItem
    'listenerCollection.Add Listener
  ElseIf TypeOf ctItem Is MSForms.ComboBox Then
    Set Listener = New Listener
    Set Listener.cbox = ctItem
    listenerCollection.Add Listener
  End If
Next ctItem

If Listener Is Nothing Then
  Set Listener = New Listener
End If
If Not Listener Is Nothing Then
  Listener.StartListening Me
End If
Set ctItem = Nothing: Set ctMultiPageItem = Nothing
End Sub



'Private Sub Label1_Click()
'    Link = "http://www.whitehouse.gov"
'    On Error GoTo NoCanDo
'    ActiveWorkbook.FollowHyperlink address:=Link, NewWindow:=True
'    Unload Me
'    Exit Sub
'NoCanDo:
'    MsgBox "Cannot open " & Link
'End Sub
'To create a "mail to" hyperlink, use a statement like this:
'
'    Link = "mailto:president@whitehouse.gov"


'Private Sub ComboBoxPopulate()
'Dim ctl As MSForms.Control
'Dim mpctl As MSForms.Control
'Dim Page As MSForms.Page
'Dim ChildKeys As Variant
'Dim dic As New cDictionary
'Dim strItem As String
'Dim i As Integer, j As Integer, q As Integer
'Dim strArr() As String
'For Each ctl In Me.Controls
'  If TypeOf ctl Is MSForms.MultiPage Then
'    For Each Page In ctl.Pages
'      For Each mpctl In Page.Controls
'        If (TypeOf mpctl Is MSForms.ComboBox) Then
'          mpctl.ListIndex = -1
'          mpctl.text = vbNullString
'          mpctl.RowSource = ""
'          If Left(mpctl.Name, Len(mpctl.Name) - 1) = "cmbController" Then
'            With mpctl
'                Set dic = dicSource.ObjectItem("контролер")
'                ChildKeys = dic.keys
'                For i = 0 To dic.Size - 1
'                  strItem = dic.Item(ChildKeys(i))
'                  If InStr(strItem, "|") <> 0 Then strItem = Left(strItem, InStr(strItem, "|") - 1)
'                  .AddItem strItem
'                Next i
'            End With
'          ElseIf Left(mpctl.Name, Len(mpctl.Name) - 1) = "cmbWeighHouse" Then
'            With mpctl
'              Set dic = dicSource.ObjectItem("вагова")
'              ChildKeys = dic.keys
'              For i = 0 To dic.Size - 1
'                strItem = dic.Item(ChildKeys(i))
'                If InStr(strItem, "|") <> 0 Then strItem = Left(strItem, InStr(strItem, "|") - 1)
'                .AddItem strItem
'              Next i
'            End With
'          ElseIf Left(mpctl.Name, Len(mpctl.Name) - 1) = "cmbAutoID" Then
'            With mpctl
'              Erase strArr
'              Set dic = dicSource.ObjectItem("авто")
'              ChildKeys = dic.keys
'              For i = 0 To dic.Size - 1
'                q = CountIn(dic.Item(ChildKeys(i)), "|")
'                If q > 0 Then
'                  strArr = Split(dic.Item(ChildKeys(i)), "|")
'                    .AddItem
'                    For j = 0 To 1
'                      If InStr(strArr(j), " ") <> 0 Then
'                        .list(i, j) = Left(strArr(j), InStr(strArr(j), " ") - 1)
'                      Else:
'                        .list(i, j) = strArr(j)
'                      End If
'                    Next j
'                Else:
'                  .AddItem dic.Item(ChildKeys(i))
'                End If
'              Next i
'            End With
'          ElseIf Left(mpctl.Name, Len(mpctl.Name) - 1) = "cmbTrailer" Then
'            With mpctl
'              Set dic = dicSource.ObjectItem("причіп")
'              ChildKeys = dic.keys
'              For i = 0 To dic.Size - 1
'                strItem = dic.Item(ChildKeys(i))
'                If InStr(strItem, "|") <> 0 Then strItem = Left(strItem, InStr(strItem, "|") - 1)
'                .AddItem strItem
'              Next i
'            End With
'          ElseIf Left(mpctl.Name, Len(mpctl.Name) - 1) = "cmbField" Then
'            With mpctl
'              Erase strArr
'              Set dic = dicSource.ObjectItem("поле")
'              ChildKeys = dic.keys
'              For i = 0 To dic.Size - 1
'                q = CountIn(dic.Item(ChildKeys(i)), "|")
'                If q > 0 Then
'                  strArr = Split(dic.Item(ChildKeys(i)), "|")
'                    .AddItem
'                    For j = 0 To 1
'                      .list(i, j) = strArr(j)
'                    Next j
'                Else:
'                  .AddItem dic.Item(ChildKeys(i))
'                End If
'              Next i
'            End With
'          ElseIf Left(mpctl.Name, Len(mpctl.Name) - 1) = "cmbCargo" Then
'            With mpctl
'              Set dic = dicSource.ObjectItem("культура")
'              ChildKeys = dic.keys
'              For i = 0 To dic.Size - 1
'                strItem = dic.Item(ChildKeys(i))
'                If InStr(strItem, "|") <> 0 Then strItem = Left(strItem, InStr(strItem, "|") - 1)
'                .AddItem strItem
'              Next i
'            End With
'          ElseIf Left(mpctl.Name, Len(mpctl.Name) - 1) = "cmbCategory" Then
'            With mpctl
'              Set dic = dicSource.ObjectItem("категорія")
'              ChildKeys = dic.keys
'              For i = 0 To dic.Size - 1
'                strItem = dic.Item(ChildKeys(i))
'                If InStr(strItem, "|") <> 0 Then strItem = Left(strItem, InStr(strItem, "|") - 1)
'                .AddItem strItem
'              Next i
'            End With
'          ElseIf Left(mpctl.Name, Len(mpctl.Name) - 1) = "cmbWarehouse" Then
'            With mpctl
'              Set dic = dicSource.ObjectItem("тік")
'              ChildKeys = dic.keys
'              For i = 0 To dic.Size - 1
'                strItem = dic.Item(ChildKeys(i))
'                If InStr(strItem, "|") <> 0 Then strItem = Left(strItem, InStr(strItem, "|") - 1)
'                .AddItem strItem
'              Next i
'            End With
'          ElseIf Left(mpctl.Name, Len(mpctl.Name) - 1) = "cmbExcelType" Then
'            With mpctl
'              .Clear
'              .AddItem ".xls"
'              .AddItem ".xlsx"
'            End With
'          End If
'        End If
'      Next mpctl
'    Next Page
'  ElseIf (TypeOf ctl Is MSForms.ComboBox) Then
'  End If
'Next ctl
'If Not ctl Is Nothing Then Set ctl = Nothing
'If Not mpctl Is Nothing Then Set mpctl = Nothing
'If Not Page Is Nothing Then Set Page = Nothing
'If Not dic Is Nothing Then Set dic = Nothing
'End Sub


Private Function className() As String
  className = TypeName(Me)
End Function

Private Sub ChangeFormLanguage()
  On Error GoTo FailExit
  Dim myLang As EntryLanguage
  Dim longString As String
  Set myLang = NewEntry(langStorage)
  Set lRecords = NewRecords(myLang)
  Set myLang = lRecords.Item("frmSettings"):      Me.Caption = myLang
  With Me.MultiPage.Pages("MainPage")
    Set myLang = lRecords.Item("MainPage"):       .Caption = myLang
    Set myLang = lRecords.Item("frUser"):         .frUser.Caption = myLang
    Set myLang = lRecords.Item("lblLogin"):       .frUser.lblLogin.Caption = myLang
    Set myLang = lRecords.Item("lblLoginPwd"):    .frUser.lblLoginPwd.Caption = myLang
    Set myLang = lRecords.Item("chkLoginPwdRemember"): .frUser.chkLoginPwdRemember.Caption = myLang
    Set myLang = lRecords.Item("lblStatus"):      .frUser.lblStatus.Caption = myLang
    Set myLang = lRecords.Item("cmdChangeUser"):  .frUser.cmdChangeUser.Caption = myLang
    Set myLang = lRecords.Item("frMenu"):         .frMenu.Caption = myLang
    Set myLang = lRecords.Item("lblSession"):     .frMenu.lblSession.Caption = myLang
    Set myLang = lRecords.Item("lblMin"):         .frMenu.lblMin.Caption = myLang
    Set myLang = lRecords.Item("chkMenuAuto"):    .frMenu.chkMenuAuto.Caption = myLang
    Set myLang = lRecords.Item("lblShortCut"):    .frMenu.lblShortCut.Caption = myLang
    Set myLang = lRecords.Item("lblSharedWork"):  .lblSharedWork.Caption = myLang
    Set myLang = lRecords.Item("optAutonomous"):  .optAutonomous.Caption = myLang
    Set myLang = lRecords.Item("optWebSync"):     .optWebSync.Caption = myLang
    Set myLang = lRecords.Item("frPrivateKey"):   .frPrivateKey.Caption = myLang
    Set myLang = lRecords.Item("lblPrivateKeyPath"): .frPrivateKey.lblPrivateKeyPath.Caption = myLang
    Set myLang = lRecords.Item("cmdRequestKey"):  .frPrivateKey.cmdRequestKey.Caption = myLang
    Set myLang = lRecords.Item("lblPrivateKeyLife")
    longString = myLang & " "
    'change:
    longString = longString & "31/12/2017" & " "
    Set myLang = lRecords.Item("lblPrivateKeyLifeInfo1")
    longString = longString & myLang & " "
    longString = longString & "XXX" & " "
    Set myLang = lRecords.Item("lblPrivateKeyLifeInfo2")
    longString = longString & myLang & " "
    .frPrivateKey.lblPrivateKeyLife.Caption = longString
  End With
  Set myLang = Nothing
  Exit Sub
FailExit:
End Sub



Private Sub listener_OnEnter(ctrl As MSForms.control)
If Not ctrl Is Nothing Then
  If Not (TypeOf ctrl Is MSForms.TextBox Or TypeOf ctrl Is MSForms.ComboBox Or TypeOf ctrl Is MSForms.ListBox) Then Exit Sub
  ctrl.BackColor = &HC8FFFF
  ctrl.BorderColor = &H8000000D
End If
End Sub

Private Sub listener_OnExit(ctrl As MSForms.control, Cancel As Boolean)
If Not ctrl Is Nothing Then
  If Not (TypeOf ctrl Is MSForms.TextBox Or TypeOf ctrl Is MSForms.ComboBox Or TypeOf ctrl Is MSForms.ListBox) Then Exit Sub
  ctrl.BackColor = &H80000005
  ctrl.BorderColor = &H80000006
End If
End Sub



Private Sub cbLanguage_Change()

End Sub

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If KeyAscii = 27 Then Unload Me
  'If KeyAscii = 13 Then Call cmdUpdate_Click
End Sub

Private Sub cmdOK_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  If KeyAscii = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub
