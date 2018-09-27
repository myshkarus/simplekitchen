VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Settings class for simpleKitchen project"
Option Explicit

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' simpleKitchen VBA v.0.1.0
' (c) Mykhailo Shpilienko - https://github.com/myshkarus/simplekitchen
'----------------------------------------------------------------------------------
' Initial settings for simpleKitchen VBA project
'
' Detailed description:
'   Setup and restore general settings. Class creates config storage and
'   ?????? checks for project worksheets.
'   Executed on Workbook_Open event before Ribbon.customUIonLoad event.
'
' Errors:
' Status: testing
'
' TODO: clear settings
'
' @class Settings, v.0.3.0
' @author myshkarus
' @hqinside.com
' @license - n/a
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'list of worksheets in Russian
Private Const SHEET_NAMES As String = "старт;продукты;блюда;справочники"
'list of worksheets in Ukrainian
'Private Const SHEET_NAMES As String = "старт;продукти;блюда;довідники"
'list of worksheets in English
'Private Const SHEET_NAMES As String = "startpage;products;dishes;references"

'preinstalled values:
Private Const TITLE_CONFIG_TABLE As String = "Settings;Value" 'config table heading
Private Const TITLE_SHEET_TABLES As String = "Code;Value;Ctls;Opt;DevStage" 'worksheet's list table heading

'named ranges to be watched:
Private Const WATCH As String = "MAC"
Private Const TO_REBUILD As String = "torebuild"

'values for named ranges:
Private Const SHEET_ID As String = "SheetID"
Private Const SETTING_ITEMS As String = "databasepath;computer;MAC;torebuild;ribbonpointer;tabs;formulabar;scrollbar;headings"

Private Type settingsPage
   tSheet As Worksheet
End Type

Private this As settingsPage

'flag to show if the settings are Ready
Private blnReady As Boolean
'dictionary to store current settings temporary
Private mySettings As Dictionary

'------------------------------------
'PUBLIC PROPERTIES
'------------------------------------

Public Property Get Ready() As Boolean
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Get the state of the Settings class
   ' True if settings are set up successfully
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Ready = blnReady
End Property

'------------------------------------
'PRIVATE FUNCTIONS
'------------------------------------

Private Function ConfigStorageNew() As Worksheet
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Return new worksheet with formatted config tables
   ' Depends on SettingsAdded() function and RestoreSettings sub
   ' Out: Worksheet object
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Dim wSheet As Worksheet
   On Error GoTo 0                               'to avoid gaps in execution
   Set wSheet = ThisWorkbook.Worksheets.Add
   If Not wSheet Is Nothing Then
      wSheet.name = SHEET_SETTINGS
      If SettingsAdded(wSheet) Then
         Call RestoreSettings
         Set ConfigStorageNew = wSheet
      End If
      Set wSheet = Nothing
   End If
End Function

Private Function ShouldRebuild() As Boolean
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Check if the project should be rebuilded from scratch
   ' Depends on ShouldUpdate(), ManualRebuild(), RangeExists() functions and UpdateProject sub
   ' Out: Boolean
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Dim takeAction As Boolean
   takeAction = Not RangeExists(WATCH) Or ShouldUpdate Or ManualRebuild
   If takeAction Then
      Call UpdateProject
   End If
   ShouldRebuild = takeAction
End Function

Private Function ShouldUpdate() As Boolean
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Check whether the PC has been changed (by MAC)
   ' Depends on SysInfo.Mac
   ' Out: Boolean
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   On Error GoTo FailExit
   ShouldUpdate = SysInfo.Mac <> Range(WATCH).value
   Exit Function
FailExit:
   ShouldUpdate = True
End Function

Private Function ManualRebuild() As Boolean
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Check if manual flag to rebuild is in existence
   ' Out: Boolean
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   On Error GoTo FailExit
   ManualRebuild = Range(TO_REBUILD).value
   Exit Function
FailExit:
   ManualRebuild = False
End Function

Private Function SettingsDictionary(ByVal wSheet As Worksheet) As Dictionary
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Return dictionary with all old settings
   ' Depends on Dictionary class
   ' Out: Dictionary object
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Dim rng As Range
   Dim strSearchValue As String
   strSearchValue = Split(TITLE_CONFIG_TABLE, ";")(0)
   Set rng = wSheet.Cells.Find(strSearchValue, wSheet.Cells(1, 1)).Offset(1, 0)
   With New Dictionary
      If Not rng Is Nothing Then
         Do While rng.value <> Empty
            .Add rng.Offset(0, 1).value, rng.value
            Set rng = rng.Offset(1, 0)
         Loop
      End If
      Set SettingsDictionary = .Self
   End With
End Function

Private Function SettingsAdded(ByVal wSheet As Worksheet) As Boolean
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Routine procedure to create and fill in config tables
   ' Depends on RangeFormat and NamedRangeSetup subs
   ' In: target worksheet
   ' Out: Boolean
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Dim Title As Variant
   Dim modulesItems As Variant
   Dim settingItems As Variant
   Dim rng As Range
   Dim i As Integer
   On Error GoTo FailExit
   If Not wSheet Is Nothing Then
      wSheet.Cells.ClearFormats
      
      Title = Split(TITLE_CONFIG_TABLE, ";")
      Set rng = wSheet.Cells(1, 1).Offset(1, 1).Resize(1, UBound(Title) + 1)
      rng.value = Title
      Call RangeFormat(rng)
      Erase Title
      
      settingItems = Split(SETTING_ITEMS, ";")
      Set rng = rng.Offset(1, 0).Resize(UBound(settingItems) + 1, 1)
      rng.value = Application.WorksheetFunction.Transpose(settingItems)
    
      Title = Split(TITLE_SHEET_TABLES, ";")
      Set rng = rng.Offset(UBound(settingItems) + 2, 0).Resize(1, UBound(Title) + 1)
      rng.value = Title
      Call RangeFormat(rng)
      Erase Title
      Erase settingItems
      
      modulesItems = Split(SHEET_NAMES, ";")
      Set rng = rng.Offset(1, 0).Resize(1, 1)
      For i = 0 To UBound(modulesItems)
         rng.Offset(i, 0).value = i + 1: rng.Offset(i, 0).HorizontalAlignment = xlLeft
         rng.Offset(i, 1).value = modulesItems(i)
      Next i
      Erase modulesItems

      With wSheet.Cells
         .Font.name = "Courier New"
         .Font.size = 9
         .HorizontalAlignment = xlLeft
         .EntireColumn.AutoFit
      End With
      wSheet.Cells(1, 1).ColumnWidth = 4
   
      Call NamedRangeSetup(wSheet)               'check named ranges or create new
   End If

   SettingsAdded = True
   If Not rng Is Nothing Then Set rng = Nothing
   Exit Function
FailExit:
   If Not rng Is Nothing Then Set rng = Nothing
   SettingsAdded = False
End Function

'------------------------------------
'PRIVATE PROCEDURES
'------------------------------------

Private Sub Class_Initialize()
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Class entry point
   ' Depends on ConfigStorageNew, ShouldRebuild functions and KillMe, SysInfo.SetSystemParameter subs
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   On Error GoTo FailExit
    
   'always set up system parameters
   Call SysInfo.SetSystemParameter
   On Error Resume Next
   Set this.tSheet = ThisWorkbook.Worksheets(SHEET_SETTINGS)
   On Error GoTo 0
      
   If this.tSheet Is Nothing Then
      Set this.tSheet = ConfigStorageNew
   Else
      If ShouldRebuild Then
         Call KillMe
         Set this.tSheet = Nothing
         Set this.tSheet = ConfigStorageNew
      End If
   End If
      
   If Not this.tSheet Is Nothing Then
      blnReady = True
   Else: GoTo FailExit
   End If
   Exit Sub
FailExit:
   MsgBox "Worksheet '" & SHEET_SETTINGS & "' with metadata is not found!" & vbCrLf & _
          "Contact the developer", vbCritical Or vbOKOnly, "System error"
   End                                           'destruct class if it fails
End Sub

Private Sub SaveSettings()
   Set mySettings = SettingsDictionary(this.tSheet)
End Sub

Private Sub RestoreSettings()
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Retrieve data from SettingsDictionary and fill in named ranges
   ' Depends on public UpdateSettings() sub
   ' mapping:
   '         settingItems(0) - "databasepath"
   '         settingItems(1) - "computer"
   '         settingItems(2) - "MAC"
   '         settingItems(3) - "torebuild"
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Dim i As Integer
   Dim settingItems As Variant
   settingItems = Split(SETTING_ITEMS, ";")
   If Not mySettings Is Nothing Then
      With mySettings
         For i = 0 To .size - 1
            Select Case .keys()(i)
            Case settingItems(0):
            Case settingItems(1):
            Case settingItems(2):
            Case settingItems(3):
            Case Else: Call UpdateSettings(.keys()(i), .item(.keys()(i)))
            End Select
         Next i
      End With
        
   End If
   Call UpdateSettings(settingItems(1), SysInfo.ComputerName)
   Call UpdateSettings(settingItems(2), SysInfo.Mac)
   Call UpdateSettings(settingItems(3), False)
   Erase settingItems
   Set mySettings = Nothing
End Sub

Private Sub KillMe()
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Delete sheet with settings
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   If Not this.tSheet Is Nothing Then
      Call SaveSettings
      Application.DisplayAlerts = False
      this.tSheet.Delete
   End If
End Sub

Private Sub UpdateProject()
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Rebuild project in full
   ' Depends on public RebuildProject sub
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Call RebuildProject
End Sub

Private Sub Class_Terminate()
   'uncomment after testing!
   'ThisWorkbook.Worksheets(SHEET_SETTINGS).visible = xlSheetVeryHidden
   If Not mySettings Is Nothing Then Set mySettings = Nothing
   Set this.tSheet = Nothing
End Sub

Private Sub NamedRangeSetup(ByVal wSheet As Worksheet)
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Check existence or create new named range to store config parameters
   ' Depends on public AddNamedRange sub and public RangeExists() function
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Dim rng As Range
   Dim settingItems As Variant
   Dim strSearchValue As String
   Dim i As Integer
   
   'named range to be used in MetaData class
   If Not RangeExists(SHEET_ID) Then
      strSearchValue = Split(TITLE_SHEET_TABLES, ";")(0)
      Set rng = wSheet.Cells.Find(strSearchValue, wSheet.Cells(1, 1))
      If Not rng Is Nothing Then Call AddNamedRange(SHEET_ID, rng.AddressLocal)
   End If
   
   'named range for settings items
   settingItems = Split(SETTING_ITEMS, ";")
   For i = 0 To UBound(settingItems)
      If Not RangeExists(settingItems(i)) Then
         Set rng = wSheet.Cells.Find(settingItems(i), wSheet.Cells(1, 1)).Offset(0, 1)
         If Not rng Is Nothing Then Call AddNamedRange(settingItems(i), rng.AddressLocal)
      End If
   Next i
    
   Erase settingItems
   Set rng = Nothing
End Sub

Private Sub RangeFormat(ByVal rng As Range)
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Formatting headings of config tables
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   If Not rng Is Nothing Then
      With rng
         .Interior.Color = RGB(100, 100, 100)
         .Font.Color = vbWhite
      End With
   End If
End Sub
