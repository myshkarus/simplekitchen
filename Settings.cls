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
' simpleKitchen VBA v.0.2.0
' (c) Mykhailo Shpilienko - https://github.com/myshkarus/simplekitchen
'----------------------------------------------------------------------------------
' Initial settings for simpleKitchen VBA project
'
' Detailed description:
'   Setup and restore general settings. Class creates config storage and checks
'   for project worksheets.
'   Executed on Workbook_Open event before Ribbon.customUIonLoad event.
'
' Errors:
' Status: testing
'
' @class Settings
' @author myshkarus
' @hqinside.com
' @license - n/a
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'list of worksheets (modules) in Russian
Private Const SHEET_NAMES As String = "продукты;блюда;справочники"
'list of worksheets (modules) in Ukrainian
'Private Const SHEET_NAMES As String = "продукти;блюда;довідники"
'list of worksheets (modules) in English
'Private Const SHEET_NAMES As String = "products;dishes;references"

'preinstalled values:
Private Const SHEET_SETTINGS As String = "settings" 'storage Worksheet name
Private Const TITLE_CONFIG_TABLE As String = "Settings;Value" 'config table heading
Private Const TITLE_SHEET_TABLES As String = "Code;Value;Ctls;Opt;DevStage" 'modules table heading

'values for named ranges:
Private Const SHEET_ID As String = "SheetID"
Private Const SETTING_ITEMS As String = "databasepath;ribbonpointer"

'flag to show if the settings are Ready
Private blnReady As Boolean

'------------------------------------
'PUBLIC PROPERTIES
'------------------------------------

Public Property Get Ready() As Boolean
Attribute Ready.VB_Description = "Get the state of the Settings class"
Attribute Ready.VB_UserMemId = 0
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Get the state of the Settings class
' True if the settings are set up successfully
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Ready = blnReady
End Property

'------------------------------------
'PRIVATE FUNCTIONS
'------------------------------------

Private Function ConfigStorage() As Worksheet
Attribute ConfigStorage.VB_Description = "Return worksheet with predefined name or create new worksheet to store settings"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Return worksheet with predefined name or create new worksheet to store settings
' Depends on ConfigStorageNew function
' Out: Worksheet object
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   If SheetExist(SHEET_SETTINGS) Then
      Set ConfigStorage = ThisWorkbook.Worksheets(SHEET_SETTINGS)
   Else
      Set ConfigStorage = ConfigStorageNew
   End If
End Function

Private Function ConfigStorageNew() As Worksheet
Attribute ConfigStorageNew.VB_Description = "Return new worksheet with formatted config tables"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Return new worksheet with formatted config tables
' Depends on SettingsAdd() sub
' Out: Worksheet object
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Dim wSheet As Worksheet
   On Error GoTo 0 'to avoid gaps in execution
   Set wSheet = ThisWorkbook.Worksheets.Add
   If Not wSheet Is Nothing Then
      wSheet.name = SHEET_SETTINGS
      Call SettingsAdd(wSheet)        'only if a new config storage is created
      Set ConfigStorageNew = wSheet
      Set wSheet = Nothing
   End If
End Function

'------------------------------------
'PRIVATE PROCEDURES
'------------------------------------

Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "Class entry point"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Class entry point
' Depends on function ConfigStorage and sub NamedRangeSetup
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Dim wSheet As Worksheet
   On Error GoTo FailExit
   Set wSheet = ConfigStorage 'check existence or create new storage
   If Not wSheet Is Nothing Then
      Call NamedRangeSetup 'check named ranges or create new
   Else: GoTo FailExit
   End If
   blnReady = True
   Set wSheet = Nothing
   Exit Sub
FailExit:
   MsgBox "Worksheet 'settings' with metadata is not found!" & vbCrLf & _
          "Contact the developer", vbCritical Or vbOKOnly, "System error"
   End  'destruct class if it fails
End Sub

Private Sub Class_Terminate()
Attribute Class_Terminate.VB_Description = "Class termination"
'uncomment after testing!
'ThisWorkbook.Worksheets(SHEET_SETTINGS).visible = xlSheetVeryHidden
End Sub

Private Sub NamedRangeSetup()
Attribute NamedRangeSetup.VB_Description = "Check existence or create new named range to store config parameters"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Check existence or create new named range to store config parameters
' Depends on public sub AddNamedRange
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Dim rng As Range
   Dim valueToSearch As String
   Dim wSheet As Worksheet
   
   Set wSheet = ThisWorkbook.Worksheets(SHEET_SETTINGS)
   'named range to be used in MetaData class
   If Not RangeExists(SHEET_ID) Then
      valueToSearch = Split(TITLE_SHEET_TABLES, ";")(0)
      Set rng = wSheet.Cells.Find(valueToSearch, wSheet.Cells(1, 1))
      If Not rng Is Nothing Then Call AddNamedRange(SHEET_ID, rng.AddressLocal)
   End If
   'named range for database path string (eg: D:\Database\cookinghouse.accdb)
   valueToSearch = Split(SETTING_ITEMS, ";")(0)
   If Not RangeExists(valueToSearch) Then
      Set rng = wSheet.Cells.Find(valueToSearch, wSheet.Cells(1, 1)).Offset(0, 1)
      If Not rng Is Nothing Then Call AddNamedRange(valueToSearch, rng.AddressLocal)
   End If
   'named range for RibbonPointer (used to restore Ribbon)
   valueToSearch = Split(SETTING_ITEMS, ";")(1)
   If Not RangeExists(valueToSearch) Then
      Set rng = wSheet.Cells.Find(valueToSearch, wSheet.Cells(1, 1)).Offset(0, 1)
      If Not rng Is Nothing Then Call AddNamedRange(valueToSearch, rng.AddressLocal)
   End If
   Set rng = Nothing
   Set wSheet = Nothing
End Sub

Private Sub SettingsAdd(ByVal wSheet As Worksheet)
Attribute SettingsAdd.VB_Description = "Routine procedure to create and fill in config tables"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Routine procedure to create and fill in config tables
' In: target worksheet (wshConfig cannot yet to be used)
' Depends on RangeFormat sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Dim title As Variant
   Dim modulesItems As Variant
   Dim settingItems As Variant
   Dim rng As Range
   Dim i As Integer
   If Not wSheet Is Nothing Then
      With wSheet.Cells
         .ClearFormats
         .Font.name = "Courier New"
         .Font.size = 9
      End With
      
      title = Split(TITLE_CONFIG_TABLE, ";")
      Set rng = wSheet.Cells(1, 1).Offset(1, 1).Resize(1, UBound(title) + 1)
      rng.value = title
      Call RangeFormat(rng)
      Erase title
      
      settingItems = Split(SETTING_ITEMS, ";")
      Set rng = rng.Offset(1, 0).Resize(UBound(settingItems) + 1, 1)
      rng.value = Application.WorksheetFunction.Transpose(settingItems)
    
      title = Split(TITLE_SHEET_TABLES, ";")
      Set rng = rng.Offset(UBound(settingItems) + 2, 0).Resize(1, UBound(title) + 1)
      rng.value = title
      Call RangeFormat(rng)
      Erase title
      Erase settingItems
      
      modulesItems = Split(SHEET_NAMES, ";")
      Set rng = rng.Offset(1, 0).Resize(1, 1)
      For i = 0 To UBound(modulesItems)
         rng.Offset(i, 0).value = i + 1: rng.Offset(i, 0).HorizontalAlignment = xlLeft
         rng.Offset(i, 1).value = modulesItems(i)
      Next i
      Erase modulesItems
      wSheet.Cells.EntireColumn.AutoFit
      wSheet.Cells(1, 1).ColumnWidth = 4
   End If
   If Not rng Is Nothing Then Set rng = Nothing
End Sub

Private Sub RangeFormat(ByVal rng As Range)
Attribute RangeFormat.VB_Description = "Formatting headings of config tables"
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
