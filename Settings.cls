VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' simpleKitchen VBA v.0.1.0
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
Private Const modules As String = "продукты;блюда;справочники"
'list of worksheets (modules) in Ukrainian
'Private Const modules As String = "продукти;блюда;довідники"
'list of worksheets (modules) in English
'Private Const modules As String = "products;dishes;references"

'preinstalled values:
Private Const config As String = "settings" 'storage Worksheet name
Private Const configTblTitle As String = "Settings;Value" 'config table heading
Private Const modulesTblTitle As String = "Code;Value;Ctls;Opt;DevStage" 'modules table heading
Private Const modulesTblSubTitle As String = "[table code];[sheet name];[number of controls]" 'modules table subheading

'values for named ranges:
Private Const tableID As String = "SheetID"
Private Const db As String = "databasepath"
Private Const rp As String = "ribbonpointer"

'flag to show if the settings are completed
Private setupCompleted As Boolean
'worksheet (storage) used to store all settings
Private wshConfig As Worksheet

'helper enumeration to format settings tables
Private Enum Style
   title
   subTitle
End Enum

'------------------------------------
'PUBLIC PROPERTIES
'------------------------------------

Friend Property Get Completed() As Boolean
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Get the state of the Settings class
' True if the settings are set up successfully
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Completed = setupCompleted
End Property

'------------------------------------
'PRIVATE FUNCTIONS
'------------------------------------

Private Function ConfigStorage() As Worksheet
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Return worksheet with predefined name or create new worksheet to store settings;
' Depends on NewConfigStorage function
' Out: Worksheet object
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   If SheetExist(config) Then
      Set ConfigStorage = ThisWorkbook.Worksheets(config)
   Else
      Set ConfigStorage = NewConfigStorage
   End If
End Function

Private Function NewConfigStorage() As Worksheet
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Return new worksheet with formatted config tables
' Depends on CreateConfigTables() sub
' Out: Worksheet object
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Dim wsh As Worksheet
   On Error GoTo 0 'to avoid gaps in execution
   Set wsh = ThisWorkbook.Worksheets.Add
   If Not wsh Is Nothing Then
      wsh.name = config
      Call CreateConfigTables(wsh)        'only if a new config storage is created
      Set NewConfigStorage = wsh
      Set wsh = Nothing
   End If
End Function

'------------------------------------
'PRIVATE PROCEDURES
'------------------------------------

Private Sub Class_Initialize()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Class entry point
' Depends on function ConfigStorage and sub SetupNamedRange
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   On Error GoTo FailExit
   Set wshConfig = ConfigStorage 'check existence or create new storage
   If Not wshConfig Is Nothing Then
      Call SetupNamedRange 'check named ranges or create new
   Else: GoTo FailExit
   End If
   setupCompleted = True
   Exit Sub
FailExit:
   MsgBox "Worksheet 'settings' with metadata is not found!" & vbCrLf & _
          "Contact the developer", vbCritical Or vbOKOnly, "System error"
   End  'destruct class if it fails
End Sub

Private Sub Class_Terminate()
'uncomment after testing!
'wshConfig.visible = xlSheetVeryHidden
   If Not wshConfig Is Nothing Then
      Set wshConfig = Nothing
   End If
End Sub

Private Sub SetupNamedRange()
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Check existence or create new named range to store config parameters
' Depends on public sub AddNamedRange
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Dim rng As Range
   Dim valueToSearch As String
   'named range to be used in MetaData class
   If Not RangeExists(tableID) Then
      valueToSearch = Split(modulesTblTitle, ";")(0)
      Set rng = wshConfig.Cells.Find(valueToSearch, wshConfig.Cells(1, 1))
      If Not rng Is Nothing Then Call AddNamedRange(tableID, rng.AddressLocal)
   End If
   'named range for database path string (eg: D:\Database\cookinghouse.accdb)
   If Not RangeExists(db) Then
      Set rng = wshConfig.Cells.Find(db, wshConfig.Cells(1, 1)).Offset(0, 1)
      If Not rng Is Nothing Then Call AddNamedRange(db, rng.AddressLocal)
   End If
   'named range for RibbonPointer (used to restore Ribbon)
   If Not RangeExists(rp) Then
      Set rng = wshConfig.Cells.Find(rp, wshConfig.Cells(1, 1)).Offset(0, 1)
      If Not rng Is Nothing Then Call AddNamedRange(rp, rng.AddressLocal)
   End If
   Set rng = Nothing
End Sub

Private Sub CreateConfigTables(ByVal wSheet As Worksheet)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Routine procedure to create and fill in config tables
' In: target worksheet (wshConfig cannot yet to be used)
' Depends on FormatTable sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Dim title() As String
   Dim modulesItems() As String
   Dim rng As Range
   Dim i As Integer
   If Not wSheet Is Nothing Then
      With wSheet.Cells
         .ClearFormats
         .Font.name = "Courier New"
         .Font.size = 9
      End With
      
      title = Split(configTblTitle, ";")
      Set rng = wSheet.Cells(1, 1).Offset(1, 1).Resize(1, UBound(title) + 1)
      rng.value = title
      Call FormatTable(rng, Style.title)
      Erase title
      Set rng = rng.Offset(1, 0).Resize(1, 1): rng.value = db
      Set rng = rng.Offset(1, 0).Resize(1, 1): rng.value = rp
   
      title = Split(modulesTblTitle, ";")
      Set rng = rng.Offset(2, 0).Resize(1, UBound(title) + 1)
      rng.value = title
      Call FormatTable(rng, Style.title)
      Erase title
      
      title = Split(modulesTblSubTitle, ";")
      Set rng = rng.Offset(1, 0).Resize(1, UBound(title) + 1)
      rng.value = title
      Call FormatTable(rng, Style.subTitle)
      Erase title
      
      modulesItems = Split(modules, ";")
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

Private Sub FormatTable(ByVal rng As Range, ByVal st As Style)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Formatting headings of config tables
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   If Not rng Is Nothing Then
      With rng
         Select Case st
         Case Style.title
            .Interior.Color = RGB(100, 100, 100)
            .Font.Color = vbWhite
         Case Style.subTitle
            .Font.Color = RGB(153, 0, 0)
            .Font.size = 8
         End Select
      End With
   End If
End Sub
