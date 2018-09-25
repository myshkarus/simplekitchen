VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sheetBuilderStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' simpleKitchen VBA v.0.2.0
' (c) Mykhailo Shpilienko - https://github.com/myshkarus/simplekitchen
'----------------------------------------------------------------------------------
' Worksheet Start builder for simpleKitchen VBA project
'
' Detailed description:
'   Build from scratch Start worksheet with navigational buttons or kill Start
'   page if one needs it
'
'TODO: hook application state to adjust Start page wallpaper and shape "about"
'
' Errors:
' Status: development and testing
'
' @class sheetBuilderStart
' @author myshkarus
' @hqinside.com
' @license - n/a
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'preinstalled button's captions:
Private Const NAV_BUTTONS As String = "планирование закупок;технологическая карта;подготовка меню;склад;отчеты;справочники"
'Start page name in Russian
Private Const SHEET_HOME As String = "старт"
Private Const IMG_WALLPAPER As String = "wallpaper"

Private Type startPage
   tSheet As Worksheet
End Type

Private this As startPage

'------------------------------------
'PUBLIC FUNCTIONS
'------------------------------------

Public Function Build() As sheetBuilderStart
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Return starting worksheet with navigational buttons
   ' Depends on .....
   ' Out: Worksheet object
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   
   If this.tSheet Is Nothing Then
      On Error GoTo 0                            'to avoid gaps in execution
      Set this.tSheet = ThisWorkbook.Worksheets.Add
      this.tSheet.name = SHEET_HOME
      Call BuildHomePage
   End If
   
   On Error GoTo FailStartupExit
   Call SetActiveWindow
   Call AdjustWallpaperSize
   
   Set Build = Self                              'do nothing if Start page exists already
   Exit Function
FailStartupExit:
   MsgBox "Start sheet '" & SHEET_HOME & "' is not created!" & vbCrLf & _
          "Contact the developer", vbCritical Or vbOKOnly, "System error"
   End                                           'destruct class if it fails
End Function

'------------------------------------
'PUBLIC PROCEDURES
'------------------------------------

Public Sub Activate()
   this.tSheet.Activate
End Sub

'------------------------------------
'PRIVATE PROPERTIES
'------------------------------------

Private Property Get Self() As sheetBuilderStart
   Set Self = Me
End Property

'------------------------------------
'PRIVATE PROCEDURES
'------------------------------------

Private Sub Class_Initialize()
   If SheetExist(SHEET_HOME) Then
      Set this.tSheet = ThisWorkbook.Worksheets(SHEET_HOME)
   End If
End Sub

Private Sub Class_Terminate()
   Call ProtectData(this.tSheet)
   Set this.tSheet = Nothing
End Sub

Private Sub SetActiveWindow()
   Application.DisplayFormulaBar = False
   With ActiveWindow
      .WindowState = xlMaximized
      .DisplayHeadings = False
      .DisplayWorkbookTabs = False
      .DisplayHorizontalScrollBar = False
      .DisplayGridlines = False
   End With
End Sub

Private Sub AdjustWallpaperSize()
   Dim wallpaper As Shape
   On Error GoTo FailExit
   Set wallpaper = this.tSheet.Shapes(IMG_PICTURE)
   With wallpaper
      .height = Application.UsableHeight + 15
      .Width = .height / 4
   End With
   With this.tSheet.Shapes("about")
      .LockAspectRatio = msoFalse
      .Left = wallpaper.Left
      .top = wallpaper.height - cs.Large - 25
      .Width = wallpaper.Width
      .height = cs.Large
   End With
   Set wallpaper = Nothing
   Exit Sub
FailExit:
   Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub BuildHomePage()
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Building home page from scratch
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Dim Button As IControl
   Dim label As IControl
   Dim rng As Range
   Dim image As Shape
   Dim marked As String
   Dim btnBuilder As ButtonBuilder
   Dim lblBuilder As LabelBuilder
   Dim caption As Variant
   Dim imgName As String
   Dim about As String
   Dim gap As Long
   
   On Error GoTo FailExit
   
   Application.ScreenUpdating = False
   
   about = SysInfo.AddInName & " " & SysInfo.AddInVersion & vbCrLf & Chr(169) & " " & URL_AUTHOR
   Set btnBuilder = NewButtonBuilder(this.tSheet)
   Set lblBuilder = NewLabelBuilder(this.tSheet)
   marked = "nav"
   this.tSheet.Activate
   this.tSheet.Range("A1").EntireColumn.Hidden = True
   
   ThisWorkbook.Worksheets(SHEET_RESOURCES).Shapes(IMG_WALLPAPER).Copy
   this.tSheet.Paste
   
   imgName = IMG_PICTURE
   
   With this.tSheet.Shapes(IMG_WALLPAPER)
      .LockAspectRatio = msoFalse
      .height = Application.UsableHeight + 15
      .Width = .height / 4
      .Left = this.tSheet.Range("A1").Left
      .top = this.tSheet.Range("A1").top
      .Placement = xlFreeFloating
      .name = imgName
   End With
    
   Set image = this.tSheet.Shapes(imgName)
   Set Button = btnBuilder.caption(about) _
      .name("about") _
      .OfPlaceAndSize(0, 0, 1, 1) _
      .Style(ControlStyle.csElegant) _
      .Movable(False) _
      .fontSize(fs.ButtonSmall) _
      .visible(True) _
      .Build

   caption = Split(NAV_BUTTONS, ";")
   gap = cs.Middle
   Set rng = this.tSheet.Cells(1, this.tSheet.Shapes(imgName).BottomRightCell.column).Offset(2, 1)
   
   Set Button = btnBuilder.caption(caption(0)) _
      .mark(marked) _
      .OfPlaceAndSize(rng.Left, rng.top, 500, cs.Double) _
      .Style(ControlStyle.csNavigation) _
      .Movable(False) _
      .fontSize(fs.Title) _
      .visible(True) _
      .Build
   Set Button = btnBuilder.caption(caption(1)) _
      .mark(marked) _
      .OfPlaceAndSize(Button.Left, Button.Bottom + gap, 500, cs.Double) _
      .Style(ControlStyle.csNavigation) _
      .Movable(False) _
      .fontSize(fs.Title) _
      .visible(True) _
      .Build
   Set Button = btnBuilder.caption(caption(2)) _
      .mark(marked) _
      .OfPlaceAndSize(Button.Left, Button.Bottom + gap, 500, cs.Double) _
      .Style(ControlStyle.csNavigation) _
      .Movable(False) _
      .fontSize(fs.Title) _
      .visible(True) _
      .Build
   Set Button = btnBuilder.caption(caption(3)) _
      .mark(marked) _
      .OfPlaceAndSize(Button.Left, Button.Bottom + gap, 500, cs.Double) _
      .Style(ControlStyle.csNavigation) _
      .Movable(False) _
      .fontSize(fs.Title) _
      .visible(True) _
      .Build
   Set Button = btnBuilder.caption(caption(4)) _
      .mark(marked) _
      .OfPlaceAndSize(Button.Left, Button.Bottom + gap, 500, cs.Double) _
      .Style(ControlStyle.csNavigation) _
      .Movable(False) _
      .fontSize(fs.Title) _
      .visible(True) _
      .Build
   Set Button = btnBuilder.caption(caption(5)) _
      .mark(marked) _
      .OfPlaceAndSize(Button.Left, Button.Bottom + gap, 500, cs.Double) _
      .Style(ControlStyle.csNavigation) _
      .Movable(False) _
      .fontSize(fs.Title) _
      .visible(True) _
      .Build
   Application.ScreenUpdating = True
   this.tSheet.Range("A2").Select
  
   Set btnBuilder = Nothing
   Set lblBuilder = Nothing
   Set Button = Nothing
   Set label = Nothing
   Set rng = Nothing
   Set image = Nothing
   Exit Sub
FailExit:
   Err.Raise Err.Number, Err.Source, Err.description
End Sub


