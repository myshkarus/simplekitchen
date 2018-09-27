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
' simpleKitchen VBA v.0.1.0
' (c) Mykhailo Shpilienko - https://github.com/myshkarus/simplekitchen
'----------------------------------------------------------------------------------
' Worksheet Start builder for simpleKitchen VBA project
'
' Detailed description:
'   Build Start worksheet from scratch with navigational buttons or kill Start
'   page if one needs it
'
' TODO: hook application state to adjust Start page wallpaper and shape "about"
'
' Status: testing
'
' @class sheetBuilderStart, v.0.3.0
' @author myshkarus
' @hqinside.com
' @license - n/a
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'preinstalled button's captions:
Private Const NAV_BUTTONS As String = "планирование закупок;технологическая карта;подготовка меню;склад;отчеты;справочники"
'button's macros:
' !!! TODO: develop macros
Private Const NAV_MACROS As String = "xxx;aaa;bbb;ccc;ddd;Activate"

'Start page name in Russian or Ukrainian
Private Const SHEET_HOME As String = "старт"
'Start page name in English
'Private Const SHEET_HOME As String = "start"
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
   ' Build start worksheet
   ' Depends on BuildHomePage and AdjustWallpaperSize subs
   ' Out: SheetBuilderStart object (oneself)
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   
   On Error Resume Next                          'to avoid gaps in execution
   If this.tSheet Is Nothing Then
      Set this.tSheet = ThisWorkbook.Worksheets.Add
      this.tSheet.name = SHEET_HOME
      'here build the interface
      Call BuildHomePage
   End If
   On Error GoTo 0
   
   If this.tSheet Is Nothing Then
      GoTo FailExit ' if something goes wrong
   End If
   'always adjust wallpaper size and place shape "about"
   Call AdjustWallpaperSize
   
   Set Build = Self
   Exit Function
FailExit:
   MsgBox "Start sheet '" & SHEET_HOME & "' is not created!" & vbCrLf & _
          "Contact the developer", vbCritical Or vbOKOnly, "System error"
   End                                           'destruct class if it fails
End Function

'------------------------------------
'PUBLIC PROCEDURES
'------------------------------------

Public Sub Activate()
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Activate start worksheet
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   If Not this.tSheet Is Nothing Then
      this.tSheet.Activate
   End If
End Sub

Public Sub Kill()
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Delete start worksheet
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   If Not this.tSheet Is Nothing Then
      Application.DisplayAlerts = False
      this.tSheet.Delete
   End If
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
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Class entry point
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   On Error Resume Next
   Set this.tSheet = ThisWorkbook.Worksheets(SHEET_HOME) ' try to set this.tSheet
   On Error GoTo 0
End Sub

Private Sub Class_Terminate()
   If Not this.tSheet Is Nothing Then
      Call ProtectData(this.tSheet)              ' always protect sheet
      Set this.tSheet = Nothing
   End If
End Sub

Private Sub SetActiveWindow()
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Apply some Application and ActiveWindow properties to adjust wallpaper and
   ' shape "about" size correctly
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   ' Adjust for different monitor resolution
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Dim wallpaper As Shape
   On Error GoTo FailExit
   Call SetActiveWindow
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
   ' Depends on: NewButtonBuilder, NewLabelBuilder constructors, SysInfo class,
   '             IControl interface
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   Dim Button As IControl
   Dim label As IControl
   Dim rng As Range
   Dim image As Shape
   Dim marked As String
   Dim btnBuilder As ButtonBuilder
   Dim lblBuilder As LabelBuilder
   Dim caption As Variant
   Dim macro As Variant
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
   macro = Split(NAV_MACROS, ";")
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
      .OnAction(macro(5)) _
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


