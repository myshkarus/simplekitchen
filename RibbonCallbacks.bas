Attribute VB_Name = "RibbonCallbacks"
Option Explicit	'Потребовать явного объявления всех переменных в файле

' (компонент: customUI, атрибут: onLoad), 2010+
Sub OnLoad(ribbon As IRibbonUI)
    'Объявите глобальную переменную объекта ленты: Public gobjRibbon As IRibbonUI
    Set gobjRibbon = ribbon
End Sub

'tabSimpleKitchen (компонент: tab, атрибут: getLabel), 2010+
'gpHome (компонент: group, атрибут: getLabel), 2010+
'btnHome (компонент: button, атрибут: getLabel), 2010+
'gpSettings (компонент: group, атрибут: getLabel), 2010+
'btnSettings (компонент: button, атрибут: getLabel), 2010+
Sub GetLabel(control As IRibbonControl, ByRef label)
    label = "label элемента " + control.ID
End Sub

'tabSimpleKitchen (компонент: tab, атрибут: getKeytip), 2010+
Sub GetKeyTip(control As IRibbonControl, ByRef keytip)
    keytip = "ЁПТ"
End Sub

'gpHome (компонент: group, атрибут: getVisible), 2010+
'btnHome (компонент: button, атрибут: getVisible), 2010+
'gpSettings (компонент: group, атрибут: getVisible), 2010+
'btnSettings (компонент: button, атрибут: getVisible), 2010+
Sub GetVisible(control As IRibbonControl, ByRef visible)
    visible = True
End Sub

'gpHome (компонент: group, атрибут: getScreentip), 2010+
'btnHome (компонент: button, атрибут: getScreentip), 2010+
'gpSettings (компонент: group, атрибут: getScreentip), 2010+
'btnSettings (компонент: button, атрибут: getScreentip), 2010+
Sub GetScreenTip(control As IRibbonControl, ByRef screentip)
    screentip = "screentip элемента " + control.ID
End Sub

'gpHome (компонент: group, атрибут: getSupertip), 2010+
'btnHome (компонент: button, атрибут: getSupertip), 2010+
'gpSettings (компонент: group, атрибут: getSupertip), 2010+
'btnSettings (компонент: button, атрибут: getSupertip), 2010+
Sub GetSuperTip(control As IRibbonControl, ByRef supertip)
    supertip = "supertip элемента " + control.ID
End Sub

'btnHome (компонент: button, атрибут: getSize), 2010+
'btnSettings (компонент: button, атрибут: getSize), 2010+
Sub GetSize(control As IRibbonControl, ByRef size)
    size = large
End Sub

'btnHome (компонент: button, атрибут: onAction), 2010+
'btnSettings (компонент: button, атрибут: onAction), 2010+
Sub ButtonOnAction(control As IRibbonControl)
    MsgBox "Сработала функция onAction элемента " + control.ID
End Sub
