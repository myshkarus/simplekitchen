Attribute VB_Name = "RibbonCallbacks"
Option Explicit	'����������� ������ ���������� ���� ���������� � �����

' (���������: customUI, �������: onLoad), 2010+
Sub OnLoad(ribbon As IRibbonUI)
    '�������� ���������� ���������� ������� �����: Public gobjRibbon As IRibbonUI
    Set gobjRibbon = ribbon
End Sub

'tabSimpleKitchen (���������: tab, �������: getLabel), 2010+
'gpHome (���������: group, �������: getLabel), 2010+
'btnHome (���������: button, �������: getLabel), 2010+
'gpSettings (���������: group, �������: getLabel), 2010+
'btnSettings (���������: button, �������: getLabel), 2010+
Sub GetLabel(control As IRibbonControl, ByRef label)
    label = "label �������� " + control.ID
End Sub

'tabSimpleKitchen (���������: tab, �������: getKeytip), 2010+
Sub GetKeyTip(control As IRibbonControl, ByRef keytip)
    keytip = "���"
End Sub

'gpHome (���������: group, �������: getVisible), 2010+
'btnHome (���������: button, �������: getVisible), 2010+
'gpSettings (���������: group, �������: getVisible), 2010+
'btnSettings (���������: button, �������: getVisible), 2010+
Sub GetVisible(control As IRibbonControl, ByRef visible)
    visible = True
End Sub

'gpHome (���������: group, �������: getScreentip), 2010+
'btnHome (���������: button, �������: getScreentip), 2010+
'gpSettings (���������: group, �������: getScreentip), 2010+
'btnSettings (���������: button, �������: getScreentip), 2010+
Sub GetScreenTip(control As IRibbonControl, ByRef screentip)
    screentip = "screentip �������� " + control.ID
End Sub

'gpHome (���������: group, �������: getSupertip), 2010+
'btnHome (���������: button, �������: getSupertip), 2010+
'gpSettings (���������: group, �������: getSupertip), 2010+
'btnSettings (���������: button, �������: getSupertip), 2010+
Sub GetSuperTip(control As IRibbonControl, ByRef supertip)
    supertip = "supertip �������� " + control.ID
End Sub

'btnHome (���������: button, �������: getSize), 2010+
'btnSettings (���������: button, �������: getSize), 2010+
Sub GetSize(control As IRibbonControl, ByRef size)
    size = large
End Sub

'btnHome (���������: button, �������: onAction), 2010+
'btnSettings (���������: button, �������: onAction), 2010+
Sub ButtonOnAction(control As IRibbonControl)
    MsgBox "��������� ������� onAction �������� " + control.ID
End Sub
