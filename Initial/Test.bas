Attribute VB_Name = "Test"
Option Explicit


Sub test2()
Dim r As ProductView
Dim reff As ReferencesView

Set r = New ProductView
r.Interface(InterfaceType.ProductList).Activate

Set reff = New ReferencesView
reff.Interface(InterfaceType.ReferenceProductType).Activate

End Sub

