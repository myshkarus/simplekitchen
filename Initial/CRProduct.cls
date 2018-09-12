VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type TController
  tProduct As MProduct
  tUI As VProduct
End Type

Private this As TController

Public Property Get Self() As MProduct
  Set Self = Me
End Property

Public Sub UpdateProduct()
    this.tProduct.UpdateProductRecord this.tUI.GetDisplayedProductRecord
End Sub

Public Sub FindProduct(ByVal ID As String)
    If this.tProduct.GetProductRecord(ID) Then
        this.tUI.ClearProductDisplay
        this.tUI.DisplayProductRecord this.tProduct
    Else
        this.tUI.ClearProductDisplay
    End If
End Sub

Private Sub Class_Initialize()
    Set this.tProduct = New MProduct
    Set this.tUI = New VProduct
End Sub

Private Sub Class_Terminate()
    Set this.tProduct = Nothing
    Set this.tUI = Nothing
End Sub

