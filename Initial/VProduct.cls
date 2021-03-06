VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iProduct As IMetaData

Private Type VProduct
  sheet As Worksheet
End Type

Private this As VProduct

Public Sub ClearProductDisplay()
    Application.ScreenUpdating = False
    With this.sheet
    .Range("product_���").value = Empty
    .Range("product_���").value = Empty
    .Range("product_�������").value = Empty
    .Range("product_���������").value = Empty
    .Range("product_��������").value = Empty
    .Range("product_������������").value = Empty
    .Range("product_�����").value = Empty
    .Range("product_����").value = Empty
    .Range("product_��������").value = Empty
    End With
    Application.ScreenUpdating = True
End Sub

Public Property Get Self() As MProduct
  Set Self = Me
End Property

Public Sub DisplayProductRecord(product As MProduct)
    Application.ScreenUpdating = False
    With this.sheet
    .Range("product_���").value = product.ProductType
    .Range("product_���").value = product.Code
    .Range("product_�������").value = product.Name
    .Range("product_���������").value = product.Measure
    .Range("product_��������").value = product.Description
    .Range("product_������������").value = product.Energy
    .Range("product_�����").value = product.Protein
    .Range("product_����").value = product.Fat
    .Range("product_��������").value = product.Carbohydrate
    End With
    Application.ScreenUpdating = True
End Sub

Public Function GetDisplayedProductRecord() As MProduct
  Set GetDisplayedProductRecord = New MProduct
  With this.sheet
    GetDisplayedProductRecord.ProductType = .Range("product_���").value
    GetDisplayedProductRecord.Code = .Range("product_���").value
    GetDisplayedProductRecord.Name = .Range("product_�������").value
    GetDisplayedProductRecord.Measure = .Range("product_���������").value
    GetDisplayedProductRecord.Description = .Range("product_��������").value
    GetDisplayedProductRecord.Energy = .Range("product_������������").value
    GetDisplayedProductRecord.Protein = .Range("product_�����").value
    GetDisplayedProductRecord.Fat = .Range("product_����").value
    GetDisplayedProductRecord.Carbohydrate = .Range("product_��������").value
  End With
End Function

Private Sub Class_Initialize()
  Set iProduct = New MetaData
  Set this.sheet = iProduct.Create(product)
End Sub

Private Sub Class_Terminate()
  Set iProduct = Nothing
  Set this.sheet = Nothing
End Sub
