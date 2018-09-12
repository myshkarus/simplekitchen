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
    .Range("product_вид").value = Empty
    .Range("product_код").value = Empty
    .Range("product_продукт").value = Empty
    .Range("product_измерение").value = Empty
    .Range("product_описание").value = Empty
    .Range("product_калорийность").value = Empty
    .Range("product_белки").value = Empty
    .Range("product_жиры").value = Empty
    .Range("product_углеводы").value = Empty
    End With
    Application.ScreenUpdating = True
End Sub

Public Property Get Self() As MProduct
  Set Self = Me
End Property

Public Sub DisplayProductRecord(product As MProduct)
    Application.ScreenUpdating = False
    With this.sheet
    .Range("product_вид").value = product.ProductType
    .Range("product_код").value = product.Code
    .Range("product_продукт").value = product.Name
    .Range("product_измерение").value = product.Measure
    .Range("product_описание").value = product.Description
    .Range("product_калорийность").value = product.Energy
    .Range("product_белки").value = product.Protein
    .Range("product_жиры").value = product.Fat
    .Range("product_углеводы").value = product.Carbohydrate
    End With
    Application.ScreenUpdating = True
End Sub

Public Function GetDisplayedProductRecord() As MProduct
  Set GetDisplayedProductRecord = New MProduct
  With this.sheet
    GetDisplayedProductRecord.ProductType = .Range("product_вид").value
    GetDisplayedProductRecord.Code = .Range("product_код").value
    GetDisplayedProductRecord.Name = .Range("product_продукт").value
    GetDisplayedProductRecord.Measure = .Range("product_измерение").value
    GetDisplayedProductRecord.Description = .Range("product_описание").value
    GetDisplayedProductRecord.Energy = .Range("product_калорийность").value
    GetDisplayedProductRecord.Protein = .Range("product_белки").value
    GetDisplayedProductRecord.Fat = .Range("product_жиры").value
    GetDisplayedProductRecord.Carbohydrate = .Range("product_углеводы").value
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
