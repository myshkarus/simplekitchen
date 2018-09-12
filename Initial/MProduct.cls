VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iProduct As IMetaData

Private Type product
  TCode As String
  TType As String
  TName As String
  TMeasure As String
  TDescription As String
  TEnergy As Double
  TProtein As Integer
  TFat As Integer
  TCarbohydrate As Integer
  sheet As Worksheet
End Type

Private this As product

Public Property Get Code() As String
  Code = this.TCode
End Property

Public Property Let Code(ByVal value As String)
  this.TCode = value
End Property

Public Property Get ProductType() As String
  ProductType = this.TType
End Property

Public Property Let ProductType(ByVal value As String)
  this.TType = value
End Property

Public Property Get Name() As String
  Name = this.TName
End Property

Public Property Let Name(ByVal value As String)
  this.TName = value
End Property

Public Property Get Measure() As String
  Measure = this.TMeasure
End Property

Public Property Let Measure(ByVal value As String)
  this.TMeasure = value
End Property

Public Property Get Description() As String
  Description = this.TDescription
End Property

Public Property Let Description(ByVal value As String)
  this.TDescription = value
End Property

Public Property Get Energy() As Double
  Energy = this.TEnergy
End Property

Public Property Let Energy(ByVal value As Double)
  this.TEnergy = value
End Property

Public Property Get Protein() As Integer
  Protein = this.TProtein
End Property

Public Property Let Protein(ByVal value As Integer)
  this.TProtein = value
End Property

Public Property Get Fat() As Integer
  Fat = this.TFat
End Property

Public Property Let Fat(ByVal value As Integer)
  this.TFat = value
End Property

Public Property Get Carbohydrate() As Integer
  Carbohydrate = this.TCarbohydrate
End Property

Public Property Let Carbohydrate(ByVal value As Integer)
  this.TCarbohydrate = value
End Property

Public Property Get Self() As MProduct
  Set Self = Me
End Property

Public Function UpdateProductRecord(ByVal pr As MProduct) As Boolean
  UpdateProductRecord = False
  With this.sheet
    .Range("product_вид").value = pr.ProductType
    .Range("product_код").value = pr.Code
    .Range("product_продукт").value = pr.Name
    .Range("product_измерение").value = pr.Measure
    .Range("product_описание").value = pr.Description
    .Range("product_калорийность").value = pr.Energy
    .Range("product_белки").value = pr.Protein
    .Range("product_жиры").value = pr.Fat
    .Range("product_углеводы").value = pr.Carbohydrate
  End With
  UpdateProductRecord = True
End Function

Public Function GetProductRecord(Optional ByVal ID As String) As Boolean
  GetProductRecord = False
  With this.sheet
    Me.ProductType = .Range("product_вид").value
    Me.Code = .Range("product_код").value
    Me.Name = .Range("product_продукт").value
    Me.Measure = .Range("product_измерение").value
    Me.Description = .Range("product_описание").value
    Me.Energy = .Range("product_калорийность").value
    Me.Protein = .Range("product_белки").value
    Me.Fat = .Range("product_жиры").value
    Me.Carbohydrate = .Range("product_углеводы").value
  End With
  GetProductRecord = True
End Function

Private Sub Class_Initialize()
  Set iProduct = New MetaData
  Set this.sheet = iProduct.Create(product)
End Sub

Private Sub Class_Terminate()
  Set iProduct = Nothing
  Set this.sheet = Nothing
End Sub

'Public Sub ProductForm()
'  Dim pView As ProductView
'  Set pView = New ProductView
'  pView.Interface(InterfaceType.ProductForm).Activate
'  Set pView = Nothing
'End Sub
