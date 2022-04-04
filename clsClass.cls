VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public name As String

Public productCollection As Collection
Public attributeCollection As Collection

Public Sub addProductToCollection(ByRef product As clsProduct)
    productCollection.Add product, product.id
End Sub

Public Sub addAttributeToCollection(ByRef attribut As clsAttribute)
    attributeCollection.Add attribut, attribut.id
End Sub

Private Sub Class_Initialize()
    Set productCollection = New Collection
    Set attributeCollection = New Collection
debug.print "dupsko"
    End Sub