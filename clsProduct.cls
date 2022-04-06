VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public id As String
Public productClass As String

Public attributesCollection As Collection
Public basicInfoHeadersCollection As Collection

Public Sub addBasicInfoHeaderToCollection(ByRef basicInfoHeader As clsBasicInfoHeader)
    If Len(basicInfoHeader.id) > 0 Then
        basicInfoHeadersCollection.Add basicInfoHeader, basicInfoHeader.id
    End If
End Sub

Public Sub addAttributeToCollection(ByRef attribut As clsAttribute)
    attributesCollection.Add attribut, attribut.id
End Sub

Private Sub Class_Initialize()
    Set attributesCollection = New Collection
    Set basicInfoHeadersCollection = New Collection
End Sub

Public Sub addValueIntoPropertyBasedOnHeaderID(ByVal headerID As String, ByVal propertyValue As String)
    
    Select Case headerID
        Case "<ID>", "Maintenance record ID"
            id = propertyValue
        Case "CLS_LINK_KeskoSchema", "Tech schema classification", "<Kesko Attribute Schema>"
            productClass = propertyValue
    End Select

End Sub

