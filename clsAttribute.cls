VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAttribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public name As String
Public id As String
Public attributeType As String
Public attributeMaxLength As String
Public attributeFormat As String
Public attributeValue As String
Public LOV As Object

Public attributeKeyValue As String
