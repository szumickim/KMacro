Attribute VB_Name = "createUploadFileMod"
Option Explicit
 
Private exportPath As String
Private context As String

Private Sub mainImport()

    Dim timeStart As Double
    timeStart = Timer
    
    exportPath = getExportFilePath '"C:\Users\plocitom\Documents\TomaszP\Justyna\K-Macro\export.xlsx"
    
    If Len(exportPath) > 0 Then
    
        Application.ScreenUpdating = False
            Dim classCollection As New Collection
            
            gatherDataFromTemplate classCollection
            
            getKeyValueForChoiceTypeAttributeValue classCollection
            
            ExtractToXML classCollection
        
        Application.ScreenUpdating = True
        
        MsgBox "Done"
        
    End If
    
    Debug.Print Format(Timer - timeStart, "0.0")

End Sub

Private Sub getKeyValueForChoiceTypeAttributeValue(ByRef classCollection As Collection)
    
    'DATA FIELDS SHEET
    Const ADEONA_ID_COLUMN As Long = 1
    Const FIELD_TYPE_COLUMN As Long = 10
    
    Dim wsDataFields As Worksheet
    Set wsDataFields = ThisWorkbook.Sheets("Data fields")
    
    Dim productClass As clsClass, product As clsProduct, attribut As clsAttribute, attributeRow As Long
    
    For Each productClass In classCollection
        For Each product In productClass.productCollection
            For Each attribut In product.attributesCollection
                If Len(attribut.attributeValue) > 0 Then
                    With wsDataFields
                        On Error Resume Next
                            attributeRow = 0
                            attributeRow = WorksheetFunction.Match(attribut.id, .Columns(ADEONA_ID_COLUMN), 0)
                        On Error GoTo 0
                        
                        If attributeRow > 0 Then
                            attribut.attributeType = .Cells(attributeRow, FIELD_TYPE_COLUMN).Value
                            
                            If LCase$(attribut.attributeType) = CHOICE_ATTRIBUTE_TYPE Then
                                getTheKey attribut, productClass.name
                            End If
                        End If
                    End With
                End If
            Next
        Next
    Next

End Sub
Private Sub getTheKey(ByRef attribut As clsAttribute, ByRef productClassName As String)
    
    Const METADATA_ID_COLUMN As Long = 1
    Const GLOBAL_OPTION_COLUMN As Long = 10
    Const SCHEMA_KEYS_COLUMN As Long = 12
    Const GLOBAL_VALUES_COLUMN As Long = 14
    
    Dim wsListSpecification As Worksheet
    Set wsListSpecification = ThisWorkbook.Sheets("Selection list specifications")
    
    With wsListSpecification
        Dim attributeRow As Long
        attributeRow = WorksheetFunction.Match(attribut.id, .Columns(METADATA_ID_COLUMN), 0)
        
        Dim valuesColumn As Long
        If LCase$(.Cells(attributeRow, GLOBAL_OPTION_COLUMN).Value) = "x" Then
            valuesColumn = GLOBAL_VALUES_COLUMN
        Else
            valuesColumn = getLOVColumn(context)
        End If
        
        Dim i As Long
        i = attributeRow + 1 '+1 because attributes starts row below product class
        Do While Len(.Cells(i, SCHEMA_KEYS_COLUMN).Value) > 0
            If InStr(1, .Cells(i, SCHEMA_KEYS_COLUMN).Value, productClassName, vbTextCompare) > 0 Then
                If LCase$(.Cells(i, valuesColumn).Value) = LCase$(attribut.attributeValue) Then
                    attribut.attributeKeyValue = .Cells(i, getLOVColumn("key")).Value
                    Exit Do
                End If
            End If
            i = i + 1
        Loop
        
    End With
    
End Sub

Private Sub gatherDataFromTemplate(ByRef classCollection As Collection)

    Dim wbExport As Workbook
    Set wbExport = Workbooks.Open(exportPath)
    
    Dim ws As Worksheet
    For Each ws In wbExport.Sheets
        
        If ws.name <> "Summary" And ws.name <> "No class" Then
            gatherData classCollection, ws, TEMPLATE_START_ROW
            
            If Len(context) = 0 Then
                context = getTemplateContext(ws)
            End If
        End If
        
    Next
    
    wbExport.Close 0
    
End Sub

Private Function getTemplateContext(ByVal wsTemplate As Worksheet) As String
    
    With wsTemplate
        Dim firstAttributeColumn As Long
        firstAttributeColumn = .Rows(ID_ROW).Find(ATTRIBUTE_PHRASE).Column
        
        Dim firstAttributeName As String
        firstAttributeName = .Cells(NAME_ROW, firstAttributeColumn).Value
    End With
    
    Dim langArray() As Variant
    langArray = Array("en", "fi", "se")
    
    With ThisWorkbook.Sheets("Data fields")
        Dim lang As Variant, attributeRow As Long
        For Each lang In langArray
            On Error Resume Next
                attributeRow = 0
                attributeRow = WorksheetFunction.Match(firstAttributeName, .Columns(getNameColumnBasedOnLanguage(CStr(lang))), 0)
            On Error GoTo 0
            If attributeRow > 0 Then
                getTemplateContext = lang
                Exit For
            End If
        Next
    End With

End Function

Private Sub ExtractToXML(ByVal classCollection As Collection)
    
    Const PRODUCTS_TAG As String = "Products"
    Const PRODUCT_TAG As String = "Product"
    Const VALUES_TAG As String = "Values"
    Const VALUE_TAG As String = "Value"
    Const ID_TAG As String = "ID"
    Const ATTRIBUTE_ID_ATTRIBUTE As String = "AttributeID"
    Const USER_TYPE_ATTRIBUTE As String = "UserTypeID"
    
    Dim objDoc As MSXML2.DOMDocument60
    Set objDoc = New MSXML2.DOMDocument60
    objDoc.resolveExternals = True
    
    'GENERAL CONST INFO
    '---------------------------------
    Dim objNode As MSXML2.IXMLDOMNode
    Set objNode = objDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
    Set objNode = objDoc.InsertBefore(objNode, objDoc.ChildNodes.Item(0))
    
    'STEP-ProductInformation tag
    Dim objRoot As MSXML2.IXMLDOMElement
    Set objRoot = objDoc.createElement("STEP-ProductInformation")
    Set objDoc.DocumentElement = objRoot
    objRoot.setAttribute "WorkspaceID", "Main"
    objRoot.setAttribute "ContextID", context & "-" & UCase$(context)
    objRoot.setAttribute "UseContextLocale", "false"
    'objRoot.setAttribute "AutoApprove", "Y"
    '---------------------------------
    
    'Products tag
    Dim productsNode As MSXML2.IXMLDOMElement
    Set productsNode = objDoc.createElement(PRODUCTS_TAG)
    objRoot.appendChild productsNode
    
    Dim productClass As clsClass, product As clsProduct, productNode As MSXML2.IXMLDOMElement, attribut As clsAttribute, _
        attributeNode As MSXML2.IXMLDOMElement, valuesNode As MSXML2.IXMLDOMElement, basicInfoHeader As clsBasicInfoHeader, _
        basicHeaderNode As MSXML2.IXMLDOMElement, cDataSection As MSXML2.IXMLDOMCDATASection
    
    For Each productClass In classCollection
        For Each product In productClass.productCollection
            'Product tag
            Set productNode = objDoc.createElement(PRODUCT_TAG)
            productNode.setAttribute ID_TAG, product.id
            productNode.setAttribute USER_TYPE_ATTRIBUTE, "PRD_OBJ_mainRecord"
            productsNode.appendChild productNode
            
            'Values tag
            Set valuesNode = objDoc.createElement(VALUES_TAG)
            productNode.appendChild valuesNode
            
            For Each basicInfoHeader In product.basicInfoHeadersCollection
                'Value tag - Basic Info Header
                Select Case CStr(basicInfoHeader.id)
                Case "Short Description Common", "Long Description Common", "Marketing Name", "SEO Text"
                    If Len(basicInfoHeader.val) > 0 Then
                        Set basicHeaderNode = objDoc.createElement(VALUE_TAG)
                        basicHeaderNode.setAttribute ATTRIBUTE_ID_ATTRIBUTE, BASIC_HEADER_PHRASE & Replace(Replace(basicInfoHeader.id, " ", vbNullString), "SEO", "Seo")
                        valuesNode.appendChild basicHeaderNode
                        
                        '<![CDATA[ section
                        Set cDataSection = objDoc.createCDATASection(VALUE_TAG)
                        cDataSection.Data = basicInfoHeader.val
                        
                        basicHeaderNode.appendChild cDataSection
                    End If
                End Select
            Next
            
            For Each attribut In product.attributesCollection
                'Value tag - attributes
                Set attributeNode = objDoc.createElement(VALUE_TAG)
                attributeNode.setAttribute ATTRIBUTE_ID_ATTRIBUTE, ATTRIBUTE_PHRASE & attribut.id
                attributeNode.Text = attribut.attributeValue
                If LCase$(attribut.attributeType) = CHOICE_ATTRIBUTE_TYPE Then
                    attributeNode.setAttribute ID_TAG, attribut.attributeKeyValue
                End If
                
                valuesNode.appendChild attributeNode
            Next
        Next
    Next
    
    'objDoc.Save ThisWorkbook.Path & "\KPIM Import " & " " & Format(Now, "ddMMyyyy-hhmm") & ".xml"
    objDoc.Save createNameForOutput(exportPath, "Import")
End Sub

