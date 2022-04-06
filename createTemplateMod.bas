Attribute VB_Name = "createTemplateMod"
Option Explicit

Public Const NAME_ROW As Long = 2
Public Const ID_ROW As Long = 1
Public Const IMPORT_START_ROW As Long = 2
Public Const TEMPLATE_START_ROW As Long = 3

Public Const IMPORT_FILE_TYPE As Long = 1
Public Const TEMPLATE_FILE_TYPE As Long = 2

Public Const ATTRIBUTE_PHRASE As String = "ATR_KS_"
Public Const BASIC_HEADER_PHRASE As String = "ATR_"
Public Const CHOICE_ATTRIBUTE_TYPE As String = "choice"

Const NUMBER_ATTRIBUTE_TYPE As String = "number"
Const TEXT_ATTRIBUTE_TYPE As String = "text"
Const LONG_TEXT_ATTRIBUTE_TYPE As String = "long text"
Const STRING_TO_REPLACE As String = "variable"
Const EAN_PHRASE As String = "Primary Identification" '"EAN"
Const KESKO_ATTR_SCHEMA As String = "<Kesko Attribute Schema>"

Private exportPath As String
Private language As String
    
Private Sub mainTemplate()
    
    Dim timeStart As Double
    timeStart = Timer
    
    exportPath = getExportFilePath '"C:\Users\plocitom\Documents\TomaszP\Justyna\K-Macro\export.xlsx"
    
    If Len(exportPath) > 0 Then
    
        Application.ScreenUpdating = False
            
            Progression.Show
                Dim classCollection As New Collection
                
                Progression.Text2 = "Gather data from export file"
                gatherDataFromExport classCollection
                
                Progression.Text2 = "Gather data from RautaKesko file"
                gatherDataFromRautakeskoDataFields classCollection
                
                Progression.Text2 = "Create Template"
                createTemplate classCollection
            
            Unload Progression
            
        Application.ScreenUpdating = True
        
        MsgBox "Done"
        
    End If
    
    Debug.Print Format(Timer - timeStart, "0.0")
    
End Sub

Private Sub createTemplate(ByVal classCollection As Collection)

    Dim wbTemplate As Workbook
    Set wbTemplate = Workbooks.Add
    
    'Create class worksheets
    Progression.Text2 = "Create class worksheets"
    
    Dim productClass As clsClass, wsClass As Worksheet, i As Long, productCount As Long, totalProductCount As Long
    For Each productClass In classCollection
        Set wsClass = wbTemplate.Sheets.Add(before:=wbTemplate.Sheets(1))
        wsClass.name = productClass.name
        productCount = productClass.productCollection.Count
        
        fillNewClassSheetWithBasicHeaders wsClass, productCount, TEMPLATE_FILE_TYPE
        
        fillNewClassSheetWithAllClassAttribues wsClass, productClass
        
        fillNewClassSheetWithProductValues wsClass, productClass, TEMPLATE_START_ROW, TEMPLATE_FILE_TYPE
        
        formatNewClassSheet wsClass, productCount, TEMPLATE_FILE_TYPE
        
        totalProductCount = totalProductCount + productCount
        i = i + 1
        progress i, classCollection.Count
    Next
    
    'Create summary sheet
    Progression.Text2 = "Create Summary worksheet"
    Set wsClass = wbTemplate.Sheets.Add(after:=wbTemplate.Sheets(wbTemplate.Sheets.Count - 1))
    wsClass.name = "Summary"
    fillNewClassSheetWithBasicHeaders wsClass, totalProductCount, TEMPLATE_FILE_TYPE
    
    i = TEMPLATE_START_ROW
    For Each productClass In classCollection
        fillNewClassSheetWithProductValues wsClass, productClass, i, TEMPLATE_FILE_TYPE
        'i = i + productClass.productCollection.Count
    Next
    
    formatNewClassSheet wsClass, totalProductCount, TEMPLATE_FILE_TYPE
    
    'Save Template
    saveTemplateAs wsClass
    
End Sub

Public Sub saveTemplateAs(ByVal wsClass As Worksheet)
    Dim wbClass As Workbook
    Set wbClass = wsClass.Parent
    
    'Delete unnecessary sheet
    Application.DisplayAlerts = False
        On Error Resume Next
            wbClass.Sheets(wbClass.Sheets.Count).Delete
        On Error GoTo 0
    Application.DisplayAlerts = True
        
    'wbClass.SaveAs exportPath & "KPIM Template " & fileName & " " & Format(Now, "ddMMyyyy-hhmm") & ".xlsx"
    wbClass.SaveAs createNameForOutput(exportPath, "Template")
    
    wbClass.Close 0

End Sub

Public Function createNameForOutput(ByVal filePath As String, ByVal outputType As String) As String

    Dim fileName As String
    fileName = Mid$(filePath, InStrRev(filePath, "\") + 1)
    fileName = Left$(fileName, InStr(fileName, ".") - 1)
    
    Dim folderPath As String
    folderPath = Left$(filePath, InStrRev(filePath, "\"))
    
    If outputType = "Template" Then
        createNameForOutput = folderPath & "KPIM " & outputType & " " & fileName & " " & Format(Now, "ddMMyyyy-hhmm") & ".xlsx"
    ElseIf outputType = "Import" Then
        If InStr(1, fileName, "KPIM Template") > 0 Then
            fileName = Replace(fileName, "Template", "Import")
        Else
            fileName = "KPIM Import " & fileName
        End If
        
        createNameForOutput = folderPath & fileName & ".xml" ' & " " & Format(Now, "ddMMyyyy-hhmm") & ".xml"
    End If
    
End Function

Public Sub formatNewClassSheet(ByRef wsClass As Worksheet, ByVal productCount As Long, ByVal fileType As Long)

    With wsClass
        Dim firstAttributeColumn As Long
        On Error Resume Next
            firstAttributeColumn = .Rows(ID_ROW).Find(ATTRIBUTE_PHRASE).Column
        On Error GoTo 0
        
        If firstAttributeColumn = 0 Then
            firstAttributeColumn = .Cells(ID_ROW, .Columns.Count).End(xlToLeft).Column + 1
        End If

        'INFO COLUMNS WIDTH
        Dim j As Long
        For j = 1 To firstAttributeColumn - 1
            .Columns(j).ColumnWidth = 20
        Next
        
        On Error Resume Next
            Dim eanColumn As Long
            eanColumn = .Rows(ID_ROW).Find(EAN_PHRASE).Column
            
            Dim lastRow As Long
            lastRow = .Cells(.Rows.Count, eanColumn).End(xlUp).Row
            .Range(.Cells(TEMPLATE_START_ROW, eanColumn), .Cells(lastRow, eanColumn)).NumberFormat = "0"
        On Error GoTo 0
        
        Dim lastCol As Long
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        
        'ATTRIBUTE COLUMNS WIDTH
        For j = firstAttributeColumn To lastCol
            .Columns(j).ColumnWidth = 15
        Next
        
        If fileType = TEMPLATE_FILE_TYPE Then
            With .Rows(NAME_ROW)
                .WrapText = True
                .RowHeight = 35
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            'HEADER ROW BOLD
            .Range(.Cells(NAME_ROW, 1), .Cells(NAME_ROW, lastCol)).Font.Bold = True
            With .Range(.Cells(ID_ROW, 1), .Cells(ID_ROW, lastCol)).Font
                .Bold = True
                .Color = vbWhite
            End With
            
            With .Rows(TEMPLATE_START_ROW & ":" & (TEMPLATE_START_ROW + productCount - 1))
                If wsClass.name = "Summary" Then
                    .RowHeight = 30
                Else
                    .RowHeight = 15
                End If
                .VerticalAlignment = xlCenter
            End With
        End If
        
    End With

End Sub

Public Sub fillNewClassSheetWithProductValues(ByRef wsClass As Worksheet, ByRef productClass As clsClass, ByRef i As Long, ByVal fileType As Long)
    
    With wsClass
        Dim firstAttributeColumn As Long
        On Error Resume Next
            firstAttributeColumn = .Rows(ID_ROW).Find(ATTRIBUTE_PHRASE).Column
        On Error GoTo 0
        
        If firstAttributeColumn = 0 Then
            firstAttributeColumn = .Cells(ID_ROW, .Columns.Count).End(xlToLeft).Column + 1
        End If
        
        Dim product As clsProduct, attribut As clsAttribute, attributColumn As Long, _
            tempHeader As String, j As Long, lastCol As Long
        
        For Each product In productClass.productCollection
            For j = 1 To firstAttributeColumn - 1
                tempHeader = .Cells(ID_ROW, j).Value
                On Error Resume Next
                    .Cells(i, j).Value = product.basicInfoHeadersCollection(tempHeader).val
                On Error GoTo 0
            Next
            
            If wsClass.name <> "Summary" Then
                For Each attribut In product.attributesCollection
                    On Error Resume Next
                        attributColumn = 0
                        attributColumn = WorksheetFunction.Match(ATTRIBUTE_PHRASE & attribut.id, .Rows(ID_ROW), 0)
                    On Error GoTo 0
                    
                    If attributColumn > 0 Then
                        With .Cells(i, attributColumn)
                            .Value = attribut.attributeValue
                        End With
                    ElseIf attributColumn = 0 And fileType = IMPORT_FILE_TYPE Then
                        lastCol = .Cells(ID_ROW, .Columns.Count).End(xlToLeft).Column + 1
                        '.Cells(NAME_ROW, lastCol).Value = attribut.name
                        .Cells(ID_ROW, lastCol).Value = ATTRIBUTE_PHRASE & attribut.id
                        .Cells(i, lastCol).Value = attribut.attributeValue
                    
    'OUT OF CLASS ATTRIBUTES
    '                ElseIf attributColumn = 0 'If attribute out of class
    '                    lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
    '                    .Cells(NAME_ROW, lastCol + 1).Value = attribut.name
    '                    .Cells(ID_ROW, lastCol + 1).Value = attribut.id
    '                    .Cells(i, lastCol + 1).Value = attribut.attributeValue
                    End If
                Next
            End If
            
            i = i + 1
        Next
    End With

End Sub

Private Sub fillNewClassSheetWithAllClassAttribues(ByRef wsClass As Worksheet, ByRef productClass As clsClass)
    
    'Data fields sheet
    Const TECH_ATTRIBUTE_START_ROW As Long = 110
    Const TECH_ATTRIBUTE_END_ROW As Long = 691
    
    Const TECH_ATTRIBUTE_COLUMN As Long = 1
    
    Dim listSeparator As String
    listSeparator = "," 'Application.International(xlListSeparator)
    
    Dim productCount As Long
    productCount = productClass.productCollection.Count
    
    With wsClass
        Dim firstAttributeColumn As Long
        On Error Resume Next
            firstAttributeColumn = .Rows(ID_ROW).Find(ATTRIBUTE_PHRASE).Column
        On Error GoTo 0
        
        If firstAttributeColumn = 0 Then
            firstAttributeColumn = .Cells(ID_ROW, .Columns.Count).End(xlToLeft).Column + 1
        End If
        
        Dim j As Long
        j = firstAttributeColumn
        
        Dim attribut As clsAttribute, LOVDictionary As Object, techAttributeRow As Long, _
            conditionalFormula As String, k As Long
        
        For Each attribut In productClass.attributeCollection
            'Check if attribute is "technical attribute" - exclude Atmosferic image, safety instruction, etc
            With ThisWorkbook.Sheets("Data fields")
                On Error Resume Next
                    techAttributeRow = 0
                    techAttributeRow = WorksheetFunction.Match(Trim$(attribut.id), .Range(.Cells(TECH_ATTRIBUTE_START_ROW, TECH_ATTRIBUTE_COLUMN), .Cells(TECH_ATTRIBUTE_END_ROW, TECH_ATTRIBUTE_COLUMN)), 0)
                On Error GoTo 0
            End With
            
            If techAttributeRow > 0 Then
                .Cells(NAME_ROW, j).Value = Trim$(attribut.name)
                .Cells(ID_ROW, j).Value = Trim$(ATTRIBUTE_PHRASE & attribut.id)
                
                'Add some color for values & text formatting
                With .Range(.Cells(TEMPLATE_START_ROW, j), .Cells(TEMPLATE_START_ROW + productCount - 1, j))
                    .Interior.Color = RGB(174, 211, 252)
                    .NumberFormat = "@"
                End With
                
                If LCase$(attribut.attributeType) = NUMBER_ATTRIBUTE_TYPE Then
                    .Range(.Cells(TEMPLATE_START_ROW, j), .Cells(TEMPLATE_START_ROW + productCount - 1, j)).NumberFormat = attribut.attributeFormat
                    conditionalFormula = "=AND(ISNUMBER(" & STRING_TO_REPLACE & ")=FALSE" & listSeparator & "LEN(" & STRING_TO_REPLACE & ")>0)"
                    
                    addValidationListsAndConditionalFormattingToProducts wsClass, productCount, j, , conditionalFormula
                    
                ElseIf LCase$(attribut.attributeType) = TEXT_ATTRIBUTE_TYPE Or LCase$(attribut.attributeType) = LONG_TEXT_ATTRIBUTE_TYPE Then
                    If LenB(attribut.attributeMaxLength) <> 0 Then
                        If InStrB(attribut.attributeMaxLength, " ") Then
                            conditionalFormula = "=LEN(" & STRING_TO_REPLACE & ") > " & Split(attribut.attributeMaxLength, " ")(1)
                        Else
                            conditionalFormula = "=LEN(" & STRING_TO_REPLACE & ") > " & attribut.attributeMaxLength
                        End If
                        
                        addValidationListsAndConditionalFormattingToProducts wsClass, productCount, j, , conditionalFormula
                    End If
                    
                ElseIf LCase$(attribut.attributeType) = CHOICE_ATTRIBUTE_TYPE Then
                    If Trim$(attribut.name) <> "Size" And Trim$(attribut.id) <> 756 Then
                        Set LOVDictionary = setInAlfabeticOrder(attribut.LOV)
                    Else
                        Set LOVDictionary = attribut.LOV
                    End If
                    
                    With .Range(.Cells(TEMPLATE_START_ROW, j), .Cells(TEMPLATE_START_ROW + productCount - 1, j))
                        .NumberFormat = "0.#"
                    End With
                    
                    addValidationListsAndConditionalFormattingToProducts wsClass, productCount, j, LOVDictionary
                    
                End If
                
                j = j + 1
                k = k + 1
                Progression.Text2 = "Create class worksheets: attributes " & k & "/" & productClass.attributeCollection.Count
                DoEvents
            End If
        Next
    End With

End Sub

Private Sub addValidationListsAndConditionalFormattingToProducts(ByRef wsClass As Worksheet, ByVal productCount As Long, ByVal j As Long, Optional ByRef LOVDictionary As Object, Optional ByRef conditionalFormulaBase As String)
    
    If Not LOVDictionary Is Nothing Then
        Dim listSeparator As String
        listSeparator = "," 'Application.International(xlListSeparator)
        
        'Create: list of values for validation & conditional formula for conditional formatting
        Dim i As Long, LOVDicValue As String
        For i = 0 To LOVDictionary.Count - 1
            If IsNumeric(LOVDictionary.items()(i)) Then
                LOVDicValue = LOVDictionary.items()(i)
            Else
                LOVDicValue = """" & LOVDictionary.items()(i) & """"
            End If
            
            If i > 0 Then
                conditionalFormulaBase = conditionalFormulaBase & listSeparator & STRING_TO_REPLACE & "<>" & LOVDicValue
            Else
                conditionalFormulaBase = STRING_TO_REPLACE & "<>" & LOVDicValue
            End If
        Next
        
        conditionalFormulaBase = "=AND(" & conditionalFormulaBase & ")"
    
        'Add list of values UNDER products
        Dim LOVStartRow As Long
        LOVStartRow = TEMPLATE_START_ROW + productCount + 1
        For i = 0 To LOVDictionary.Count - 1
            With wsClass.Cells(LOVStartRow + i, j)
                .Value = LOVDictionary.items()(i)
                '.NumberFormat = "0.#"
            End With
        Next
    End If
    
    With wsClass
        Dim cellAddress As String, conditionalFormulaInput As String
        For i = TEMPLATE_START_ROW To (TEMPLATE_START_ROW + productCount - 1)
            cellAddress = .Range(.Cells(i, j), .Cells(i, j)).Address
            conditionalFormulaInput = Replace(conditionalFormulaBase, STRING_TO_REPLACE, cellAddress)
            
            With .Cells(i, j)
                If Not LOVDictionary Is Nothing Then
                    'Validation
                    With .Validation
                        .Delete
                        .Add Type:=xlValidateList, Operator:=xlBetween, _
                            Formula1:="='" & wsClass.name & "'!" & wsClass.Range(wsClass.Cells(LOVStartRow, j), wsClass.Cells(LOVStartRow + LOVDictionary.Count - 1, j)).Address
                        .IgnoreBlank = True
                        .InCellDropdown = True
                        .ShowInput = True
                        .ShowError = True
                    End With
                End If
                
                If Len(conditionalFormulaBase) > 0 Then
                    'Conditional formatting
                    With .FormatConditions.Add(Type:=xlExpression, _
                        Formula1:=ConvertToLocalizedFormula(conditionalFormulaInput, ThisWorkbook.Sheets("Headers").Range("C1")))
                        .Font.Color = vbRed
                        .StopIfTrue = True
                    End With
                End If
            End With
        Next
        
    End With

End Sub

Private Function ConvertToLocalizedFormula(ByVal formulaToConvert As String, ByRef tempCell As Range) As String

    tempCell.Formula = formulaToConvert
    ConvertToLocalizedFormula = tempCell.FormulaLocal
    
    tempCell.Clear
    
End Function

Public Sub fillNewClassSheetWithBasicHeaders(ByRef wsClass As Worksheet, ByRef productCount As Long, ByRef filetypeColumn As Long)
    Const ID_COLUMN As Long = 13
    
    Dim wsHeaders As Worksheet
    Set wsHeaders = ThisWorkbook.Sheets("Headers")
    
    With wsHeaders
        Dim lastRow As Long
        lastRow = .Cells(.Rows.Count, filetypeColumn).End(xlUp).Row
        
        Dim headerArray As Variant
        headerArray = WorksheetFunction.Transpose(.Range(.Cells(1, filetypeColumn), .Cells(lastRow, filetypeColumn)).Value)
    End With
    
    Dim columnCounter As Long
    columnCounter = 1
    
    Dim j As Long
    For j = LBound(headerArray) To UBound(headerArray)
        With wsHeaders
            wsClass.Cells(ID_ROW, columnCounter).Value = headerArray(j)
            
            If filetypeColumn = 2 Then
                wsClass.Cells(NAME_ROW, columnCounter).Value = .Cells(WorksheetFunction.Match(CStr(headerArray(j)), .Columns(ID_COLUMN), 0), getLOVColumn(language)).Value
                
                If headerArray(j) = "Short Description Common" Then
                    createShortDescriptionCommonPackage wsClass, columnCounter, productCount
                End If
            End If
        End With
        columnCounter = columnCounter + 1
    Next
    
End Sub

Private Sub createShortDescriptionCommonPackage(ByRef wsClass As Worksheet, ByRef columnCounter As Long, ByRef productCount As Long)
    
    Const SHORT_DESC_PACKAGE_COL As Long = 4
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    With ThisWorkbook.Sheets("Headers")
        Dim lastRow As Long
        lastRow = .Cells(.Rows.Count, SHORT_DESC_PACKAGE_COL).End(xlUp).Row
        
        Dim i As Long
        For i = 1 To lastRow
            dict.Add .Cells(i, SHORT_DESC_PACKAGE_COL).Value, .Cells(i, SHORT_DESC_PACKAGE_COL + 1).Value & ";" & .Cells(i, SHORT_DESC_PACKAGE_COL + 2).Value
        Next
    End With
    
    Dim j As Long, formulaValue As String, conditionValue As String, conditionalFormula As String
    
    With wsClass
        For j = 0 To dict.Count - 1
            columnCounter = columnCounter + 1
            .Cells(NAME_ROW, columnCounter).Value = dict.keys()(j)
            formulaValue = Split(dict.items()(j), ";")(0)
            conditionValue = Split(dict.items()(j), ";")(1)
            
            For i = TEMPLATE_START_ROW To TEMPLATE_START_ROW + productCount - 1
                .Cells(i, columnCounter).Formula = "=" & formulaValue
            Next
            
            If Len(conditionValue) > 0 Then
                conditionalFormula = "=" & STRING_TO_REPLACE & ">" & conditionValue
                addValidationListsAndConditionalFormattingToProducts wsClass, productCount, columnCounter, , conditionalFormula
            End If
            
        Next
    End With

End Sub

Public Function getLOVColumn(ByVal lang As String) As Long
    'Columns based on "Selection list specifications" from Rautekesko data
    Const KEY_COLUMN As Long = 13
    Const GLOBAL_VERSION_COLUMN As Long = 14
    Const FINNISH_VERSION_COLUMN As Long = 15
    Const SWEDISH_VERSION_COLUMN As Long = 16
    Const ESTONIAN_VERSION_COLUMN As Long = 17
    Const LITHUANIAN_VERSION_COLUMN As Long = 18
    Const RUSSIAN_VERSION_COLUMN As Long = 19
    
    With ThisWorkbook.Sheets("Main")
        
        Select Case lang
            Case "en"
                getLOVColumn = GLOBAL_VERSION_COLUMN
            Case "fi"
                getLOVColumn = FINNISH_VERSION_COLUMN
            Case "se"
                getLOVColumn = SWEDISH_VERSION_COLUMN
            Case "key"
                getLOVColumn = KEY_COLUMN
                
            Case "ee"
                getLOVColumn = ESTONIAN_VERSION_COLUMN
            Case "lv"
                getLOVColumn = LITHUANIAN_VERSION_COLUMN
            Case "ru"
                getLOVColumn = RUSSIAN_VERSION_COLUMN
            Case Else
                getLOVColumn = GLOBAL_VERSION_COLUMN
        End Select
        
    End With

End Function

Private Sub gatherDataFromExport(ByRef classCollection As Collection)
    
    Dim wbExport As Workbook
    Set wbExport = Workbooks.Open(exportPath)
    
    Dim wsExport As Worksheet
    Set wsExport = wbExport.Sheets(1)
    
    language = ThisWorkbook.Sheets("Main").Range("I7").Value
    
    gatherData classCollection, wsExport, IMPORT_START_ROW
        
    wbExport.Close 0
    Set wbExport = Nothing
    
End Sub

Private Sub gatherDataFromRautakeskoDataFields(ByRef classCollection As Collection)

    'ATTRIBUTE SCHEMAS SHEET
    Const SCHEMA_KEY_COLUMN As Long = 1
    Const FIELD_NAME_COLUMN As Long = 4
    Const FIELD_ID_COLUMN As Long = 5
    
    'DATA FIELDS SHEET
    Const ADEONA_ID_COLUMN As Long = 1
    Const FIELD_TYPE_COLUMN As Long = 10
    Const MAX_LENGTH_COLUMN As Long = 11
    Const FORMAT_COLUMN As Long = 12
    
    Dim wsAttributeSchema As Worksheet
    Set wsAttributeSchema = ThisWorkbook.Sheets("Attribute schemas")
    
    Dim wsDataFields As Worksheet
    Set wsDataFields = ThisWorkbook.Sheets("Data fields")
    
    Dim nameColumn As Long
    nameColumn = getNameColumnBasedOnLanguage(language)
    'lovColumn = getLOVColumn(language)
    
    Dim schemaKeyRow As Long, attributeRow As Long, attribut As clsAttribute, productClass As clsClass, i As Long
        
    For Each productClass In classCollection
        'Gather all attributes to classes
        With wsAttributeSchema
            On Error Resume Next
                'schemaKeyRow = .Columns(SCHEMA_KEY_COLUMN).Find(productClass.name).Row
                schemaKeyRow = WorksheetFunction.Match(productClass.name, .Columns(SCHEMA_KEY_COLUMN), 0)
            On Error GoTo 0
            
            If schemaKeyRow = 0 Then
                MsgBox "There is no such class in Rautakesko: " & productClass.name
            End If
            
            i = schemaKeyRow + 1 '+1 because attributes starts row below product class
            If schemaKeyRow <> 0 Then
                Do While Len(.Cells(i, FIELD_NAME_COLUMN).Value) > 0
                    Set attribut = New clsAttribute
                    'attribut.name = .Cells(i, FIELD_NAME_COLUMN).Value
                    attribut.id = .Cells(i, FIELD_ID_COLUMN).Value
                    productClass.addAttributeToCollection attribut
                    
                    i = i + 1
                Loop
                        'Gather the rest information for attributes, like its type etc
                With wsDataFields
                    For Each attribut In productClass.attributeCollection
                        On Error Resume Next
                            attributeRow = 0
                            attributeRow = WorksheetFunction.Match(attribut.id, .Columns(ADEONA_ID_COLUMN), 0)
                        On Error GoTo 0
                        
                        If attributeRow > 0 Then
                            attribut.name = .Cells(attributeRow, nameColumn).Value
                            attribut.attributeType = .Cells(attributeRow, FIELD_TYPE_COLUMN).Value
                            attribut.attributeMaxLength = .Cells(attributeRow, MAX_LENGTH_COLUMN).Value
                            attribut.attributeFormat = .Cells(attributeRow, FORMAT_COLUMN).Value
                            
                            If LCase$(attribut.attributeType) = CHOICE_ATTRIBUTE_TYPE Then
                                gatherLOV attribut, productClass.name
                            End If
                        End If
                    Next
                End With
            End If
            schemaKeyRow = 0
        End With
        
    Next
    
End Sub

Public Function getNameColumnBasedOnLanguage(ByVal lang As String) As Long
    'Data Fields sheet
    Const ENGLISH_NAME_COLUMN As Long = 6
    Const FINNISH_NAME_COLUMN As Long = 7
    Const SWEDISH_NAME_COLUMN As Long = 46
    
    Select Case lang
        Case "en"
            getNameColumnBasedOnLanguage = ENGLISH_NAME_COLUMN
        Case "fi"
            getNameColumnBasedOnLanguage = FINNISH_NAME_COLUMN
        Case "se"
            getNameColumnBasedOnLanguage = SWEDISH_NAME_COLUMN
        Case Else
            getNameColumnBasedOnLanguage = ENGLISH_NAME_COLUMN
    End Select

End Function

Private Sub gatherLOV(ByRef attribut As clsAttribute, ByRef productClassName As String)
    
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
            valuesColumn = getLOVColumn(language)
        End If
        
        Dim LOVDictionary As Object
        Set LOVDictionary = CreateObject("Scripting.Dictionary")
        
        Dim i As Long
        i = attributeRow + 1 '+1 because attributes starts row below product class
        Do While Len(.Cells(i, SCHEMA_KEYS_COLUMN).Value) > 0
            If InStr(1, .Cells(i, SCHEMA_KEYS_COLUMN).Value, productClassName, vbTextCompare) > 0 Then
                If Len(.Cells(i, valuesColumn).Value) > 0 Then 'problemy z t³umaczeniem w Rautakesko --> np. puste wartosci w innych jezykach ni¿ EN
                    LOVDictionary.Add .Cells(i, valuesColumn).Value, .Cells(i, valuesColumn).Value
                End If
            End If
            i = i + 1
        Loop
        
    End With
    
    Set attribut.LOV = LOVDictionary

End Sub

Private Function ifInCollection(ByVal col As Collection, ByVal name As String) As Boolean
    
    Dim obj As Object
    
    ifInCollection = True
    
    On Error Resume Next
        Set obj = col(name)
        If Err.Number <> 0 Then
            ifInCollection = False
        End If
    On Error GoTo 0

End Function

Public Function getExportFilePath() As String

    With Application.FileDialog(msoFileDialogOpen)
        On Error GoTo errorHandler
        .AllowMultiSelect = False
        .Filters.Add "Excel Files", "*.xl*"
        .ButtonName = "Select"
        .Show
        getExportFilePath = .SelectedItems.Item(1)
    End With
    
errorHandler:

End Function

Private Function setInAlfabeticOrder(ByRef LOVDictionary As Object) As Object
    
    Dim LOVDictCount As Long
    LOVDictCount = LOVDictionary.Count - 1
    
    Dim MyArray() As Variant
    ReDim MyArray(0 To LOVDictCount)
    
    Dim i As Long
    For i = 0 To LOVDictCount
        MyArray(i) = LOVDictionary.items()(i)
    Next
    
    Dim j As Long, Temp As Variant
    
    For i = LBound(MyArray) To UBound(MyArray) - 1
        For j = i + 1 To UBound(MyArray)
            If MyArray(i) > MyArray(j) Then
                Temp = MyArray(j)
                MyArray(j) = MyArray(i)
                MyArray(i) = Temp
            End If
        Next j
    Next i
    
    Set LOVDictionary = CreateObject("Scripting.Dictionary")
    
    For i = 0 To LOVDictCount
        LOVDictionary.Add MyArray(i), MyArray(i)
    Next
    
    Set setInAlfabeticOrder = LOVDictionary
    
End Function

Public Sub gatherData(ByRef classCollection As Collection, ByVal ws As Worksheet, ByVal startProductRow As Long)
    
    With ws
        Dim lastRow As Long
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        Dim lastCol As Long
        lastCol = .Cells(ID_ROW, .Columns.Count).End(xlToLeft).Column
        
        'in case of EMPTY export - only basic info data, no attributes
        Dim firstAttributesColumn As Long, lastBasicInfoColumn As Long
        On Error Resume Next
            firstAttributesColumn = .Rows(ID_ROW).Find(ATTRIBUTE_PHRASE).Column
            If firstAttributesColumn = 0 Then
                lastBasicInfoColumn = lastCol
            Else
                lastBasicInfoColumn = firstAttributesColumn - 1
            End If
        On Error GoTo 0
        
        Dim i As Long, product As clsProduct, productClass As clsClass, _
            attribut As clsAttribute, j As Long, basicInfoHeader As clsBasicInfoHeader
        
        For i = startProductRow To lastRow
            Set product = New clsProduct
            
            For j = 1 To lastBasicInfoColumn
                Set basicInfoHeader = New clsBasicInfoHeader
                basicInfoHeader.id = .Cells(ID_ROW, j).Value
                basicInfoHeader.val = .Cells(i, j).Value
                
                product.addValueIntoPropertyBasedOnHeaderID basicInfoHeader.id, basicInfoHeader.val
                product.addBasicInfoHeaderToCollection basicInfoHeader
            Next
            
            If firstAttributesColumn > 0 Then
                For j = firstAttributesColumn To lastCol
                    If LenB(.Cells(i, j).Value) <> 0 Then
                        Set attribut = New clsAttribute
                        attribut.id = Replace(.Cells(ID_ROW, j).Value, ATTRIBUTE_PHRASE, vbNullString) 'Split(.Cells(ID_ROW, j).Value, "_")(2)
                        attribut.attributeValue = .Cells(i, j).Value
                        product.addAttributeToCollection attribut
                    End If
                Next
            End If
            
            
            If Len(product.productClass) = 0 Then
                If ifInCollection(product.basicInfoHeadersCollection, KESKO_ATTR_SCHEMA) Then
                    product.productClass = "No class"
                    'if blank class - JUST IN CASE
                    'in case no class column in template
                    'product.productClass = ws.name
                Else
                    product.productClass = ws.name
                End If
            End If
            
            If ifInCollection(classCollection, product.productClass) Then
                Set productClass = classCollection(product.productClass)
                productClass.addProductToCollection product
            Else
                Set productClass = New clsClass
                productClass.name = product.productClass
                productClass.addProductToCollection product
                classCollection.Add productClass, productClass.name
            End If
            
        Next
    End With

End Sub


