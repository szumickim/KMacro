Attribute VB_Name = "uploadRautakeskoDataFile"
Option Explicit

Private Sub uploadRautakesko()

    Const ATTRIBUTE_SCHEMAS As String = "Attribute schemas"
    Const SELECTION_LIST As String = "Selection list specifications"
    Const DATA_FIELDS As String = "Data fields"
    Const VISIBILITY As Long = 0
    Const MAIN_SHEET As String = "Main"
    
    Dim rautaKeskoFilePath As String
    rautaKeskoFilePath = getExportFilePath
    
    If Len(rautaKeskoFilePath) > 0 Then
        Application.ScreenUpdating = False
        
        Dim wbRautakesko As Workbook
        Set wbRautakesko = Workbooks.Open(rautaKeskoFilePath)
        
        Dim wsAttribute As Worksheet
        On Error Resume Next
            Set wsAttribute = wbRautakesko.Sheets(ATTRIBUTE_SCHEMAS)
            If Err.Number <> 0 Then
                wbRautakesko.Close 0
                Application.ScreenUpdating = True
                
                MsgBox "This is not Rautakesko data file!"
                
                Exit Sub
            End If
        On Error GoTo 0
        
        Dim wbThisWb As Workbook
        Set wbThisWb = ThisWorkbook
        
        With wbThisWb
            Application.DisplayAlerts = False
                .Sheets(ATTRIBUTE_SCHEMAS).Delete
                .Sheets(SELECTION_LIST).Delete
                .Sheets(DATA_FIELDS).Delete
            Application.DisplayAlerts = True
        
            wbRautakesko.Sheets(ATTRIBUTE_SCHEMAS).Copy after:=.Sheets(1)
            wbRautakesko.Sheets(SELECTION_LIST).Copy after:=.Sheets(1)
            wbRautakesko.Sheets(DATA_FIELDS).Copy after:=.Sheets(1)
            
            wbRautakesko.Close 0
        
            .Sheets(ATTRIBUTE_SCHEMAS).Visible = VISIBILITY
            .Sheets(SELECTION_LIST).Visible = VISIBILITY
            .Sheets(DATA_FIELDS).Visible = VISIBILITY
        
            .Sheets(MAIN_SHEET).Range("K3").Value = getRautakeskoFileVersion(rautaKeskoFilePath)
        End With
        
        Application.ScreenUpdating = True
        
        MsgBox "Rautakesko file updated succesfully!"
        
    End If
    

End Sub

Private Function getRautakeskoFileVersion(ByRef filePath As String) As String

    Dim lastSlashPosition As Long
    lastSlashPosition = InStrRev(filePath, "\")
    
    Dim fileName As String
    fileName = Mid$(filePath, lastSlashPosition + 1)
    
    getRautakeskoFileVersion = fileName
    
End Function
