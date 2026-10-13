Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal target As Range)
     
    Dim mfgCell As Range
    Dim ProdDesc As Range
    On Error GoTo ErrHandler
    
    DisableAppUpdates
    UnprotectActiveSheet
    
    Set mfgCell = GetFieldByLabel(Me.Name, "C", "Manufacturing date")
    Set ProdDesc = GetFieldByLabel(Me.Name, "C", "Product description")
     
    If Not ProdDesc Is Nothing Then
    If Not Intersect(target, ProdDesc) Is Nothing Then
        LoadProductFields ("Batch Number Generation")
        'Helpers.ClearProductFields
       
    End If
    End If
   
    
    If Not mfgCell Is Nothing Then
        If Not Intersect(target, mfgCell) Is Nothing Then
        
        Call CalculateExpiryDate(Me.Name)
       
        End If
    End If
    
CleanExit:
    ProtectActiveSheet
    RestoreAppState
    Exit Sub

ErrHandler:
    
    MsgBox "Error Batch number generatioin change event: " & Err.Description, vbCritical
    Resume CleanExit
    
End Sub


Private Sub Worksheet_Activate()
    DisableAppUpdates
    UnprotectActiveSheet

    If Date <> LastDate Then
        LastDate = Date
       GetFieldByLabel("Batch Number Generation", "C", "Batch number").value = GenerateBatchNumber
    End If
    
    ProtectActiveSheet
    RestoreAppState
End Sub