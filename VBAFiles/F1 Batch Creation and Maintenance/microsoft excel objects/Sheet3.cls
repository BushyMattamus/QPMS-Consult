Attribute VB_Name = "Sheet3"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal target As Range)
    Dim mfgCell As Range
    Dim batchNum As Range
    Dim shelfLife As Range
    On Error GoTo ErrHandler
    DisableAppUpdates
    UnprotectActiveSheet
  
    Set mfgCell = GetFieldByLabel(Me.Name, "C", "Manufacturing date")
    Set shelfLife = GetFieldByLabel(Me.Name, "C", "Shelf Life(Days/Months/Years)")
    
    If shelfLife <> "as per manufacturer" Then
    
        If Not Intersect(target, mfgCell) Is Nothing Then
            Call CalculateExpiryDate(Me.Name)
        End If
    End If
     
    Set batchNum = GetFieldByLabel(Me.Name, "C", "Batch number")
   
    If Not Intersect(target, batchNum) Is Nothing Then
        'Helpers.ClearProductFields
        Range("D14").ClearContents
        LoadBatchForMaintenance
        
    End If
   
     
CleanExit:
    ProtectActiveSheet
    RestoreAppState
    Exit Sub

ErrHandler:
    
    MsgBox "Error Batch number maintenance change event: " & Err.Description, vbCritical
    Resume CleanExit
    
End Sub