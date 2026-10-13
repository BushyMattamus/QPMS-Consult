Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrHandler
    
    Dim productDescription As String
    Dim pmiNumber As String
    
    ' --- Prevent cascading events ---
    If Target.Cells.Count > 1 Then Exit Sub
    
    AppStateUtils.DisableAppUpdates
    UnprotectActiveSheet
    
    ' --- Handle cell I5 change (Product Description) ---
    If Target.Address = "$I$5" Then
        Range("I6:I10").ClearContents
        Range("D5").value = Target.value
        
        ' --- Trigger D5 event logic ---
        If Len(Trim$(Target.value)) > 0 Then
            productDescription = Trim$(Target.value)
            LookupPMINumberByProductDescription productDescription
            
            ' --- Trigger D6 event logic after PMI Number is populated ---
            pmiNumber = Trim$(Range("D6").value)
            If Len(pmiNumber) > 0 Then
                LookupProductDescriptionByPMINumber pmiNumber
            End If
        Else
            Range("D6").ClearContents
        End If
        
        GoTo CleanExit
    End If
    
    ' --- Handle cell I6 change (Batch Number) ---
    If Target.Address = "$I$6" Then
        If Len(Trim$(Target.value)) = 0 Then
            Range("I6:I9").ClearContents
            GoTo CleanExit
        End If
        
        LoadBatchDetailsFromMaster
        
        Me.Calculate
        
        Range("D5").value = Range("I5").value
        Range("D6").value = Range("I7").value
        
        ' --- Trigger D6 event logic after PMI Number is populated ---
        pmiNumber = Trim$(Range("D6").value)
        If Len(pmiNumber) > 0 Then
            LookupProductDescriptionByPMINumber pmiNumber
        End If
        
        GoTo CleanExit
    End If
    
    ' --- Handle cell D5 change (Product Description manual entry) ---
    If Target.Address = "$D$5" Then
        If Len(Trim$(Target.value)) > 0 Then
            productDescription = Trim$(Target.value)
            LookupPMINumberByProductDescription productDescription
            
            ' --- Trigger D6 event logic after PMI Number is populated ---
            pmiNumber = Trim$(Range("D6").value)
            If Len(pmiNumber) > 0 Then
                LookupProductDescriptionByPMINumber pmiNumber
            End If
        Else
            Range("D6").ClearContents
        End If
        
        GoTo CleanExit
    End If
    
    ' --- Handle cell D6 change (PMI Number manual entry) ---
    If Target.Address = "$D$6" Then
        If Len(Trim$(Target.value)) > 0 Then
            pmiNumber = Trim$(Target.value)
            LookupProductDescriptionByPMINumber pmiNumber
        Else
            Range("D5").ClearContents
        End If
        
        GoTo CleanExit
    End If
    
    GoTo CleanExit

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Worksheet_Change"
    GoTo CleanExit

CleanExit:
    FormUtils.ProtectActiveSheet
    AppStateUtils.RestoreAppState
End Sub