Attribute VB_Name = "SpecOdxMemRs"
Option Explicit

' ==============================================================================
' Module: SpecOdxMemRs
' Purpose: Contains comprehensive unit test specifications for OdxMemRs class,
'          including initialization, field classification, navigation,
'          binary field change detection, and simultaneous batch operations.
' ==============================================================================

''
' Register and run all OdxMemRs test specifications.
'
' @param aSuite TestSuite instance
Public Sub RegisterSpecs(aSuite As TestSuite)
    SpecOdxMemRsInitialization aSuite
    SpecOdxMemRsFieldManagementAndClassification aSuite
    SpecOdxMemRsNavigationAndModification aSuite
    SpecOdxMemRsStateTrackingAndChangeDetection aSuite
    SpecOdxMemRsBinaryFieldChange aSuite
    SpecOdxMemRsSimultaneousBatchOperations aSuite
    SpecOdxMemRsSetOperationFilter aSuite
    SpecOdxMemRsCloningAndBatchCommit aSuite
End Sub

''
' 1. Tests for initialization, lifecycle methods, and basic properties.
'
' @param aSuite TestSuite instance
Private Sub SpecOdxMemRsInitialization(aSuite As TestSuite)
    With aSuite.Test("TestOdxMemRs - Initialization and Lifecycle")
        Dim oRs As OdxMemRs
        Set oRs = New OdxMemRs
        
        .NotUndefined oRs
        .IsUndefined oRs.UnderlyingRecordset
        
        oRs.InitRs
        .NotUndefined oRs.UnderlyingRecordset
        .IsEqual oRs.State, MemStateClosed
        
        oRs.AddIdField
        
        oRs.OpenRs
        .IsEqual oRs.State, MemStateOpen
        .IsEqual oRs.RecordCount, 0
        .IsEqual (oRs.BOF And oRs.EOF), True
        
        oRs.CloseRs
        .IsEqual oRs.State, MemStateClosed
    End With
End Sub

''
' 2. Tests for schema setup, field additions, and field classification helpers.
'
' @param aSuite TestSuite instance
Private Sub SpecOdxMemRsFieldManagementAndClassification(aSuite As TestSuite)
    With aSuite.Test("TestOdxMemRs - Field Management and Classification")
        Dim oRs As OdxMemRs
        Set oRs = New OdxMemRs
        
        oRs.InitRs
        oRs.AddIdField
        oRs.AddExternalIdField "ext_id"
        oRs.AddStatusFields
        oRs.AddDataBoundField "name", MemVarWChar, 100
        oRs.AddLocalField "local_flag", MemBoolean
        oRs.OpenRs
        
        .IsEqual oRs.ExistsField("id"), True
        .IsEqual oRs.ExistsField("ext_id"), True
        .IsEqual oRs.ExistsField("name"), True
        .IsEqual oRs.ExistsField("local_flag"), True
        .IsEqual oRs.ExistsField("non_existent"), False
        
        .IsEqual oRs.IsIdField("id"), True
        .IsEqual oRs.IsExternalIdField("ext_id"), True
        .IsEqual oRs.IsControlField("_ope_"), True
        .IsEqual oRs.IsControlField("_del_"), True
        .IsEqual oRs.IsControlField("_syn_"), True
        .IsEqual oRs.IsDataBoundField("name"), True
        .IsEqual oRs.IsLocalField("local_flag"), True
        .IsEqual oRs.IsRequiredField("ext_id"), True
        
        Dim colNames As Collection
        Set colNames = oRs.FieldNames
        .NotUndefined colNames
        .IsEqual colNames.Count >= 6, True
        
        oRs.CloseRs
    End With
End Sub

''
' 3. Tests for record navigation, modification, and find utilities.
'
' @param aSuite TestSuite instance
Private Sub SpecOdxMemRsNavigationAndModification(aSuite As TestSuite)
    With aSuite.Test("TestOdxMemRs - Navigation and Modification")
        Dim oRs As OdxMemRs
        Set oRs = New OdxMemRs
        
        oRs.InitRs
        oRs.AddIdField
        oRs.AddStatusFields
        oRs.AddDataBoundField "score", MemInteger
        oRs.OpenRs
        
        oRs.AddNew
        oRs.FieldValue("id") = 1
        oRs.FieldValue("score") = 85
        oRs.Update
        
        oRs.AddNew
        oRs.FieldValue("id") = 2
        oRs.FieldValue("score") = 95
        oRs.Update
        
        .IsEqual oRs.RecordCount, 2
        
        oRs.MoveFirst
        .IsEqual oRs.FieldValue("id"), 1
        
        oRs.MoveNext
        .IsEqual oRs.FieldValue("id"), 2
        
        oRs.MoveLast
        .IsEqual oRs.FieldValue("id"), 2
        
        oRs.MoveFirst
        oRs.Find "id = 1"
        .IsEqual oRs.FieldValue("score"), 85
        
        oRs.FindFirstOrAddNew 2
        .IsEqual oRs.FieldValue("score"), 95
        
        oRs.FindFirstOrAddNew 3
        .IsEqual oRs.FieldValue("id"), 3
        .IsUndefined oRs.FieldValue("score")
        oRs.Update
        .IsEqual oRs.RecordCount, 3
        
        oRs.CloseRs
    End With
End Sub

''
' 4. Tests for state tracking, sync flags, and deletion flags.
'
' @param aSuite TestSuite instance
Private Sub SpecOdxMemRsStateTrackingAndChangeDetection(aSuite As TestSuite)
    With aSuite.Test("TestOdxMemRs - State Tracking and Change Detection")
        Dim oRs As OdxMemRs
        Set oRs = New OdxMemRs
        
        oRs.InitRs
        oRs.AddIdField
        oRs.AddStatusFields
        oRs.AddDataBoundField "title", MemVarWChar, 50
        oRs.OpenRs
        
        oRs.AddNew
        oRs.FieldValue("id") = 10
        oRs.FieldValue("title") = "Original Title"
        oRs.Update
        
        oRs.ResolveAllStatuses
        .IsEqual oRs.IsSynced(), False
        .IsEqual oRs.IsMarkedForDelete(), False
        .IsEqual oRs.GetOperationFieldValue(), MemOpeAdded
        
        oRs.MarkAsSynced
        .IsEqual oRs.IsSynced(), True
        oRs.MarkAsUnsynced
        .IsEqual oRs.IsSynced(), False
        
        oRs.MarkRecordForDelete
        .IsEqual oRs.IsMarkedForDelete(), True
        oRs.UnmarkRecordForDelete
        .IsEqual oRs.IsMarkedForDelete(), False
        
        oRs.FieldValue("title") = "Modified Title"
        .IsEqual oRs.IsFieldValueChanged("title"), True
        
        oRs.CloseRs
    End With
End Sub

''
' 5. Tests field value change detection specifically for binary fields (byte arrays).
'
' @param aSuite TestSuite instance
Private Sub SpecOdxMemRsBinaryFieldChange(aSuite As TestSuite)
    With aSuite.Test("TestOdxMemRs - Binary Field Change Detection")
        Dim oRs As OdxMemRs
        Set oRs = New OdxMemRs
        
        oRs.InitRs
        oRs.AddIdField
        oRs.AddStatusFields
        oRs.AddDataBoundField "bin_data", MemVarBinary, 255
        oRs.OpenRs
        
        Dim abOriginal() As Byte
        abOriginal = "OdooVbaBinaryTest"
        
        oRs.AddNew
        oRs.FieldValue("id") = 1
        oRs.FieldValue("bin_data") = abOriginal
        oRs.Update
        
        oRs.ResolveAllStatuses
        oRs.CommitBatchChanges
        
        .IsEqual oRs.IsFieldValueChanged("bin_data"), False
        .IsEqual oRs.GetOperationFieldValue(), MemOpeUnchanged
        
        Dim abSame() As Byte
        abSame = "OdooVbaBinaryTest"
        
        oRs.FieldValue("bin_data") = abSame
        .IsEqual oRs.IsFieldValueChanged("bin_data"), False
        
        Dim abDifferent() As Byte
        abDifferent = "ModifiedBinaryData"
        
        oRs.FieldValue("bin_data") = abDifferent
        .IsEqual oRs.IsFieldValueChanged("bin_data"), True
        
        oRs.Update
        oRs.ResolveAllStatuses
        .IsEqual oRs.GetOperationFieldValue(), MemOpeModified
        
        oRs.CloseRs
    End With
End Sub

''
' 6. Tests simultaneous operations (Add, Update, Delete) after invoking CommitBatchChanges
' with multiple rows.
'
' @param aSuite TestSuite instance
Private Sub SpecOdxMemRsSimultaneousBatchOperations(aSuite As TestSuite)
    With aSuite.Test("TestOdxMemRs - Simultaneous Add, Update, and Delete After CommitBatchChanges")
        Dim oRs As OdxMemRs
        Set oRs = New OdxMemRs
        
        oRs.InitRs
        oRs.AddIdField
        oRs.AddStatusFields
        oRs.AddDataBoundField "name", MemVarWChar, 50
        oRs.OpenRs
        
        ' Prepare multiple rows (Row 1 and Row 2)
        oRs.AddNew
        oRs.FieldValue("id") = 1
        oRs.FieldValue("name") = "Record One"
        oRs.Update
        
        oRs.AddNew
        oRs.FieldValue("id") = 2
        oRs.FieldValue("name") = "Record Two"
        oRs.Update
        
        oRs.ResolveAllStatuses
        oRs.CommitBatchChanges
        
        ' Verify initial committed state
        oRs.MoveFirst
        Do Until oRs.EOF
            .IsEqual oRs.FieldValue("_ope_"), MemOpeUnchanged
            oRs.MoveNext
        Loop
        
        ' Perform simultaneous Add, Update, and Delete
        oRs.MoveFirst
        oRs.Find "id = 1"
        oRs.MarkRecordForDelete
        
        oRs.MoveFirst
        oRs.Find "id = 2"
        oRs.FieldValue("name") = "Record Two Updated"
        oRs.Update
        
        oRs.AddNew
        oRs.FieldValue("id") = 3
        oRs.FieldValue("name") = "Record Three Added"
        oRs.Update
        
        ' Resolve statuses and verify each operation
        Dim bResolved As Boolean
        bResolved = oRs.ResolveAllStatuses()
        .IsEqual bResolved, True
        
        ' Verify Row 1: Deleted
        oRs.MoveFirst
        oRs.Find "id = 1"
        .IsEqual oRs.IsMarkedForDelete(), True
        .IsEqual oRs.GetOperationFieldValue(), MemOpeDeleted
        
        ' Verify Row 2: Modified
        oRs.MoveFirst
        oRs.Find "id = 2"
        .IsEqual oRs.IsMarkedForDelete(), False
        .IsEqual oRs.GetOperationFieldValue(), MemOpeModified
        
        ' Verify Row 3: Added
        oRs.MoveFirst
        oRs.Find "id = 3"
        .IsEqual oRs.IsMarkedForDelete(), False
        .IsEqual oRs.GetOperationFieldValue(), MemOpeAdded
        
        oRs.CloseRs
    End With
End Sub

''
' 7. Tests SetOperationFilter method for filtering records by operation status.
'
' @param aSuite TestSuite instance
Private Sub SpecOdxMemRsSetOperationFilter(aSuite As TestSuite)
    With aSuite.Test("TestOdxMemRs - SetOperationFilter")
        Dim oRs As OdxMemRs
        Set oRs = New OdxMemRs
        
        oRs.InitRs
        oRs.AddIdField
        oRs.AddStatusFields
        oRs.AddDataBoundField "name", MemVarWChar, 50
        oRs.OpenRs
        
        ' Row 1: Will be committed then Unchanged
        oRs.AddNew
        oRs.FieldValue("id") = 1
        oRs.FieldValue("name") = "Unchanged Record"
        oRs.Update
        
        ' Row 2: Will be committed then modified
        oRs.AddNew
        oRs.FieldValue("id") = 2
        oRs.FieldValue("name") = "Modified Record"
        oRs.Update
        
        oRs.ResolveAllStatuses
        oRs.CommitBatchChanges
        
        ' Modify Row 2
        oRs.MoveFirst
        oRs.Find "id = 2"
        oRs.FieldValue("name") = "Modified Record Updated"
        oRs.Update
        
        ' Add Row 3: Another added record
        oRs.AddNew
        oRs.FieldValue("id") = 3
        oRs.FieldValue("name") = "Another Added Record"
        oRs.Update
        
        oRs.ResolveAllStatuses
        
        ' Total records before filter should be 3
        .IsEqual oRs.RecordCount, 3
        
        ' Test filtering by MemOpeAdded (should find record 3)
        oRs.SetOperationFilter MemOpeAdded
        .IsEqual oRs.RecordCount, 1
        
        ' Test filtering by MemOpeModified (should find record 2)
        oRs.SetOperationFilter MemOpeModified
        .IsEqual oRs.RecordCount, 1
        .IsEqual oRs.FieldValue("id"), 2
        .IsEqual oRs.FieldValue("name"), "Modified Record Updated"
        
        ' Clear filter and verify all records are accessible again
        oRs.ClearFilter
        .IsEqual oRs.RecordCount, 3
        
        oRs.CloseRs
    End With
End Sub

''
' 8. Tests for cloning, filtered clones, and batch commit operations.
'
' @param aSuite TestSuite instance
Private Sub SpecOdxMemRsCloningAndBatchCommit(aSuite As TestSuite)
    With aSuite.Test("TestOdxMemRs - Cloning and Batch Commit")
        Dim oRs As OdxMemRs
        Dim oClone As OdxMemRs
        Set oRs = New OdxMemRs
        
        oRs.InitRs
        oRs.AddIdField
        oRs.AddStatusFields
        oRs.AddDataBoundField "val", MemInteger
        oRs.OpenRs
        
        oRs.AddNew
        oRs.FieldValue("id") = 1
        oRs.FieldValue("val") = 500
        oRs.Update
        
        oRs.ResolveAllStatuses
        
        Set oClone = oRs.Clone(False)
        .NotUndefined oClone
        .IsEqual oClone.RecordCount, 1
        
        oClone.SetOperationFilter MemOpeAdded
        .IsEqual oClone.RecordCount, 1
        
        oClone.SetOperationFilter MemOpeModified
        .IsEqual oClone.RecordCount, 0
        
        oClone.SetOperationFilter MemOpeDeleted
        .IsEqual oClone.RecordCount, 0
        
        oClone.CloseRs
        
        oRs.CommitBatchChanges
        .IsEqual oRs.FieldValue("_ope_"), MemOpeUnchanged
        
        oRs.CloseRs
    End With
End Sub
