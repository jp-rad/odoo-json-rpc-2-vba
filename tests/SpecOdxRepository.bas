Attribute VB_Name = "SpecOdxRepository"
' External API - odoo-json-rpc-2-vba
'
' MIT License
'
' Copyright (c) 2022-2027 jp-rad
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'

Option Explicit

' ------------------------------------------------------------------------------
' Module: SpecOdxRepository
' Purpose: Contains unit test specifications for OdxRepository external ID and CRUD operations.
' Standard:
'   - On success: Returns a Dictionary containing {"result": <payload>}
'   - On failure: Raises a runtime error.
' ------------------------------------------------------------------------------

''
' Register and run all OdxRepository test specifications across different repository types.
'
' @param aSuite TestSuite instance
' @param aType2 Boolean flag indicating whether to use Type2 repository
''
Public Sub RegisterSpecs(ByRef aSuite As TestSuite, ByVal aType2 As Boolean)
    SpecOdxRepositoryExternalIdMapping aSuite, aType2
    SpecRegisterExternalIdLifecycle aSuite, aType2
    SpecUpsertRecordByExternalId aSuite, aType2
    SpecUnlinkRecordByExternalId aSuite, aType2
    SpecGenerateExternalId aSuite, aType2
End Sub

' ------------------------------------------------------------------------------
' Test Specifications
' ------------------------------------------------------------------------------

''
' Tests for OdxRepository External ID mapping and resolution functions.
' Verifies dynamic lookup and name matching since res_ids vary across environments.
'
' @param aSuite TestSuite instance
' @param aType2 Boolean flag indicating whether to use Type2 repository
''
Private Sub SpecOdxRepositoryExternalIdMapping(ByRef aSuite As TestSuite, ByVal aType2 As Boolean)
    Dim oRepo As OdxRepository
    
    Set oRepo = CreateOdxRepository(aType2)
    
    With aSuite.Test("TestOdxRepository - External ID and Record Resolution")
        
        Dim oResolvedModelId As Dictionary
        Dim oPartnerRecords As Collection
        Dim oFieldsToRead As New Collection
        
        oFieldsToRead.Add "name"
        
        ' ----------------------------------------------------------------------
        ' 1. Test: Resolve from External ID ("base.res_partner_2")
        ' Verify model and read record to confirm the name is "Acme Corporation"
        ' ----------------------------------------------------------------------
        Set oResolvedModelId = oRepo.GetModelAndIdFromExternalId("base.res_partner_2")
        .NotUndefined oResolvedModelId
        If Not oResolvedModelId Is Nothing Then
            .IsEqual oResolvedModelId("model"), "res.partner"
        
            ' Read record using the resolved res_id to check name dynamically
            Set oPartnerRecords = oRepo.ReadRecord("res.partner", oResolvedModelId("res_id"), oFieldsToRead)
            .IsEqual oPartnerRecords.Count, 1
            .IsEqual CStr(oPartnerRecords(1)("name")), "Acme Corporation"
        
        End If
        
        ' ----------------------------------------------------------------------
        ' 2. Test: Resolve from External ID ("base.res_partner_address_31")
        ' Verify record name is "Addison Olson"
        ' ----------------------------------------------------------------------
        Set oResolvedModelId = oRepo.GetModelAndIdFromExternalId("base.res_partner_address_31")
        .NotUndefined oResolvedModelId
        If Not oResolvedModelId Is Nothing Then
            .IsEqual oResolvedModelId("model"), "res.partner"
            
            Set oPartnerRecords = oRepo.ReadRecord("res.partner", oResolvedModelId("res_id"), oFieldsToRead)
            .IsEqual oPartnerRecords.Count, 1
            .IsEqual CStr(oPartnerRecords(1)("name")), "Addison Olson"
        
        End If
        
        ' ----------------------------------------------------------------------
        ' 3. Test: Search by domain to find a record, then test GetExternalIdFromModelAndId
        ' Finds "Douglas Fletcher" dynamically, gets its External ID, and verifies it
        ' ----------------------------------------------------------------------
        Dim oDomain As New Collection
        Dim oCond As New Collection
        oCond.Add "name": oCond.Add "=": oCond.Add "Douglas Fletcher"
        oDomain.Add oCond
        
        Dim oRecordIds As Collection
        Set oRecordIds = oRepo.SearchRecord("res.partner", oDomain, , 1)
        
        If oRecordIds.Count > 0 Then
            Dim nDynamicResId As Long
            nDynamicResId = oRecordIds(1)
            
            Dim sExternalId As String
            sExternalId = oRepo.GetExternalIdFromModelAndId("res.partner", nDynamicResId)
            .IsEqual sExternalId, "base.res_partner_address_3"
        End If
        
    End With
End Sub

''
' Tests for RegisterExternalId and related lifecycle (create, register, verify, and cleanup).
'
' @param aSuite TestSuite instance
' @param aType2 Boolean flag indicating whether to use Type2 repository
''
Private Sub SpecRegisterExternalIdLifecycle(ByRef aSuite As TestSuite, ByVal aType2 As Boolean)
    Dim oRepo As OdxRepository
    
    Set oRepo = CreateOdxRepository(aType2)
    
    With aSuite.Test("TestRegisterExternalId - Lifecycle and Cleanup")
        
        Dim oTestPartnerVals As New Dictionary
        Dim oCreatedIds As Collection
        Dim sTestModel As String
        Dim sTestExternalId As String
        Dim nPartnerId As Long
        Dim oResDict As Dictionary
        
        sTestModel = "res.partner"
        sTestExternalId = "SpecOdxRepository.test_partner_lifecycle_1"
        
        ' ----------------------------------------------------------------------
        ' 1. Create a test record in res.partner
        ' ----------------------------------------------------------------------
        oTestPartnerVals.Add "name", "Test Partner (Unit Test)"
        oTestPartnerVals.Add "is_company", False
        
        Set oCreatedIds = oRepo.CreateRecord(sTestModel, oTestPartnerVals)
        .NotUndefined oCreatedIds
        .IsEqual oCreatedIds.Count, 1
        
        nPartnerId = oCreatedIds(1)
        
        ' ----------------------------------------------------------------------
        ' 2. Register External ID using RegisterExternalId
        ' ----------------------------------------------------------------------
        Dim bIsRegistered As Boolean
        bIsRegistered = oRepo.RegisterExternalId(sTestModel, nPartnerId, sTestExternalId)
        .IsEqual bIsRegistered, True
        
        ' ----------------------------------------------------------------------
        ' 3. Verify resolution from External ID to Model and ID
        ' ----------------------------------------------------------------------
        Set oResDict = oRepo.GetModelAndIdFromExternalId(sTestExternalId)
        .NotUndefined oResDict
        .IsEqual oResDict("model"), sTestModel
        .IsEqual oResDict("res_id"), nPartnerId
        
        ' ----------------------------------------------------------------------
        ' 4. Cleanup: Delete the test partner record
        ' ----------------------------------------------------------------------
        Dim bIsDeleted As Boolean
        bIsDeleted = oRepo.UnlinkRecord(sTestModel, nPartnerId)
        .IsEqual bIsDeleted, True
        
        ' ----------------------------------------------------------------------
        ' 5. Verify that the External ID is also removed or no longer resolves
        ' ----------------------------------------------------------------------
        Dim oResolvedAfterDelete As Dictionary
        Set oResolvedAfterDelete = oRepo.GetModelAndIdFromExternalId(sTestExternalId)
        
        If Not oResolvedAfterDelete Is Nothing Then
            .IsEqual oResolvedAfterDelete.Count, 0
        Else
            .IsUndefined oResolvedAfterDelete
        End If
        
    End With
End Sub

''
' Tests for UpsertRecordByExternalId (Create and Update operations based on External ID).
'
' @param aSuite TestSuite instance
' @param aType2 Boolean flag indicating whether to use Type2 repository
''
Private Sub SpecUpsertRecordByExternalId(ByRef aSuite As TestSuite, ByVal aType2 As Boolean)
    Dim oRepo As OdxRepository
    
    Set oRepo = CreateOdxRepository(aType2)
    
    With aSuite.Test("TestUpsertRecordByExternalId - Upsert and Status Verification")
        
        Dim sTestModel As String
        Dim sTestExternalId As String
        Dim oMainDict As New Dictionary
        Dim oPartnerVals As New Dictionary
        Dim oResultRes As Dictionary
        Dim oFieldsToRead As New Collection
        Dim oPartnerRecords As Collection
        
        sTestModel = "res.partner"
        sTestExternalId = "SpecOdxRepository.test_upsert_partner_1"
        oFieldsToRead.Add "name"
        
        ' ----------------------------------------------------------------------
        ' 0. Pre-cleanup (just in case previous test left orphaned record)
        ' ----------------------------------------------------------------------
        Dim oExistingMap As Dictionary
        Set oExistingMap = oRepo.GetModelAndIdFromExternalId(sTestExternalId)
        If Not oExistingMap Is Nothing Then
            oRepo.UnlinkRecord sTestModel, oExistingMap("res_id")
        End If
        
        ' ----------------------------------------------------------------------
        ' 1. Test: Create action via Upsert
        ' ----------------------------------------------------------------------
        oPartnerVals.Add "name", "Upsert Initial Name"
        oMainDict.Add sTestExternalId, oPartnerVals
        
        Set oResultRes = oRepo.UpsertRecordByExternalId(sTestModel, oMainDict)
        
        .NotUndefined oResultRes
        .IsEqual oResultRes.Exists(sTestExternalId), True
        .IsEqual CStr(oResultRes(sTestExternalId)("action")), "created"
        
        Dim nCreatedId As Long
        nCreatedId = CLng(oResultRes(sTestExternalId)("id"))
        .NotEqual nCreatedId, 0
        
        ' Verify the created record name
        Set oPartnerRecords = oRepo.ReadRecord(sTestModel, nCreatedId, oFieldsToRead)
        .IsEqual CStr(oPartnerRecords(1)("name")), "Upsert Initial Name"
        
        ' ----------------------------------------------------------------------
        ' 2. Test: Update action via Upsert (using the same external ID)
        ' ----------------------------------------------------------------------
        Dim oUpdateDict As New Dictionary
        Dim oUpdateVals As New Dictionary
        oUpdateVals.Add "name", "Upsert Updated Name"
        oUpdateDict.Add sTestExternalId, oUpdateVals
        
        Set oResultRes = oRepo.UpsertRecordByExternalId(sTestModel, oUpdateDict)
        
        .NotUndefined oResultRes
        .IsEqual oResultRes.Exists(sTestExternalId), True
        .IsEqual CStr(oResultRes(sTestExternalId)("action")), "updated"
        .IsEqual CLng(oResultRes(sTestExternalId)("id")), nCreatedId
        
        ' Verify the updated record name
        Set oPartnerRecords = oRepo.ReadRecord(sTestModel, nCreatedId, oFieldsToRead)
        .IsEqual CStr(oPartnerRecords(1)("name")), "Upsert Updated Name"
        
        ' ----------------------------------------------------------------------
        ' 3. Cleanup: Delete the test partner record
        ' ----------------------------------------------------------------------
        Dim bIsDeleted As Boolean
        bIsDeleted = oRepo.UnlinkRecord(sTestModel, nCreatedId)
        .IsEqual bIsDeleted, True
        
    End With
End Sub

''
' Tests for UnlinkRecordByExternalId (Delete operations based on External ID(s) supporting String and Collection).
'
' @param aSuite TestSuite instance
' @param aType2 Boolean flag indicating whether to use Type2 repository
''
Private Sub SpecUnlinkRecordByExternalId(ByRef aSuite As TestSuite, ByVal aType2 As Boolean)
    Dim oRepo As OdxRepository
    
    Set oRepo = CreateOdxRepository(aType2)
    
    With aSuite.Test("TestUnlinkRecordByExternalId - Unlink and Status Verification")
        
        Dim sTestModel As String
        Dim sExtId1 As String, sExtId2 As String, sExtId3 As String
        Dim oTestModelVals As Dictionary
        Dim oCreatedIds As Collection
        Dim oResultRes As Dictionary
        Dim oExistingMap As Dictionary
        
        sTestModel = "res.partner"
        sExtId1 = "SpecOdxRepository.test_unlink_partner_1"
        sExtId2 = "SpecOdxRepository.test_unlink_partner_2"
        sExtId3 = "SpecOdxRepository.test_unlink_partner_3"
        
        ' ----------------------------------------------------------------------
        ' 0. Pre-cleanup (just in case previous test left orphaned records)
        ' ----------------------------------------------------------------------
        Set oExistingMap = oRepo.GetModelAndIdFromExternalId(sExtId1)
        If Not oExistingMap Is Nothing Then oRepo.UnlinkRecord sTestModel, oExistingMap("res_id")
        
        Set oExistingMap = oRepo.GetModelAndIdFromExternalId(sExtId2)
        If Not oExistingMap Is Nothing Then oRepo.UnlinkRecord sTestModel, oExistingMap("res_id")
        
        Set oExistingMap = oRepo.GetModelAndIdFromExternalId(sExtId3)
        If Not oExistingMap Is Nothing Then oRepo.UnlinkRecord sTestModel, oExistingMap("res_id")
        
        ' ----------------------------------------------------------------------
        ' 1. Setup: Create 3 test records and register their External IDs
        ' ----------------------------------------------------------------------
        Set oTestModelVals = New Dictionary
        oTestModelVals.Add "name", "Unlink Test Partner A"
        Set oCreatedIds = oRepo.CreateRecord(sTestModel, oTestModelVals)
        oRepo.RegisterExternalId sTestModel, oCreatedIds(1), sExtId1
        
        oTestModelVals("name") = "Unlink Test Partner B"
        Set oCreatedIds = oRepo.CreateRecord(sTestModel, oTestModelVals)
        oRepo.RegisterExternalId sTestModel, oCreatedIds(1), sExtId2
        
        oTestModelVals("name") = "Unlink Test Partner C"
        Set oCreatedIds = oRepo.CreateRecord(sTestModel, oTestModelVals)
        oRepo.RegisterExternalId sTestModel, oCreatedIds(1), sExtId3
        
        ' ----------------------------------------------------------------------
        ' 2. Test: Unlink the 1st record individually using a String External ID
        ' ----------------------------------------------------------------------
        Set oResultRes = oRepo.UnlinkRecordByExternalId(sTestModel, sExtId1)
        
        .NotUndefined oResultRes
        .IsEqual oResultRes.Exists(sExtId1), True
        .IsEqual CStr(oResultRes(sExtId1)("action")), "unlinked"
        
        ' ----------------------------------------------------------------------
        ' 3. Test: Unlink all 3 records using a Collection (sExtId1 is already gone -> error, sExtId2 & sExtId3 -> success)
        ' ----------------------------------------------------------------------
        Dim oExtIdCol As New Collection
        oExtIdCol.Add sExtId1  ' Already deleted -> should result in an error
        oExtIdCol.Add sExtId2  ' Still exists -> success (unlinked)
        oExtIdCol.Add sExtId3  ' Still exists -> success (unlinked)
        
        Set oResultRes = oRepo.UnlinkRecordByExternalId(sTestModel, oExtIdCol)
        
        .NotUndefined oResultRes
        
        ' Verification 1: The 1st record should return an error since it was already unlinked
        .IsEqual oResultRes.Exists(sExtId1), True
        .IsEqual CStr(oResultRes(sExtId1)("action")), "error"
        .NotUndefined CStr(oResultRes(sExtId1)("error_message"))
        
        ' Verification 2: The 2nd record should be successfully unlinked
        .IsEqual oResultRes.Exists(sExtId2), True
        .IsEqual CStr(oResultRes(sExtId2)("action")), "unlinked"
        
        ' Verification 3: The 3rd record should be successfully unlinked
        .IsEqual oResultRes.Exists(sExtId3), True
        .IsEqual CStr(oResultRes(sExtId3)("action")), "unlinked"
        
    End With
End Sub

''
' Tests generating a unique external ID.
' Verifies that the generated External ID starts with the model name representation "__rpc_vba__.res_partner_".
'
' @param aSuite TestSuite instance
' @param aType2 Boolean flag indicating whether to use Type2 repository
''
Private Sub SpecGenerateExternalId(ByRef aSuite As TestSuite, ByVal aType2 As Boolean)
    Dim oRepo As OdxRepository
    
    Set oRepo = CreateOdxRepository(aType2)
    
    With aSuite.Test("Generate Unique External ID (GenerateExternalId)")
        
        On Error Resume Next
        Dim sExternalId As String
        sExternalId = oRepo.GenerateExternalId("res.partner")
        Dim nErrNum As Long
        nErrNum = Err.Number
        On Error GoTo 0
        
        .IsEqual nErrNum, 0, "GenerateExternalId call should not raise an error"
        If nErrNum = 0 Then
            .IsOk (Len(sExternalId) > 0), "Generated External ID should not be empty"
            ' Verify that the generated External ID starts with the model name representation "__rpc_vba__.res_partner_"
            .IsOk (InStr(sExternalId, "__rpc_vba__.res_partner_") = 1), "Generated External ID should start with '__rpc_vba__.res_partner_': " & sExternalId
        End If
    End With
End Sub
