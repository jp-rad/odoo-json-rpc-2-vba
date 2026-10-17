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
' Purpose: Contains unit test specifications for OdxRepository
' Standard:
'   - On success: Returns a Dictionary containing {"result": <payload>}
'   - On failure: Raises a runtime error.
' ------------------------------------------------------------------------------

''
' Register and run all OdxRepository test specifications across different repository types.
'
' @param Suite TestSuite instance
''
Public Sub RegisterSpecs(Suite As TestSuite)
    SpecOdxRepositoryExternalIdMapping Suite, False
    SpecOdxRepositoryExternalIdMapping Suite, True
    SpecRegisterExternalIdLifecycle Suite, False
    SpecRegisterExternalIdLifecycle Suite, True
    SpecUpsertRecordByExternalId Suite, False
    SpecUpsertRecordByExternalId Suite, True
    SpecUnlinkRecordByExternalId Suite, False
    SpecUnlinkRecordByExternalId Suite, True
End Sub

' ------------------------------------------------------------------------------
' Test Specifications
' ------------------------------------------------------------------------------

''
' Tests for OdxRepository External ID mapping and resolution functions.
' Verifies dynamic lookup and name matching since res_ids vary across environments.
'
' @param Suite TestSuite instance
' @param aType2 Boolean flag indicating whether to use Type2 repository
''
Private Sub SpecOdxRepositoryExternalIdMapping(Suite As TestSuite, aType2 As Boolean)
    Dim repo As OdxRepository
    
    Set repo = CreateOdxRepository(aType2)
    
    With Suite.Test("TestOdxRepository - External ID and Record Resolution")
        
        Dim resolvedModelId As Dictionary
        Dim partnerRecords As Collection
        Dim fieldsToRead As New Collection
        
        fieldsToRead.Add "name"
        
        ' ----------------------------------------------------------------------
        ' 1. Test: Resolve from External ID ("base.res_partner_2")
        ' Verify model and read record to confirm the name is "Acme Corporation"
        ' ----------------------------------------------------------------------
        Set resolvedModelId = repo.GetModelAndIdFromExternalId("base.res_partner_2")
        .NotUndefined resolvedModelId
        .IsEqual resolvedModelId("model"), "res.partner"
        
        ' Read record using the resolved res_id to check name dynamically
        Set partnerRecords = repo.ReadRecord("res.partner", resolvedModelId("res_id"), fieldsToRead)
        .IsEqual partnerRecords.Count, 1
        .IsEqual CStr(partnerRecords(1)("name")), "Acme Corporation"
        
        ' ----------------------------------------------------------------------
        ' 2. Test: Resolve from External ID ("base.res_partner_address_31")
        ' Verify record name is "Addison Olson"
        ' ----------------------------------------------------------------------
        Set resolvedModelId = repo.GetModelAndIdFromExternalId("base.res_partner_address_31")
        .NotUndefined resolvedModelId
        .IsEqual resolvedModelId("model"), "res.partner"
        
        Set partnerRecords = repo.ReadRecord("res.partner", resolvedModelId("res_id"), fieldsToRead)
        .IsEqual partnerRecords.Count, 1
        .IsEqual CStr(partnerRecords(1)("name")), "Addison Olson"
        
        ' ----------------------------------------------------------------------
        ' 3. Test: Search by domain to find a record, then test GetExternalIdFromModelAndId
        ' Finds "Douglas Fletcher" dynamically, gets its External ID, and verifies it
        ' ----------------------------------------------------------------------
        Dim domain As New Collection
        Dim cond As New Collection
        cond.Add "name": cond.Add "=": cond.Add "Douglas Fletcher"
        domain.Add cond
        
        Dim recordIds As Collection
        Set recordIds = repo.SearchRecord("res.partner", domain, , 1)
        
        If recordIds.Count > 0 Then
            Dim dynamicResId As Long
            dynamicResId = recordIds(1)
            
            Dim externalId As String
            externalId = repo.GetExternalIdFromModelAndId("res.partner", dynamicResId)
            .IsEqual externalId, "base.res_partner_address_3"
        End If
        
    End With
End Sub

''
' Tests for RegisterExternalId and related lifecycle (create, register, verify, and cleanup).
'
' @param Suite TestSuite instance
' @param aType2 Boolean flag indicating whether to use Type2 repository
''
Private Sub SpecRegisterExternalIdLifecycle(Suite As TestSuite, aType2 As Boolean)
    Dim repo As OdxRepository
    
    Set repo = CreateOdxRepository(aType2)
    
    With Suite.Test("TestRegisterExternalId - Lifecycle and Cleanup")
        
        Dim testPartnerVals As New Dictionary
        Dim createdIds As Collection
        Dim testModel As String
        Dim testExternalId As String
        Dim partnerId As Long
        Dim resDict As Dictionary
        
        testModel = "res.partner"
        testExternalId = "SpecOdxRepository.test_partner_lifecycle_1"
        
        ' ----------------------------------------------------------------------
        ' 1. Create a test record in res.partner
        ' ----------------------------------------------------------------------
        testPartnerVals.Add "name", "Test Partner (Unit Test)"
        testPartnerVals.Add "is_company", False
        
        Set createdIds = repo.CreateRecord(testModel, testPartnerVals)
        .NotUndefined createdIds
        .IsEqual createdIds.Count, 1
        
        partnerId = createdIds(1)
        
        ' ----------------------------------------------------------------------
        ' 2. Register External ID using RegisterExternalId
        ' ----------------------------------------------------------------------
        Dim isRegistered As Boolean
        isRegistered = repo.RegisterExternalId(testModel, partnerId, testExternalId)
        .IsEqual isRegistered, True
        
        ' ----------------------------------------------------------------------
        ' 3. Verify resolution from External ID to Model and ID
        ' ----------------------------------------------------------------------
        Set resDict = repo.GetModelAndIdFromExternalId(testExternalId)
        .NotUndefined resDict
        .IsEqual resDict("model"), testModel
        .IsEqual resDict("res_id"), partnerId
        
        ' ----------------------------------------------------------------------
        ' 4. Cleanup: Delete the test partner record
        ' (Odoo typically cascades or handles ir.model.data deletion on record unlink,
        '  or you can explicitly unlink from ir.model.data if required by your setup)
        ' ----------------------------------------------------------------------
        Dim isDeleted As Boolean
        isDeleted = repo.UnlinkRecord(testModel, partnerId)
        .IsEqual isDeleted, True
        
        ' ----------------------------------------------------------------------
        ' 5. Verify that the External ID is also removed or no longer resolves
        ' ----------------------------------------------------------------------
        Dim resolvedAfterDelete As Dictionary
        Set resolvedAfterDelete = repo.GetModelAndIdFromExternalId(testExternalId)
        
        ' Depending on Odoo's cascade delete behavior for ir.model.data,
        ' it should return Nothing or empty after the target record is deleted.
        If Not resolvedAfterDelete Is Nothing Then
            ' If ir.model.data record remains orphaned, ensure its res_id no longer exists or handle accordingly
            .IsEqual resolvedAfterDelete.Count, 0
        Else
            .IsUndefined resolvedAfterDelete
        End If
        
    End With
End Sub

''
' Tests for UpsertRecordByExternalId (Create and Update operations based on External ID).
'
' @param Suite TestSuite instance
' @param aType2 Boolean flag indicating whether to use Type2 repository
''
Private Sub SpecUpsertRecordByExternalId(Suite As TestSuite, aType2 As Boolean)
    Dim repo As OdxRepository
    
    Set repo = CreateOdxRepository(aType2)
    
    With Suite.Test("TestUpsertRecordByExternalId - Upsert and Status Verification")
        
        Dim testModel As String
        Dim testExternalId As String
        Dim mainDict As New Dictionary
        Dim partnerVals As New Dictionary
        Dim resultRes As Dictionary
        Dim fieldsToRead As New Collection
        Dim partnerRecords As Collection
        
        testModel = "res.partner"
        testExternalId = "SpecOdxRepository.test_upsert_partner_1"
        fieldsToRead.Add "name"
        
        ' ----------------------------------------------------------------------
        ' 0. Pre-cleanup (just in case previous test left orphaned record)
        ' ----------------------------------------------------------------------
        Dim existingMap As Dictionary
        Set existingMap = repo.GetModelAndIdFromExternalId(testExternalId)
        If Not existingMap Is Nothing Then
            repo.UnlinkRecord testModel, existingMap("res_id")
        End If
        
        ' ----------------------------------------------------------------------
        ' 1. Test: Create action via Upsert
        ' ----------------------------------------------------------------------
        partnerVals.Add "name", "Upsert Initial Name"
        mainDict.Add testExternalId, partnerVals
        
        Set resultRes = repo.UpsertRecordByExternalId(testModel, mainDict)
        
        .NotUndefined resultRes
        .IsEqual resultRes.Exists(testExternalId), True
        .IsEqual CStr(resultRes(testExternalId)("action")), "created"
        
        Dim createdId As Long
        createdId = CLng(resultRes(testExternalId)("id"))
        .NotEqual createdId, 0
        
        ' Verify the created record name
        Set partnerRecords = repo.ReadRecord(testModel, createdId, fieldsToRead)
        .IsEqual CStr(partnerRecords(1)("name")), "Upsert Initial Name"
        
        ' ----------------------------------------------------------------------
        ' 2. Test: Update action via Upsert (using the same external ID)
        ' ----------------------------------------------------------------------
        Dim updateDict As New Dictionary
        Dim updateVals As New Dictionary
        updateVals.Add "name", "Upsert Updated Name"
        updateDict.Add testExternalId, updateVals
        
        Set resultRes = repo.UpsertRecordByExternalId(testModel, updateDict)
        
        .NotUndefined resultRes
        .IsEqual resultRes.Exists(testExternalId), True
        .IsEqual CStr(resultRes(testExternalId)("action")), "updated"
        .IsEqual CLng(resultRes(testExternalId)("id")), createdId
        
        ' Verify the updated record name
        Set partnerRecords = repo.ReadRecord(testModel, createdId, fieldsToRead)
        .IsEqual CStr(partnerRecords(1)("name")), "Upsert Updated Name"
        
        ' ----------------------------------------------------------------------
        ' 3. Cleanup: Delete the test partner record
        ' ----------------------------------------------------------------------
        Dim isDeleted As Boolean
        isDeleted = repo.UnlinkRecord(testModel, createdId)
        .IsEqual isDeleted, True
        
    End With
End Sub

''
' Tests for UnlinkRecordByExternalId (Delete operations based on External ID(s) supporting String and Collection).
'
' @param Suite TestSuite instance
' @param aType2 Boolean flag indicating whether to use Type2 repository
''
Private Sub SpecUnlinkRecordByExternalId(Suite As TestSuite, aType2 As Boolean)
    Dim repo As OdxRepository
    
    Set repo = CreateOdxRepository(aType2)
    
    With Suite.Test("TestUnlinkRecordByExternalId - Unlink and Status Verification")
        
        Dim testModel As String
        Dim extId1 As String, extId2 As String, extId3 As String
        Dim testModelVals As Dictionary
        Dim createdIds As Collection
        Dim resultRes As Dictionary
        Dim existingMap As Dictionary
        
        testModel = "res.partner"
        extId1 = "SpecOdxRepository.test_unlink_partner_1"
        extId2 = "SpecOdxRepository.test_unlink_partner_2"
        extId3 = "SpecOdxRepository.test_unlink_partner_3"
        
        ' ----------------------------------------------------------------------
        ' 0. Pre-cleanup (just in case previous test left orphaned records)
        ' ----------------------------------------------------------------------
        Set existingMap = repo.GetModelAndIdFromExternalId(extId1)
        If Not existingMap Is Nothing Then repo.UnlinkRecord testModel, existingMap("res_id")
        
        Set existingMap = repo.GetModelAndIdFromExternalId(extId2)
        If Not existingMap Is Nothing Then repo.UnlinkRecord testModel, existingMap("res_id")
        
        Set existingMap = repo.GetModelAndIdFromExternalId(extId3)
        If Not existingMap Is Nothing Then repo.UnlinkRecord testModel, existingMap("res_id")
        
        ' ----------------------------------------------------------------------
        ' 1. Setup: Create 3 test records and register their External IDs
        ' ----------------------------------------------------------------------
        Set testModelVals = New Dictionary
        testModelVals.Add "name", "Unlink Test Partner A"
        Set createdIds = repo.CreateRecord(testModel, testModelVals)
        repo.RegisterExternalId testModel, createdIds(1), extId1
        
        testModelVals("name") = "Unlink Test Partner B"
        Set createdIds = repo.CreateRecord(testModel, testModelVals)
        repo.RegisterExternalId testModel, createdIds(1), extId2
        
        testModelVals("name") = "Unlink Test Partner C"
        Set createdIds = repo.CreateRecord(testModel, testModelVals)
        repo.RegisterExternalId testModel, createdIds(1), extId3
        
        ' ----------------------------------------------------------------------
        ' 2. Test: Unlink the 1st record individually using a String External ID
        ' ----------------------------------------------------------------------
        Set resultRes = repo.UnlinkRecordByExternalId(testModel, extId1)
        
        .NotUndefined resultRes
        .IsEqual resultRes.Exists(extId1), True
        .IsEqual CStr(resultRes(extId1)("action")), "unlinked"
        
        ' ----------------------------------------------------------------------
        ' 3. Test: Unlink all 3 records using a Collection (extId1 is already gone -> error, extId2 & extId3 -> success)
        ' ----------------------------------------------------------------------
        Dim extIdCol As New Collection
        extIdCol.Add extId1  ' Already deleted -> should result in an error
        extIdCol.Add extId2  ' Still exists -> success (unlinked)
        extIdCol.Add extId3  ' Still exists -> success (unlinked)
        
        Set resultRes = repo.UnlinkRecordByExternalId(testModel, extIdCol)
        
        .NotUndefined resultRes
        
        ' Verification 1: The 1st record should return an error since it was already unlinked
        .IsEqual resultRes.Exists(extId1), True
        .IsEqual CStr(resultRes(extId1)("action")), "error"
        .NotUndefined CStr(resultRes(extId1)("error_message"))
        
        ' Verification 2: The 2nd record should be successfully unlinked
        .IsEqual resultRes.Exists(extId2), True
        .IsEqual CStr(resultRes(extId2)("action")), "unlinked"
        
        ' Verification 3: The 3rd record should be successfully unlinked
        .IsEqual resultRes.Exists(extId3), True
        .IsEqual CStr(resultRes(extId3)("action")), "unlinked"
        
    End With
End Sub
