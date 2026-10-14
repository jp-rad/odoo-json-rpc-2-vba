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
    Dim Tests As New TestSuite
    Dim Test As TestCase
    Dim repo As OdxRepository
    
    Set repo = CreateOdxRepository(aType2)
    
    With Suite.Test("TestOdxRepository - External ID and Record Resolution")
        Set Test = Tests.Test("should resolve external ids and match partner names dynamically")
        
        Dim resolvedModelId As Dictionary
        Dim partnerRecords As Collection
        Dim fieldsToRead As New Collection
        
        fieldsToRead.Add "name"
        
        ' ----------------------------------------------------------------------
        ' 1. Test: Resolve from External ID ("base.res_partner_2")
        ' Verify model and read record to confirm the name is "Acme Corporation"
        ' ----------------------------------------------------------------------
        Set resolvedModelId = repo.GetModelAndIdFromExternalId("base.res_partner_2")
        Test.NotUndefined resolvedModelId
        Test.IsEqual resolvedModelId("model"), "res.partner"
        
        ' Read record using the resolved res_id to check name dynamically
        Set partnerRecords = repo.ReadRecord("res.partner", resolvedModelId("res_id"), fieldsToRead)
        Test.IsEqual partnerRecords.Count, 1
        Test.IsEqual CStr(partnerRecords(1)("name")), "Acme Corporation"
        
        ' ----------------------------------------------------------------------
        ' 2. Test: Resolve from External ID ("base.res_partner_address_31")
        ' Verify record name is "Addison Olson"
        ' ----------------------------------------------------------------------
        Set resolvedModelId = repo.GetModelAndIdFromExternalId("base.res_partner_address_31")
        Test.NotUndefined resolvedModelId
        Test.IsEqual resolvedModelId("model"), "res.partner"
        
        Set partnerRecords = repo.ReadRecord("res.partner", resolvedModelId("res_id"), fieldsToRead)
        Test.IsEqual partnerRecords.Count, 1
        Test.IsEqual CStr(partnerRecords(1)("name")), "Addison Olson"
        
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
            Test.IsEqual externalId, "base.res_partner_address_3"
        End If
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
End Sub

''
' Tests for RegisterExternalId and related lifecycle (create, register, verify, and cleanup).
'
' @param Suite TestSuite instance
' @param aType2 Boolean flag indicating whether to use Type2 repository
''
Private Sub SpecRegisterExternalIdLifecycle(Suite As TestSuite, aType2 As Boolean)
    Dim Tests As New TestSuite
    Dim Test As TestCase
    Dim repo As OdxRepository
    
    Set repo = CreateOdxRepository(aType2)
    
    With Suite.Test("TestRegisterExternalId - Lifecycle and Cleanup")
        Set Test = Tests.Test("should register, verify, and cleanup external id and partner record")
        
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
        Test.NotUndefined createdIds
        Test.IsEqual createdIds.Count, 1
        
        partnerId = createdIds(1)
        
        ' ----------------------------------------------------------------------
        ' 2. Register External ID using RegisterExternalId
        ' ----------------------------------------------------------------------
        Dim isRegistered As Boolean
        isRegistered = repo.RegisterExternalId(testModel, partnerId, testExternalId)
        Test.IsEqual isRegistered, True
        
        ' ----------------------------------------------------------------------
        ' 3. Verify resolution from External ID to Model and ID
        ' ----------------------------------------------------------------------
        Set resDict = repo.GetModelAndIdFromExternalId(testExternalId)
        Test.NotUndefined resDict
        Test.IsEqual resDict("model"), testModel
        Test.IsEqual resDict("res_id"), partnerId
        
        ' ----------------------------------------------------------------------
        ' 4. Cleanup: Delete the test partner record
        ' (Odoo typically cascades or handles ir.model.data deletion on record unlink,
        '  or you can explicitly unlink from ir.model.data if required by your setup)
        ' ----------------------------------------------------------------------
        Dim isDeleted As Boolean
        isDeleted = repo.UnlinkRecord(testModel, partnerId)
        Test.IsEqual isDeleted, True
        
        ' ----------------------------------------------------------------------
        ' 5. Verify that the External ID is also removed or no longer resolves
        ' ----------------------------------------------------------------------
        Dim resolvedAfterDelete As Dictionary
        Set resolvedAfterDelete = repo.GetModelAndIdFromExternalId(testExternalId)
        
        ' Depending on Odoo's cascade delete behavior for ir.model.data,
        ' it should return Nothing or empty after the target record is deleted.
        If Not resolvedAfterDelete Is Nothing Then
            ' If ir.model.data record remains orphaned, ensure its res_id no longer exists or handle accordingly
            Test.IsEqual resolvedAfterDelete.Count, 0
        Else
            Test.IsUndefined resolvedAfterDelete
        End If
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
End Sub
