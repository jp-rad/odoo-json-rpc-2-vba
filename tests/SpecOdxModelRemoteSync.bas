Attribute VB_Name = "SpecOdxModelRemoteSync"
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
    SpecOdxModelRemoteSync Suite, False
    SpecOdxModelRemoteSync Suite, True
End Sub

''
' Tests for OdxModel remote synchronization methods:
' FetchSearchAndReadRemote, UpsertCurrentRecordRemote, and UnlinkCurrentRecordRemote.
'
' @param Suite TestSuite instance
' @param aType2 Boolean flag indicating whether to use Type2 repository
''
Private Sub SpecOdxModelRemoteSync(Suite As TestSuite, aType2 As Boolean)
    Dim Tests As New TestSuite
    Dim Test As TestCase
    Dim repo As OdxRepository
    Dim model As OdxModel
    
    Set repo = CreateOdxRepository(aType2)
    Set model = New OdxModel
    model.InitBase repo, "res.partner"
    model.AddExternalIdField
    model.AddField "name"
    model.AddField "is_company"
    
    With Suite.Test("TestOdxModelRemoteSync - Fetch, Upsert, and Unlink Current Record")
        Set Test = Tests.Test("should fetch, upsert, and unlink current record successfully")
        
        Dim testExternalId As String
        Dim existingMap As Dictionary
        Dim domain As New Collection
        Dim cond As New Collection
        Dim fetchCount As Long
        Dim upsertRes As Dictionary
        Dim unlinkRes As Dictionary
        
        testExternalId = "SpecOdxModel.test_remote_sync_1"
        
        ' ----------------------------------------------------------------------
        ' 0. Pre-cleanup (remove orphaned records if any)
        ' ----------------------------------------------------------------------
        Set existingMap = repo.GetModelAndIdFromExternalId(testExternalId)
        If Not existingMap Is Nothing Then
            repo.UnlinkRecord "res.partner", existingMap("res_id")
        End If
        
        ' ----------------------------------------------------------------------
        ' 1. Test: Fetch records remotely using FetchSearchAndReadRemote
        ' ----------------------------------------------------------------------
        fetchCount = model.FetchSearchAndReadRemote()
        Test.IsEqual (fetchCount >= 0), True
        
        ' ----------------------------------------------------------------------
        ' 2. Test: Add a new record locally and upsert via UpsertCurrentRecordRemote
        ' ----------------------------------------------------------------------
        Dim rs As ADODB.Recordset
        Set rs = model.Recordset
        
        rs.AddNew
        rs("name").Value = "OdxModel Sync Test Partner"
        rs("is_company").Value = False
        rs(model.ExternalIdFieldName).Value = testExternalId
        rs.Update
        
        ' Move to the newly added record to make it the current record
        rs.MoveLast
        
        ' Upsert current record (should trigger create)
        Set upsertRes = model.UpsertCurrentRecordRemote()
        
        Test.NotUndefined upsertRes
        Test.IsEqual CStr(upsertRes("action")), "created"
        Test.NotEqual CLng(upsertRes("id")), 0
        
        Dim createdId As Long
        createdId = CLng(upsertRes("id"))
        Test.IsEqual CLng(rs("id").Value), createdId
        
        ' ----------------------------------------------------------------------
        ' 3. Test: Modify local data and upsert again (should trigger update)
        ' ----------------------------------------------------------------------
        rs("name").Value = "OdxModel Sync Test Partner Updated"
        rs.Update
        
        Set upsertRes = model.UpsertCurrentRecordRemote()
        
        Test.NotUndefined upsertRes
        Test.IsEqual CStr(upsertRes("action")), "updated"
        Test.IsEqual CLng(upsertRes("id")), createdId
        
        ' ----------------------------------------------------------------------
        ' 4. Test: Unlink current record using UnlinkCurrentRecordRemote
        ' ----------------------------------------------------------------------
        Set unlinkRes = model.UnlinkCurrentRecordRemote()
        
        Test.NotUndefined unlinkRes
        Test.IsEqual CStr(unlinkRes("action")), "unlinked"
        Test.IsEqual CLng(unlinkRes("id")), createdId
                
        .IsEqual Test.Result, TestResultType.Pass
    End With
End Sub
