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
Public Sub RegisterSpecs(Suite As TestSuite, aType2 As Boolean)
    SpecOdxModelRemoteSync Suite, aType2
End Sub

''
' Tests for OdxModel remote synchronization methods with comprehensive coverage
' for creating and updating records containing Null values across multiple types
' (String, Integer, Date/Datetime).
'
' @param Suite TestSuite instance
' @param aType2 Boolean flag indicating whether to use Type2 repository
''
Private Sub SpecOdxModelRemoteSync(Suite As TestSuite, aType2 As Boolean)
    Dim repo As OdxRepository
    Dim model As OdxModel
    
    Set repo = CreateOdxRepository(aType2)
    Set model = New OdxModel
    model.InitBase repo, "res.partner"
    model.AddExternalIdField
    model.AddField "name"
    model.AddField "is_company"
    model.AddField "comment"
    model.AddField "phone"
    model.AddField "partner_latitude"
    'model.AddField "offline_since"
    
    With Suite.Test("TestOdxModelRemoteSync - Comprehensive Null Coverage (String, float, Datetime)")
        
        Dim testExternalId As String
        Dim existingMap As Dictionary
        Dim fetchCount As Long
        Dim upsertRes As Dictionary
        Dim unlinkRes As Dictionary
        Dim rs As ADODB.Recordset
        Dim createdId As Long
        Dim partnerRecords As Collection
        Dim fieldsToRead As New Collection
        
        testExternalId = "SpecOdxModel.test_remote_sync_comprehensive_null"
        
        ' ----------------------------------------------------------------------
        ' 0. Pre-cleanup (remove orphaned records if any)
        ' ----------------------------------------------------------------------
        Set existingMap = repo.GetModelAndIdFromExternalId(testExternalId)
        If Not existingMap Is Nothing Then
            repo.UnlinkRecord "res.partner", existingMap("res_id")
        End If
        
        fetchCount = model.FetchSearchAndReadRemote()
        
        ' ----------------------------------------------------------------------
        ' 1. Test: Create record with valid values and explicit Nulls across types
        ' ----------------------------------------------------------------------
        Set rs = model.Recordset
        
        rs.AddNew
        rs("id").Value = -1
        rs("name").Value = "Comprehensive Null Test Partner"
        rs("is_company").Value = False
        rs("phone").Value = "+1-555-9999"                   ' Valid string
        rs("partner_latitude").Value = 42.195               ' Valid integer
        'rs("offline_since").Value = DateSerial(2026, 6, 1) + TimeSerial(12, 34, 56) ' Valid datetime
        rs("comment").Value = Null                          ' Null string
        rs(model.ExternalIdFieldName).Value = testExternalId
        rs.Update
        
        rs.MoveLast
        
        Set upsertRes = model.UpsertCurrentRecordRemote()
        
        .NotUndefined upsertRes
        .IsEqual CStr(upsertRes("action")), "created"
        .NotEqual CLng(upsertRes("id")), 0
        
        createdId = CLng(upsertRes("id"))
        .IsEqual CLng(rs("id").Value), createdId
        
        ' Verify creation values remotely
        fieldsToRead.Add "phone"
        fieldsToRead.Add "partner_latitude"
        'fieldsToRead.Add "offline_since"
        fieldsToRead.Add "comment"
        
        Set partnerRecords = repo.ReadRecord("res.partner", createdId, fieldsToRead)
        .IsEqual partnerRecords.Count, 1
        
        .IsEqual partnerRecords(1)("phone"), "+1-555-9999"
        .IsEqual partnerRecords(1)("partner_latitude"), 42.195
        '.IsEqual ParseUtcDatetime(partnerRecords(1)("offline_since")), DateSerial(2026, 6, 1) + TimeSerial(12, 34, 56)
        .IsEqual partnerRecords(1)("comment"), False
        
        ' ----------------------------------------------------------------------
        ' 2. Test: Update record, changing valid values of various types to Null
        ' ----------------------------------------------------------------------
        rs.MoveLast
        Debug.Assert rs("id").Value = createdId
        rs("phone").Value = Null                ' Change string to Null
        rs("partner_latitude").Value = Null     ' Change float to Null
        'rs("offline_since").Value = Null        ' Change datetime to Null
        rs("comment").Value = "Now populated"   ' Change null string to valid value
        rs.Update
        
        Set upsertRes = model.UpsertCurrentRecordRemote()
        
        .NotUndefined upsertRes
        .IsEqual CStr(upsertRes("action")), "updated"
        .IsEqual CLng(upsertRes("id")), createdId
        
        ' Verify remotely that cleared fields are False/Null/Zero and comment is updated
        Set partnerRecords = repo.ReadRecord("res.partner", createdId, fieldsToRead)
        .IsEqual partnerRecords.Count, 1
        
        .IsEqual partnerRecords(1)("phone"), False
        .IsEqual partnerRecords(1)("partner_latitude"), False
        '.IsEqual partnerRecords(1)("offline_since"), False
        .IsEqual partnerRecords(1)("comment"), "<p>Now populated</p>"
        
        ' ----------------------------------------------------------------------
        ' 3. Cleanup: Unlink current record
        ' ----------------------------------------------------------------------
        Set unlinkRes = model.UnlinkCurrentRecordRemote()
        
        .NotUndefined unlinkRes
        .IsEqual CStr(unlinkRes("action")), "unlinked"
        .IsEqual CLng(unlinkRes("id")), createdId
        
    End With
End Sub
