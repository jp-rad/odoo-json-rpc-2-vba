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

Private Const CTEMP_FOLDER As String = "C:\Temp"
Private Const CTEST_PDF    As String = "test.pdf"
''
' Register and run all OdxRepository test specifications across different repository types.
'
' @param Suite TestSuite instance
''
Public Sub RegisterSpecs(Suite As TestSuite, aType2 As Boolean)
    SpecOdxModelRemoteSync Suite, aType2
    SpecVbaJsonSampleRemoteSync Suite, aType2
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
        '.IsEqual ParseDatetime(partnerRecords(1)("offline_since")), DateSerial(2026, 6, 1) + TimeSerial(12, 34, 56)
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

Private Sub SaveFile(aContentField As ADODB.Field, aFilenameField As ADODB.Field, aPath As String)
    Dim bContent() As Byte
    Dim sFilename As String
    
    bContent = aContentField.Value
    
    If IsNull(aFilenameField.Value) Then
        sFilename = ""
    Else
        sFilename = Trim(aFilenameField.Value)
    End If
    If sFilename = "" Then
        sFilename = "image_file_" & CStr(UBound(bContent) + 1) & "." & OdJsonRpc.GetImageFormat(bContent)
    End If
    sFilename = aPath & "\" & sFilename
    
    OdJsonRpc.SaveBinaryFile sFilename, bContent
    
    Debug.Print "[SpecOdxModelRemoteSync] SaveFile: " & sFilename
End Sub

''
' Tests for OdxModel remote synchronization methods using vba_json.sample
' with comprehensive coverage across ALL Odoo data types.
'
' @param Suite TestSuite instance
' @param aType2 Boolean flag indicating whether to use Type2 repository
''
Private Sub SpecVbaJsonSampleRemoteSync(Suite As TestSuite, aType2 As Boolean)
    Dim repo As OdxRepository
    Dim model As OdxModel
    
    If Dir(CTEMP_FOLDER, vbDirectory) = "" Then
        MkDir CTEMP_FOLDER
    End If
    
    Set repo = CreateOdxRepository(aType2)
    Set model = New OdxModel
    model.InitBase repo, "vba_json.sample"
    model.AddExternalIdField
    
    model.AddField "name"
    
    model.AddField "boolean_field"
    model.AddField "char_field"
    model.AddField "float_field"
    model.AddField "integer_field"
    
    model.AddField "html_field"
    model.AddField "currency_id"
    model.AddField "monetary_field"
    model.AddField "selection_field"
    model.AddField "text_field"
    
    model.AddField "date_field"
    model.AddField "datetime_field"
    
    model.AddField "many2one_field"
    model.AddField "one2many_field"
    model.AddField "many2many_field"
    
    model.AddField "reference_field"
    model.AddField "ref_model"
    model.AddField "ref_id"
    
    model.AddField "binary_field"
    model.AddField "image_field"
    
    With Suite.Test("TestVbaJsonSampleRemoteSync - Comprehensive All Field Types Coverage")
        
        Dim testExternalId As String
        Dim existingMap As Dictionary
        Dim fetchCount As Long
        Dim upsertRes As Dictionary
        Dim unlinkRes As Dictionary
        Dim rs As ADODB.Recordset
        Dim createdId As Long
        Dim sampleRecords As Collection
        Dim fieldsToRead As New Collection
        
        testExternalId = "SpecVbaJsonSample.test_remote_sync_all_types"
        
        ' ----------------------------------------------------------------------
        ' 0. Pre-cleanup (remove orphaned records if any)
        ' ----------------------------------------------------------------------
        Set existingMap = repo.GetModelAndIdFromExternalId(testExternalId)
        If Not existingMap Is Nothing Then
            repo.UnlinkRecord "vba_json.sample", existingMap("res_id")
        End If
        
        fetchCount = model.FetchSearchAndReadRemote()
        
        If fetchCount > 0 Then
            Set rs = model.Recordset
            SaveFile rs.Fields("binary_field"), rs.Fields("binary_field_filename_"), CTEMP_FOLDER
            SaveFile rs.Fields("image_field"), rs.Fields("image_field_filename_"), CTEMP_FOLDER
        End If
        
        ' ----------------------------------------------------------------------
        ' 1. Test: Create record with ALL data types
        ' ----------------------------------------------------------------------
        Set rs = model.Recordset
        
        rs.AddNew
        rs("id").Value = -1
        rs("name").Value = "VBA Sample All Types Test Record"
        
        ' Basic fields
        rs("boolean_field").Value = True
        rs("char_field").Value = "All Types Char"
        rs("float_field").Value = 789.12
        rs("integer_field").Value = 100
        
        ' Advanced fields
        rs("html_field").Value = "<p>Hello <b>Odoo 20</b> HTML</p>"
        rs("currency_id").Value = 1
        rs("monetary_field").Value = 2500.5
        rs("selection_field").Value = "in_progress"
        rs("text_field").Value = "Line 1: Text field" & vbCrLf & "Line 2: Multi-line text"
        
        ' Date(time) fields
        rs("date_field").Value = DateSerial(2026, 10, 18)
        rs("datetime_field").Value = DateSerial(2026, 10, 18) + TimeSerial(15, 30, 0)   ' +9 hours for JST
        
        ' Relational fields
        rs("many2one_field").Value = 1
        rs("many2many_field").Value = "[1, 4, 3, 2, 5]"
        
        ' Pseudo-relational fields (Reference)
        rs("reference_field").Value = 1
        rs("reference_field_model_name_").Value = "res.partner"
        rs("ref_model").Value = "res.users"
        rs("ref_id").Value = 1
        
        ' Binary fields
        rs("binary_field").Value = OdJsonRpc.LoadBinaryFile(CTEMP_FOLDER & "\" & CTEST_PDF)
        rs("binary_field_filename_").Value = CTEST_PDF
        
        ' rs("image_field").Value = ...
        
        rs(model.ExternalIdFieldName).Value = testExternalId
        
        rs.Update
        
        rs.MoveLast
        
        Set upsertRes = model.UpsertCurrentRecordRemote()
        
        .NotUndefined upsertRes
        .IsEqual CStr(upsertRes("action")), "created"
        .NotEqual CLng(upsertRes("id")), 0
        
        createdId = CLng(upsertRes("id"))
        .IsEqual CLng(rs("id").Value), createdId
        
        ' ----------------------------------------------------------------------
        ' 2. Verify creation values remotely (Reading all fields back)
        ' ----------------------------------------------------------------------
        fieldsToRead.Add "boolean_field"
        fieldsToRead.Add "char_field"
        fieldsToRead.Add "float_field"
        fieldsToRead.Add "integer_field"
        
        fieldsToRead.Add "html_field"
        fieldsToRead.Add "monetary_field"
        fieldsToRead.Add "selection_field"
        fieldsToRead.Add "text_field"
        
        fieldsToRead.Add "date_field"
        fieldsToRead.Add "datetime_field"
    
        fieldsToRead.Add "many2one_field"
        fieldsToRead.Add "one2many_field"
        fieldsToRead.Add "many2many_field"
        
        fieldsToRead.Add "reference_field"
        fieldsToRead.Add "ref_model"
        fieldsToRead.Add "ref_id"
        
        fieldsToRead.Add "binary_field"
        fieldsToRead.Add "image_field"
        
        Set sampleRecords = repo.ReadRecord("vba_json.sample", createdId, fieldsToRead)
        .IsEqual sampleRecords.Count, 1
        
        .IsEqual sampleRecords(1)("boolean_field"), True
        .IsEqual sampleRecords(1)("char_field"), "All Types Char"
        .IsEqual sampleRecords(1)("float_field"), 789.12
        .IsEqual sampleRecords(1)("integer_field"), 100
        
        .IsEqual sampleRecords(1)("html_field"), "<p>Hello <b>Odoo 20</b> HTML</p>"
        .IsEqual sampleRecords(1)("monetary_field"), 2500.5
        .IsEqual sampleRecords(1)("selection_field"), "in_progress"
        .IsEqual sampleRecords(1)("text_field"), "Line 1: Text field" & vbCrLf & "Line 2: Multi-line text"
        
        .IsEqual sampleRecords(1)("date_field"), "2026-10-18"
        .IsEqual sampleRecords(1)("datetime_field"), "2026-10-18 06:30:00"  ' UTC (-9 hours)
        
        .IsEqual ConvertToJson(sampleRecords(1)("many2one_field")), "[1,""YourCompany""]"
        .IsEqual ConvertToJson(sampleRecords(1)("many2many_field")), "[5,4,3,2,1]"
        
        .IsEqual sampleRecords(1)("reference_field"), "res.partner,1"
        .IsEqual sampleRecords(1)("ref_model"), "res.users"
        .IsEqual sampleRecords(1)("ref_id"), 1
        
        .IsEqual sampleRecords(1)("binary_field")("filename"), CTEST_PDF
        .IsEqual sampleRecords(1)("image_field"), False
        
        ' ----------------------------------------------------------------------
        ' 3. Test: Update and verify
        ' ----------------------------------------------------------------------
        rs.MoveLast
        rs("selection_field").Value = "done"
        rs("integer_field").Value = 200
        rs("binary_field").Value = Null
        rs.Update
        
        Set upsertRes = model.UpsertCurrentRecordRemote()
        .IsEqual CStr(upsertRes("action")), "updated"
        
        Set sampleRecords = repo.ReadRecord("vba_json.sample", createdId, fieldsToRead)
        .IsEqual sampleRecords(1)("selection_field"), "done"
        .IsEqual sampleRecords(1)("integer_field"), 200
        .IsEqual sampleRecords(1)("binary_field"), False
        
        ' ----------------------------------------------------------------------
        ' 4. Cleanup: Unlink current record
        ' ----------------------------------------------------------------------
        Set unlinkRes = model.UnlinkCurrentRecordRemote()
        .NotUndefined unlinkRes
        .IsEqual CStr(unlinkRes("action")), "unlinked"
        .IsEqual CLng(unlinkRes("id")), createdId
        
    End With
End Sub
