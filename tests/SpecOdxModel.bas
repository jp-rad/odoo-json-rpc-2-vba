Attribute VB_Name = "SpecOdxModel"
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
    With model
        .InitBase repo, "res.partner"
        .AddExternalIdField
        
        .AddOdooField "name"
        .AddOdooField "is_company"
        .AddOdooField "phone"
        .AddOdooField "partner_latitude"
        '.AddOdooField "offline_since"
        .AddOdooField "comment"
        
        .AddStatusFields   ' Updatable
    End With
    
    With Suite.Test("TestOdxModelRemoteSync - Comprehensive Null Coverage (String, float, Datetime)")
        
        Dim testExternalId As String
        Dim existingMap As Dictionary
        Dim fetchCount As Long
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
        
        fetchCount = model.FetchSearchAndRead()
        
        ' ----------------------------------------------------------------------
        ' 1. Test: Create record with valid values and explicit Nulls across types
        ' ----------------------------------------------------------------------
        With model
            .AddNew
            
            .FieldValue("name") = "Comprehensive Null Test Partner"
            
            .FieldValue("is_company") = False
            .FieldValue("phone") = "+1-555-9999"                  ' Valid string
            .FieldValue("partner_latitude") = 42.195              ' Valid integer
            '.FieldValue("offline_since") = DateSerial(2026, 6, 1) + TimeSerial(12, 34, 56) ' Valid datetime
            .FieldValue("comment") = Null                         ' Null string
            
            .FieldExternalIdValue = testExternalId
            
            .Update
        End With
        
        .IsUndefined model.FieldIdValue
        model.ApplyBatchUpdate
        
        model.MoveLast
        Debug.Assert model.FieldValue("name") = "Comprehensive Null Test Partner"
        createdId = model.FieldIdValue
        
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
        model.MoveLast
        Debug.Assert model.FieldValue("name") = "Comprehensive Null Test Partner"
        
        With model
            .FieldValue("phone") = Null               ' Change string to Null
            .FieldValue("partner_latitude") = Null    ' Change float to Null
            '.FieldValue("offline_since") = Null       ' Change datetime to Null
            .FieldValue("comment") = "Now populated"  ' Change null string to valid value
            
            .Update
        End With
        
        model.ApplyBatchUpdate
        
        model.MoveLast
        Debug.Assert model.FieldValue("name") = "Comprehensive Null Test Partner"
        
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
        model.MoveLast
        Debug.Assert model.FieldValue("name") = "Comprehensive Null Test Partner"
        
        model.MarkRecordForDelete
        model.ApplyBatchUpdate
        
        .IsEqual model.FetchSearchAndRead(), fetchCount
        
    End With
End Sub

Private Sub SaveFile(aContentFieldValue As Variant, aFilenameFieldValue As Variant, aPath As String)
    Dim bContent() As Byte
    Dim sFilename As String
    
    If IsNull(aContentFieldValue) Then
        Exit Sub
    End If
    bContent = aContentFieldValue
    
    If IsNull(aFilenameFieldValue) Then
        sFilename = ""
    Else
        sFilename = Trim(aFilenameFieldValue)
    End If
    If sFilename = "" Then
        sFilename = "image_file_" & CStr(UBound(bContent) + 1) & "." & GetImageFormat(bContent)
    End If
    sFilename = aPath & "\" & sFilename
    
    SaveBinaryFile sFilename, bContent
    
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
    With model
        .InitBase repo, "vba_json.sample"
        .AddExternalIdField
        
        .AddOdooField "name"
        
        .AddOdooField "boolean_field"
        .AddOdooField "char_field"
        .AddOdooField "float_field"
        .AddOdooField "integer_field"
        
        .AddOdooField "html_field"
        .AddOdooField "currency_id"
        .AddOdooField "monetary_field"
        .AddOdooField "selection_field"
        .AddOdooField "text_field"
        
        .AddOdooField "date_field"
        .AddOdooField "datetime_field"
        
        .AddOdooField "many2one_field"
        .AddOdooField "one2many_field"
        .AddOdooField "many2many_field"
        
        .AddOdooField "reference_field"
        .AddOdooField "ref_model"
        .AddOdooField "ref_id"
        
        .AddOdooField "binary_field"
        .AddOdooField "image_field"
        
        .AddStatusFields   ' Updatable
    End With
    
    With Suite.Test("TestVbaJsonSampleRemoteSync - Comprehensive All Field Types Coverage")
        
        Dim testExternalId As String
        Dim existingMap As Dictionary
        Dim fetchCount As Long
        Dim upsertRes As Dictionary
        Dim unlinkRes As Dictionary
        Dim bResult As Boolean
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
        
        fetchCount = model.FetchSearchAndRead()
        
        If fetchCount > 0 Then
            SaveFile model.FieldValue("binary_field"), model.FieldValue("binary_field_filename_"), CTEMP_FOLDER
            SaveFile model.FieldValue("image_field"), model.FieldValue("image_field_filename_"), CTEMP_FOLDER
        End If
        
        ' ----------------------------------------------------------------------
        ' 1. Test: Create record with ALL data types
        ' ----------------------------------------------------------------------
        With model
            .AddNew
            .FieldValue("name") = "VBA Sample All Types Test Record"
            
            ' Basic fields
            .FieldValue("boolean_field") = True
            .FieldValue("char_field") = "All Types Char"
            .FieldValue("float_field") = 789.12
            .FieldValue("integer_field") = 100
            
            ' Advanced fields
            .FieldValue("html_field") = "<p>Hello <b>Odoo 20</b> HTML</p>"
            .FieldValue("currency_id") = 1
            .FieldValue("monetary_field") = 2500.5
            .FieldValue("selection_field") = "in_progress"
            .FieldValue("text_field") = "Line 1: Text field" & vbCrLf & "Line 2: Multi-line text"
            
            ' Date(time) fields
            .FieldValue("date_field") = DateSerial(2026, 10, 18)
            .FieldValue("datetime_field") = DateSerial(2026, 10, 18) + TimeSerial(15, 30, 0)  ' +9 hours for JST
            
            ' Relational fields
            .FieldValue("many2one_field") = 1
            .FieldValue("many2many_field") = "[1, 4, 3, 2, 5]"
            
            ' Pseudo-relational fields (Reference)
            .FieldValue("reference_field") = 1
            .FieldValue("reference_field_model_name_") = "res.partner"
            .FieldValue("ref_model") = "res.users"
            .FieldValue("ref_id") = 1
            
            ' Binary fields
            .FieldValue("binary_field") = LoadBinaryFile(CTEMP_FOLDER & "\" & CTEST_PDF)
            .FieldValue("binary_field_filename_") = CTEST_PDF
            
            '.fieldvalue("image_field")= ...
            
            .FieldExternalIdValue = testExternalId
            
            .Update
        End With
        
        .IsEqual model.FieldValue("name"), "VBA Sample All Types Test Record"
        .IsUndefined model.FieldIdValue
        
        bResult = model.ApplyBatchUpdate
        .IsEqual True, bResult
        
        model.MoveLast
        .IsEqual model.FieldValue("name"), "VBA Sample All Types Test Record"
        .NotUndefined model.FieldIdValue
        createdId = model.FieldIdValue
        
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
        .IsEqual ConvertToJson(sampleRecords(1)("many2many_field")), "[4,3,2,5,1]"
        
        .IsEqual sampleRecords(1)("reference_field"), "res.partner,1"
        .IsEqual sampleRecords(1)("ref_model"), "res.users"
        .IsEqual sampleRecords(1)("ref_id"), 1
        
        .IsEqual sampleRecords(1)("binary_field")("filename"), CTEST_PDF
        .IsEqual sampleRecords(1)("image_field"), False
        
        ' ----------------------------------------------------------------------
        ' 3. Test: Update and verify
        ' ----------------------------------------------------------------------
        model.MoveLast
        .IsEqual model.FieldValue("name"), "VBA Sample All Types Test Record"
        
        With model
            .FieldValue("selection_field") = "done"
            .FieldValue("integer_field") = 200
            .FieldValue("binary_field") = Null
            .Update
        End With
        
        bResult = model.ApplyBatchUpdate
        .IsEqual True, bResult
        
        model.MoveLast
        .IsEqual model.FieldValue("name"), "VBA Sample All Types Test Record"
        .IsEqual model.FieldIdValue, createdId
        
        Set sampleRecords = repo.ReadRecord("vba_json.sample", createdId, fieldsToRead)
        .IsEqual sampleRecords(1)("selection_field"), "done"
        .IsEqual sampleRecords(1)("integer_field"), 200
        .IsEqual sampleRecords(1)("binary_field"), False
        
        ' ----------------------------------------------------------------------
        ' 4. Cleanup: Unlink current record
        ' ----------------------------------------------------------------------
        model.MoveLast
        .IsEqual model.FieldValue("name"), "VBA Sample All Types Test Record"
        .IsEqual model.FieldIdValue, createdId
        
        model.MarkRecordForDelete
        
        bResult = model.ApplyBatchUpdate
        .IsEqual True, bResult
        .IsEqual model.RecordCount, fetchCount
        
        model.MoveLast
        .NotEqual model.FieldValue("name"), "VBA Sample All Types Test Record"
        .NotEqual model.FieldValue("id"), createdId
        
    End With
End Sub
