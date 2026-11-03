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
' Module: SpecOdxModel.bas
' Purpose: Contains unit test specifications for OdxModel class, including
'          backend communication, null handling, external ID configurations,
'          and operation filter tests.
' ------------------------------------------------------------------------------

Private Const CTEMP_FOLDER As String = "C:\Temp"
Private Const CTEST_PDF    As String = "test.pdf"

''
' Registers and runs all OdxModel test specifications across different repository types.
'
' @param aSuite TestSuite instance
' @param aType2 Boolean flag indicating whether to use Type2 repository
''
Public Sub RegisterSpecs(ByRef aSuite As TestSuite, ByVal aType2 As Boolean)
    SpecOdxModelBackendComm aSuite, aType2
    SpecOdxModelVbaJsonSampleTypes aSuite, aType2
    SpecOdxModelNullExternalId aSuite, aType2
    SpecOdxModelWithoutExternalIdField aSuite, aType2
End Sub

''
' Tests OdxModel backend communication methods with comprehensive coverage
' for creating and updating records containing Null values across multiple types
' (String, Float, Datetime).
'
' @param aSuite TestSuite instance
' @param aType2 Boolean flag indicating whether to use Type2 repository
''
Private Sub SpecOdxModelBackendComm(ByRef aSuite As TestSuite, ByVal aType2 As Boolean)
    Dim oRepo As OdxRepository
    Dim oModel As OdxModel
    
    Set oRepo = CreateOdxRepository(aType2)
    Set oModel = New OdxModel
    With oModel
        .InitBase oRepo, "res.partner"
        .AddExternalIdField
        
        .AddOdooField "name"
        .AddOdooField "is_company"
        .AddOdooField "phone"
        .AddOdooField "partner_latitude"
        .AddOdooField "comment"
        
        .AddStatusFields   ' Updatable
    End With
    
    With aSuite.Test("SpecOdxModelBackendComm - Comprehensive Null Coverage (String, Float, Datetime)")
        
        Dim sTestExternalId As String
        Dim dicExistingMap As Dictionary
        Dim nFetchCount As Long
        Dim nCreatedId As Long
        Dim colPartnerRecords As Collection
        Dim colFieldsToRead As New Collection
        
        sTestExternalId = "SpecOdxModel.test_backend_comprehensive_null"
        
        ' ----------------------------------------------------------------------
        ' 0. Pre-cleanup (remove orphaned records if any)
        ' ----------------------------------------------------------------------
        Set dicExistingMap = oRepo.GetModelAndIdFromExternalId(sTestExternalId)
        If Not dicExistingMap Is Nothing Then
            oRepo.UnlinkRecord "res.partner", dicExistingMap("res_id")
        End If
        
        nFetchCount = oModel.FetchSearchAndRead()
        
        ' ----------------------------------------------------------------------
        ' 1. Test: Create record with valid values and explicit Nulls across types
        ' ----------------------------------------------------------------------
        With oModel
            .AddNew
            
            .FieldValue("name") = "Comprehensive Null Test Partner"
            
            .FieldValue("is_company") = False
            .FieldValue("phone") = "+1-555-9999"         ' Valid string
            .FieldValue("partner_latitude") = 42.195     ' Valid float
            .FieldValue("comment") = Null                ' Null string
            
            .FieldExternalIdValue = sTestExternalId
            
            .Update
        End With
        
        .IsUndefined oModel.FieldIdValue
        oModel.ApplyBatchUpdate
        
        oModel.MoveLast
        Debug.Assert oModel.FieldValue("name") = "Comprehensive Null Test Partner"
        nCreatedId = oModel.FieldIdValue
        
        ' Verify creation values via backend communication
        colFieldsToRead.Add "phone"
        colFieldsToRead.Add "partner_latitude"
        colFieldsToRead.Add "comment"
        
        Set colPartnerRecords = oRepo.ReadRecord("res.partner", nCreatedId, colFieldsToRead)
        .IsEqual colPartnerRecords.Count, 1
        
        .IsEqual colPartnerRecords(1)("phone"), "+1-555-9999"
        .IsEqual colPartnerRecords(1)("partner_latitude"), 42.195
        .IsEqual colPartnerRecords(1)("comment"), False
        
        ' ----------------------------------------------------------------------
        ' 2. Test: Update record, changing valid values of various types to Null
        ' ----------------------------------------------------------------------
        oModel.MoveLast
        Debug.Assert oModel.FieldValue("name") = "Comprehensive Null Test Partner"
        
        With oModel
            .FieldValue("phone") = Null           ' Change string to Null
            .FieldValue("partner_latitude") = Null    ' Change float to Null
            .FieldValue("comment") = "Now populated"  ' Change null string to valid value
            
            .Update
        End With
        
        oModel.ApplyBatchUpdate
        
        oModel.MoveLast
        Debug.Assert oModel.FieldValue("name") = "Comprehensive Null Test Partner"
        
        ' Verify via backend that cleared fields are False/Null/Zero and comment is updated
        Set colPartnerRecords = oRepo.ReadRecord("res.partner", nCreatedId, colFieldsToRead)
        .IsEqual colPartnerRecords.Count, 1
        
        .IsEqual colPartnerRecords(1)("phone"), False
        .IsEqual colPartnerRecords(1)("partner_latitude"), False
        .IsEqual colPartnerRecords(1)("comment"), "<p>Now populated</p>"
        
        ' ----------------------------------------------------------------------
        ' 3. Cleanup: Unlink current record
        ' ----------------------------------------------------------------------
        oModel.MoveLast
        Debug.Assert oModel.FieldValue("name") = "Comprehensive Null Test Partner"
        
        oModel.MarkRecordForDelete
        oModel.ApplyBatchUpdate
        
        .IsEqual oModel.FetchSearchAndRead(), nFetchCount
        
    End With
End Sub

''
' Tests OdxModel when ExternalId is Null/empty.
'
' @param aSuite TestSuite instance
' @param aType2 Boolean flag indicating whether to use Type2 repository
''
Private Sub SpecOdxModelNullExternalId(ByRef aSuite As TestSuite, ByVal aType2 As Boolean)
    Dim oRepo As OdxRepository
    Dim oModel As OdxModel
    
    Set oRepo = CreateOdxRepository(aType2)
    Set oModel = New OdxModel
    With oModel
        .InitBase oRepo, "res.partner"
        .AddExternalIdField
        
        .AddOdooField "name"
        .AddOdooField "phone"
        
        .AddStatusFields   ' Updatable
    End With
    
    With aSuite.Test("SpecOdxModelNullExternalId - Create and Update with Null ExternalId")
        Dim sTestName As String
        Dim nCreatedId As Long
        Dim nFetchCount As Long
        
        sTestName = "Null ExternalId Test Partner"
        nFetchCount = oModel.FetchSearchAndRead()
        
        ' 1. Create record with explicit Null ExternalId
        With oModel
            .AddNew
            .FieldValue("name") = sTestName
            .FieldValue("phone") = "+1-555-1111"
            .FieldExternalIdValue = Null
            .Update
        End With
        
        oModel.ApplyBatchUpdate
        
        oModel.MoveLast
        .IsEqual oModel.FieldValue("name"), sTestName
        .NotUndefined oModel.FieldIdValue
        .NotUndefined oModel.FieldExternalIdValue
        .IsEqual InStr(1, oModel.FieldExternalIdValue, "__rpc_vba__.res_partner_"), 1
        nCreatedId = oModel.FieldIdValue
        
        ' 2. Update record without external ID
        With oModel
            .FieldValue("phone") = "+1-555-2222"
            .Update
        End With
        
        oModel.ApplyBatchUpdate
        
        oModel.MoveLast
        .IsEqual oModel.FieldIdValue, nCreatedId
        
        ' 3. Cleanup
        oModel.MoveLast
        oModel.MarkRecordForDelete
        oModel.ApplyBatchUpdate
        
        .IsEqual oModel.FetchSearchAndRead(), nFetchCount
    End With
End Sub

''
' Tests OdxModel when AddExternalIdField is NOT called (updating by ID).
'
' @param aSuite TestSuite instance
' @param aType2 Boolean flag indicating whether to use Type2 repository
''
Private Sub SpecOdxModelWithoutExternalIdField(ByRef aSuite As TestSuite, ByVal aType2 As Boolean)
    Dim oRepo As OdxRepository
    Dim oModel As OdxModel
    
    Set oRepo = CreateOdxRepository(aType2)
    Set oModel = New OdxModel
    With oModel
        .InitBase oRepo, "res.partner"
        ' Note: intentionally NOT calling .AddExternalIdField
        
        .AddOdooField "name"
        .AddOdooField "phone"
        
        .AddLocalField "local", MemLongVarWChar, MemSizeVariable
        
        .AddStatusFields   ' Updatable
    End With
    
    With aSuite.Test("SpecOdxModelWithoutExternalIdField - Update and Create by ID")
        Dim sTestName As String
        Dim sUpdatedName As String
        Dim nCreatedId As Long
        Dim nFetchCount As Long
        
        sTestName = "No ExternalId Field Test Partner"
        sUpdatedName = "No ExternalId Field Updated Partner"
        
        nFetchCount = oModel.FetchSearchAndRead()
        
        ' 1. Create record without external id field definition
        With oModel
            .AddNew
            .FieldValue("name") = sTestName
            .FieldValue("phone") = "+1-555-3333"
            .Update
        End With
        
        oModel.ApplyBatchUpdate
        
        oModel.MoveLast
        .IsEqual oModel.FieldValue("name"), sTestName
        .NotUndefined oModel.FieldIdValue
        nCreatedId = oModel.FieldIdValue
        
        ' 2. Update the record by ID (relying on internal ID tracking)
        With oModel
            .FieldValue("name") = sUpdatedName
            .Update
        End With
        
        oModel.ApplyBatchUpdate
        
        oModel.MoveLast
        .IsEqual oModel.FieldValue("name"), sUpdatedName
        .IsEqual oModel.FieldIdValue, nCreatedId
        .IsUndefined oModel.FieldValue("local")
        
        ' 3. Only local fields have changed, with no updates to Odoo fields.
        With oModel
            .FieldValue("local") = "test"
            .Update
        End With
        
        oModel.ApplyBatchUpdate
        
        ' 4. Cleanup: Delete record
        oModel.MoveLast
        oModel.MarkRecordForDelete
        oModel.ApplyBatchUpdate
        
        .IsEqual oModel.FetchSearchAndRead(), nFetchCount
    End With
End Sub

Private Sub SaveFile(ByVal aContentFieldValue As Variant, ByVal aFilenameFieldValue As Variant, ByVal aPath As String)
    Dim abContent() As Byte
    Dim sFilename As String
    
    If IsNull(aContentFieldValue) Then
        Exit Sub
    End If
    abContent = aContentFieldValue
    
    If IsNull(aFilenameFieldValue) Then
        sFilename = ""
    Else
        sFilename = Trim(aFilenameFieldValue)
    End If
    If sFilename = "" Then
        sFilename = "image_file_" & CStr(UBound(abContent) + 1) & "." & GetImageFormat(abContent)
    End If
    sFilename = aPath & "\" & sFilename
    
    SaveBinaryFile sFilename, abContent
    
    Debug.Print "[SpecOdxModelBackendComm] SaveFile: " & sFilename
End Sub

''
' Tests OdxModel backend communication methods using vba_json.sample
' with comprehensive coverage across ALL Odoo data types.
'
' @param aSuite TestSuite instance
' @param aType2 Boolean flag indicating whether to use Type2 repository
''
Private Sub SpecOdxModelVbaJsonSampleTypes(ByRef aSuite As TestSuite, ByVal aType2 As Boolean)
    Dim oRepo As OdxRepository
    Dim oModel As OdxModel
    
    If Dir(CTEMP_FOLDER, vbDirectory) = "" Then
        MkDir CTEMP_FOLDER
    End If
    
    Set oRepo = CreateOdxRepository(aType2)
    Set oModel = New OdxModel
    With oModel
        .InitBase oRepo, "vba_json.sample"
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
    
    With aSuite.Test("SpecOdxModelVbaJsonSampleTypes - Comprehensive All Field Types Coverage")
        
        Dim sTestExternalId As String
        Dim dicExistingMap As Dictionary
        Dim nFetchCount As Long
        Dim bResult As Boolean
        Dim nCreatedId As Long
        Dim colSampleRecords As Collection
        Dim colFieldsToRead As New Collection
        
        sTestExternalId = "SpecOdxModel.test_backend_all_types"
        
        ' ----------------------------------------------------------------------
        ' 0. Pre-cleanup (remove orphaned records if any)
        ' ----------------------------------------------------------------------
        Set dicExistingMap = oRepo.GetModelAndIdFromExternalId(sTestExternalId)
        If Not dicExistingMap Is Nothing Then
            oRepo.UnlinkRecord "vba_json.sample", dicExistingMap("res_id")
        End If
        
        nFetchCount = oModel.FetchSearchAndRead()
        
        If nFetchCount > 0 Then
            SaveFile oModel.FieldValue("binary_field"), oModel.FieldValue("binary_field_filename_"), CTEMP_FOLDER
            SaveFile oModel.FieldValue("image_field"), oModel.FieldValue("image_field_filename_"), CTEMP_FOLDER
        End If
        
        ' ----------------------------------------------------------------------
        ' 1. Test: Create record with ALL data types
        ' ----------------------------------------------------------------------
        With oModel
            .AddNew
            .FieldValue("name") = "VBA Sample All Types Test Record"
            
            .FieldValue("boolean_field") = True
            .FieldValue("char_field") = "All Types Char"
            .FieldValue("float_field") = 789.12
            .FieldValue("integer_field") = 100
            
            .FieldValue("html_field") = "<p>Hello <b>Odoo 20</b> HTML</p>"
            .FieldValue("currency_id") = 1
            .FieldValue("monetary_field") = 2500.5
            .FieldValue("selection_field") = "in_progress"
            .FieldValue("text_field") = "Line 1: Text field" & vbCrLf & "Line 2: Multi-line text"
            
            .FieldValue("date_field") = DateSerial(2026, 10, 18)
            .FieldValue("datetime_field") = DateSerial(2026, 10, 18) + TimeSerial(15, 30, 0)
            
            .FieldValue("many2one_field") = 1
            .FieldValue("many2many_field") = "[1, 4, 3, 2, 5]"
            
            .FieldValue("reference_field") = 1
            .FieldValue("reference_field_model_name_") = "res.partner"
            .FieldValue("ref_model") = "res.users"
            .FieldValue("ref_id") = 1
            
            .FieldValue("binary_field") = LoadBinaryFile(CTEMP_FOLDER & "\" & CTEST_PDF)
            .FieldValue("binary_field_filename_") = CTEST_PDF
            
            .FieldExternalIdValue = sTestExternalId
            
            .Update
        End With
        
        .IsEqual oModel.FieldValue("name"), "VBA Sample All Types Test Record"
        .IsUndefined oModel.FieldIdValue
        
        bResult = oModel.ApplyBatchUpdate
        .IsEqual True, bResult
        
        oModel.MoveLast
        .IsEqual oModel.FieldValue("name"), "VBA Sample All Types Test Record"
        .NotUndefined oModel.FieldIdValue
        nCreatedId = oModel.FieldIdValue
        
        ' ----------------------------------------------------------------------
        ' 2. Verify creation values via backend (Reading all fields back)
        ' ----------------------------------------------------------------------
        colFieldsToRead.Add "boolean_field"
        colFieldsToRead.Add "char_field"
        colFieldsToRead.Add "float_field"
        colFieldsToRead.Add "integer_field"
        colFieldsToRead.Add "html_field"
        colFieldsToRead.Add "monetary_field"
        colFieldsToRead.Add "selection_field"
        colFieldsToRead.Add "text_field"
        colFieldsToRead.Add "date_field"
        colFieldsToRead.Add "datetime_field"
        colFieldsToRead.Add "many2one_field"
        colFieldsToRead.Add "one2many_field"
        colFieldsToRead.Add "many2many_field"
        colFieldsToRead.Add "reference_field"
        colFieldsToRead.Add "ref_model"
        colFieldsToRead.Add "ref_id"
        colFieldsToRead.Add "binary_field"
        colFieldsToRead.Add "image_field"
        
        Set colSampleRecords = oRepo.ReadRecord("vba_json.sample", nCreatedId, colFieldsToRead)
        .IsEqual colSampleRecords.Count, 1
        
        .IsEqual colSampleRecords(1)("boolean_field"), True
        .IsEqual colSampleRecords(1)("char_field"), "All Types Char"
        .IsEqual colSampleRecords(1)("float_field"), 789.12
        .IsEqual colSampleRecords(1)("integer_field"), 100
        
        .IsEqual colSampleRecords(1)("html_field"), "<p>Hello <b>Odoo 20</b> HTML</p>"
        .IsEqual colSampleRecords(1)("monetary_field"), 2500.5
        .IsEqual colSampleRecords(1)("selection_field"), "in_progress"
        .IsEqual colSampleRecords(1)("text_field"), "Line 1: Text field" & vbCrLf & "Line 2: Multi-line text"
        
        .IsEqual colSampleRecords(1)("date_field"), "2026-10-18"
        .IsEqual colSampleRecords(1)("datetime_field"), "2026-10-18 06:30:00"
        
        .IsEqual ConvertToJson(colSampleRecords(1)("many2one_field")), "[1,""YourCompany""]"
        .IsEqual ConvertToJson(colSampleRecords(1)("many2many_field")), "[4,3,2,5,1]"
        
        .IsEqual colSampleRecords(1)("reference_field"), "res.partner,1"
        .IsEqual colSampleRecords(1)("ref_model"), "res.users"
        .IsEqual colSampleRecords(1)("ref_id"), 1
        
        .IsEqual colSampleRecords(1)("binary_field")("filename"), CTEST_PDF
        .IsEqual colSampleRecords(1)("image_field"), False
        
        ' ----------------------------------------------------------------------
        ' 3. Test: Update and verify
        ' ----------------------------------------------------------------------
        oModel.MoveLast
        .IsEqual oModel.FieldValue("name"), "VBA Sample All Types Test Record"
        
        With oModel
            .FieldValue("selection_field") = "done"
            .FieldValue("integer_field") = 200
            .FieldValue("binary_field") = Null
            .Update
        End With
        
        bResult = oModel.ApplyBatchUpdate
        .IsEqual True, bResult
        
        oModel.MoveLast
        .IsEqual oModel.FieldValue("name"), "VBA Sample All Types Test Record"
        .IsEqual oModel.FieldIdValue, nCreatedId
        
        Set colSampleRecords = oRepo.ReadRecord("vba_json.sample", nCreatedId, colFieldsToRead)
        .IsEqual colSampleRecords(1)("selection_field"), "done"
        .IsEqual colSampleRecords(1)("integer_field"), 200
        .IsEqual colSampleRecords(1)("binary_field"), False
        
        ' ----------------------------------------------------------------------
        ' 4. Cleanup: Unlink current record
        ' ----------------------------------------------------------------------
        oModel.MoveLast
        .IsEqual oModel.FieldValue("name"), "VBA Sample All Types Test Record"
        .IsEqual oModel.FieldIdValue, nCreatedId
        
        oModel.MarkRecordForDelete
        
        bResult = oModel.ApplyBatchUpdate
        .IsEqual True, bResult
        .IsEqual oModel.RecordCount, nFetchCount
        
        oModel.MoveLast
        .NotEqual oModel.FieldValue("name"), "VBA Sample All Types Test Record"
        .NotEqual oModel.FieldValue("id"), nCreatedId
        
    End With
End Sub
