Attribute VB_Name = "SpecOd2Repository"
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
' Module: SpecOd2Repository
' Purpose: Contains unit test specifications for Od2Repository
' Standard:
'   - On success: Returns a Dictionary containing {"result": <payload>}
'   - On failure: Raises a runtime error.
' ------------------------------------------------------------------------------

''
' Register and run all Od2Repository test specifications.
'
' @param Suite TestSuite instance
''
Public Sub RegisterSpecs(Suite As TestSuite)
    SpecUninitializedGuard Suite
    SpecGetCurrentUserId Suite
    SpecExecuteModelMethod Suite
    SpecGetModelFields Suite
    SpecSearchRecord Suite
    SpecReadRecord Suite
    SpecSearchReadRecord Suite
    SpecRepositoryCrudLifecycle Suite
End Sub

' ------------------------------------------------------------------------------
' Test Specifications
' ------------------------------------------------------------------------------

''
' Tests the uninitialized guard check for Od2Repository.
'
' @param Suite TestSuite instance
''
Private Sub SpecUninitializedGuard(Suite As TestSuite)
    With Suite.Test("Uninitialized Guard Check")
        On Error Resume Next
        
        Dim objRepo As New Od2Repository
        objRepo.GetModelFields "res.partner"
        
        Dim errNum As Long
        errNum = Err.Number - vbObjectError
        On Error GoTo 0
        
        .IsEqual errNum, 1001, "Should raise uninitialized error code (1001)"
    End With
End Sub

''
' Tests fetching and caching the current user ID.
'
' @param Suite TestSuite instance
''
Private Sub SpecGetCurrentUserId(Suite As TestSuite)
    With Suite.Test("Get Current User ID (context_get)")
        Dim objClient As Od2Client
        Set objClient = SpecRunner.CreateOd2Client
        
        Dim objRepo As New Od2Repository
        objRepo.Initialize objClient
        
        ' 1. First fetch via RPC
        Dim Uid As Long
        Uid = objRepo.GetCurrentUserId()
        
        .IsOk (Uid > 0), "Should retrieve a valid user ID (> 0)"
        
        ' 2. Second fetch from internal cache
        Dim cachedUid As Long
        cachedUid = objRepo.GetCurrentUserId()
        
        .IsEqual cachedUid, Uid, "Cached User ID should match initial User ID"
    End With
End Sub

''
' Tests generic model method execution passthrough.
'
' @param Suite TestSuite instance
''
Private Sub SpecExecuteModelMethod(Suite As TestSuite)
    With Suite.Test("Generic ExecuteModelMethod Passthrough")
        Dim objClient As Od2Client
        Set objClient = SpecRunner.CreateOd2Client
        
        Dim objRepo As New Od2Repository
        objRepo.Initialize objClient
        
        Dim paramsDict As New Dictionary
        paramsDict.Add "domain", New Collection
        paramsDict.Add "limit", 1
        
        On Error Resume Next
        Dim resDict As Dictionary
        Set resDict = objRepo.ExecuteModelMethod("res.partner", "search_read", paramsDict)
        Dim errNum As Long
        errNum = Err.Number
        On Error GoTo 0
        
        .IsEqual errNum, 0, "ExecuteModelMethod call should not raise an error"
        If errNum = 0 Then
            Dim records As Collection
            Set records = resDict("result")
            .IsOk (records.Count > 0), "Expected at least 1 record"
        End If
    End With
End Sub

''
' Tests fetching model field definitions (fields_get).
'
' @param Suite TestSuite instance
''
Private Sub SpecGetModelFields(Suite As TestSuite)
    With Suite.Test("Get Model Fields (fields_get)")
        Dim objClient As Od2Client
        Set objClient = SpecRunner.CreateOd2Client
        
        Dim objRepo As New Od2Repository
        objRepo.Initialize objClient
        
        Dim targetFields As New Collection
        targetFields.Add "name"
        targetFields.Add "email"
        
        Dim targetAttrs As New Collection
        targetAttrs.Add "string"
        targetAttrs.Add "type"
        targetAttrs.Add "required"
        
        On Error Resume Next
        Dim resDict As Dictionary
        Set resDict = objRepo.GetModelFields("res.partner", targetFields, targetAttrs)
        Dim errNum As Long
        errNum = Err.Number
        On Error GoTo 0
        
        .IsEqual errNum, 0, "GetModelFields call should not raise an error"
        If errNum = 0 Then
            Dim fieldsDict As Dictionary
            Set fieldsDict = resDict("result")
            .IsOk fieldsDict.Exists("name"), "Key 'name' missing in field definitions"
            .IsOk fieldsDict.Exists("email"), "Key 'email' missing in field definitions"
        End If
    End With
End Sub

''
' Tests searching for record IDs (search).
'
' @param Suite TestSuite instance
''
Private Sub SpecSearchRecord(Suite As TestSuite)
    With Suite.Test("Search Records (search)")
        Dim objClient As Od2Client
        Set objClient = SpecRunner.CreateOd2Client
        
        Dim objRepo As New Od2Repository
        objRepo.Initialize objClient
        
        On Error Resume Next
        Dim resDict As Dictionary
        Set resDict = objRepo.SearchRecord("res.partner", aLimit:=2)
        Dim errNum As Long
        errNum = Err.Number
        On Error GoTo 0
        
        .IsEqual errNum, 0, "SearchRecord call should not raise an error"
        If errNum = 0 Then
            Dim ids As Collection
            Set ids = resDict("result")
            .IsOk (ids.Count > 0), "Should return at least 1 record ID"
        End If
    End With
End Sub

''
' Tests reading record details (read).
'
' @param Suite TestSuite instance
''
Private Sub SpecReadRecord(Suite As TestSuite)
    With Suite.Test("Read Records (read)")
        Dim objClient As Od2Client
        Set objClient = SpecRunner.CreateOd2Client
        
        Dim objRepo As New Od2Repository
        objRepo.Initialize objClient
        
        Dim targetFields As New Collection
        targetFields.Add "name"
        targetFields.Add "email"
        
        On Error Resume Next
        Dim resDict As Dictionary
        Set resDict = objRepo.ReadRecord("res.partner", 1, targetFields)
        Dim errNum As Long
        errNum = Err.Number
        On Error GoTo 0
        
        .IsEqual errNum, 0, "ReadRecord call should not raise an error"
        If errNum = 0 Then
            Dim records As Collection
            Set records = resDict("result")
            .IsEqual records.Count, 1, "Expected exactly 1 record"
        End If
    End With
End Sub

''
' Tests searching and reading records in a single call (search_read).
'
' @param Suite TestSuite instance
''
Private Sub SpecSearchReadRecord(Suite As TestSuite)
    With Suite.Test("Search and Read Records (search_read)")
        Dim objClient As Od2Client
        Set objClient = SpecRunner.CreateOd2Client
        
        Dim objRepo As New Od2Repository
        objRepo.Initialize objClient
        
        Dim domain As New Collection
        Dim cond As New Collection
        cond.Add "is_company"
        cond.Add "="
        cond.Add True
        domain.Add cond
        
        Dim targetFields As New Collection
        targetFields.Add "name"
        
        On Error Resume Next
        Dim resDict As Dictionary
        Set resDict = objRepo.SearchReadRecord("res.partner", domain, targetFields, aLimit:=5)
        Dim errNum As Long
        errNum = Err.Number
        On Error GoTo 0
        
        .IsEqual errNum, 0, "SearchReadRecord call should not raise an error"
        If errNum = 0 Then
            Dim records As Collection
            Set records = resDict("result")
            .IsOk (records.Count > 0), "Expected at least 1 record"
        End If
    End With
End Sub

''
' Tests repository CRUD lifecycle operations (Create, Write, Unlink).
'
' @param Suite TestSuite instance
''
Private Sub SpecRepositoryCrudLifecycle(Suite As TestSuite)
    With Suite.Test("Repository CRUD Lifecycle (Create, Write, Unlink)")
        Dim objClient As Od2Client
        Set objClient = SpecRunner.CreateOd2Client
        
        Dim objRepo As New Od2Repository
        objRepo.Initialize objClient
        
        ' --- 1. CreateRecord ---
        Dim newVals As New Dictionary
        newVals.Add "name", "VBA_Test_Partner_" & Format$(Now, "hhmmss")
        newVals.Add "email", "dummy.test@example.com"
        
        On Error Resume Next
        Dim resCreate As Dictionary
        Set resCreate = objRepo.CreateRecord("res.partner", newVals)
        Dim errCreate As Long
        errCreate = Err.Number
        On Error GoTo 0
        
        .IsEqual errCreate, 0, "CreateRecord should not raise an error"
        
        Dim newId As Long
        If errCreate = 0 Then
            Dim createdData As Variant
            If IsObject(resCreate("result")) Then
                Set createdData = resCreate("result")
            Else
                createdData = resCreate("result")
            End If
            
            If IsObject(createdData) Then
                If TypeOf createdData Is Collection Then
                    Dim colData As Collection
                    Set colData = createdData
                    newId = CLng(colData(1))
                End If
            Else
                newId = CLng(createdData)
            End If
            .IsOk (newId > 0), "Created ID is invalid"
        End If
        
        ' --- 2. WriteRecord ---
        If newId > 0 Then
            Dim updateVals As New Dictionary
            updateVals.Add "email", "dummy.updated@example.com"
            
            On Error Resume Next
            Dim resWrite As Dictionary
            Set resWrite = objRepo.WriteRecord("res.partner", newId, updateVals)
            Dim errWrite As Long
            errWrite = Err.Number
            On Error GoTo 0
            
            .IsEqual errWrite, 0, "WriteRecord should not raise an error"
            If errWrite = 0 Then
                .IsOk CBool(resWrite("result")), "WriteRecord should return True"
            End If
        End If
        
        ' --- 3. UnlinkRecord ---
        If newId > 0 Then
            On Error Resume Next
            Dim resUnlink As Dictionary
            Set resUnlink = objRepo.UnlinkRecord("res.partner", newId)
            Dim errUnlink As Long
            errUnlink = Err.Number
            On Error GoTo 0
            
            .IsEqual errUnlink, 0, "UnlinkRecord should not raise an error"
            If errUnlink = 0 Then
                .IsOk CBool(resUnlink("result")), "UnlinkRecord should return True"
            End If
        End If
    End With
End Sub
