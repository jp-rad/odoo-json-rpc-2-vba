Attribute VB_Name = "SpecOd2Client"
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
' Module: SpecOd2Client
' Purpose: Contains unit test specifications for Od2Client
' ------------------------------------------------------------------------------

''
' Register and run all Od2Client test specifications.
'
' @param Suite TestSuite instance
''
Public Sub RegisterSpecs(Suite As TestSuite)
    SpecInitializationGuard Suite
    SpecGetVersion Suite
    SpecJson2CrudFlow Suite
End Sub

' ------------------------------------------------------------------------------
' Test Specifications
' ------------------------------------------------------------------------------

''
' Tests the initialization guard check for Od2Client.
'
' @param Suite TestSuite instance
''
Private Sub SpecInitializationGuard(Suite As TestSuite)
    With Suite.Test("Initialization Guard Check")
        On Error Resume Next
        
        Dim objClient As Od2Client
        Set objClient = NewOd2Client
        objClient.ExecuteModelMethod "res.partner", "search_read"
        
        Dim errNum As Long
        errNum = Err.Number - vbObjectError
        On Error GoTo 0
        
        .IsEqual errNum, 1001, "Should raise uninitialized error code (1001)"
    End With
End Sub

''
' Tests the GetVersion API method for Od2Client.
'
' @param Suite TestSuite instance
''
Private Sub SpecGetVersion(Suite As TestSuite)
    With Suite.Test("Get Version API (GET /web/version)")
        Dim objClient As Od2Client
        Set objClient = SpecRunner.CreateOd2Client
        
        Dim res As Dictionary
        Set res = objClient.GetVersion()
        
        .IsOk res.Exists("result"), "GetVersion HTTP call failed"
        
        If res.Exists("result") Then
            Dim resultDict As Dictionary
            Set resultDict = res("result")
            .IsOk resultDict.Exists("version_info"), "Key 'version_info' missing in response"
            .IsOk resultDict.Exists("version"), "Key 'version' missing in response"
            If resultDict.Exists("version") Then
                Debug.Print "TESTS - SpecOd2Client.SpecGetVersion: " & resultDict("version")
            End If
        End If
    End With
End Sub

''
' Tests External JSON 2 CRUD operations (Create, SearchRead, Unlink) for Od2Client.
'
' @param Suite TestSuite instance
''
Private Sub SpecJson2CrudFlow(Suite As TestSuite)
    With Suite.Test("External JSON 2 CRUD Operations")
        Dim objClient As Od2Client
        Set objClient = SpecRunner.CreateOd2Client
        
        ' --- 1. Create ---
        Dim createVals As New Dictionary
        createVals.Add "name", "VBA_Test_Partner2_" & Format$(Now, "hhmmss")
        createVals.Add "email", "vba_test_partner2@example.com"
        
        Dim valsList As New Collection
        valsList.Add createVals
        
        Dim createParams As New Dictionary
        createParams.Add "vals_list", valsList
        
        Dim resCreate As Dictionary
        Set resCreate = objClient.ExecuteModelMethod("res.partner", "create", createParams)
        
        .IsOk resCreate.Exists("result"), "Create method failed"
        
        Dim newId As Long
        Dim createdData As Variant
        
        If IsObject(resCreate("result")) Then
            Set createdData = resCreate("result")
        Else
            createdData = resCreate("result")
        End If
        
        If TypeOf createdData Is Collection Then
            newId = CLng(createdData(1))
        Else
            newId = CLng(createdData)
        End If
        
        .IsOk (newId > 0), "Created ID is invalid"

        ' --- 2. SearchRead ---
        Dim domainCol As New Collection
        Dim cond As New Collection
        cond.Add "id"
        cond.Add "="
        cond.Add newId
        domainCol.Add cond
        
        Dim fieldsCol As New Collection
        fieldsCol.Add "id"
        fieldsCol.Add "name"
        fieldsCol.Add "email"
        
        Dim searchParams As New Dictionary
        searchParams.Add "domain", domainCol
        searchParams.Add "fields", fieldsCol
        
        Dim resSearch As Dictionary
        Set resSearch = objClient.ExecuteModelMethod("res.partner", "search_read", searchParams)
        
        .IsOk resSearch.Exists("result"), "SearchRead method failed"
        
        Dim records As Collection
        Set records = resSearch("result")
        .IsEqual records.Count, 1, "Expected exactly 1 record"

        ' --- 3. Unlink ---
        Dim idsList As New Collection
        idsList.Add newId
        
        Dim unlinkParams As New Dictionary
        unlinkParams.Add "ids", idsList
        
        Dim resUnlink As Dictionary
        Set resUnlink = objClient.ExecuteModelMethod("res.partner", "unlink", unlinkParams, True)
        
        .IsOk resUnlink("result"), "Unlink method failed"
    End With
End Sub
