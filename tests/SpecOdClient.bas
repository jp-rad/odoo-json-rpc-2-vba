Attribute VB_Name = "SpecOdClient"
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
' Module: SpecOdClient
' Purpose: Contains unit test specifications for OdClient (JSON-RPC)
' ------------------------------------------------------------------------------

''
' Register and run all OdClient test specifications.
'
' @param Suite TestSuite instance
''
Public Sub RegisterSpecs(Suite As TestSuite)
    SpecInitializationGuard Suite
    SpecGetVersion Suite
    SpecJsonRpcCrudFlow Suite
End Sub

' ------------------------------------------------------------------------------
' Test Specifications
' ------------------------------------------------------------------------------

''
' Tests the initialization guard check for OdClient.
'
' @param Suite TestSuite instance
''
Private Sub SpecInitializationGuard(Suite As TestSuite)
    With Suite.Test("Initialization Guard Check")
        On Error Resume Next
        
        Dim objClient As OdClient
        Set objClient = NewOdClient
        
        Dim args As New Collection
        objClient.ExecuteKw "res.partner", "search_read", args
        
        Dim errNum As Long
        errNum = Err.Number - vbObjectError
        On Error GoTo 0
        
        .IsEqual errNum, 1001, "Should raise uninitialized error code (1001)"
    End With
End Sub

''
' Tests the GetVersion API method for OdClient.
'
' @param Suite TestSuite instance
''
Private Sub SpecGetVersion(Suite As TestSuite)
    With Suite.Test("Get Version API (common service -> version)")
        Dim objClient As OdClient
        Set objClient = SpecRunner.CreateOdClient
        
        Dim res As Dictionary
        Set res = objClient.GetVersion()
        
        .IsOk res.Exists("result"), "GetVersion JSON-RPC call failed"
        
        If res.Exists("result") Then
            Dim resultDict As Dictionary
            Set resultDict = res("result")
            .IsOk resultDict.Exists("server_version"), "Key 'server_version' missing in response"
            If resultDict.Exists("server_version") Then
                Debug.Print "TESTS - SpecOdClient.SpecGetVersion: " & resultDict("server_version")
            End If
        End If
    End With
End Sub

''
' Tests JSON-RPC CRUD operations (Authenticate, Create, SearchRead, Unlink) for OdClient.
'
' @param Suite TestSuite instance
''
Private Sub SpecJsonRpcCrudFlow(Suite As TestSuite)
    With Suite.Test("JSON-RPC CRUD Operations (Authenticate, Create, SearchRead, Unlink)")
        Dim objClient As OdClient
        Set objClient = SpecRunner.CreateOdClient
        
        ' --- 1. Create ---
        Dim createVals As New Dictionary
        createVals.Add "name", "VBA_Test_Partner_" & Format$(Now, "hhmmss")
        createVals.Add "email", "vba_test_partner@example.com@example.com"
        
        Dim argsCreate As New Collection
        argsCreate.Add createVals
        
        Dim resCreate As Dictionary
        Set resCreate = objClient.ExecuteKw("res.partner", "create", argsCreate)
        
        .IsOk resCreate.Exists("result"), "Create method failed"
        
        Dim newId As Long
        newId = CLng(resCreate("result"))
        .IsOk (newId > 0), "Created ID is invalid"

        ' --- 2. SearchRead ---
        Dim domainCol As New Collection
        Dim cond As New Collection
        cond.Add "id"
        cond.Add "="
        cond.Add newId
        domainCol.Add cond
        
        Dim searchArgs As New Collection
        searchArgs.Add domainCol
        
        Dim searchKwargs As New Dictionary
        Dim fieldsCol As New Collection
        fieldsCol.Add "id"
        fieldsCol.Add "name"
        fieldsCol.Add "email"
        searchKwargs.Add "fields", fieldsCol
        
        Dim resSearch As Dictionary
        Set resSearch = objClient.ExecuteKw("res.partner", "search_read", searchArgs, searchKwargs)
        
        .IsOk resSearch.Exists("result"), "SearchRead method failed"
        
        Dim records As Collection
        Set records = resSearch("result")
        .IsEqual records.Count, 1, "Expected exactly 1 record"

        ' --- 3. Unlink ---
        Dim idsList As New Collection
        idsList.Add newId
        
        Dim unlinkArgs As New Collection
        unlinkArgs.Add idsList
        
        Dim resUnlink As Dictionary
        Set resUnlink = objClient.ExecuteKw("res.partner", "unlink", unlinkArgs)
        
        .IsOk resUnlink("result"), "Unlink method failed"
    End With
End Sub
