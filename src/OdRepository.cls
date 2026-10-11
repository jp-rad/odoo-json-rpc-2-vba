VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

' ==============================================================================
' Class: OdRepository
' Purpose: Repository abstraction layer for Odoo 16 JSON-RPC API operations
' Architecture Note: Wraps OdClient to provide higher-level ORM-style methods.
' Specification:
'   - On success: Returns a Dictionary containing {"result": <payload>}
'   - On error: Raises a runtime exception (Err.Raise) with error details
' ==============================================================================

Private m_client As OdClient
Private m_userId As Long

' ------------------------------------------------------------------------------
' Initialization & Validation
' ------------------------------------------------------------------------------

Private Sub Class_Initialize()
    m_userId = 0
End Sub

Private Sub Class_Terminate()
    Set m_client = Nothing
End Sub

''
' Initializes the repository with an authenticated or ready-to-use OdClient.
'
' @param aClient Authenticated or initialized OdClient instance
'
' @return OdRepository Initialized OdRepository instance
''
Public Function Initialize(aClient As OdClient) As OdRepository
    If aClient Is Nothing Then
        Err.Raise ERR_API_EXECUTION, "OdRepository", "Provided client instance is Nothing."
    End If
    Set m_client = aClient
    m_userId = 0
    
    Set Initialize = Me
End Function

''
' Check if the repository has been initialized
''
Public Property Get IsInitialized() As Boolean
    IsInitialized = Not (m_client Is Nothing)
End Property

''
' Ensures client is initialized before executing operations
''
Private Sub CheckClientInitialized()
    If Not Me.IsInitialized Then
        Err.Raise ERR_UNINITIALIZED, "OdRepository", "Repository is not initialized with an OdClient instance."
    End If
End Sub

' ------------------------------------------------------------------------------
' Private Helper Methods
' ------------------------------------------------------------------------------

''
' Wraps a payload into a Dictionary with "result" key
'
' @param aPayload Value or object to encapsulate
' @return Dictionary {"result": aPayload}
''
Private Function WrapResult(aPayload As Variant) As Dictionary
    Dim resDict As New Dictionary
    
    If IsObject(aPayload) Then
        Set resDict("result") = aPayload
    Else
        resDict("result") = aPayload
    End If
    Set WrapResult = resDict
End Function

' ------------------------------------------------------------------------------
' Repository Operations
' ------------------------------------------------------------------------------

''
' Executes an arbitrary model-level method on Odoo via JSON-RPC execute_kw.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aMethod Name of the model method to invoke (e.g., "search_read")
' @param aArgs Positional arguments for the method (Collection)
' @param aKwargs keyword arguments for the method (Dictionary)
' @return Dictionary {"result": Result payload}
''
Public Function ExecuteModelMethod( _
    aModel As String, _
    aMethod As String, _
    aArgs As Collection, _
    aKwargs As Dictionary _
) As Dictionary
    CheckClientInitialized
    
    ' Invoke OdClient execute_kw method (errors are raised directly via Err.Raise from client)
    Set ExecuteModelMethod = m_client.ExecuteKw(aModel, aMethod, aArgs, aKwargs)
End Function

''
' Retrieves field definitions and attributes for a specified model via fields_get.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aFields Optional Collection of field names to retrieve
' @param aAttributes Optional Collection of attribute names to return per field
' @return Dictionary {"result": Field definitions mapping (Dictionary)}
''
Public Function GetModelFields( _
    aModel As String, _
    Optional aFields As Collection, _
    Optional aAttributes As Collection _
) As Dictionary
    Dim argsList As New Collection
    Dim argsDict As New Dictionary
    
    If Not aFields Is Nothing Then
        If aFields.Count > 0 Then argsDict.Add "allfields", aFields
    End If
    
    If Not aAttributes Is Nothing Then
        If aAttributes.Count > 0 Then argsDict.Add "attributes", aAttributes
    End If
    
    Set GetModelFields = ExecuteModelMethod(aModel, "fields_get", argsList, argsDict)
End Function

''
' Searches records matching domain criteria and returns a list of matching record IDs.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aDomain Optional Collection of domain search conditions
' @param aOffset Optional Long specifying number of records to skip (default: 0)
' @param aLimit Optional Long specifying maximum number of records to return (default: 0 = unlimited)
' @param aOrder Optional String specifying sort order (e.g., "name ascii, id desc")
' @return Dictionary {"result": Collection of record IDs}
''
Public Function SearchRecord( _
    aModel As String, _
    Optional aDomain As Collection, _
    Optional aOffset As Long = 0, _
    Optional aLimit As Long = 0, _
    Optional aOrder As String = "" _
) As Dictionary
    Dim argsList As New Collection
    Dim argsDict As New Dictionary
    
    If aDomain Is Nothing Then
        argsList.Add New Collection
    Else
        argsList.Add aDomain
    End If
    
    If aOffset > 0 Then argsDict.Add "offset", aOffset
    If aLimit > 0 Then argsDict.Add "limit", aLimit
    If Len(aOrder) > 0 Then argsDict.Add "order", aOrder
    
    Set SearchRecord = ExecuteModelMethod(aModel, "search", argsList, argsDict)
End Function

''
' Reads specific field values for given record ID(s).
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aRecordId Single Long ID or Collection of Long IDs
' @param aFields Optional Collection of field names to read
' @return Dictionary {"result": Collection of record dictionaries}
''
Public Function ReadRecord( _
    aModel As String, _
    aRecordId As Variant, _
    Optional aFields As Collection _
) As Dictionary
    Dim argsList As New Collection
    Dim argsDict As New Dictionary

    If TypeOf aRecordId Is Collection Then
        argsList.Add aRecordId
    Else
        Dim idList As New Collection
        idList.Add aRecordId
        argsList.Add idList
    End If
    If Not aFields Is Nothing Then
        If aFields.Count > 0 Then argsDict.Add "fields", aFields
    End If
    
    Set ReadRecord = ExecuteModelMethod(aModel, "read", argsList, argsDict)
End Function

''
' Searches and reads record field values directly in a single RPC call.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aDomain Optional Collection of domain search conditions
' @param aFields Optional Collection of field names to read
' @param aOffset Optional Long specifying number of records to skip (default: 0)
' @param aLimit Optional Long specifying maximum number of records to return (default: 0 = unlimited)
' @param aOrder Optional String specifying sort order (e.g., "id desc")
' @return Dictionary {"result": Collection of record Dictionaries}
''
Public Function SearchReadRecord( _
    aModel As String, _
    Optional aDomain As Collection, _
    Optional aFields As Collection, _
    Optional aOffset As Long = 0, _
    Optional aLimit As Long = 0, _
    Optional aOrder As String = "" _
) As Dictionary
    Dim argsList As New Collection
    Dim argsDict As New Dictionary
    
    If aDomain Is Nothing Then
        argsList.Add New Collection
    Else
        argsList.Add aDomain
    End If
    
    If Not aFields Is Nothing Then
        If aFields.Count > 0 Then argsDict.Add "fields", aFields
    End If
    If aOffset > 0 Then argsDict.Add "offset", aOffset
    If aLimit > 0 Then argsDict.Add "limit", aLimit
    If Len(aOrder) > 0 Then argsDict.Add "order", aOrder
    
    Set SearchReadRecord = ExecuteModelMethod(aModel, "search_read", argsList, argsDict)
End Function

''
' Creates a new record on the specified model.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aValues Dictionary or Collection of field name-value pairs for the new record
' @return Dictionary {"result": Created Record ID (Long) or Collection of IDs}
''
Public Function CreateRecord(aModel As String, aValues As Variant) As Dictionary
    Dim argsList As New Collection
    Dim argsDict As New Dictionary
    
    If TypeOf aValues Is Collection Then
        Set argsList = aValues
    Else
        argsList.Add aValues
    End If
    
    Set CreateRecord = ExecuteModelMethod(aModel, "create", argsList, argsDict)
End Function

''
' Updates an existing record by ID with new key-value pairs.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aRecordId Single Long ID or Collection of Long IDs
' @param aValues Dictionary of field name-value pairs to update
' @return Dictionary {"result": Boolean True on success}
''
Public Function WriteRecord(aModel As String, aRecordId As Variant, aValues As Dictionary) As Dictionary
    Dim argsList As New Collection
    Dim argsDict As New Dictionary
    
    If TypeOf aRecordId Is Collection Then
        Set argsList = aRecordId
    Else
        argsList.Add aRecordId
    End If
    argsDict.Add "vals", aValues
    
    Set WriteRecord = ExecuteModelMethod(aModel, "write", argsList, argsDict)
End Function

''
' Deletes record(s) by ID.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aRecordId Single Long ID or Collection of Long IDs
' @return Dictionary {"result": Boolean True on success}
''
Public Function UnlinkRecord(aModel As String, aRecordId As Variant) As Dictionary
    Dim argsList As New Collection
    Dim argsDict As New Dictionary
    
    If TypeOf aRecordId Is Collection Then
        Set argsList = aRecordId
    Else
        argsList.Add aRecordId
    End If
    
    Set UnlinkRecord = ExecuteModelMethod(aModel, "unlink", argsList, argsDict)
End Function

''
' Retrieve current authenticated User ID by RPC call to res.users/context_get.
' Caches the retrieved User ID after the first successful call.
'
' @return Long User ID
''
Public Function GetCurrentUserId() As Long
    If m_userId > 0 Then
        GetCurrentUserId = m_userId
        Exit Function
    End If
    
    Dim argsList As New Collection
    Dim argsDict As New Dictionary
        
    Dim resDict As Dictionary
    Set resDict = ExecuteModelMethod("res.users", "context_get", argsList, argsDict)
    
    If Not resDict Is Nothing Then
        If resDict.Exists("result") Then
            If TypeOf resDict("result") Is Dictionary Then
                Dim resultDict As Dictionary
                Set resultDict = resDict("result")
                If resultDict.Exists("uid") Then
                    m_userId = CLng(resultDict("uid"))
                    GetCurrentUserId = m_userId
                    Exit Function
                End If
            End If
        End If
    End If
    
    Err.Raise ERR_API_EXECUTION, "OdRepository", "Failed to retrieve authenticated user ID from Odoo response."
End Function
