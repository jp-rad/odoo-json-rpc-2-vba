VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Od2Repository"
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
' Class: Od2Repository
' Purpose: Repository abstraction layer for Odoo 19 External JSON-2 API operations
' Architecture Note: Wraps Od2Client to provide higher-level ORM-style methods.
' Specification:
'   - On success: Returns a Dictionary containing {"result": <payload>}
'   - On error: Raises a runtime exception (Err.Raise) with error details
' ==============================================================================

Private m_client As Od2Client
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
' Initializes the repository with an authenticated or ready-to-use Od2Client.
'
' @param aClient Authenticated or initialized Od2Client instance
'
' @return Od2Repository Initialized Od2Repository instance
''
Public Function Initialize(aClient As Od2Client) As Od2Repository
    If aClient Is Nothing Then
        Err.Raise ERR_API_EXECUTION, "Od2Repository", "Provided client instance is Nothing."
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
        Err.Raise ERR_UNINITIALIZED, "Od2Repository", "Repository is not initialized with an Od2Client instance."
    End If
End Sub

' ------------------------------------------------------------------------------
' Repository Operations
' ------------------------------------------------------------------------------

''
' Executes an arbitrary model-level method on Odoo via the bound client.
' Allows calling custom model methods or action buttons directly from the repository layer.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aMethod Name of the model method to invoke (e.g., "search_read")
' @param aParams Optional Dictionary of RPC arguments
' @param aExpectPrimitive Optional Boolean specifying whether a primitive/single value is expected
' @return Dictionary {"result": Result payload}
''
Public Function ExecuteModelMethod( _
    aModel As String, _
    aMethod As String, _
    Optional aParams As Dictionary, _
    Optional aExpectPrimitive As Boolean = False _
) As Dictionary
    CheckClientInitialized
    
    ' Od2Client handles JSON-2 execution, error raising, and primitive extraction
    Set ExecuteModelMethod = m_client.ExecuteModelMethod(aModel, aMethod, aParams, aExpectPrimitive)
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
    Dim paramsDict As New Dictionary
    
    If Not aFields Is Nothing Then
        If aFields.Count > 0 Then paramsDict.Add "allfields", aFields
    End If
    
    If Not aAttributes Is Nothing Then
        If aAttributes.Count > 0 Then paramsDict.Add "attributes", aAttributes
    End If
    
    Set GetModelFields = ExecuteModelMethod(aModel, "fields_get", paramsDict)
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
    Dim paramsDict As New Dictionary
    
    If Not aDomain Is Nothing Then
        paramsDict.Add "domain", aDomain
    Else
        paramsDict.Add "domain", New Collection
    End If
    
    If aOffset > 0 Then paramsDict.Add "offset", aOffset
    If aLimit > 0 Then paramsDict.Add "limit", aLimit
    If Len(aOrder) > 0 Then paramsDict.Add "order", aOrder
    
    Set SearchRecord = ExecuteModelMethod(aModel, "search", paramsDict)
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
    Dim paramsDict As New Dictionary

    Dim idList As Collection
    If TypeOf aRecordId Is Collection Then
        Set idList = aRecordId
    Else
        Set idList = New Collection
        idList.Add aRecordId
    End If
    paramsDict.Add "ids", idList
    If Not aFields Is Nothing Then
        If aFields.Count > 0 Then paramsDict.Add "fields", aFields
    End If
    
    Set ReadRecord = ExecuteModelMethod(aModel, "read", paramsDict)
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
    Dim paramsDict As New Dictionary
    
    If Not aDomain Is Nothing Then
        paramsDict.Add "domain", aDomain
    Else
        paramsDict.Add "domain", New Collection
    End If
    If Not aFields Is Nothing Then
        If aFields.Count > 0 Then paramsDict.Add "fields", aFields
    End If
    If aOffset > 0 Then paramsDict.Add "offset", aOffset
    If aLimit > 0 Then paramsDict.Add "limit", aLimit
    If Len(aOrder) > 0 Then paramsDict.Add "order", aOrder
    
    Set SearchReadRecord = ExecuteModelMethod(aModel, "search_read", paramsDict)
End Function

''
' Creates a new record on the specified model.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aValues Dictionary or Collection of field name-value pairs for the new record
' @return Dictionary {"result": Created Record ID (Long) or Collection of IDs}
''
Public Function CreateRecord(aModel As String, aValues As Variant) As Dictionary
    Dim paramsDict As New Dictionary
    
    Dim valsList As Collection
    If TypeOf aValues Is Collection Then
        Set valsList = aValues
    Else
        Set valsList = New Collection
        valsList.Add aValues
    End If
    paramsDict.Add "vals_list", valsList
    
    Set CreateRecord = ExecuteModelMethod(aModel, "create", paramsDict)
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
    Dim paramsDict As New Dictionary
    
    Dim idList As Collection
    If TypeOf aRecordId Is Collection Then
        Set idList = aRecordId
    Else
        Set idList = New Collection
        idList.Add aRecordId
    End If
    paramsDict.Add "ids", idList
    paramsDict.Add "vals", aValues
    
    Set WriteRecord = ExecuteModelMethod(aModel, "write", paramsDict, True)
End Function

''
' Deletes record(s) by ID.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aRecordId Single Long ID or Collection of Long IDs
' @return Dictionary {"result": Boolean True on success}
''
Public Function UnlinkRecord(aModel As String, aRecordId As Variant) As Dictionary
    Dim paramsDict As New Dictionary
    
    Dim idList As Collection
    If TypeOf aRecordId Is Collection Then
        Set idList = aRecordId
    Else
        Set idList = New Collection
        idList.Add aRecordId
    End If
    paramsDict.Add "ids", idList
    
    Set UnlinkRecord = ExecuteModelMethod(aModel, "unlink", paramsDict, True)
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
    
    Dim resDict As Dictionary
    Set resDict = ExecuteModelMethod("res.users", "context_get")
    
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
    
    Err.Raise ERR_API_EXECUTION, "Od2Repository", "Failed to retrieve authenticated user ID from Odoo response."
End Function
