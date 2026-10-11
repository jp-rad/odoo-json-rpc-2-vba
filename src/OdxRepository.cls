VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdxRepository"
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

''
' ==============================================================================
' Class: OdxRepository
' ==============================================================================
' Description:
'   Acts as a unified wrapper and repository facade for interacting with Odoo.
'   It delegates repository operations dynamically to either an underlying
'   `OdRepository` or `Od2Repository` instance based on the initialization type,
'   providing a common, abstracted interface for client code.
'
' Features:
'   - Dynamic delegation of Odoo RPC model operations (Search, Read, Write, etc.)
'   - Unified initialization and validation handling
'   - Model schema caching and deep-copy utilities
' ==============================================================================
''

Private Const ERR_UNINITIALIZED As Long = vbObjectError + 1003
Private Const ERR_API_EXECUTION  As Long = vbObjectError + 1004

Private m_Type1 As OdRepository
Private m_Type2 As Od2Repository
Private m_SchemaCache As Dictionary

' ------------------------------------------------------------------------------
' Initialization & Validation
' ------------------------------------------------------------------------------

''
' Initializes the wrapper with either an OdRepository or Od2Repository instance.
'
' @param aRepository Repository instance (OdRepository or Od2Repository)
''
Public Sub Initialize(aRepository As Object)
    If TypeOf aRepository Is OdRepository Then
        Set m_Type1 = aRepository
    ElseIf TypeOf aRepository Is Od2Repository Then
        Set m_Type2 = aRepository
    Else
        Err.Raise ERR_API_EXECUTION, "OdxRepository", "Provided repository instance is invalid or Nothing."
    End If
    Set m_SchemaCache = New Dictionary
End Sub

Private Property Get IsType1() As Boolean
    IsType1 = Not (m_Type1 Is Nothing)
End Property

Private Property Get IsType2() As Boolean
    IsType2 = Not (m_Type2 Is Nothing)
End Property

''
' Checks if the repository has been initialized with a valid instance.
'
' @return Boolean True if initialized, False otherwise
''
Public Property Get IsInitialized() As Boolean
    IsInitialized = IsType1 Or IsType2
End Property

''
' Ensures the repository is initialized before executing operations.
''
Private Sub CheckClientInitialized()
    If Not Me.IsInitialized Then
        Err.Raise ERR_UNINITIALIZED, "OdxRepository", "Repository is not initialized with a repository instance."
    End If
End Sub

' ------------------------------------------------------------------------------
' Repository & Schema Operations
' ------------------------------------------------------------------------------

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
    CheckClientInitialized
    If IsType1 Then
        Set GetModelFields = m_Type1.GetModelFields(aModel, aFields, aAttributes)
    ElseIf IsType2 Then
        Set GetModelFields = m_Type2.GetModelFields(aModel, aFields, aAttributes)
    End If
End Function

''
' Retrieves field definitions for a model, utilizing an internal cache.
'
' @param aModelName Target Odoo model name
' @return Dictionary Field definitions mapping
''
Public Function FetchModelFields(aModelName As String) As Dictionary
    Dim dicModelFields As Dictionary
    
    If m_SchemaCache.Exists(aModelName) Then
        Set dicModelFields = m_SchemaCache(aModelName)
        GoTo ExitProc
    End If
    
    Set dicModelFields = GetModelFields(aModelName)
    Set m_SchemaCache(aModelName) = dicModelFields

ExitProc:
    Set FetchModelFields = dicModelFields
End Function

''
' Performs a deep copy of a Dictionary object using JSON serialization.
'
' @param dic Target Dictionary to copy
' @return Dictionary Cloned Dictionary instance
''
Private Function DeepCopyDictionary(dic As Dictionary) As Dictionary
    Dim sJson As String
    Debug.Assert dic.CompareMode = BinaryCompare
    
    sJson = WebHelpers.ConvertToJson(dic)
    Set DeepCopyDictionary = WebHelpers.ParseJson(sJson)
    
    Debug.Assert DeepCopyDictionary.CompareMode = dic.CompareMode
End Function

''
' Retrieves a deep-copied schema definition for a specified model.
'
' @param aModelName Target Odoo model name
' @return Dictionary Cloned model fields dictionary
''
Public Function GetModelSchema(aModelName As String) As Dictionary
    Set GetModelSchema = DeepCopyDictionary(FetchModelFields(aModelName))
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
    CheckClientInitialized
    If IsType1 Then
        Set SearchRecord = m_Type1.SearchRecord(aModel, aDomain, aOffset, aLimit, aOrder)
    ElseIf IsType2 Then
        Set SearchRecord = m_Type2.SearchRecord(aModel, aDomain, aOffset, aLimit, aOrder)
    End If
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
    CheckClientInitialized
    If IsType1 Then
        Set ReadRecord = m_Type1.ReadRecord(aModel, aRecordId, aFields)
    ElseIf IsType2 Then
        Set ReadRecord = m_Type2.ReadRecord(aModel, aRecordId, aFields)
    End If
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
    CheckClientInitialized
    If IsType1 Then
        Set SearchReadRecord = m_Type1.SearchReadRecord(aModel, aDomain, aFields, aOffset, aLimit, aOrder)
    ElseIf IsType2 Then
        Set SearchReadRecord = m_Type2.SearchReadRecord(aModel, aDomain, aFields, aOffset, aLimit, aOrder)
    End If
End Function

''
' Creates a new record on the specified model.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aValues Dictionary or Collection of field name-value pairs for the new record
' @return Dictionary {"result": Created Record ID (Long) or Collection of IDs}
''
Public Function CreateRecord(aModel As String, aValues As Variant) As Dictionary
    CheckClientInitialized
    If IsType1 Then
        Set CreateRecord = m_Type1.CreateRecord(aModel, aValues)
    ElseIf IsType2 Then
        Set CreateRecord = m_Type2.CreateRecord(aModel, aValues)
    End If
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
    CheckClientInitialized
    If IsType1 Then
        Set WriteRecord = m_Type1.WriteRecord(aModel, aRecordId, aValues)
    ElseIf IsType2 Then
        Set WriteRecord = m_Type2.WriteRecord(aModel, aRecordId, aValues)
    End If
End Function

''
' Deletes record(s) by ID.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aRecordId Single Long ID or Collection of Long IDs
' @return Dictionary {"result": Boolean True on success}
''
Public Function UnlinkRecord(aModel As String, aRecordId As Variant) As Dictionary
    CheckClientInitialized
    If IsType1 Then
        Set UnlinkRecord = m_Type1.UnlinkRecord(aModel, aRecordId)
    ElseIf IsType2 Then
        Set UnlinkRecord = m_Type2.UnlinkRecord(aModel, aRecordId)
    End If
End Function

''
' Retrieves the current authenticated User ID.
'
' @return Long User ID
''
Public Function GetCurrentUserId() As Long
    CheckClientInitialized
    If IsType1 Then
        GetCurrentUserId = m_Type1.GetCurrentUserId()
    ElseIf IsType2 Then
        GetCurrentUserId = m_Type2.GetCurrentUserId()
    End If
End Function
