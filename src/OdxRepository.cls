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
'   - Model schema caching utilities
' ==============================================================================
''

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
'
' @return OdxRepository Initialized OdxRepository instance
''
Public Function Initialize(aRepository As Object) As OdxRepository
    If TypeOf aRepository Is OdRepository Then
        Set m_Type1 = aRepository
    ElseIf TypeOf aRepository Is Od2Repository Then
        Set m_Type2 = aRepository
    Else
        Err.Raise ERR_API_EXECUTION, "OdxRepository", "Provided repository instance is invalid or Nothing."
    End If
    Set m_SchemaCache = New Dictionary
    
    Set Initialize = Me
End Function

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
' Executes an arbitrary model-level method on Odoo via the bound repository.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aMethod Name of the model method to invoke (e.g., "search_read", custom methods)
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
    If IsType1 Then
        Set ExecuteModelMethod = m_Type1.ExecuteModelMethod(aModel, aMethod, New Collection, aParams)
    ElseIf IsType2 Then
        Set ExecuteModelMethod = m_Type2.ExecuteModelMethod(aModel, aMethod, aParams, aExpectPrimitive)
    End If
End Function

''
' Executes an arbitrary model-level method and returns the object/dictionary result directly.
'
' @param aModel Target Odoo model name
' @param aMethod Name of the model method to invoke
' @param aParams Optional Dictionary of RPC arguments
' @return Object (Dictionary or Collection) extracted from the "result" key
''
Public Function ExecuteModelMethodObject( _
    aModel As String, _
    aMethod As String, _
    Optional aParams As Dictionary _
) As Object
    Dim dicResponse As Dictionary
    Set dicResponse = ExecuteModelMethod(aModel, aMethod, aParams, False)
    
    If Not dicResponse Is Nothing Then
        If dicResponse.Exists("result") Then
            If IsObject(dicResponse("result")) Then
                Set ExecuteModelMethodObject = dicResponse("result")
            End If
        End If
    End If
End Function

''
' Executes an arbitrary model-level method and returns the primitive result directly.
'
' @param aModel Target Odoo model name
' @param aMethod Name of the model method to invoke
' @param aParams Optional Dictionary of RPC arguments
' @return Variant Primitive value (Long, String, Boolean, Double, etc.) extracted from the "result" key
''
Public Function ExecuteModelMethodPrimitive( _
    aModel As String, _
    aMethod As String, _
    Optional aParams As Dictionary _
) As Variant
    Dim dicResponse As Dictionary
    Set dicResponse = ExecuteModelMethod(aModel, aMethod, aParams, True)
    
    If Not dicResponse Is Nothing Then
        If dicResponse.Exists("result") Then
            If IsObject(dicResponse("result")) Then
                Set ExecuteModelMethodPrimitive = dicResponse("result")
            Else
                ExecuteModelMethodPrimitive = dicResponse("result")
            End If
        End If
    End If
End Function

''
' Retrieves field definitions and attributes for a specified model via fields_get.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aFields Optional Collection of field names to retrieve
' @param aAttributes Optional Collection of attribute names to return per field
' @return Dictionary Field definitions mapping
''
Public Function GetModelFields( _
    aModel As String, _
    Optional aFields As Collection, _
    Optional aAttributes As Collection _
) As Dictionary
    CheckClientInitialized
    If IsType1 Then
        Set GetModelFields = m_Type1.GetModelFields(aModel, aFields, aAttributes)("result")
    ElseIf IsType2 Then
        Set GetModelFields = m_Type2.GetModelFields(aModel, aFields, aAttributes)("result")
    End If
End Function

''
' Retrieves model schema field definitions, utilizing an internal cache.
'
' @param aModelName Target Odoo model name
' @return Dictionary Field definitions mapping
''
Public Function ModelSchema(aModelName As String) As Dictionary
    Dim dicModelFields As Dictionary
    
    If m_SchemaCache.Exists(aModelName) Then
        Set dicModelFields = m_SchemaCache(aModelName)
        GoTo ExitProc
    End If
    
    Set dicModelFields = GetModelFields(aModelName)
    Set m_SchemaCache(aModelName) = dicModelFields

ExitProc:
    Set ModelSchema = dicModelFields
End Function

''
' Searches records matching domain criteria and returns a list of matching record IDs.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aDomain Optional Collection of domain search conditions
' @param aOffset Optional Long specifying number of records to skip (default: 0)
' @param aLimit Optional Long specifying maximum number of records to return (default: 0 = unlimited)
' @param aOrder Optional String specifying sort order (e.g., "name ascii, id desc")
' @return Collection Collection of Record Long IDs}
''
Public Function SearchRecord( _
    aModel As String, _
    Optional aDomain As Collection, _
    Optional aOffset As Long = 0, _
    Optional aLimit As Long = 0, _
    Optional aOrder As String = "" _
) As Collection
    CheckClientInitialized
    If IsType1 Then
        Set SearchRecord = m_Type1.SearchRecord(aModel, aDomain, aOffset, aLimit, aOrder)("result")
    ElseIf IsType2 Then
        Set SearchRecord = m_Type2.SearchRecord(aModel, aDomain, aOffset, aLimit, aOrder)("result")
    End If
End Function

''
' Reads specific field values for given record ID(s).
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aRecordId Single Long ID or Collection of Long IDs
' @param aFields Optional Collection of field names to read
' @return Collection Collection of Record Dictionaries
''
Public Function ReadRecord( _
    aModel As String, _
    aRecordId As Variant, _
    Optional aFields As Collection _
) As Collection
    CheckClientInitialized
    If IsType1 Then
        Set ReadRecord = m_Type1.ReadRecord(aModel, aRecordId, aFields)("result")
    ElseIf IsType2 Then
        Set ReadRecord = m_Type2.ReadRecord(aModel, aRecordId, aFields)("result")
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
' @return Collection Collection of Record Dictionaries
''
Public Function SearchReadRecord( _
    aModel As String, _
    Optional aDomain As Collection, _
    Optional aFields As Collection, _
    Optional aOffset As Long = 0, _
    Optional aLimit As Long = 0, _
    Optional aOrder As String = "" _
) As Collection
    CheckClientInitialized
    If IsType1 Then
        Set SearchReadRecord = m_Type1.SearchReadRecord(aModel, aDomain, aFields, aOffset, aLimit, aOrder)("result")
    ElseIf IsType2 Then
        Set SearchReadRecord = m_Type2.SearchReadRecord(aModel, aDomain, aFields, aOffset, aLimit, aOrder)("result")
    End If
End Function

''
' Creates a new record on the specified model.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aValues Dictionary or Collection of field name-value pairs for the new record
' @return Collection Collection of Record Long IDs
''
Public Function CreateRecord(aModel As String, aValues As Variant) As Collection
    CheckClientInitialized
    If IsType1 Then
        Set CreateRecord = m_Type1.CreateRecord(aModel, aValues)("result")
    ElseIf IsType2 Then
        Set CreateRecord = m_Type2.CreateRecord(aModel, aValues)("result")
    End If
End Function

''
' Updates an existing record by ID with new key-value pairs.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aRecordId Single Long ID or Collection of Long IDs
' @param aValues Dictionary of field name-value pairs to update
' @return Boolean True on success
''
Public Function WriteRecord(aModel As String, aRecordId As Variant, aValues As Dictionary) As Boolean
    CheckClientInitialized
    If IsType1 Then
        WriteRecord = m_Type1.WriteRecord(aModel, aRecordId, aValues)("result")
    ElseIf IsType2 Then
        WriteRecord = m_Type2.WriteRecord(aModel, aRecordId, aValues)("result")
    End If
End Function

''
' Deletes record(s) by ID.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aRecordId Single Long ID or Collection of Long IDs
' @return Boolean True on success
''
Public Function UnlinkRecord(aModel As String, aRecordId As Variant) As Boolean
    CheckClientInitialized
    If IsType1 Then
        UnlinkRecord = m_Type1.UnlinkRecord(aModel, aRecordId)("result")
    ElseIf IsType2 Then
        UnlinkRecord = m_Type2.UnlinkRecord(aModel, aRecordId)("result")
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
