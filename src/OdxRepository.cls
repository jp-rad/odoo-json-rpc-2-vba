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

' ------------------------------------------------------------------------------
' Module: OdxRepository.cls
' Purpose:
'   Acts as a unified wrapper and repository facade for interacting with Odoo.
'   It delegates repository operations dynamically to either an underlying
'   `OdRepository` or `Od2Repository` instance based on the initialization type,
'   providing a common, abstracted interface for client code.
' ------------------------------------------------------------------------------

Private mType1 As OdRepository
Private mType2 As Od2Repository
Private mSchemaCache As Dictionary

' ==========================================
' Initialization & Validation Methods
' ==========================================

''
' Initializes the wrapper with either an OdRepository or Od2Repository instance.
'
' @param aRepository Repository instance (OdRepository or Od2Repository)
' @return OdxRepository Initialized OdxRepository instance
Public Function Initialize(ByRef aRepository As Object) As OdxRepository
    If TypeOf aRepository Is OdRepository Then
        Set mType1 = aRepository
    ElseIf TypeOf aRepository Is Od2Repository Then
        Set mType2 = aRepository
    Else
        Err.Raise CERR_API_EXECUTION, "OdxRepository", "Provided repository instance is invalid or Nothing."
    End If
    Set mSchemaCache = New Dictionary
    
    Set Initialize = Me
End Function

''
' Checks if the repository is type 1.
'
' @return Boolean True if type 1, False otherwise
Private Property Get IsType1() As Boolean
    IsType1 = Not (mType1 Is Nothing)
End Property

''
' Checks if the repository is type 2.
'
' @return Boolean True if type 2, False otherwise
Private Property Get IsType2() As Boolean
    IsType2 = Not (mType2 Is Nothing)
End Property

''
' Checks if the repository has been initialized with a valid instance.
'
' @return Boolean True if initialized, False otherwise
Public Property Get IsInitialized() As Boolean
    IsInitialized = IsType1 Or IsType2
End Property

''
' Ensures the repository is initialized before executing operations.
Private Sub CheckClientInitialized()
    If Not Me.IsInitialized Then
        Err.Raise CERR_UNINITIALIZED, "OdxRepository", "Repository is not initialized with a repository instance."
    End If
End Sub

' ==========================================
' Repository & Schema Operations
' ==========================================

''
' Executes an arbitrary model-level method on Odoo via the bound repository.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aMethod Name of the model method to invoke (e.g., "search_read", custom methods)
' @param aParams Optional Dictionary of RPC arguments
' @param aExpectPrimitive Optional Boolean specifying whether a primitive/single value is expected
' @return Dictionary {"result": Result payload}
Public Function ExecuteModelMethod( _
    ByRef aModel As String, _
    ByRef aMethod As String, _
    Optional ByVal aParams As Dictionary, _
    Optional ByVal aExpectPrimitive As Boolean = False _
) As Dictionary
    CheckClientInitialized
    If IsType1 Then
        Set ExecuteModelMethod = mType1.ExecuteModelMethod(aModel, aMethod, New Collection, aParams)
    ElseIf IsType2 Then
        Set ExecuteModelMethod = mType2.ExecuteModelMethod(aModel, aMethod, aParams, aExpectPrimitive)
    End If
End Function

''
' Executes an arbitrary model-level method and returns the object/dictionary result directly.
'
' @param aModel Target Odoo model name
' @param aMethod Name of the model method to invoke
' @param aParams Optional Dictionary of RPC arguments
' @return Object (Dictionary or Collection) extracted from the "result" key
Public Function ExecuteModelMethodObject( _
    ByRef aModel As String, _
    ByRef aMethod As String, _
    Optional ByVal aParams As Dictionary _
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
Public Function ExecuteModelMethodPrimitive( _
    ByRef aModel As String, _
    ByRef aMethod As String, _
    Optional ByVal aParams As Dictionary _
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
Public Function GetModelFields( _
    ByRef aModel As String, _
    Optional ByVal aFields As Collection, _
    Optional ByVal aAttributes As Collection _
) As Dictionary
    CheckClientInitialized
    If IsType1 Then
        Set GetModelFields = mType1.GetModelFields(aModel, aFields, aAttributes)("result")
    ElseIf IsType2 Then
        Set GetModelFields = mType2.GetModelFields(aModel, aFields, aAttributes)("result")
    End If
End Function

''
' Retrieves model schema field definitions, utilizing an internal cache.
'
' @param aModelName Target Odoo model name
' @return Dictionary Field definitions mapping
Public Function ModelSchema(ByRef aModelName As String) As Dictionary
    Dim dicModelFields As Dictionary
    
    If mSchemaCache.Exists(aModelName) Then
        Set dicModelFields = mSchemaCache(aModelName)
        GoTo ExitProc
    End If
    
    Set dicModelFields = GetModelFields(aModelName)
    Set mSchemaCache(aModelName) = dicModelFields

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
' @return Collection Collection of Record Long IDs
Public Function SearchRecord( _
    ByRef aModel As String, _
    Optional ByVal aDomain As Collection, _
    Optional ByVal aOffset As Long = 0, _
    Optional ByVal aLimit As Long = 0, _
    Optional ByVal aOrder As String = "" _
) As Collection
    CheckClientInitialized
    If IsType1 Then
        Set SearchRecord = mType1.SearchRecord(aModel, aDomain, aOffset, aLimit, aOrder)("result")
    ElseIf IsType2 Then
        Set SearchRecord = mType2.SearchRecord(aModel, aDomain, aOffset, aLimit, aOrder)("result")
    End If
End Function

''
' Reads specific field values for given record ID(s).
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aRecordId Single Long ID or Collection of Long IDs
' @param aFields Optional Collection of field names to read
' @return Collection Collection of Record Dictionaries
Public Function ReadRecord( _
    ByRef aModel As String, _
    ByRef aRecordId As Variant, _
    Optional ByVal aFields As Collection _
) As Collection
    CheckClientInitialized
    If IsType1 Then
        Set ReadRecord = mType1.ReadRecord(aModel, aRecordId, aFields)("result")
    ElseIf IsType2 Then
        Set ReadRecord = mType2.ReadRecord(aModel, aRecordId, aFields)("result")
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
Public Function SearchReadRecord( _
    ByRef aModel As String, _
    Optional ByVal aDomain As Collection, _
    Optional ByVal aFields As Collection, _
    Optional ByVal aOffset As Long = 0, _
    Optional ByVal aLimit As Long = 0, _
    Optional ByVal aOrder As String = "" _
) As Collection
    CheckClientInitialized
    If IsType1 Then
        Set SearchReadRecord = mType1.SearchReadRecord(aModel, aDomain, aFields, aOffset, aLimit, aOrder)("result")
    ElseIf IsType2 Then
        Set SearchReadRecord = mType2.SearchReadRecord(aModel, aDomain, aFields, aOffset, aLimit, aOrder)("result")
    End If
End Function

''
' Creates a new record on the specified model.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aValues Dictionary or Collection of field name-value pairs for the new record
' @return Collection Collection of Record Long IDs
Public Function CreateRecord(ByRef aModel As String, ByRef aValues As Variant) As Collection
    CheckClientInitialized
    If IsType1 Then
        Dim dicRes As Dictionary
        Set dicRes = mType1.CreateRecord(aModel, aValues)
        If TypeOf dicRes("result") Is Collection Then
            Set CreateRecord = dicRes("result")
        Else
            Set CreateRecord = New Collection
            CreateRecord.Add dicRes("result")
        End If
    ElseIf IsType2 Then
        Set CreateRecord = mType2.CreateRecord(aModel, aValues)("result")
    End If
End Function

''
' Updates an existing record by ID with new key-value pairs.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aRecordId Single Long ID or Collection of Long IDs
' @param aValues Dictionary of field name-value pairs to update
' @return Boolean True on success
Public Function WriteRecord(ByRef aModel As String, ByRef aRecordId As Variant, ByRef aValues As Dictionary) As Boolean
    CheckClientInitialized
    If IsType1 Then
        WriteRecord = mType1.WriteRecord(aModel, aRecordId, aValues)("result")
    ElseIf IsType2 Then
        WriteRecord = mType2.WriteRecord(aModel, aRecordId, aValues)("result")
    End If
End Function

''
' Deletes record(s) by ID.
'
' @param aModel Target Odoo model name (e.g., "res.partner")
' @param aRecordId Single Long ID or Collection of Long IDs
' @return Boolean True on success
Public Function UnlinkRecord(ByRef aModel As String, ByRef aRecordId As Variant) As Boolean
    CheckClientInitialized
    If IsType1 Then
        UnlinkRecord = mType1.UnlinkRecord(aModel, aRecordId)("result")
    ElseIf IsType2 Then
        UnlinkRecord = mType2.UnlinkRecord(aModel, aRecordId)("result")
    End If
End Function

''
' Retrieves the current authenticated User ID.
'
' @return Long User ID
Public Function GetCurrentUserId() As Long
    CheckClientInitialized
    If IsType1 Then
        GetCurrentUserId = mType1.GetCurrentUserId()
    ElseIf IsType2 Then
        GetCurrentUserId = mType2.GetCurrentUserId()
    End If
End Function

' ==========================================
' Odoo External ID Management
' ==========================================

''
' Retrieves Model and Record ID from an External ID.
'
' @param aExternalId External ID string (e.g., "base_module.partner_root_1")
' @return Dictionary containing "model" and "res_id", or Nothing if not found.
Public Function GetModelAndIdFromExternalId(ByRef aExternalId As String) As Dictionary
    Dim sSplitParts() As String
    Dim sModuleName As String
    Dim sObjName As String
    Dim colDomain As New Collection
    Dim colCond1 As New Collection
    Dim colCond2 As New Collection
    Dim colRes As Collection
    Dim dicResult As New Dictionary
    
    sSplitParts = SplitExternalId(aExternalId)
    
    sModuleName = sSplitParts(0)
    sObjName = sSplitParts(1)
    
    ' Query ir.model.data using module and name
    colCond1.Add "module": colCond1.Add "=": colCond1.Add sModuleName
    colDomain.Add colCond1
    
    colCond2.Add "name": colCond2.Add "=": colCond2.Add sObjName
    colDomain.Add colCond2
    
    Dim colFieldsList As New Collection
    colFieldsList.Add "model"
    colFieldsList.Add "res_id"
    
    Set colRes = SearchReadRecord("ir.model.data", colDomain, colFieldsList)
    
    If colRes.Count > 0 Then
        Dim dicRec As Dictionary
        Set dicRec = colRes(1)
        
        dicResult.Add "model", CStr(dicRec("model"))
        dicResult.Add "res_id", CLng(dicRec("res_id"))
        Set GetModelAndIdFromExternalId = dicResult
    End If
End Function

''
' Retrieves an existing External ID from Model and Record ID.
'
' @param aModel Model name (e.g., "res.partner")
' @param aResId Database ID (Long)
' @return External ID string, or empty string if not found.
Public Function GetExternalIdFromModelAndId(ByRef aModel As String, ByRef aResId As Long) As String
    Dim colDomain As New Collection
    Dim colCond1 As New Collection
    Dim colCond2 As New Collection
    Dim colRes As Collection
    Dim colFieldsList As New Collection
        
    GetExternalIdFromModelAndId = ""
    
    colCond1.Add "model": colCond1.Add "=": colCond1.Add aModel
    colDomain.Add colCond1
    
    colCond2.Add "res_id": colCond2.Add "=": colCond2.Add aResId
    colDomain.Add colCond2
    
    colFieldsList.Add "module"
    colFieldsList.Add "name"
    
    Set colRes = SearchReadRecord("ir.model.data", colDomain, colFieldsList)
    
    If colRes.Count > 0 Then
        Dim dicRec As Dictionary
        Set dicRec = colRes(1)
        
        GetExternalIdFromModelAndId = CStr(dicRec("module")) & "." & CStr(dicRec("name"))
    End If
End Function

''
' Registers a new External ID in ir.model.data for a given Model and Record ID (Insert Only).
' - If the Model and Record ID (model + res_id) already has a mapping, an exception is thrown.
'
' @param aModel Model name (e.g., "res.partner")
' @param aResId Database ID (Long)
' @param aExternalId External ID string to register (e.g., "base_module.partner_root_1")
' @return Boolean True on successful registration
Public Function RegisterExternalId(ByRef aModel As String, ByRef aResId As Long, ByRef aExternalId As String) As Boolean
    Dim colFieldsList As New Collection
    Dim colDomain As New Collection
    Dim colCond1 As New Collection
    Dim colCond2 As New Collection
    Dim colRes As Collection
    Dim sSplitParts() As String
    Dim sModuleName As String
    Dim sObjName As String
    Dim dicVals As New Dictionary
        
    colCond1.Add "model": colCond1.Add "=": colCond1.Add aModel
    colDomain.Add colCond1
    
    colCond2.Add "res_id": colCond2.Add "=": colCond2.Add aResId
    colDomain.Add colCond2
    
    colFieldsList.Add "id"
    
    Set colRes = SearchReadRecord("ir.model.data", colDomain, colFieldsList)
    
    ' Raise an exception if a mapping for (model, res_id) already exists
    If colRes.Count > 0 Then
        Err.Raise CERR_API_EXECUTION, "OdxRepository", "Mapping already exists: Model '" & aModel & "' with ID " & aResId & " is already registered with an External ID."
    End If
    
    sSplitParts = SplitExternalId(aExternalId)
    sModuleName = sSplitParts(0)
    sObjName = sSplitParts(1)
        
    dicVals.Add "module", sModuleName
    dicVals.Add "name", sObjName
    dicVals.Add "model", aModel
    dicVals.Add "res_id", aResId
    
    Dim colCreatedRes As Collection
    Set colCreatedRes = CreateRecord("ir.model.data", dicVals)
    
    RegisterExternalId = True
End Function

''
' Upserts records for a given model based on External IDs ("module.name") without throwing exceptions on row-level errors.
'
' @param aModel Model name (e.g., "res.partner")
' @param aValuesDict Nested Dictionary where Key = External ID ("module.name"), Value = Dictionary of field values
' @return Dictionary Key = External ID, Value = Dictionary containing "id" (Long), "action" (String: "created", "updated", or "error"), and optional "error_message"
Public Function UpsertRecordByExternalId(ByRef aModel As String, ByRef aValuesDict As Dictionary) As Dictionary
    Dim vKey As Variant
    Dim sExternalId As String
    Dim dicFieldValues As Dictionary
    Dim dicModelAndId As Dictionary
    Dim nResId As Long
    Dim bWriteSuccess As Boolean
    Dim colCreatedRes As Collection
    Dim bRegSuccess As Boolean
    
    Dim dicResult As New Dictionary
    Dim dicRecordInfo As Dictionary
    
    For Each vKey In aValuesDict.Keys
        sExternalId = CStr(vKey)
        Set dicFieldValues = aValuesDict(sExternalId)
        
        Set dicRecordInfo = New Dictionary
        
        On Error GoTo RecordError
        
        Set dicModelAndId = GetModelAndIdFromExternalId(sExternalId)
        
        If Not dicModelAndId Is Nothing Then
            If dicModelAndId("model") <> aModel Then
                Err.Raise vbObjectError + 101, "OdxRepository", "Model mismatch."
            End If
            
            nResId = dicModelAndId("res_id")
            bWriteSuccess = WriteRecord(aModel, nResId, dicFieldValues)
            
            dicRecordInfo.Add "id", nResId
            dicRecordInfo.Add "action", "updated"
        Else
            Set colCreatedRes = CreateRecord(aModel, dicFieldValues)
            nResId = colCreatedRes(1)
            bRegSuccess = RegisterExternalId(aModel, nResId, sExternalId)
            
            dicRecordInfo.Add "id", nResId
            dicRecordInfo.Add "action", "created"
        End If
        
        GoTo NextRecord

RecordError:
        dicRecordInfo.RemoveAll
        dicRecordInfo.Add "id", 0
        dicRecordInfo.Add "action", "error"
        dicRecordInfo.Add "error_message", Err.Description
        Err.Clear

NextRecord:
        On Error GoTo 0
        dicResult.Add sExternalId, dicRecordInfo
    Next vKey
    
    Set UpsertRecordByExternalId = dicResult
End Function

''
' Unlinks (deletes) records for a given model based on External ID(s).
'
' @param aModel Model name (e.g., "res.partner")
' @param aExternalId String or Collection of Strings ("module.name")
' @return Dictionary Key = External ID, Value = Result Dictionary
Public Function UnlinkRecordByExternalId(ByRef aModel As String, ByRef aExternalId As Variant) As Dictionary
    Dim colIdList As Collection
    
    If TypeOf aExternalId Is Collection Then
        Set colIdList = aExternalId
    Else
        Set colIdList = New Collection
        colIdList.Add aExternalId
    End If
    
    Dim dicResult As New Dictionary
    Dim vItem As Variant
    Dim sExtId As String
    
    For Each vItem In colIdList
        sExtId = CStr(vItem)
        Dim dicRecordInfo As Dictionary
        Dim dicModelAndId As Dictionary
        Dim nResId As Long
        Dim bUnlinkSuccess As Boolean
        
        Set dicRecordInfo = New Dictionary
        Set dicModelAndId = Nothing
        nResId = 0
        bUnlinkSuccess = False
        
        On Error GoTo UnlinkError
        
        Set dicModelAndId = GetModelAndIdFromExternalId(sExtId)
        If Err.Number <> 0 Then GoTo UnlinkError
        
        If dicModelAndId Is Nothing Then
            Err.Raise vbObjectError + 201, "OdxRepository", "External ID not found."
        End If
        
        If dicModelAndId("model") <> aModel Then
            Err.Raise vbObjectError + 202, "OdxRepository", "Model mismatch."
        End If
        
        nResId = dicModelAndId("res_id")
        
        bUnlinkSuccess = UnlinkRecord(aModel, nResId)
        If Not bUnlinkSuccess Then
            Err.Raise vbObjectError + 203, "OdxRepository", "Failed to unlink record."
        End If
        
        dicRecordInfo.Add "action", "unlinked"
        dicRecordInfo.Add "id", nResId
        GoTo NextUnlinkItem

UnlinkError:
        dicRecordInfo.RemoveAll
        dicRecordInfo.Add "action", "error"
        dicRecordInfo.Add "error_message", Err.Description
        Err.Clear

NextUnlinkItem:
        On Error GoTo 0
        dicResult.Add sExtId, dicRecordInfo
    Next vItem
    
    Set UnlinkRecordByExternalId = dicResult
End Function
