VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdxModel"
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
' Class: OdxModel
' ==============================================================================
' Description:
'   Manages an ADODB.Recordset mapped to an Odoo model, handling local data,
'   relational mappings (Many2One, One2Many, Many2Many), field schema generation,
'   and synchronizing data remotely via the unified `OdxRepository`.
' ==============================================================================
''

Private Const CBIT_MODEL          As Long = &H100000  ' bit20 - model property
Private Const CBIT_XML_ID         As Long = &H200000  ' bit21 - external id
Private Const CBIT_PK_ID          As Long = FieldAttributeEnum.adFldKeyColumn
Private Const CBIT_NULLABLE       As Long = FieldAttributeEnum.adFldMayBeNull ' Enable NULLs via adFldMayBeNull even for Odoo's required fields.
Private Const CBIT_FIELD          As Long = FieldAttributeEnum.adFldUpdatable Or FieldAttributeEnum.adFldIsNullable

Private Const CATTR_PK_ID         As Long = CBIT_PK_ID
Private Const CATTR_XML_ID        As Long = CBIT_FIELD Or CBIT_XML_ID
Private Const CATTR_MODEL         As Long = CBIT_FIELD Or CBIT_MODEL
Private Const CATTR_NON_MODEL     As Long = CBIT_FIELD

Private Const CODOO_ATTR_NAME     As String = "name"
Private Const CODOO_ATTR_STRING   As String = "string"
Private Const CODOO_ATTR_TYPE     As String = "type"
Private Const CODOO_ATTR_RELATION As String = "relation"
Private Const CODOO_ATTR_FKEY     As String = "relation_field"   ' foreign key
Private Const CODOO_ATTR_REQUIRED As String = "required"

Private Const CDISPLAY_NAME As String = "_display_name_"
Private Const CEXTERNAL_ID As String = "external_id"

' Odoo IDs are always positive integers, so we use -1 as an impossible ID.
' Filtering by this value guarantees the Recordset becomes empty.
Private Const CMISSING_ID As Long = -1

Private mRs As ADODB.Recordset

Private mRepository As OdxRepository
Private mModelName As String
Private mModelFields As Dictionary

Private mRelations As Dictionary    ' key: relation field name, value: OdxModel
Private mFKeyFieldName As String    ' for one2many: foreign key field name
Private mRelationMap As Collection  ' for many2many: key = CStr(parent id), value = collection of IDs

Private mExternalIdFieldName As String

Private Enum odxModelType
    mtBase
    mtMany2One
    mtOne2Many
    mtMany2Many
End Enum

Private mModelType As odxModelType

Public Property Get IsMany2One() As Boolean
    IsMany2One = mModelType = mtMany2One
End Property

Public Property Get IsOne2Many() As Boolean
    IsOne2Many = mModelType = mtOne2Many
End Property

Public Property Get IsMany2Many() As Boolean
    IsMany2Many = mModelType = mtMany2Many
End Property

Public Property Get ExternalIdFieldName() As String
    ExternalIdFieldName = mExternalIdFieldName
End Property

''
' Initializes the base model view with an OdxRepository and target model name.
'
' @param aRepository OdxRepository instance for Odoo communication
' @param aModelName Target Odoo model name
''
Public Sub InitBase(aRepository As OdxRepository, aModelName As String)
    Set mRepository = aRepository
    Set mRs = CreateNewRecordSet()
    mModelName = aModelName
    Set mRelations = New Dictionary
    mFKeyFieldName = ""
    Set mRelationMap = Nothing
    mModelType = mtBase
    Set mModelFields = mRepository.ModelSchema(aModelName)
End Sub

Public Sub InitMany2One(aRepository As OdxRepository, aModelName As String)
    InitBase aRepository, aModelName
    mModelType = mtMany2One
End Sub

Public Sub InitOne2Many(aRepository As OdxRepository, aModelName As String, aFKeyFieldName As String)
    InitBase aRepository, aModelName
    mFKeyFieldName = aFKeyFieldName
    AddField mFKeyFieldName
    mModelType = mtOne2Many
End Sub

Public Sub InitMany2Many(aRepository As OdxRepository, aModelName As String)
    InitBase aRepository, aModelName
    Set mRelationMap = New Collection
    mModelType = mtMany2Many
End Sub

Public Property Get ModelName() As String
    ModelName = mModelName
End Property

Public Property Get RefMe() As OdxModel
    Set RefMe = Me
End Property

Private Sub SafetyOpenRecordset()
On Error Resume Next
    mRs.Open
End Sub

Private Sub SafetyMoveFirst()
On Error Resume Next
    mRs.MoveFirst
End Sub

Public Property Get Recordset() As ADODB.Recordset
    Set Recordset = mRs
End Property

''
' Adds a field to the local Recordset and sets up related model views if relational.
'
' @param aFieldName Target Odoo field name
' @return OdxModel Related model instance if applicable, or Nothing
''
Public Function AddField(aFieldName As String) As OdxModel
    Dim oRelation As OdxModel
    Set oRelation = AddRecordsetFieldAndRelatedModel(mRs, mModelName, aFieldName)
    If Not oRelation Is Nothing Then
        Set mRelations(aFieldName) = oRelation
        Set AddField = oRelation
    End If
End Function

''
' Adds an external ID field to the local Recordset as a non-model property.
' Raises an exception if the external ID field has already been added.
'
' @param aFieldName Optional field name for the external ID (Default: CEXTERNAL_ID)
' @throws ERR_API_EXECUTION if the external ID field already exists.
''
Public Sub AddExternalIdField(Optional aFieldName As String = CEXTERNAL_ID)
    ' Check if the external ID field has already been added
    If mExternalIdFieldName <> "" Then
        Err.Raise ERR_API_EXECUTION, "AddExternalIdField", "External ID field has already been added: " & mExternalIdFieldName
    End If
    
    ' Append a non-model external ID field to the local Recordset structure
    mRs.Fields.Append Name:=aFieldName, Type:=adVarWChar, DefinedSize:=-1, Attrib:=CATTR_XML_ID
    mExternalIdFieldName = aFieldName
End Sub

''
' Adds a new field to the local disconnected Recordset structure as a non-model property.
'
' @param aFieldName   The name of the field to be added.
' @param aDataType    The ADODB data type (e.g., adInteger, adVarWChar).
' @param aDefinedSize Optional maximum length or size of the field (Default: 0 for types that do not require size).
''
Public Sub AddLocalField(aFieldName As String, aDataType As ADODB.DataTypeEnum, Optional aDefinedSize As Long = 0)
    ' Append a non-model field to the local Recordset structure
    If aDefinedSize = 0 Then
        mRs.Fields.Append Name:=aFieldName, Type:=aDataType, Attrib:=CATTR_NON_MODEL
    Else
        mRs.Fields.Append Name:=aFieldName, Type:=aDataType, DefinedSize:=aDefinedSize, Attrib:=CATTR_NON_MODEL
    End If
End Sub

''
' Clears all local records and resets relational data.
''
Public Sub ClearRecords()
    Dim ds As OdxModel
    Dim v As Variant
    
    SafetyOpenRecordset
    ClearFilter
    
    Do Until mRs.EOF
        mRs.Delete
        mRs.MoveNext
    Loop
    
    For Each v In mRelations.Keys
        Set ds = mRelations(v)
        ds.ClearRecords
    Next
End Sub

Private Sub MoveFirstOrAddNewRecord(aId As Long)
    SafetyMoveFirst
    mRs.Find "id = " & aId

    If mRs.EOF Then
        mRs.AddNew
        mRs("id").Value = aId
    End If
End Sub

Public Sub UpsertMany2One(aId As Long)
    MoveFirstOrAddNewRecord aId
    mRs.Update
End Sub

Public Sub UpsertOne2Many(aId As Long, aFkId As Long)
    MoveFirstOrAddNewRecord aFkId
    mRs(mFKeyFieldName) = aId
    mRs.Update
End Sub

Public Sub UpsertMany2Many(aId As Long, aFkId As Long)
    Dim sId As String
    Dim colIds As Collection
    
    MoveFirstOrAddNewRecord aFkId
    mRs.Update
    
    sId = CStr(aId)
    On Error Resume Next
    Set colIds = mRelationMap(sId)
    On Error GoTo 0
    
    If colIds Is Nothing Then
        Set colIds = New Collection
        mRelationMap.Add colIds, sId
    End If
    colIds.Add aFkId
End Sub

''
' Updates the current record of the local Recordset with a single record dictionary (Values).
' Safely handles object types for vVal and checks null/false beforehand.
'
' @param dicRecord Dictionary of a single record fetched from Odoo
''
Public Sub ValuesToRecord(dicRecord As Dictionary)
    Dim nId As Long
    Dim fld As ADODB.Field
    Dim sFieldName As String
    Dim sFieldType As String
    Dim vFieldValue As Variant
    
    Dim dsRelation As OdxModel
    Dim colValue As Collection
    Dim itemValue As Variant
    Dim nRelatedId As Long
    Dim sRelatedDisplayName As String
    
    nId = dicRecord("id")
    
    For Each fld In mRs.Fields
        sFieldName = fld.Name
        
        ' ID
        If IsOdooPkField(fld) Then
            Debug.Assert fld.Value = nId
            GoTo CONTINUE_FIELD
        End If
        ' External ID
        If IsExternalIdField(fld) Then
            fld.Value = mRepository.GetExternalIdFromModelAndId(mModelName, nId)
            GoTo CONTINUE_FIELD
        End If
        
        ' not odoo fields
        If Not IsOdooField(fld) Then
            GoTo CONTINUE_FIELD
        End If
        
        sFieldType = mModelFields(sFieldName)(CODOO_ATTR_TYPE)
        
        ' Check null / false
        OdJsonRpc.LetOrSetValue OdJsonRpc.ValueOrDefault(dicRecord(sFieldName), False), vFieldValue
        If VarType(vFieldValue) = vbBoolean Then
            If sFieldType <> "boolean" Then
                If vFieldValue = False Then
                    vFieldValue = Null
                End If
            End If
        End If
        
        ' odoo fields
        Select Case sFieldType
        ' Basic fields
        Case "boolean":
            fld.Value = vFieldValue
            
        Case "char":
            fld.Value = vFieldValue
            
        Case "float":
            fld.Value = vFieldValue
            
        Case "integer":
            fld.Value = vFieldValue
            
        ' Advanced fields
        Case "html":
            If IsNull(vFieldValue) Then
                fld.Value = Null
            Else
                fld.Value = OdJsonRpc.ReplaceCrLf(vFieldValue, True)
            End If
            
        Case "monetary":
            fld.Value = vFieldValue
                        
        Case "selection":
            fld.Value = vFieldValue
            
        Case "text":
            If IsNull(vFieldValue) Then
                fld.Value = Null
            Else
                fld.Value = OdJsonRpc.ReplaceCrLf(vFieldValue, True)
            End If
            
        ' Date(time) fields
        Case "date":
            If IsNull(vFieldValue) Then
                fld.Value = Null
            Else
                fld.Value = OdJsonRpc.ParseDate(vFieldValue)
            End If
            
        Case "datetime":
            If IsNull(vFieldValue) Then
                fld.Value = Null
            Else
                fld.Value = OdJsonRpc.ParseUtcDatetime(vFieldValue)
            End If
            
        ' Relational fields
        Case "many2one":
            If IsNull(vFieldValue) Then
                fld.Value = Null
                GetDisplayNameField(mRs, sFieldName).Value = Null
            Else
                Set colValue = vFieldValue
                nRelatedId = colValue(1)
                fld.Value = nRelatedId
                sRelatedDisplayName = colValue(2)
                GetDisplayNameField(mRs, sFieldName).Value = sRelatedDisplayName
                
                If mRelations.Exists(sFieldName) Then
                    Set dsRelation = mRelations(sFieldName)
                    dsRelation.UpsertMany2One nRelatedId
                End If
            End If
            
        Case "one2many":
            If IsNull(vFieldValue) Then
                fld.Value = Null
            Else
                Set colValue = vFieldValue
                fld.Value = "[" & OdJsonRpc.CollectionToCsv(colValue) & "]"
                
                If mRelations.Exists(sFieldName) Then
                    Set dsRelation = mRelations(sFieldName)
                    For Each itemValue In colValue
                        nRelatedId = itemValue
                        dsRelation.UpsertOne2Many nId, nRelatedId
                    Next itemValue
                End If
            End If
            
        Case "many2many":
            If IsNull(vFieldValue) Then
                fld.Value = Null
            Else
                Set colValue = vFieldValue
                fld.Value = "[" & CollectionToCsv(colValue) & "]"
                
                If mRelations.Exists(sFieldName) Then
                    Set dsRelation = mRelations(sFieldName)
                    For Each itemValue In colValue
                        nRelatedId = itemValue
                        dsRelation.UpsertMany2Many nId, nRelatedId
                    Next itemValue
                End If
            End If
            
        ' Pseudo-relational fields
        Case "reference"
            fld.Value = vFieldValue
            
        Case "many2one_reference"
            fld.Value = vFieldValue
            
        ' Binary fields
        Case "binary"
            If IsNull(vFieldValue) Then
                fld.Value = Null
            Else
                ' filename, content, size
                fld.Value = WebHelpers.ConvertToJson(vFieldValue)
            End If
            
        Case Else
            Debug.Print "Not Supported: " & sFieldName & " [" & sFieldType & "]"
            Debug.Assert False
        End Select
        
CONTINUE_FIELD:
    Next fld
End Sub

''
' Converts the current record of the local Recordset into a Dictionary
' suitable for Odoo API requests, mapping explicitly to each Odoo field type.
'
' @return Dictionary representing the record values
''
Public Function RecordToValues() As Dictionary
    Dim dicRecord As Dictionary
    Set dicRecord = New Dictionary
    
    Dim fld As ADODB.Field
    Dim sFieldName As String
    Dim sFieldType As String
    
    For Each fld In mRs.Fields
        sFieldName = fld.Name
        
        ' ID field
        If IsOdooPkField(fld) Then
            GoTo CONTINUE_FIELD
        End If
        
        ' External ID or non-Odoo fields
        If IsExternalIdField(fld) Or Not IsOdooField(fld) Then
            GoTo CONTINUE_FIELD
        End If
        
        ' null
        If IsNull(fld.Value) Then
            dicRecord.Add sFieldName, False
            GoTo CONTINUE_FIELD
        End If
        
        ' Explicit Case mapping symmetric to ValuesToRecord
        sFieldType = mModelFields(sFieldName)(CODOO_ATTR_TYPE)
        Select Case sFieldType
        ' Basic fields
        Case "boolean"
            dicRecord.Add sFieldName, CBool(fld.Value)
            
        Case "char"
            dicRecord.Add sFieldName, CStr(fld.Value)
            
        Case "float"
            dicRecord.Add sFieldName, CDbl(fld.Value)
            
        Case "integer"
            dicRecord.Add sFieldName, CLng(fld.Value)
            
        ' Advanced fields
        Case "html"
            dicRecord.Add sFieldName, ReplaceCrLf(fld.Value, False)
            
        Case "monetary"
            dicRecord.Add sFieldName, CDbl(fld.Value)
            
        Case "selection"
            dicRecord.Add sFieldName, CStr(fld.Value)
            
        Case "text"
            dicRecord.Add sFieldName, ReplaceCrLf(fld.Value, False)
            
        ' Date(time) fields
        Case "date"
            dicRecord.Add sFieldName, OdJsonRpc.FormatDate(fld.Value)
            
        Case "datetime"
            dicRecord.Add sFieldName, OdJsonRpc.ConvertToUtcDatetime(fld.Value)
            
        ' Relational fields
        Case "many2one"
            dicRecord.Add sFieldName, CLng(fld.Value)
            
        Case "one2many"
            ' Recordset typically stores these as CSV/JSON string (e.g., "[1, 2, 3]").
            ' Odoo write expects an array of IDs or command tuples, usually a Collection or Array.
            dicRecord.Add sFieldName, WebHelpers.ParseJson(fld.Value)
            
        Case "many2many"
            ' Recordset typically stores these as CSV/JSON string (e.g., "[1, 2, 3]").
            ' Odoo write expects an array of IDs or command tuples, usually a Collection or Array.
            dicRecord.Add sFieldName, WebHelpers.ParseJson(fld.Value)
            
        ' Pseudo-relational fields
        Case "reference"
            dicRecord.Add sFieldName, fld.Value
            
        Case "many2one_reference"
            dicRecord.Add sFieldName, fld.Value
            
        ' Binary fields
        Case "binary"
            dicRecord.Add sFieldName, WebHelpers.ParseJson(fld.Value)
            
        Case Else
            Debug.Print "Not Supported for RecordToValues: " & sFieldName & " [" & sFieldType & "]"
            Debug.Assert False
        End Select
        
CONTINUE_FIELD:
    Next fld
    
    Set RecordToValues = dicRecord
End Function

''
' Updates or inserts remote result records into the local Recordset and child relations.
'
' @param colResult Collection of record dictionaries fetched from Odoo
''
Private Sub UpsertRecordSet(colResult As Collection)
    Dim dicRecord As Dictionary
    Dim dsRelation As OdxModel
    
    Dim nId As Long
    Dim fld As ADODB.Field
    Dim isNullValue As Boolean
    Dim colValue As Collection
    Dim v As Variant
    
    Dim sFieldType As String
    
    Dim nRelatedId As Long
    Dim sRelatedDisplayName As String
        
    For Each dicRecord In colResult
        nId = dicRecord("id")
        MoveFirstOrAddNewRecord nId
        
        ValuesToRecord dicRecord
        
        mRs.Update
    Next dicRecord
    
    SafetyMoveFirst
    
    ' Fetch remote data for child relations recursively
    For Each v In mRelations.Keys
        Set dsRelation = mRelations(v)
        dsRelation.FetchRecordsRemote
    Next v
End Sub

Private Function GetFieldNames() As Collection
    Dim fld As ADODB.Field
    
    Set GetFieldNames = NewList
    For Each fld In mRs.Fields
        If IsOdooField(fld) Then
            GetFieldNames.Add fld.Name
        End If
    Next fld
End Function

''
' Executes a remote search-read operation via the repository and populates local records.
' - Raises an exception if called on a non-base (relational) model view.
'
' @param aDomain OdFilterDomain search criteria
' @param aOffset Number of records to skip
' @param aLimit Maximum number of records to retrieve
' @param aOrder Sorting order
' @return Long Number of fetched records
' @throws ERR_API_EXECUTION if called on a non-base model.
''
Public Function FetchSearchAndReadRemote( _
    Optional aDomain As Collection, _
    Optional aOffset As Long = 0, _
    Optional aLimit As Long = 0, _
    Optional aOrder As String = "" _
) As Long
    ' Ensure FetchSearchAndReadRemote is only called on the base model view
    If mModelType <> mtBase Then
        Err.Raise ERR_API_EXECUTION, "OdxModel", "FetchSearchAndReadRemote can only be called on a base model, not on relational views."
    End If
    Dim colFieldNames As Collection
    Dim colResult As Collection
    
    ClearRecords
    Set colFieldNames = GetFieldNames()
    
    ' Delegate repository search-read call
    Set colResult = mRepository.SearchReadRecord(mModelName, aDomain, colFieldNames, aOffset, aLimit, aOrder)
    
    UpsertRecordSet colResult
    FetchSearchAndReadRemote = colResult.Count
End Function

''
' Fetches remote records by IDs via the repository and updates local data.
'
' @return Long Number of fetched records
''
Public Function FetchRecordsRemote() As Long
    Dim colIds As Collection
    Dim nId As Long
    Dim colFieldNames As Collection
    Dim colResult As Collection
        
    Set colFieldNames = GetFieldNames()
    Set colIds = NewList
    
    If Not mRs.BOF Then
        mRs.MoveFirst
    End If
    
    Do Until mRs.EOF
        nId = mRs.Fields("id").Value
        colIds.Add nId
        mRs.MoveNext
    Loop
    
    If colIds.Count = 0 Then
        FetchRecordsRemote = 0
        Exit Function
    End If
    
    ' Delegate repository read call
    Set colResult = mRepository.ReadRecord(mModelName, colIds, colFieldNames)
    
    UpsertRecordSet colResult
    FetchRecordsRemote = colResult.Count
End Function

''
' Upserts the current record in the local Recordset to Odoo remotely.
' - If the external ID field is defined and populated, it uses external ID-based upsert.
' - Otherwise, it updates existing records by internal ID or creates new records.
' - For newly created records, it writes back the assigned Odoo internal ID into the local Recordset.
' - Raises an exception if the operation fails or if there is no current record.
'
' @return Dictionary containing "action" ("created" or "updated") and "id" (Long)
' @throws ERR_API_EXECUTION if no current record or if the upsert operation fails.
''
Public Function UpsertCurrentRecordRemote() As Dictionary
    Dim fieldValues As Dictionary
    Dim fld As ADODB.Field
    Dim nId As Long
    Dim extId As String
    Dim hasExternalIdField As Boolean
    Dim resultInfo As New Dictionary
    Dim isSuccess As Boolean
    Dim actionType As String
    Dim extIdDict As Dictionary
    Dim upsertResult As Dictionary
    Dim resInfo As Dictionary
    Dim errMsg As String
    Dim createdIds As Collection
    
    hasExternalIdField = mExternalIdFieldName <> ""
    
    ' Collect field values for the current record
    Set fieldValues = RecordToValues()
    
    ' Determine key and execute Upsert/Update/Create
    If hasExternalIdField Then
        ' ------------------------------------------------------------------
        ' Upsert via External ID
        ' ------------------------------------------------------------------
        extId = GetFieldValueOrDefault(mExternalIdFieldName, "")
        Set extIdDict = New Dictionary
        extIdDict.Add extId, fieldValues
        
        Set upsertResult = mRepository.UpsertRecordByExternalId(mModelName, extIdDict)
        
        isSuccess = False
        If Not upsertResult Is Nothing Then
            Set resInfo = upsertResult(extId)
            If resInfo("action") <> "error" Then
                nId = resInfo("id")
                actionType = resInfo("action")
                mRs("id").Value = nId
                isSuccess = True
            End If
        End If
        
        If Not isSuccess Then
            errMsg = "Failed to upsert record by External ID."
            If Not upsertResult Is Nothing Then
                If upsertResult.Exists(extId) Then
                    If upsertResult(extId).Exists("error_message") Then
                        errMsg = upsertResult(extId)("error_message")
                    End If
                End If
            End If
            Err.Raise ERR_API_EXECUTION, "OdxModel", errMsg
        End If
        
    Else
        ' ------------------------------------------------------------------
        ' Upsert via Internal ID (Create vs Write)
        ' ------------------------------------------------------------------
        nId = mRs("id").Value
        
        If nId > 0 Then
            ' --- Update Existing Record ---
            isSuccess = mRepository.WriteRecord(mModelName, nId, fieldValues)
            actionType = "updated"
            If Not isSuccess Then
                Err.Raise ERR_API_EXECUTION, "OdxModel", "Failed to update record with ID."
            End If
        Else
            ' --- Create New Record ---
            Set createdIds = mRepository.CreateRecord(mModelName, fieldValues)
            If Not createdIds Is Nothing Then
                nId = createdIds(1)
                mRs("id").Value = nId
                actionType = "created"
            Else
                Err.Raise ERR_API_EXECUTION, "OdxModel", "Failed to create new record."
            End If
        End If
    End If
    
    mRs.Update
    
    resultInfo.Add "action", actionType
    resultInfo.Add "id", nId
    Set UpsertCurrentRecordRemote = resultInfo
End Function

''
' Unlinks the current record in the local Recordset from Odoo remotely.
' - If the external ID field is defined and populated, it unlinks using the external ID.
' - Otherwise, it unlinks using the internal record ID.
' - Upon successful remote unlinking, the record is deleted from the local Recordset.
' - Raises an exception if the operation fails or if there is no current record.
'
' @return Dictionary containing "action" ("unlinked") and "id" (Long)
' @throws ERR_API_EXECUTION if no current record or if the unlink operation fails.
''
Public Function UnlinkCurrentRecordRemote() As Dictionary
    Dim extId As String
    Dim nId As Long
    Dim hasExternalIdField As Boolean
    Dim resultInfo As New Dictionary
    Dim isSuccess As Boolean
    Dim unlinkResult As Dictionary
    Dim resInfo As Dictionary
    Dim errMsg As String
        
    hasExternalIdField = mExternalIdFieldName <> ""
    
    If hasExternalIdField Then
        ' ------------------------------------------------------------------
        ' Unlink via External ID
        ' ------------------------------------------------------------------
        extId = GetFieldValueOrDefault(mExternalIdFieldName, "")
        Set unlinkResult = mRepository.UnlinkRecordByExternalId(mModelName, extId)
        
        isSuccess = False
        If Not unlinkResult Is Nothing Then
            Set resInfo = unlinkResult(extId)
            If CStr(resInfo("action")) = "unlinked" Then
                nId = resInfo("id")
                isSuccess = True
            End If
        End If
        
        If Not isSuccess Then
            errMsg = "Failed to unlink record by External ID."
            If Not unlinkResult Is Nothing Then
                If unlinkResult.Exists(extId) Then
                    If unlinkResult(extId).Exists("error_message") Then
                        errMsg = unlinkResult(extId)("error_message")
                    End If
                End If
            End If
            Err.Raise ERR_API_EXECUTION, "OdxModel", errMsg
        End If
        
    Else
        ' ------------------------------------------------------------------
        ' Unlink via Internal ID
        ' ------------------------------------------------------------------
        nId = mRs("id").Value
        isSuccess = mRepository.UnlinkRecord(mModelName, nId)
        If Not isSuccess Then
            Err.Raise ERR_API_EXECUTION, "OdxModel", "Failed to unlink record with ID."
        End If
    End If
        
    resultInfo.Add "action", "unlinked"
    resultInfo.Add "id", nId
    Set UnlinkCurrentRecordRemote = resultInfo
End Function

Public Sub SetFilter(aExpr As String)
    mRs.Filter = aExpr
    SafetyMoveFirst
End Sub

Public Sub SetFilterMany2One(aId As Long)
    Debug.Assert IsMany2One
    mRs.Filter = "id = " & aId
    SafetyMoveFirst
End Sub

Public Sub SetFilterOne2Many(aId As Long)
    Debug.Assert IsOne2Many
    mRs.Filter = mFKeyFieldName & " = " & aId
    SafetyMoveFirst
End Sub

Public Sub SetFilterMany2Many(aId As Long)
    Dim colIds As Collection
    Dim A() As String
    Dim i As Long
    Dim s As String

    Debug.Assert IsMany2Many
    
    On Error Resume Next
    Set colIds = mRelationMap(CStr(aId))
    On Error GoTo 0
    
    If colIds Is Nothing Then
        mRs.Filter = "id = -1"
    Else
        If colIds.Count > 0 Then
            ReDim A(1 To colIds.Count)
            For i = 1 To colIds.Count
                A(i) = "(id = " & CStr(colIds(i)) & ")"
            Next i
            s = Join(A, " OR ")
            mRs.Filter = s
        Else
            mRs.Filter = "id = -1"
        End If
    End If
    SafetyMoveFirst
End Sub

Public Sub ClearFilter()
    mRs.Filter = adFilterNone
    SafetyMoveFirst
End Sub

Private Function GetFieldValueOrDefault(aFieldName As String, aDefault As Variant) As Variant
    Dim fld As ADODB.Field
    
    If mRs.BOF Or mRs.EOF Then
        OdJsonRpc.LetOrSetValue aDefault, GetFieldValueOrDefault
        Exit Function
    End If
    
    Set fld = mRs.Fields(aFieldName)
    If IsNull(fld.Value) Then
        OdJsonRpc.LetOrSetValue aDefault, GetFieldValueOrDefault
    Else
        OdJsonRpc.LetOrSetValue fld.Value, GetFieldValueOrDefault
    End If
End Function

Public Function GetRelatedModel(aFieldName As String, Optional optNotFiltered As Boolean = False) As OdxModel
    If Not mRelations.Exists(aFieldName) Then
        Debug.Assert False
        GoTo ExitProc
    End If
    
    Set GetRelatedModel = mRelations(aFieldName)
    
    If optNotFiltered Then
        GetRelatedModel.ClearFilter
    Else
        With GetRelatedModel
            If .IsMany2One Then
                .SetFilterMany2One GetFieldValueOrDefault(aFieldName, CMISSING_ID)
            ElseIf .IsOne2Many Then
                .SetFilterOne2Many GetFieldValueOrDefault("id", CMISSING_ID)
            ElseIf .IsMany2Many Then
                .SetFilterMany2Many GetFieldValueOrDefault("id", CMISSING_ID)
            Else
                Debug.Assert False
                GoTo ExitProc
            End If
        End With
    End If
ExitProc:
    Exit Function
End Function

Private Function FormatDisplayName(aFieldName As String) As String
    FormatDisplayName = aFieldName & CDISPLAY_NAME
End Function

Private Function IsOdooPkField(fld As ADODB.Field) As Boolean
    IsOdooPkField = 0 <> (fld.attributes And CATTR_PK_ID)
End Function

Private Function IsExternalIdField(fld As ADODB.Field) As Boolean
    IsExternalIdField = 0 <> (fld.attributes And CBIT_XML_ID)
End Function

Private Function IsOdooField(fld As ADODB.Field) As Boolean
    IsOdooField = 0 <> (fld.attributes And CBIT_MODEL)
End Function

Private Function IsOdooRequiredField(fld As ADODB.Field) As Boolean
    IsOdooRequiredField = 0 = (fld.attributes And CBIT_NULLABLE)
End Function

Private Function CreateNewRecordSet() As ADODB.Recordset
    Set CreateNewRecordSet = New ADODB.Recordset
    CreateNewRecordSet.CursorLocation = adUseClient
    With CreateNewRecordSet.Fields
        ' <<PK>> id: Long
        .Append Name:="id", Type:=adInteger, Attrib:=CATTR_PK_ID
    End With
End Function

Private Function AddRecordsetField(rs As ADODB.Recordset, dicModelField As Dictionary) As ADODB.Field
    Dim sFieldName As String
    Dim sFieldType As String
    Dim nFieldAttr As Long
    
    Debug.Assert dicModelField.Exists(CODOO_ATTR_NAME)
    sFieldName = dicModelField(CODOO_ATTR_NAME)
    If dicModelField.Exists(CODOO_ATTR_REQUIRED) Then
        nFieldAttr = IIf(dicModelField(CODOO_ATTR_REQUIRED), 0, CBIT_NULLABLE)
    Else
        nFieldAttr = CBIT_NULLABLE
    End If
    nFieldAttr = nFieldAttr Or CATTR_MODEL
    Debug.Assert dicModelField.Exists(CODOO_ATTR_TYPE)
    sFieldType = dicModelField(CODOO_ATTR_TYPE)
    
    Select Case sFieldType
        ' Basic fields
        Case "boolean"
            rs.Fields.Append Name:=sFieldName, Type:=adBoolean, Attrib:=nFieldAttr
        Case "char"
            rs.Fields.Append Name:=sFieldName, Type:=adVarWChar, DefinedSize:=-1, Attrib:=nFieldAttr
        Case "float"
            rs.Fields.Append Name:=sFieldName, Type:=adDouble, Attrib:=nFieldAttr
        Case "integer"
            rs.Fields.Append Name:=sFieldName, Type:=adInteger, Attrib:=nFieldAttr
        'Advanced fields
        Case "html"
            rs.Fields.Append Name:=sFieldName, Type:=adLongVarWChar, DefinedSize:=-1, Attrib:=nFieldAttr
        Case "monetary"
            rs.Fields.Append Name:=sFieldName, Type:=adCurrency, Attrib:=nFieldAttr
        Case "selection"
            rs.Fields.Append Name:=sFieldName, Type:=adVarWChar, DefinedSize:=-1, Attrib:=nFieldAttr
        Case "text"
            rs.Fields.Append Name:=sFieldName, Type:=adLongVarWChar, DefinedSize:=-1, Attrib:=nFieldAttr
        'Date(time) fields
        Case "date"
            rs.Fields.Append Name:=sFieldName, Type:=adDBDate, Attrib:=nFieldAttr
        Case "datetime"
            rs.Fields.Append Name:=sFieldName, Type:=adDate, Attrib:=nFieldAttr
        ' Relational fields
        Case "many2one"
            rs.Fields.Append Name:=sFieldName, Type:=adInteger, Attrib:=nFieldAttr
            rs.Fields.Append Name:=FormatDisplayName(sFieldName), Type:=adVarWChar, DefinedSize:=-1, Attrib:=CATTR_NON_MODEL
        Case "one2many"
            rs.Fields.Append Name:=sFieldName, Type:=adLongVarWChar, DefinedSize:=-1, Attrib:=nFieldAttr
        Case "many2many"
            rs.Fields.Append Name:=sFieldName, Type:=adLongVarWChar, DefinedSize:=-1, Attrib:=nFieldAttr
        ' Pseudo-relational fields
        Case "reference"
            rs.Fields.Append Name:=sFieldName, Type:=adVarWChar, DefinedSize:=-1, Attrib:=nFieldAttr
        Case "many2one_reference"
            rs.Fields.Append Name:=sFieldName, Type:=adVarWChar, DefinedSize:=-1, Attrib:=nFieldAttr
        ' Binary fields
        Case "binary"   ' Binary/Image
            rs.Fields.Append Name:=sFieldName, Type:=adLongVarWChar, DefinedSize:=-1, Attrib:=nFieldAttr
        Case Else
            Debug.Print "Unknown type:", sFieldType, sFieldName
            Debug.Assert False
            rs.Fields.Append Name:=sFieldName, Type:=adVarWChar, DefinedSize:=-1, Attrib:=nFieldAttr
    End Select

ExitProc:
    Set AddRecordsetField = rs.Fields(sFieldName)
End Function

Private Function GetDisplayNameField(rs As ADODB.Recordset, aFieldName As String) As ADODB.Field
    Set GetDisplayNameField = rs.Fields(FormatDisplayName(aFieldName))
End Function

Private Function AddRecordsetFieldAndRelatedModel(aRs As ADODB.Recordset, aModelName As String, aFieldName As String) As OdxModel
    Dim dicModelField As Dictionary
    
    If Not mModelFields.Exists(aFieldName) Then
        Err.Raise ERR_API_EXECUTION, "OdxModel", "Unknown field '" & aFieldName & "' in model '" & mModelName & "'."
    End If
    Set dicModelField = mModelFields(aFieldName)
    
    AddRecordsetField aRs, dicModelField
    Set AddRecordsetFieldAndRelatedModel = CreateRelatedModelOrNothing(dicModelField)
End Function

Private Function CreateRelatedModelOrNothing(dicModelField As Dictionary) As OdxModel
    Dim oModel As OdxModel
    Dim sFieldType As String
    Dim sRelationModelName As String
    Dim sFKeyFieldName As String
    
    sRelationModelName = dicModelField(CODOO_ATTR_RELATION)
    sFieldType = dicModelField(CODOO_ATTR_TYPE)
    
    Select Case sFieldType
        ' Relational fields
        Case "many2one"
            Set oModel = New OdxModel
            oModel.InitMany2One mRepository, sRelationModelName
        Case "one2many"
            sFKeyFieldName = dicModelField(CODOO_ATTR_FKEY)
            Set oModel = New OdxModel
            oModel.InitOne2Many mRepository, sRelationModelName, sFKeyFieldName
        Case "many2many"
            Set oModel = New OdxModel
            oModel.InitMany2Many mRepository, sRelationModelName
        ' (Others)
        Case Else
            Set oModel = Nothing
    End Select

    Set CreateRelatedModelOrNothing = oModel
End Function
