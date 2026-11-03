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

' ------------------------------------------------------------------------------
' Module: OdxModel.cls
' Purpose:
'   Manages an ADODB.Recordset mapped to an Odoo model, handling local data,
'   relational mappings (Many2One, One2Many, Many2Many), field schema generation,
'   and synchronizing data remotely via the unified OdxRepository.
' ------------------------------------------------------------------------------

Private Const CODOO_ATTR_NAME As String = "name"
Private Const CODOO_ATTR_STRING As String = "string"
Private Const CODOO_ATTR_TYPE As String = "type"
Private Const CODOO_ATTR_RELATION As String = "relation"
Private Const CODOO_ATTR_FKEY As String = "relation_field"   ' foreign key
Private Const CODOO_ATTR_REQUIRED As String = "required"

Private Const CFIELD_DISPLAY_NAME As String = "_display_name_"
Private Const CFIELD_MODEL_NAME As String = "_model_name_"
Private Const CFIELD_FILENAME As String = "_filename_"

Private Const CDEFAULT_EXTERNAL_ID As String = "external_id"

Private Enum RelationModelType
    mtBase
    mtMany2One
    mtOne2Many
    mtMany2Many
End Enum

Private mRepository As OdxRepository
Private mModelFields As Dictionary
Private mRs As OdxMemRs
Private mModelName As String
Private mRelationModelType As RelationModelType
Private mExternalIdFieldName As String
Private mRelations As Dictionary    ' key: relation field name, value: OdxModel
Private mOne2ManyFkName As String   ' for one2many: foreign key field name

' ==========================================
' Initialization Methods
' ==========================================

''
' Initializes the base model with an OdxRepository and target model name.
'
' @param aRepository OdxRepository instance for Odoo communication
' @param aModelName Target Odoo model name
Public Sub InitBase(ByRef aRepository As OdxRepository, ByRef aModelName As String)
    Set mRepository = aRepository
    Set mModelFields = mRepository.ModelSchema(aModelName)
    
    Set mRs = New OdxMemRs
    mRs.InitRs
    mRs.AddIdField
    
    mModelName = aModelName
    mExternalIdFieldName = ""
    Set mRelations = New Dictionary
    mOne2ManyFkName = ""
    
    mRelationModelType = mtBase
End Sub

''
' Initializes a Many2One relational model.
'
' @param aRepository OdxRepository instance
' @param aModelName Target Odoo model name
Public Sub InitMany2One(ByRef aRepository As OdxRepository, ByRef aModelName As String)
    InitBase aRepository, aModelName
    mRelationModelType = mtMany2One
End Sub

''
' Initializes a One2Many relational model.
'
' @param aRepository OdxRepository instance
' @param aModelName Target Odoo model name
' @param aFKeyFieldName Foreign key field name linking back to parent
Public Sub InitOne2Many(ByRef aRepository As OdxRepository, ByRef aModelName As String, ByRef aFKeyFieldName As String)
    InitBase aRepository, aModelName
    mRelationModelType = mtOne2Many
    mOne2ManyFkName = aFKeyFieldName
    AddOdooField aFKeyFieldName
End Sub

''
' Initializes a Many2Many relational model.
'
' @param aRepository OdxRepository instance
' @param aModelName Target Odoo model name
Public Sub InitMany2Many(ByRef aRepository As OdxRepository, ByRef aModelName As String)
    InitBase aRepository, aModelName
    mRelationModelType = mtMany2Many
End Sub

' ==========================================
' Properties
' ==========================================

''
' Gets a value indicating whether the model is Many2One.
'
' @return Boolean True if Many2One model, False otherwise
Public Property Get IsMany2OneModel() As Boolean
    IsMany2OneModel = (mRelationModelType = mtMany2One)
End Property

''
' Gets a value indicating whether the model is One2Many.
'
' @return Boolean True if One2Many model, False otherwise
Public Property Get IsOne2ManyModel() As Boolean
    IsOne2ManyModel = (mRelationModelType = mtOne2Many)
End Property

''
' Gets a value indicating whether the model is Many2Many.
'
' @return Boolean True if Many2Many model, False otherwise
Public Property Get IsMany2ManyModel() As Boolean
    IsMany2ManyModel = (mRelationModelType = mtMany2Many)
End Property

''
' Gets the target Odoo model name.
'
' @return String Model name
Public Property Get ModelName() As String
    ModelName = mModelName
End Property

''
' Returns a reference to the current OdxModel instance.
'
' @return OdxModel Current instance
Public Property Get RefMe() As OdxModel
    Set RefMe = Me
End Property

''
' Returns a cloned ADODB.Recordset instance.
' Note: Changes to the record pointer or filters do not affect the original.
'
' @param aPreserveFilter Optional Boolean flag to preserve active filter and sort (Default: True)
' @return Object Cloned ADODB.Recordset instance
Public Property Get CloneRecordset(Optional ByVal aPreserveFilter As Boolean = True) As Object
    Dim oCloneRs As Object
    Set oCloneRs = mRs.UnderlyingRecordset.Clone
    
    If aPreserveFilter Then
        oCloneRs.Filter = mRs.UnderlyingRecordset.Filter
        oCloneRs.Sort = mRs.UnderlyingRecordset.Sort
    Else
        oCloneRs.Filter = MemFilterNone
        oCloneRs.Sort = ""
    End If
    
    Set CloneRecordset = oCloneRs
End Property

' ==========================================
' Field methods
' ==========================================

''
' Gets or sets the value of an ID field in the current record.
'
' @return Variant Field value
Public Property Get FieldIdValue() As Variant
    FieldIdValue = mRs.FieldIdValue
End Property

Public Property Let FieldIdValue(ByRef aNewValue As Variant)
    mRs.FieldIdValue = aNewValue
End Property

''
' Gets or sets the value of an External ID field in the current record.
'
' @return Variant Field value
Public Property Get FieldExternalIdValue() As Variant
    FieldExternalIdValue = mRs.FieldValue(mExternalIdFieldName)
End Property

Public Property Let FieldExternalIdValue(ByRef aNewValue As Variant)
    mRs.FieldValue(mExternalIdFieldName) = aNewValue
End Property

''
' Gets or sets the value of a specified field in the current record.
'
' @param aFieldName Target field name
' @return Variant Field value
Public Property Get FieldValue(ByRef aFieldName As String) As Variant
    FieldValue = mRs.FieldValue(aFieldName)
End Property

Public Property Let FieldValue(ByRef aFieldName As String, ByRef aNewValue As Variant)
    mRs.FieldValue(aFieldName) = aNewValue
End Property

''
' Returns a collection of all field names in the recordset.
'
' @return Collection Collection of field name strings
Public Property Get FieldNames() As Collection
    Set FieldNames = mRs.FieldNames
End Property

''
' Adds standard status and control fields (_ope_, _del_, _syn_).
Public Sub AddStatusFields()
    mRs.AddStatusFields
End Sub

''
' Adds an external ID field to the local Recordset as a non-model property.
' Raises an exception if the external ID field has already been added.
'
' @param aFieldName Optional field name for the external ID (Default: CDEFAULT_EXTERNAL_ID)
Public Sub AddExternalIdField(Optional ByVal aFieldName As String = CDEFAULT_EXTERNAL_ID)
    Dim sTargetFieldName As String
    sTargetFieldName = IIf(aFieldName = "", CDEFAULT_EXTERNAL_ID, aFieldName)
    
    mRs.AddExternalIdField sTargetFieldName
    mExternalIdFieldName = sTargetFieldName
End Sub

''
' Adds a field to the local Recordset and sets up related models if relational.
'
' @param aFieldName Target Odoo field name
' @return OdxModel Related model instance if applicable, or Nothing
Public Function AddOdooField(ByRef aFieldName As String) As OdxModel
    Dim dicModelField As Dictionary
    Dim oRelation As OdxModel
    
    If Not mModelFields.Exists(aFieldName) Then
        Err.Raise CERR_API_EXECUTION, "OdxModel", "Unknown field '" & aFieldName & "' in model '" & mModelName & "'."
    End If
    
    Set dicModelField = mModelFields(aFieldName)
    AddDataBoundFieldByModelField dicModelField
    Set oRelation = CreateRelatedModelOrNothing(dicModelField)
    If Not oRelation Is Nothing Then
        Set mRelations(aFieldName) = oRelation
    End If
    Set AddOdooField = oRelation
End Function

''
' Adds a new field to the local disconnected Recordset structure as a non-model property.
'
' @param aFieldName Target field name
' @param aDataType Field data type (MemDataType)
' @param aDefinedSize Optional field size (MemFieldSize)
' @param aAttributes Field attributes
Public Sub AddLocalField(ByRef aFieldName As String, ByRef aDataType As MemDataType, Optional ByVal aDefinedSize As MemFieldSize = MemSizeDefalut, Optional ByVal aAttributes As MemFldAttribute = 0)
    mRs.AddLocalField aFieldName, aDataType, CLng(aDefinedSize)
End Sub

' ==========================================
' Odoo field methods
' ==========================================

''
' Collects names of all Odoo-mapped fields in the Recordset.
'
' @return Collection of field name strings
Private Function GetOdooFieldNames() As Collection
    Dim vItem As Variant
    
    Set GetOdooFieldNames = New Collection
    For Each vItem In mRs.FieldNames
        If mRs.IsDataBoundField(CStr(vItem)) Then
            GetOdooFieldNames.Add vItem
        End If
    Next vItem
End Function

''
' Formats the field name for a many2one display name property.
'
' @param aFieldName Original field name
' @return String Formatted field name
Private Function FormatFieldDisplayName(ByRef aFieldName As String) As String
    FormatFieldDisplayName = aFieldName & CFIELD_DISPLAY_NAME
End Function

''
' Formats the field name for a reference model name property.
'
' @param aFieldName Original field name
' @return String Formatted field name
Private Function FormatFieldModelName(ByRef aFieldName As String) As String
    FormatFieldModelName = aFieldName & CFIELD_MODEL_NAME
End Function

''
' Formats the field name for a binary filename property.
'
' @param aFieldName Original field name
' @return String Formatted field name
Private Function FormatFieldFilename(ByRef aFieldName As String) As String
    FormatFieldFilename = aFieldName & CFIELD_FILENAME
End Function

''
' Gets a related OdxModel instance for a specified relational field.
'
' @param aFieldName Relational field name
' @param optNotFiltered Optional boolean to skip automatic filtering (Default: False)
' @return OdxModel Related model instance
Public Function GetRelatedModel(ByRef aFieldName As String, Optional ByVal optNotFiltered As Boolean = False) As OdxModel
    Dim vValOrNull As Variant
    
    If Not mRelations.Exists(aFieldName) Then
        Debug.Assert False
        GoTo ExitProc
    End If
    
    Set GetRelatedModel = mRelations(aFieldName)
    
    If optNotFiltered Then
        GetRelatedModel.ClearFilter
    Else
        With GetRelatedModel
            If .IsMany2OneModel Then
                If mRs.BOF Or mRs.EOF Then
                    vValOrNull = Null
                Else
                    vValOrNull = mRs.FieldValue(aFieldName)
                End If
                .SetMany2OneFilter vValOrNull
            ElseIf .IsOne2ManyModel Then
                If mRs.BOF Or mRs.EOF Then
                    vValOrNull = Null
                Else
                    vValOrNull = mRs.FieldIdValue
                End If
                .SetOne2ManyFilter vValOrNull
            ElseIf .IsMany2ManyModel Then
                If mRs.BOF Or mRs.EOF Then
                    vValOrNull = Null
                Else
                    vValOrNull = mRs.FieldValue(aFieldName)
                End If
                .SetMany2ManyFilter vValOrNull
            Else
                Debug.Assert False
                GoTo ExitProc
            End If
        End With
    End If

ExitProc:
End Function

''
' Adds a field to the Recordset based on Odoo model schema definition.
'
' @param dicModelField Dictionary containing model field attributes
Private Sub AddDataBoundFieldByModelField(ByRef dicModelField As Dictionary)
    Dim sFieldName As String
    Dim sFieldType As String
    Dim nFieldAttr As Long
    
    Debug.Assert dicModelField.Exists(CODOO_ATTR_NAME)
    sFieldName = dicModelField(CODOO_ATTR_NAME)
    If dicModelField.Exists(CODOO_ATTR_REQUIRED) Then
        nFieldAttr = IIf(dicModelField(CODOO_ATTR_REQUIRED), MemFldRequired, 0)
    End If
    
    Debug.Assert dicModelField.Exists(CODOO_ATTR_TYPE)
    sFieldType = dicModelField(CODOO_ATTR_TYPE)
    Select Case sFieldType
        ' Basic fields
        Case "boolean"
            mRs.AddDataBoundField sFieldName, MemBoolean, , nFieldAttr
        Case "char"
            mRs.AddDataBoundField sFieldName, MemVarWChar, -1, nFieldAttr
        Case "float"
            mRs.AddDataBoundField sFieldName, MemCurrency, , nFieldAttr
        Case "integer"
            mRs.AddDataBoundField sFieldName, MemInteger, , nFieldAttr
        ' Advanced fields
        Case "html"
            mRs.AddDataBoundField sFieldName, MemLongVarWChar, -1, nFieldAttr
        Case "monetary"
            mRs.AddDataBoundField sFieldName, MemCurrency, , nFieldAttr
        Case "selection"
            mRs.AddDataBoundField sFieldName, MemVarWChar, -1, nFieldAttr
        Case "text"
            mRs.AddDataBoundField sFieldName, MemLongVarWChar, -1, nFieldAttr
        ' Date(time) fields
        Case "date"
            mRs.AddDataBoundField sFieldName, MemDate, , nFieldAttr
        Case "datetime"
            mRs.AddDataBoundField sFieldName, MemDate, , nFieldAttr
        ' Relational fields
        Case "many2one"
            mRs.AddDataBoundField sFieldName, MemInteger, , nFieldAttr
            mRs.AddLocalField FormatFieldDisplayName(sFieldName), MemVarWChar, -1
        Case "one2many"
            mRs.AddDataBoundField sFieldName, MemLongVarWChar, -1, nFieldAttr
        Case "many2many"
            mRs.AddDataBoundField sFieldName, MemLongVarWChar, -1, nFieldAttr
        ' Pseudo-relational fields
        Case "reference"
            mRs.AddDataBoundField sFieldName, MemInteger, , nFieldAttr
            mRs.AddLocalField FormatFieldModelName(sFieldName), MemVarWChar, -1
        Case "many2one_reference"
            mRs.AddDataBoundField sFieldName, MemVarWChar, -1, nFieldAttr
        ' Binary fields
        Case "binary"
            mRs.AddDataBoundField sFieldName, MemLongVarBinary, -1, nFieldAttr
            mRs.AddLocalField FormatFieldFilename(sFieldName), MemVarWChar, -1
        Case Else
            Debug.Print "Unknown type:", sFieldType, sFieldName
            Debug.Assert False
            mRs.AddDataBoundField sFieldName, MemVarWChar, -1, nFieldAttr
    End Select
End Sub

''
' Creates a related OdxModel instance based on field definition attributes, or returns Nothing.
'
' @param dicModelField Dictionary containing field attributes
' @return OdxModel Related model instance or Nothing
Private Function CreateRelatedModelOrNothing(ByRef dicModelField As Dictionary) As OdxModel
    Dim oModel As OdxModel
    Dim sFieldType As String
    Dim sRelationModelName As String
    Dim sFKeyFieldName As String
    
    If Not dicModelField.Exists(CODOO_ATTR_RELATION) Then
        Set CreateRelatedModelOrNothing = Nothing
        Exit Function
    End If
    
    sRelationModelName = dicModelField(CODOO_ATTR_RELATION)
    sFieldType = dicModelField(CODOO_ATTR_TYPE)
    Select Case sFieldType
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
        Case Else
            Set oModel = Nothing
    End Select
    
    Set CreateRelatedModelOrNothing = oModel
End Function

' ==========================================
' Record methods
' ==========================================

''
' Returns the total record count.
'
' @return Long Record count
Public Property Get RecordCount() As Long
    RecordCount = mRs.RecordCount
End Property

''
' Returns True if the cursor is at the beginning of the recordset.
'
' @return Boolean True if BOF
Public Property Get BOF() As Boolean
    BOF = mRs.BOF
End Property

''
' Returns True if the cursor is at the end of the recordset.
'
' @return Boolean True if EOF
Public Property Get EOF() As Boolean
    EOF = mRs.EOF
End Property

''
' Adds a new record to the underlying recordset.
Public Sub AddNew()
    mRs.AddNew
End Sub

''
' Updates the current record in the underlying recordset.
Public Sub Update()
    mRs.Update
End Sub

''
' Moves the cursor to the first record.
Public Sub MoveFirst()
    mRs.MoveFirst
End Sub

''
' Moves the cursor to the last record.
Public Sub MoveLast()
    mRs.MoveLast
End Sub

''
' Moves the cursor to the previous record.
Public Sub MovePrevious()
    mRs.MovePrevious
End Sub

''
' Moves the cursor to the next record.
Public Sub MoveNext()
    mRs.MoveNext
End Sub

''
' Marks the current record for deletion.
Public Sub MarkRecordForDelete()
    mRs.MarkRecordForDelete
End Sub

''
' Unmarks the current record for deletion.
Public Sub UnmarkRecordForDelete()
    mRs.UnmarkRecordForDelete
End Sub

''
' Clears all local records and resets relational data.
Public Sub ClearRecords()
    Dim dsRelation As OdxModel
    Dim vItem As Variant
    
    mRs.OpenRs
    mRs.ClearFilter
    mRs.MoveFirst
    Do Until mRs.EOF
        mRs.Delete
        mRs.MoveNext
    Loop
    mRs.CommitBatchChanges
    
    For Each vItem In mRelations.Keys
        Set dsRelation = mRelations(vItem)
        dsRelation.ClearRecords
    Next vItem
End Sub

''
' Clears any active filter on the Recordset.
Public Sub ClearFilter()
    mRs.ClearFilter
End Sub

''
' Sets a custom filter expression on the Recordset.
'
' @param aFilterCriteria Filter expression
Public Sub SetFilter(ByRef aFilterCriteria As Variant)
    mRs.SetFilter aFilterCriteria
End Sub

''
' Clears any active sort on the recordset.
Public Sub ClearSort()
    mRs.ClearSort
End Sub

''
' Sets a sort on the recordset.
'
' @param aSortCriteria Sort criteria expression
Public Sub SetSort(ByRef aSortCriteria As Variant)
    mRs.SetSort aSortCriteria
End Sub

''
' Finds a record in the underlying ADODB.Recordset based on specified criteria.
'
' @param aFindCriteria Search criteria expression
Public Function Find(ByRef aFindCriteria As Variant)
    mRs.Find aFindCriteria
End Function

''
' Filters the Recordset for a specific Many2One record ID.
'
' @param aIdOrNull Target record ID
Public Sub SetMany2OneFilter(ByRef aIdOrNull As Variant)
    Debug.Assert mRelationModelType = mtMany2One
    Dim colIdList As New Collection
    If Not IsNull(aIdOrNull) Then
        colIdList.Add CLng(aIdOrNull)
    End If
    mRs.SetIdFilter colIdList
End Sub

''
' Filters the Recordset for a specific One2Many foreign key ID.
'
' @param aFkIdOrNull Relation record ID
Public Sub SetOne2ManyFilter(ByRef aFkIdOrNull As Variant)
    Debug.Assert mRelationModelType = mtOne2Many
    If IsNull(aFkIdOrNull) Then
        mRs.SetIdFilter New Collection
    Else
        mRs.SetFilter mOne2ManyFkName & " = " & CLng(aFkIdOrNull)
    End If
End Sub

''
' Filters the Recordset for a specific Many2Many relation IDs collection.
'
' @param aJsonOrNull Relation record IDs collection (json string)
Public Sub SetMany2ManyFilter(ByRef aJsonOrNull As Variant)
    Debug.Assert mRelationModelType = mtMany2Many
    Dim colIds As Collection
    If IsNull(aJsonOrNull) Then
        Set colIds = New Collection
    Else
        Set colIds = ParseJson(aJsonOrNull)
    End If
    mRs.SetIdFilter colIds
End Sub

''
' Links a record for a Many2One relation.
'
' @param aId Many2One record ID
Public Sub LinkMany2OneRelation(ByRef aId As Long)
    Debug.Assert mRelationModelType = mtMany2One
    mRs.FindFirstOrAddNew aId
    mRs.Update
End Sub

''
' Links a record for a One2Many relation.
'
' @param aFkId Foreign record ID
' @param aId One2Many record ID
Public Sub LinkOne2ManyRelation(ByRef aFkId As Long, ByRef aId As Long)
    Debug.Assert mRelationModelType = mtOne2Many
    mRs.FindFirstOrAddNew aId
    mRs.FieldValue(mOne2ManyFkName) = aFkId
    mRs.Update
End Sub

''
' Links a record for a Many2Many relation.
'
' @param aId Many2Many record ID
Public Sub LinkMany2ManyRelation(ByRef aId As Long)
    Debug.Assert mRelationModelType = mtMany2Many
    mRs.FindFirstOrAddNew aId
    mRs.Update
End Sub

''
' Fills field values for a Many2One relationship.
'
' @param aValueOrNull Relationship value or null
' @param aFieldName Target field name
Private Sub FillMany2OneFieldValue(ByRef aValueOrNull As Variant, ByRef aFieldName As String)
    Dim sDispFieldName As String
    Dim colValue As Collection
    Dim dsRelation As OdxModel
    
    sDispFieldName = FormatFieldDisplayName(aFieldName)
    
    If IsNull(aValueOrNull) Then
        mRs.FieldValue(aFieldName) = Null
        mRs.FieldValue(sDispFieldName) = Null
        Exit Sub
    End If
    
    Set colValue = aValueOrNull
    mRs.FieldValue(aFieldName) = colValue(1)
    mRs.FieldValue(sDispFieldName) = colValue(2)

    If mRelations.Exists(aFieldName) Then
        Set dsRelation = mRelations(aFieldName)
        dsRelation.LinkMany2OneRelation colValue(1)
    End If
End Sub

''
' Fills field values for a One2Many relationship.
'
' @param aValueOrNull Relationship value or null
' @param aFieldName Target field name
' @param aFkId Foreign key ID
Private Sub FillOne2ManyFieldValue(ByRef aValueOrNull As Variant, ByRef aFieldName As String, ByRef aFkId As Long)
    Dim colValue As Collection
    Dim vItem As Variant
    Dim dsRelation As OdxModel

    If IsNull(aValueOrNull) Then
        mRs.FieldValue(aFieldName) = Null
        Exit Sub
    End If
    
    Set colValue = aValueOrNull
    mRs.FieldValue(aFieldName) = ConvertToJson(colValue)
    
    If mRelations.Exists(aFieldName) Then
        Set dsRelation = mRelations(aFieldName)
        For Each vItem In colValue
            dsRelation.LinkOne2ManyRelation aFkId, CLng(vItem)
        Next vItem
    End If
End Sub

''
' Fills field values for a Many2Many relationship.
'
' @param aValueOrNull Relationship value or null
' @param aFieldName Target field name
Private Sub FillMany2ManyFieldValue(ByRef aValueOrNull As Variant, ByRef aFieldName As String)
    Dim colValue As Collection
    Dim vItem As Variant
    Dim dsRelation As OdxModel

    If IsNull(aValueOrNull) Then
        mRs.FieldValue(aFieldName) = Null
        Exit Sub
    End If
    
    Set colValue = aValueOrNull
    mRs.FieldValue(aFieldName) = ConvertToJson(colValue)
    
    If mRelations.Exists(aFieldName) Then
        Set dsRelation = mRelations(aFieldName)
        For Each vItem In colValue
            dsRelation.LinkMany2ManyRelation CLng(vItem)
        Next vItem
    End If
End Sub

''
' Fills field values for a reference field type.
'
' @param aValueOrNull Reference value or null
' @param aFieldName Target field name
Private Sub FillReferenceFieldValue(ByRef aValueOrNull As Variant, ByRef aFieldName As String)
    Dim sModelFieldName As String
    
    sModelFieldName = FormatFieldModelName(aFieldName)

    If IsNull(aValueOrNull) Then
        mRs.FieldValue(aFieldName) = Null
        mRs.FieldValue(sModelFieldName) = Null
        Exit Sub
    End If
    
    Dim sValueList() As String
    sValueList = Split(aValueOrNull, ",")
    Debug.Assert UBound(sValueList) > 0
    Debug.Assert IsNumeric(sValueList(1))
    
    mRs.FieldValue(aFieldName) = CLng(sValueList(1))
    mRs.FieldValue(sModelFieldName) = Trim(sValueList(0))
End Sub

''
' Fills field values for a binary field type.
'
' @param aValueOrNull Binary dictionary or null
' @param aFieldName Target field name
Private Sub FillBinaryFieldValue(ByRef aValueOrNull As Variant, ByRef aFieldName As String)
    Dim sFilenameFieldName As String
    
    sFilenameFieldName = FormatFieldFilename(aFieldName)

    If IsNull(aValueOrNull) Then
        mRs.FieldValue(aFieldName) = Null
        mRs.FieldValue(sFilenameFieldName) = Null
        Exit Sub
    End If
    
    mRs.FieldValue(aFieldName) = Base64ToBytes(CStr(aValueOrNull("content")))
    mRs.FieldValue(sFilenameFieldName) = aValueOrNull("filename")
End Sub

''
' Updates the current record of the local Recordset with a single record dictionary (Values).
'
' @param dicRecord Dictionary of a single record fetched from Odoo
Private Sub FillRecord(ByRef dicRecord As Dictionary)
    Dim nId As Long
    Dim vItem As Variant
    Dim sFieldName As String
    Dim sFieldType As String
    Dim vFieldValue As Variant
        
    nId = dicRecord("id")
    Debug.Assert mRs.FieldIdValue = nId
            
    For Each vItem In mRs.FieldNames
        sFieldName = vItem
        
        ' ID
        If mRs.IsIdField(sFieldName) Then
            GoTo CONTINUE_FIELD
        End If
        ' External ID
        If mRs.IsExternalIdField(sFieldName) Then
            mRs.FieldValue(sFieldName) = mRepository.GetExternalIdFromModelAndId(mModelName, nId)
            GoTo CONTINUE_FIELD
        End If
        
        ' not odoo fields
        If Not mRs.IsDataBoundField(sFieldName) Then
            GoTo CONTINUE_FIELD
        End If
        
        sFieldType = mModelFields(sFieldName)(CODOO_ATTR_TYPE)
        
        ' Check null / false
        LetOrSetValue ValueOrDefault(dicRecord(sFieldName), False), vFieldValue
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
        Case "boolean"
            mRs.FieldValue(sFieldName) = vFieldValue
            
        Case "char"
            mRs.FieldValue(sFieldName) = vFieldValue
            
        Case "float"
            mRs.FieldValue(sFieldName) = vFieldValue
            
        Case "integer"
            mRs.FieldValue(sFieldName) = vFieldValue
            
        ' Advanced fields
        Case "html"
            If IsNull(vFieldValue) Then
                mRs.FieldValue(sFieldName) = Null
            Else
                mRs.FieldValue(sFieldName) = NormalizeLineBreaks(CStr(vFieldValue), True)
            End If
            
        Case "monetary"
            mRs.FieldValue(sFieldName) = vFieldValue
                    
        Case "selection"
            mRs.FieldValue(sFieldName) = vFieldValue
            
        Case "text"
            If IsNull(vFieldValue) Then
                mRs.FieldValue(sFieldName) = Null
            Else
                mRs.FieldValue(sFieldName) = NormalizeLineBreaks(CStr(vFieldValue), True)
            End If
            
        ' Date(time) fields
        Case "date"
            If IsNull(vFieldValue) Then
                mRs.FieldValue(sFieldName) = Null
            Else
                mRs.FieldValue(sFieldName) = ParseDate(CStr(vFieldValue))
            End If
            
        Case "datetime"
            If IsNull(vFieldValue) Then
                mRs.FieldValue(sFieldName) = Null
            Else
                mRs.FieldValue(sFieldName) = ParseDatetime(CStr(vFieldValue))
            End If
            
        ' Relational fields
        Case "many2one"
            FillMany2OneFieldValue vFieldValue, sFieldName
            
        Case "one2many"
            FillOne2ManyFieldValue vFieldValue, sFieldName, nId
            
        Case "many2many"
            FillMany2ManyFieldValue vFieldValue, sFieldName
            
        ' Pseudo-relational fields
        Case "reference"
            FillReferenceFieldValue vFieldValue, sFieldName
            
        Case "many2one_reference"
            mRs.FieldValue(sFieldName) = vFieldValue
            
        ' Binary fields
        Case "binary"
            FillBinaryFieldValue vFieldValue, sFieldName
            
        Case Else
            Debug.Print "Not Supported: " & sFieldName & " [" & sFieldType & "]"
            Debug.Assert False
        End Select
        
CONTINUE_FIELD:
    Next vItem
End Sub

''
' Sets reference field values into the target dictionary.
'
' @param aFieldName Target field name
' @param aTargetDict Target dictionary
Private Sub SetReferenceFieldToValue(ByRef aFieldName As String, ByRef aTargetDict As Dictionary)
    Dim sModelFieldName As String
    Dim vIdValue As Variant
    Dim vModelValue As Variant
    
    sModelFieldName = FormatFieldModelName(aFieldName)
    
    vIdValue = mRs.FieldValue(aFieldName)
    vModelValue = mRs.FieldValue(sModelFieldName)
    
    If IsNull(vIdValue) Or IsNull(vModelValue) Then
        aTargetDict(aFieldName) = Null
    Else
        aTargetDict(aFieldName) = CStr(vModelValue) & "," & CStr(vIdValue)
    End If
End Sub

''
' Sets binary field values into the target dictionary.
'
' @param aFieldName Target field name
' @param aTargetDict Target dictionary
Private Sub SetBinaryFieldToValue(ByRef aFieldName As String, ByRef aTargetDict As Dictionary)
    Dim sFilenameFieldName As String
    Dim vFieldValue As Variant
    Dim vFilenameValue As Variant
    Dim binaryDict As Dictionary
    
    vFieldValue = mRs.FieldValue(aFieldName)
    
    If IsNull(vFieldValue) Then
        aTargetDict(aFieldName) = Null
    Else
        Set binaryDict = New Dictionary
        sFilenameFieldName = FormatFieldFilename(aFieldName)
        vFilenameValue = mRs.FieldValue(sFilenameFieldName)
        If Not IsNull(vFilenameValue) Then
            If Not IsEmpty(vFilenameValue) Then
                binaryDict("filename") = CStr(vFilenameValue)
            End If
        End If
        binaryDict("content") = BytesToBase64(vFieldValue)
        binaryDict("size") = UBound(vFieldValue) - LBound(vFieldValue) + 1
        
        Set aTargetDict(aFieldName) = binaryDict
    End If
End Sub

''
' Converts the current record of the local Recordset into a Dictionary.
'
' @param aOnlyChanged Optional Boolean to output only changed fields (Default: True)
' @return Dictionary representing the record values
Private Function GetRecordValuesAsDict(Optional ByVal aOnlyChanged As Boolean = True) As Dictionary
    Dim dicRecord As Dictionary
    Dim vItem As Variant
    Dim sFieldName As String
    Dim vFieldValue As Variant
    Dim sFieldType As String
    
    Set dicRecord = New Dictionary
    For Each vItem In mRs.FieldNames
        sFieldName = vItem
        
        If Not mRs.IsDataBoundField(sFieldName) Then
            GoTo CONTINUE_FIELD
        End If
        
        If aOnlyChanged Then
            If Not mRs.IsFieldValueChanged(sFieldName) Then
                GoTo CONTINUE_FIELD
            End If
        End If
        
        vFieldValue = mRs.FieldValue(sFieldName)
        
        If IsNull(vFieldValue) Then
            dicRecord.Add sFieldName, False
            GoTo CONTINUE_FIELD
        End If
        
        sFieldType = mModelFields(sFieldName)(CODOO_ATTR_TYPE)
        Select Case sFieldType
        Case "boolean"
            dicRecord.Add sFieldName, CBool(vFieldValue)
        Case "char"
            dicRecord.Add sFieldName, CStr(vFieldValue)
        Case "float"
            dicRecord.Add sFieldName, CDbl(vFieldValue)
        Case "integer"
            dicRecord.Add sFieldName, CLng(vFieldValue)
        Case "html"
            dicRecord.Add sFieldName, NormalizeLineBreaks(CStr(vFieldValue), False)
        Case "monetary"
            dicRecord.Add sFieldName, CDbl(vFieldValue)
        Case "selection"
            dicRecord.Add sFieldName, CStr(vFieldValue)
        Case "text"
            dicRecord.Add sFieldName, NormalizeLineBreaks(CStr(vFieldValue), False)
        Case "date"
            dicRecord.Add sFieldName, FormatDate(CStr(vFieldValue))
        Case "datetime"
            dicRecord.Add sFieldName, FormatDatetime(CStr(vFieldValue))
        Case "many2one"
            dicRecord.Add sFieldName, CLng(vFieldValue)
        Case "one2many"
            dicRecord.Add sFieldName, WebHelpers.ParseJson(vFieldValue)
        Case "many2many"
            dicRecord.Add sFieldName, WebHelpers.ParseJson(vFieldValue)
        Case "reference"
            SetReferenceFieldToValue sFieldName, dicRecord
        Case "many2one_reference"
            dicRecord.Add sFieldName, vFieldValue
        Case "binary"
            SetBinaryFieldToValue sFieldName, dicRecord
        Case Else
            Debug.Print "Not Supported for RecordToValues: " & sFieldName & " [" & sFieldType & "]"
            Debug.Assert False
        End Select
        
CONTINUE_FIELD:
    Next vItem

EXIT_PROC:
    Set GetRecordValuesAsDict = dicRecord
End Function

''
' Fills a collection of fetched records into the local Recordset.
'
' @param colResult Collection of record dictionaries from Odoo
Private Sub FillRecordset(ByRef colResult As Collection)
    Dim dicRecord As Dictionary
    Dim dsRelation As OdxModel
    Dim nId As Long
    Dim vItem As Variant
    
    For Each dicRecord In colResult
        nId = dicRecord("id")
        mRs.FindFirstOrAddNew nId
        FillRecord dicRecord
        mRs.Update
    Next dicRecord
    
    mRs.CommitBatchChanges
    
    mRs.MoveFirst
    
    For Each vItem In mRelations.Keys
        Set dsRelation = mRelations(vItem)
        dsRelation.FetchRead
    Next vItem
End Sub

''
' Executes a remote search-read operation via the repository and populates local records.
'
' @param aDomain OdFilterDomain search criteria
' @param aOffset Number of records to skip
' @param aLimit Maximum number of records to retrieve
' @param aOrder Sorting order
' @return Long Number of fetched records
Public Function FetchSearchAndRead( _
    Optional ByVal aDomain As Collection, _
    Optional ByVal aOffset As Long = 0, _
    Optional ByVal aLimit As Long = 0, _
    Optional ByVal aOrder As String = "" _
) As Long
    If mRelationModelType <> mtBase Then
        Err.Raise CERR_API_EXECUTION, "OdxModel", "FetchSearchAndRead can only be called on a base model, not on relational models."
    End If
    
    Dim colFieldNames As Collection
    Dim colResult As Collection
    
    ClearRecords
    Set colFieldNames = GetOdooFieldNames()
    
    Set colResult = mRepository.SearchReadRecord(mModelName, aDomain, colFieldNames, aOffset, aLimit, aOrder)
    
    FillRecordset colResult
    FetchSearchAndRead = colResult.Count
End Function

''
' Fetches remote records by IDs via the repository and updates local data.
'
' @return Long Number of fetched records
Public Function FetchRead() As Long
    Dim colIds As Collection
    Dim nId As Long
    Dim colFieldNames As Collection
    Dim colResult As Collection
        
    Set colFieldNames = GetOdooFieldNames()
    Set colIds = NewList
    
    If Not mRs.BOF Then
        mRs.MoveFirst
    End If
    
    Do Until mRs.EOF
        nId = mRs.FieldIdValue
        colIds.Add nId
        mRs.MoveNext
    Loop
    
    If colIds.Count = 0 Then
        FetchRead = 0
        Exit Function
    End If
    
    Set colResult = mRepository.ReadRecord(mModelName, colIds, colFieldNames)
    
    FillRecordset colResult
    FetchRead = colResult.Count
End Function

''
' Upserts the current record in the local Recordset to Odoo remotely.
'
' @param aOnlyChanged Optional Boolean to output only changed fields (Default: True)
' @return Dictionary containing "action" ("created" or "updated") and "id" (Long)
Private Function UpsertCurrentRecordRemote(Optional ByVal aOnlyChanged As Boolean = True) As Dictionary
    Dim fieldValues As Dictionary
    Dim nId As Long
    Dim sExtId As String
    Dim bHasExternalIdField As Boolean
    Dim resultInfo As New Dictionary
    Dim bIsSuccess As Boolean
    Dim sActionType As String
    Dim extIdDict As Dictionary
    Dim upsertResult As Dictionary
    Dim resInfo As Dictionary
    Dim sErrMsg As String
    Dim createdIds As Collection
    
    bHasExternalIdField = (mExternalIdFieldName <> "")
    Set fieldValues = GetRecordValuesAsDict(aOnlyChanged)
    
    If bHasExternalIdField Then
        sExtId = mRs.GetFieldValueOrDefault(mExternalIdFieldName, "")
        If Len(Trim(sExtId)) = 0 Then
            sExtId = mRepository.GenerateExternalId(mModelName)
            mRs.FieldValue(mExternalIdFieldName) = sExtId
        End If
        Set extIdDict = New Dictionary
        extIdDict.Add sExtId, fieldValues
        
        Set upsertResult = mRepository.UpsertRecordByExternalId(mModelName, extIdDict)
        
        bIsSuccess = False
        If Not upsertResult Is Nothing Then
            Set resInfo = upsertResult(sExtId)
            If resInfo("action") <> "error" Then
                nId = resInfo("id")
                sActionType = resInfo("action")
                If IsNull(mRs.FieldIdValue()) Then
                    mRs.FieldIdValue() = nId
                End If
                bIsSuccess = True
            End If
        End If
        
        If Not bIsSuccess Then
            sErrMsg = "Failed to upsert record by External ID."
            If Not upsertResult Is Nothing Then
                If upsertResult.Exists(sExtId) Then
                    If upsertResult(sExtId).Exists("error_message") Then
                        sErrMsg = upsertResult(sExtId)("error_message")
                    End If
                End If
            End If
            Err.Raise CERR_API_EXECUTION, "OdxModel", sErrMsg
        End If
    Else
        If IsNull(mRs.FieldIdValue) Then
            Set createdIds = mRepository.CreateRecord(mModelName, fieldValues)
            If Not createdIds Is Nothing Then
                nId = createdIds(1)
                mRs.FieldIdValue = nId
                sActionType = "created"
            Else
                Err.Raise CERR_API_EXECUTION, "OdxModel", "Failed to create new record."
            End If
        Else
            nId = mRs.FieldIdValue
            bIsSuccess = mRepository.WriteRecord(mModelName, nId, fieldValues)
            sActionType = "updated"
            If Not bIsSuccess Then
                Err.Raise CERR_API_EXECUTION, "OdxModel", "Failed to update record with ID."
            End If
        End If
    End If
    
    mRs.Update
    
    resultInfo.Add "action", sActionType
    resultInfo.Add "id", nId
    Set UpsertCurrentRecordRemote = resultInfo
End Function

''
' Unlinks the current record in the local Recordset from Odoo remotely.
'
' @return Dictionary containing "action" ("unlinked") and "id" (Long)
Private Function UnlinkCurrentRecordRemote() As Dictionary
    Dim sExtId As String
    Dim nId As Long
    Dim bHasExternalIdField As Boolean
    Dim resultInfo As New Dictionary
    Dim bIsSuccess As Boolean
    Dim unlinkResult As Dictionary
    Dim resInfo As Dictionary
    Dim sErrMsg As String
        
    bHasExternalIdField = (mExternalIdFieldName <> "")
    
    If bHasExternalIdField Then
        sExtId = mRs.GetFieldValueOrDefault(mExternalIdFieldName, "")
        Set unlinkResult = mRepository.UnlinkRecordByExternalId(mModelName, sExtId)
        
        bIsSuccess = False
        If Not unlinkResult Is Nothing Then
            Set resInfo = unlinkResult(sExtId)
            If CStr(resInfo("action")) = "unlinked" Then
                nId = resInfo("id")
                bIsSuccess = True
            End If
        End If
        
        If Not bIsSuccess Then
            sErrMsg = "Failed to unlink record by External ID."
            If Not unlinkResult Is Nothing Then
                If unlinkResult.Exists(sExtId) Then
                    If unlinkResult(sExtId).Exists("error_message") Then
                        sErrMsg = unlinkResult(sExtId)("error_message")
                    End If
                End If
            End If
            Err.Raise CERR_API_EXECUTION, "OdxModel", sErrMsg
        End If
    Else
        nId = mRs.FieldIdValue
        bIsSuccess = mRepository.UnlinkRecord(mModelName, nId)
        If Not bIsSuccess Then
            Err.Raise CERR_API_EXECUTION, "OdxModel", "Failed to unlink record with ID."
        End If
    End If
        
    resultInfo.Add "action", "unlinked"
    resultInfo.Add "id", nId
    Set UnlinkCurrentRecordRemote = resultInfo
End Function

''
' Applies batch updates (added, modified, deleted records) to the remote Odoo repository.
'
' @return Boolean True if successful, False otherwise
Public Function ApplyBatchUpdate() As Boolean
    If Not mRs.ResolveAllStatuses() Then
        ApplyBatchUpdate = False
        Exit Function
    End If
    Dim vItem As Variant
    Dim sFieldName As String
    ' Validation - required
    mRs.MoveFirst
    Do Until mRs.EOF
        If Not mRs.IsSynced Then
            Select Case mRs.GetOperationFieldValue()
            Case MemOpeAdded, MemOpeModified
                For Each vItem In mRs.FieldNames
                    sFieldName = vItem
                    If mRs.IsRequiredField(sFieldName) Then
                        If IsNull(mRs.FieldValue(sFieldName)) Then
                            Err.Raise CERR_API_EXECUTION, "OdxModel", "Invalid required field: " & sFieldName
                        End If
                    End If
                Next vItem
            Case MemOpeDeleted
                For Each vItem In mRs.FieldNames
                    sFieldName = vItem
                    If mRs.IsExternalIdField(sFieldName) Then
                        If IsNull(sFieldName) Then
                            Err.Raise CERR_API_EXECUTION, "OdxModel", "Invalid required field: " & sFieldName
                        End If
                    End If
                Next vItem
            End Select
        End If
        mRs.MoveNext
    Loop
    
    ' Deleted
    mRs.SetOperationFilter MemOpeDeleted
    mRs.MoveFirst
    Do Until mRs.EOF
        If Not mRs.IsSynced Then
            UnlinkCurrentRecordRemote
            mRs.MarkAsSynced
        End If
        mRs.MoveNext
    Loop
    
    ' Modified
    mRs.SetOperationFilter MemOpeModified
    mRs.MoveFirst
    Do Until mRs.EOF
        If Not mRs.IsSynced Then
            UpsertCurrentRecordRemote True
            mRs.MarkAsSynced
        End If
        mRs.MoveNext
    Loop
    
    ' Added
    mRs.SetOperationFilter MemOpeAdded
    mRs.MoveFirst
    Do Until mRs.EOF
        If Not mRs.IsSynced Then
            UpsertCurrentRecordRemote False
            mRs.MarkAsSynced
        End If
        mRs.MoveNext
    Loop
    
    mRs.CommitBatchChanges
    
    ApplyBatchUpdate = True
End Function
