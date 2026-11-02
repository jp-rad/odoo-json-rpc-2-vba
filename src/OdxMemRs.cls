VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdxMemRs"
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
' Module: OdxMemRs.cls
' Purpose: Encapsulates an ADODB.Recordset to provide in-memory data operations,
'          change tracking, inline array-safe field comparison, and batch updates.
' ------------------------------------------------------------------------------

' ==========================================
' Constants & Control Field Names
' ==========================================
' Physical column names used for record identification and change tracking
Private Const CFIELD_ID As String = "id"
Private Const CFIELD_OPERATION As String = "_ope_"
Private Const CFIELD_MARKED_FOR_DELETE As String = "_del_"
Private Const CFIELD_IS_SYNCED As String = "_syn_"

' Odoo IDs are always positive integers, so we use -1 as an impossible ID.
' Filtering by this value guarantees the Recordset becomes empty.
Private Const CMISSING_ID As Long = -1

' ==========================================
' Enums & Constants (Late Binding)
' ==========================================

''
' Represents the operation state of a memory record
Public Enum MemOperation
    MemOpeUnchanged = 0
    MemOpeAdded = 1
    MemOpeModified = 2
    MemOpeDeleted = 3
    MemOpeCanceled = 4
End Enum

''
' Represents ADODB record status values
Public Enum MemRecordStatus
    MemRecOK = &H0        ' Corresponds to adRecOK (&H0)
    MemRecNew = &H1       ' Corresponds to adRecNew (&H1)
    MemRecModified = &H2  ' Corresponds to adRecModified (&H2)
End Enum

''
' Represents ADODB data types for late binding
Public Enum MemDataType
    MemInteger = 3        ' Corresponds to adInteger (3)
    MemCurrency = 6       ' Corresponds to adCurrency (6)
    MemDate = 7           ' Corresponds to adDate (7)
    MemBoolean = 11       ' Corresponds to adBoolean (11)
    MemBinary = 128       ' Corresponds to adBinary (128)
    MemVarBinary = 204    ' Corresponds to adVarBinary (204)
    MemLongVarBinary = 205 ' Corresponds to adLongVarBinary (205)
    MemVarChar = 200      ' Corresponds to adVarChar (200)
    MemLongVarChar = 201  ' Corresponds to adLongVarChar (201)
    MemVarWChar = 202     ' Corresponds to adVarWChar (202)
    MemLongVarWChar = 203 ' Corresponds to adLongVarWChar (203)
End Enum

' Constants used for field size definitions (Long)
Public Enum MemFieldSize
    MemSizeDefalut = 0&
    MemSizeVariable = -1&
End Enum

''
' Represents ADODB field attributes and custom extension bits
Public Enum MemFldAttribute
    MemFldKeyColumn = &H1         ' Corresponds to adFldKeyColumn (&H1)
    MemFldUpdatable = &H4         ' Corresponds to adFldUpdatable (&H4)
    MemFldFixed = &H10            ' Corresponds to adFldFixed (&H10)
    MemFldIsNullable = &H20       ' Corresponds to adFldIsNullable (&H20)
    'MemFldMayBeNull = &H40       ' Corresponds to adFldMayBeNull (&H40)
    MemFldLong = &H80             ' Corresponds to adLongVarChar (&H80)
    MemFldControl = &H100000      ' Custom Bit: Row-level control field
    MemFldExternalId = &H200000   ' Custom Bit: External ID field identifier
    MemFldDataBound = &H400000    ' Custom Bit: Remote data-bound field
    MemFldRequired = &H800000     ' Custom Bit: Required (Not Null) field constraint
End Enum

''
' Represents ADODB object state values
Public Enum MemState
    MemStateClosed = 0  ' Corresponds to adStateClosed (0)
    MemStateOpen = 1    ' Corresponds to adStateOpen (1)
End Enum

''
' Represents ADODB cursor type values
Public Enum MemCursorType
    MemOpenStatic = 3   ' Corresponds to adOpenStatic (3)
End Enum

''
' Represents ADODB lock type values
Public Enum MemLockType
    MemLockBatchOptimistic = 4  ' Corresponds to adLockBatchOptimistic (4)
End Enum

''
' Represents ADODB filter group values
Public Enum MemFilter
    MemFilterNone = 0   ' Corresponds to adFilterNone (0)
End Enum

Private Const CMODEL_ATTR As Long = (MemFldKeyColumn Or MemFldControl Or MemFldExternalId Or MemFldDataBound)
Private Const CNOT_MODEL_ATTR As Long = Not CMODEL_ATTR

'Private mRs As ADODB.Recordset
Private mRs As Object   ' ADODB.Recordset

' ==========================================
' Lifecycle & Initialization Methods
' ==========================================

''
' Initializes the internal recordset instance.
'
' @param aSource Optional existing OdxMemRs to clone
' @return OdxMemRs Initialized instance
Public Function InitRs(Optional ByVal aSource As OdxMemRs = Nothing) As OdxMemRs
    If aSource Is Nothing Then
        Set mRs = CreateObject("ADODB.Recordset")
    Else
        Set mRs = aSource.UnderlyingRecordset.Clone
    End If
    Set InitRs = Me
End Function

''
' Safely opens the recordset if it is currently closed.
Public Sub OpenRs()
    If mRs.State = MemStateClosed Then mRs.Open
End Sub

''
' Closes the recordset.
Public Sub CloseRs()
    If mRs.State <> MemStateClosed Then mRs.Close
End Sub

''
' Creates and returns a new OdxMemRs instance containing a clone of the underlying Recordset.
'
' @param aPreserveFilter Boolean flag indicating whether to preserve the active filter and sort
' @return OdxMemRs Cloned OdxMemRs instance
Public Function Clone(Optional ByVal aPreserveFilter As Boolean = True) As OdxMemRs
    Dim oNewMemRs As OdxMemRs
    Set oNewMemRs = New OdxMemRs
    Set Clone = oNewMemRs.InitRs(Me)
    If aPreserveFilter Then
        oNewMemRs.SetFilter mRs.Filter
        oNewMemRs.SetSort mRs.Sort
    Else
        oNewMemRs.ClearFilter
        oNewMemRs.ClearSort
    End If
End Function

' ==========================================
' Properties
' ==========================================

''
' Returns the underlying ADODB.Recordset object.
'
' @return Object Underlying ADODB.Recordset instance
Public Property Get UnderlyingRecordset() As Object
    Set UnderlyingRecordset = mRs
End Property

''
' Returns the state of the recordset.
'
' @return Long Recordset state (MemState)
Public Property Get State() As MemState
    State = mRs.State
End Property

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
' Returns a collection of all field names in the recordset.
'
' @return Collection Collection of field name strings
Public Property Get FieldNames() As Collection
    Dim oFld As Object  ' ADODB.Field
    Dim colNames As New Collection
    For Each oFld In mRs.Fields
        colNames.Add oFld.Name
    Next oFld
    Set FieldNames = colNames
End Property

''
' Gets or sets the value of a specified field in the current record.
'
' @param aFieldName Target field name
' @return Variant Field value
Public Property Get FieldValue(ByRef aFieldName As String) As Variant
    FieldValue = mRs.Fields(aFieldName).Value
End Property

Public Property Let FieldValue(ByRef aFieldName As String, ByRef aNewValue As Variant)
    mRs.Fields(aFieldName).Value = aNewValue
End Property

''
' Determines if a specific field has been modified in the current record,
' supporting scalar values as well as byte arrays via inline comparison.
'
' @param aFieldName Target field name
' @return Boolean True if changed, False otherwise
Public Function IsFieldValueChanged(ByRef aFieldName As String) As Boolean
    Dim vCurr As Variant
    Dim vOrig As Variant
     
    With mRs.Fields(aFieldName)
        vCurr = .Value
        vOrig = .OriginalValue
    End With
     
    If IsNull(vCurr) And IsNull(vOrig) Then
        IsFieldValueChanged = False
    ElseIf IsNull(vCurr) Or IsNull(vOrig) Then
        IsFieldValueChanged = True
    ElseIf IsBinaryField(aFieldName) Then
        If LBound(vCurr) <> LBound(vOrig) Then
            IsFieldValueChanged = True
        ElseIf UBound(vCurr) <> UBound(vOrig) Then
            IsFieldValueChanged = True
        Else
            Dim nIndex As Long
            Dim bIsDiff As Boolean
            bIsDiff = False
            For nIndex = LBound(vCurr) To UBound(vCurr)
                If vCurr(nIndex) <> vOrig(nIndex) Then
                    bIsDiff = True
                    Exit For
                End If
            Next nIndex
            IsFieldValueChanged = bIsDiff
        End If
    Else
        IsFieldValueChanged = (vCurr <> vOrig)
    End If
End Function

''
' Returns the attributes of a specified field.
'
' @param aFieldName Target field name
' @return MemFldAttribute Field attributes enum value
Public Property Get FieldAttributes(ByRef aFieldName As String) As MemFldAttribute
    FieldAttributes = mRs.Fields(aFieldName).Attributes
End Property

''
' Returns the data type of a specified field.
'
' @param aFieldName Target field name
' @return MemDataType Field data type enum value
Public Property Get FieldType(ByRef aFieldName As String) As MemDataType
    FieldType = mRs.Fields(aFieldName).Type
End Property

''
' Gets the field value or a default value if the field is null.
'
' @param aFieldName Target field name
' @param aDefault Default value to return if null
' @return Variant Field value or default
Public Function GetFieldValueOrDefault(ByRef aFieldName As String, ByRef aDefault As Variant) As Variant
    With mRs.Fields(aFieldName)
        If IsNull(.Value) Then
            LetOrSetValue aDefault, GetFieldValueOrDefault
        Else
            LetOrSetValue .Value, GetFieldValueOrDefault
        End If
    End With
End Function

''
' Checks whether a specified field exists in the recordset.
'
' @param aFieldName Target field name
' @return Boolean True if field exists, False otherwise
Public Function ExistsField(ByRef aFieldName As String) As Boolean
    On Error GoTo ErrHandler
    Dim vDummy As Variant
    vDummy = mRs.Fields(aFieldName).Type
    ExistsField = True
ExitProc:
    Exit Function
ErrHandler:
    ExistsField = False
    Resume ExitProc
End Function

''
' Gets or sets the value of an ID field in the current record.
'
' @return Variant Field value
Public Property Get FieldIdValue() As Variant
    FieldIdValue = mRs.Fields(CFIELD_ID).Value
End Property

Public Property Let FieldIdValue(ByRef aNewValue As Variant)
    mRs.Fields(CFIELD_ID).Value = aNewValue
End Property

' ==========================================
' Schema & Field Management Methods
' ==========================================

''
' Generic helper to add a field to the Recordset with specified attributes,
' automatically ensuring that null values and updatability are permitted.
'
' @param aFieldName Field name
' @param aDataType ADODB data type enum value
' @param aDefinedSize Maximum length or size
' @param aAttributes Field attributes
Private Sub AddField(ByRef aFieldName As String, ByRef aDataType As MemDataType, ByRef aDefinedSize As MemFieldSize, ByRef aAttributes As MemFldAttribute)
    aAttributes = aAttributes Or MemFldIsNullable Or MemFldUpdatable
    mRs.Fields.Append Name:=aFieldName, Type:=aDataType, DefinedSize:=aDefinedSize, Attrib:=aAttributes
End Sub

''
' Adds the primary key ID field to the recordset.
Public Sub AddIdField()
    AddField CFIELD_ID, MemInteger, MemSizeDefalut, MemFldKeyColumn Or MemFldControl
End Sub

''
' Adds an external ID field to the recordset.
'
' @param aFieldName Custom field name for External ID
Public Sub AddExternalIdField(ByRef aFieldName As String)
    Dim vItem As Variant
    Dim sName As String
    For Each vItem In FieldNames
        sName = vItem
        If IsExternalIdField(sName) Then
            Err.Raise CERR_API_EXECUTION, "AddExternalIdField", "External ID field has already been added: '" & sName & "'"
        End If
    Next vItem
    AddField aFieldName, MemLongVarWChar, MemSizeVariable, MemFldControl Or MemFldExternalId Or MemFldRequired
End Sub

''
' Adds a standard data-bound field with automatic mask processing.
'
' @param aFieldName Target field name
' @param aDataType Field data type (MemDataType)
' @param aDefinedSize Field size (MemFieldSize)
' @param aAttributes Field attributes
Public Sub AddDataBoundField(ByRef aFieldName As String, ByRef aDataType As MemDataType, Optional ByVal aDefinedSize As MemFieldSize = MemSizeDefalut, Optional ByVal aAttributes As MemFldAttribute)
    aAttributes = aAttributes And CNOT_MODEL_ATTR
    AddField aFieldName, aDataType, aDefinedSize, MemFldDataBound Or aAttributes
End Sub

''
' Adds a local-only field with automatic mask processing.
'
' @param aFieldName Target field name
' @param aDataType Field data type (MemDataType)
' @param aDefinedSize Field size (MemFieldSize)
' @param aAttributes Field attributes
Public Sub AddLocalField(ByRef aFieldName As String, ByRef aDataType As MemDataType, Optional ByVal aDefinedSize As MemFieldSize = MemSizeDefalut, Optional ByVal aAttributes As MemFldAttribute)
    aAttributes = aAttributes And CNOT_MODEL_ATTR
    AddField aFieldName, aDataType, aDefinedSize, aAttributes
End Sub

''
' Adds standard status and control fields (_ope_, _del_, _syn_).
Public Sub AddStatusFields()
    AddField CFIELD_OPERATION, MemInteger, MemSizeDefalut, MemFldControl
    AddField CFIELD_MARKED_FOR_DELETE, MemBoolean, MemSizeDefalut, MemFldControl
    AddField CFIELD_IS_SYNCED, MemBoolean, MemSizeDefalut, MemFldControl
End Sub

' ==========================================
' Field Classification Methods
' ==========================================

''
' Determines if the specified field is a control field.
'
' @param aFieldName Field name
' @return Boolean True if control field, False otherwise
Public Function IsControlField(ByRef aFieldName As String) As Boolean
    IsControlField = ((FieldAttributes(aFieldName) And MemFldControl) <> 0)
End Function

''
' Determines if the specified field is an ID field.
'
' @param aFieldName Field name
' @return Boolean True if ID field, False otherwise
Public Function IsIdField(ByRef aFieldName As String) As Boolean
    IsIdField = ((FieldAttributes(aFieldName) And MemFldKeyColumn) <> 0)
End Function

''
' Determines if the specified field is an external ID field.
'
' @param aFieldName Field name
' @return Boolean True if external ID field, False otherwise
Public Function IsExternalIdField(ByRef aFieldName As String) As Boolean
    IsExternalIdField = ((FieldAttributes(aFieldName) And MemFldExternalId) <> 0)
End Function

''
' Determines if the specified field is a data-bound field.
'
' @param aFieldName Field name
' @return Boolean True if data-bound field, False otherwise
Public Function IsDataBoundField(ByRef aFieldName As String) As Boolean
    IsDataBoundField = ((FieldAttributes(aFieldName) And MemFldDataBound) <> 0)
End Function

''
' Determines if the specified field is a local-only field.
'
' @param aFieldName Field name
' @return Boolean True if local field, False otherwise
Public Function IsLocalField(ByRef aFieldName As String) As Boolean
    IsLocalField = ((FieldAttributes(aFieldName) And CMODEL_ATTR) = 0)
End Function

''
' Determines if the specified field is a required field.
'
' @param aFieldName Field name
' @return Boolean True if required field, False otherwise
Public Function IsRequiredField(ByRef aFieldName As String) As Boolean
    IsRequiredField = ((FieldAttributes(aFieldName) And MemFldRequired) <> 0)
End Function

''
' Determines if a field type is binary (array).
'
' @param aFieldName Field name
' @return Boolean True if binary, False otherwise
Private Function IsBinaryField(ByRef aFieldName As String) As Boolean
    Select Case FieldType(aFieldName)
        Case MemBinary, MemVarBinary, MemLongVarBinary
            IsBinaryField = True
        Case Else
            IsBinaryField = False
    End Select
End Function

' ==========================================
' Record Navigation Methods
' ==========================================

''
' Safely moves the cursor to the first record if records exist.
Public Sub MoveFirst()
    If Not (mRs.BOF And mRs.EOF) Then mRs.MoveFirst
End Sub

''
' Safely moves the cursor to the last record if records exist.
Public Sub MoveLast()
    If Not (mRs.BOF And mRs.EOF) Then mRs.MoveLast
End Sub

''
' Moves the cursor to the previous record.
Public Sub MovePrevious()
    If Not mRs.BOF Then mRs.MovePrevious
End Sub

''
' Moves the cursor to the next record.
Public Sub MoveNext()
    If Not mRs.EOF Then mRs.MoveNext
End Sub

''
' Clears any active filter on the recordset.
Public Sub ClearFilter()
    mRs.Filter = MemFilterNone
End Sub

''
' Sets a filter on the recordset.
'
' @param aFilterCriteria Filter criteria string or array
Public Sub SetFilter(ByRef aFilterCriteria As Variant)
    mRs.Filter = aFilterCriteria
End Sub

''
' Clears any active sort on the recordset.
Public Sub ClearSort()
    mRs.Sort = ""
End Sub

''
' Sets a sort on the recordset.
'
' @param aSortCriteria Sort criteria string
Public Sub SetSort(ByRef aSortCriteria As Variant)
    mRs.Sort = aSortCriteria
End Sub

''
' Finds a record in the underlying ADODB.Recordset based on specified criteria.
'
' @param aFindCriteria Search criteria expression
Public Sub Find(ByRef aFindCriteria As Variant)
    mRs.Find aFindCriteria
End Sub

''
' Filters the Recordset for a record ID.
'
' @param aIdList Collection of Target record IDs
Public Sub SetIdFilter(ByRef aIdList As Collection)
    Dim arrCriteria() As String
    Dim vItem As Variant
    Dim i As Long
     
    If aIdList.Count > 0 Then
        ReDim arrCriteria(1 To aIdList.Count)
        For i = LBound(arrCriteria) To UBound(arrCriteria)
            arrCriteria(i) = CFIELD_ID & " = " & aIdList(i)
        Next i
        mRs.Filter = Join(arrCriteria, " OR ")
    Else
        mRs.Filter = CFIELD_ID & " = " & CMISSING_ID
    End If
End Sub

' ==========================================
' Record Modification Methods
' ==========================================

''
' Begins adding a new record.
Public Sub AddNew()
    mRs.AddNew
End Sub

''
' Updates the current record.
Public Sub Update()
    mRs.Update
End Sub

''
' Deletes the current record.
Public Sub Delete()
    mRs.Delete
End Sub

''
' Moves to a record by ID or creates a new record if it does not exist.
'
' @param aId Primary record ID
Public Sub FindFirstOrAddNew(ByRef aId As Long)
    MoveFirst
    mRs.Find CFIELD_ID & " = " & aId
    If mRs.EOF Then
        mRs.AddNew
        mRs.Fields(CFIELD_ID).Value = aId
    End If
End Sub

' ==========================================
' Status Resolution & Batch Processing
' ==========================================

''
' Checks whether all standard status fields exist in the recordset.
'
' @return Boolean True if all status fields exist, False otherwise
Private Function ExistsStatusFields() As Boolean
    ExistsStatusFields = ExistsField(CFIELD_OPERATION) _
                     And ExistsField(CFIELD_MARKED_FOR_DELETE) _
                     And ExistsField(CFIELD_IS_SYNCED)
End Function

''
' Checks if the current record is marked as synced.
'
' @return Boolean True if synced, False otherwise
Public Function IsSynced() As Boolean
    IsSynced = False
    If Not IsNull(mRs.Fields(CFIELD_IS_SYNCED).Value) Then
        IsSynced = CBool(mRs.Fields(CFIELD_IS_SYNCED).Value)
    End If
End Function

''
' Marks the current record as synced.
Public Sub MarkAsSynced()
    mRs.Fields(CFIELD_IS_SYNCED).Value = True
End Sub

''
' Marks the current record as unsynced.
Public Sub MarkAsUnsynced()
    mRs.Fields(CFIELD_IS_SYNCED).Value = False
End Sub

''
' Checks if the current record is marked for deletion.
'
' @return Boolean True if marked for delete, False otherwise
Public Function IsMarkedForDelete() As Boolean
    IsMarkedForDelete = False
    If Not IsNull(mRs.Fields(CFIELD_MARKED_FOR_DELETE).Value) Then
        IsMarkedForDelete = CBool(mRs.Fields(CFIELD_MARKED_FOR_DELETE).Value)
    End If
End Function

''
' Marks the current record for deletion.
Public Sub MarkRecordForDelete()
    mRs.Fields(CFIELD_MARKED_FOR_DELETE).Value = True
End Sub

''
' Removes the deletion mark from the current record.
Public Sub UnmarkRecordForDelete()
    mRs.Fields(CFIELD_MARKED_FOR_DELETE).Value = False
End Sub

''
' Returns the operation field value or a default unchanged state.
'
' @return MemOperation Operation status enum value
Public Function GetOperationFieldValue() As MemOperation
    GetOperationFieldValue = GetFieldValueOrDefault(CFIELD_OPERATION, MemOpeUnchanged)
End Function

''
' Sets a filter on the recordset based on the specified operation status.
'
' @param aOpe Operation status enum value (MemOperation)
Public Sub SetOperationFilter(ByRef aOpe As MemOperation)
    mRs.Filter = CFIELD_OPERATION & " = " & CStr(aOpe)
End Sub

''
' Computes and updates the operation status for the current record.
'
' @return MemOperation Resolved operation status
Private Function ResolveCurrentStatus() As MemOperation
    ResolveCurrentStatus = MemOpeUnchanged
    Debug.Assert Not (mRs.BOF Or mRs.EOF)
     
    If IsNull(mRs.Fields(CFIELD_OPERATION).Value) Then
        mRs.Fields(CFIELD_OPERATION).Value = MemOpeUnchanged
    End If
     
    If IsNull(mRs.Fields(CFIELD_MARKED_FOR_DELETE).Value) Then
        mRs.Fields(CFIELD_MARKED_FOR_DELETE).Value = False
    End If
     
    If IsNull(mRs.Fields(CFIELD_IS_SYNCED).Value) Then
        mRs.Fields(CFIELD_IS_SYNCED).Value = False
    End If
     
    Dim bDel As Boolean
    bDel = IsMarkedForDelete()
     
    Dim nOpe As MemOperation
    Select Case mRs.Status
        Case MemRecNew
            If bDel Then
                nOpe = MemOpeCanceled
            Else
                nOpe = MemOpeAdded
            End If
        Case MemRecModified
            If bDel Then
                nOpe = MemOpeDeleted
            Else
                nOpe = MemOpeModified
            End If
        Case Else
            nOpe = MemOpeUnchanged
    End Select
    
    If mRs.Fields(CFIELD_OPERATION).Value <> nOpe Then
        mRs.Fields(CFIELD_OPERATION).Value = nOpe
    End If
     
    ResolveCurrentStatus = nOpe
End Function

''
' Iterates through all records and resolves their operation statuses.
'
' @return Boolean True if successful, False if status fields are missing
Public Function ResolveAllStatuses() As Boolean
    If Not ExistsStatusFields() Then
        ResolveAllStatuses = False
        Exit Function
    End If
    ClearFilter
    MoveFirst
    Do Until EOF
        ResolveCurrentStatus
        MoveNext
    Loop
    MoveFirst
    ResolveAllStatuses = True
End Function

''
' Clears filters, commits deletions, resets tracking statuses to unchanged,
' and executes a batch update.
Public Sub CommitBatchChanges()
    If ExistsStatusFields() Then
        ClearFilter
        MoveFirst
        Do Until EOF
            If IsMarkedForDelete() Then
                Delete
            Else
                ' Reset
                MarkAsUnsynced
                UnmarkRecordForDelete
                FieldValue(CFIELD_OPERATION) = MemOpeUnchanged
            End If
            MoveNext
        Loop
    End If
    mRs.UpdateBatch
    MoveFirst
End Sub
