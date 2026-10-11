VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdFilterCriteria"
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
' Class: OdFilterCriteria
' ==============================================================================
' Description:
'   Represents an individual search domain criterion condition (field, operator, value)
'   for Odoo ORM search queries, providing fluid helper methods for various comparison operators.
' ==============================================================================
''

'
' Search domains - ORM API
'
' A search domain is a first-order logical predicate used for filtering and searching recordsets.
' You combine simple conditions on a field expression with logical operators.
'
' https://www.odoo.com/documentation/master/developer/reference/backend/orm.html#search-domains
'

Private mFieldExpr As String
Private mOperator As String
Private mValue As Variant   ' string, number, boolean, null or OdFilterDomain

Private Const CERR_FIELD_EMPTY     As Long = vbObjectError + 1000
Private Const CERR_OPERATOR_EMPTY  As Long = vbObjectError + 1001
Private Const CERR_VALUE_MISSING   As Long = vbObjectError + 1002
Private Const CERR_INVALID_OBJECT  As Long = vbObjectError + 1003

''
' Validates that criteria members (field, operator, value) are properly set and typed.
''
Private Sub ValidateMembers()
    If Len(Trim(mFieldExpr)) = 0 Then
        Err.Raise CERR_FIELD_EMPTY, "OdFilterCriteria", "Field expression must not be empty."
    End If

    If Len(Trim(mOperator)) = 0 Then
        Err.Raise CERR_OPERATOR_EMPTY, "OdFilterCriteria", "Operator must not be empty." & vbCrLf & "[FieldExpr = '" & mFieldExpr & "']"
    End If

    If IsMissing(mValue) Then
        Err.Raise CERR_VALUE_MISSING, "OdFilterCriteria", "Value expression is Missing." & vbCrLf & "[FieldExpr = '" & mFieldExpr & "']"
    End If

    If IsObject(mValue) Then
        If Not mValue Is Nothing Then
            If Not TypeOf mValue Is OdFilterDomain And Not TypeOf mValue Is Collection Then
                Err.Raise CERR_INVALID_OBJECT, "OdFilterCriteria", "Object value must be of type OdFilterDomain or Collection." & vbCrLf & "[FieldExpr = '" & mFieldExpr & "']"
            End If
        End If
    End If
End Sub

''
' Serializes the search criterion into a JSON array string representation.
'
' @return String JSON representation of the criteria tuple
''
Public Function ToJson() As String
    Dim sValue As String
    Dim filterDomain As OdFilterDomain
    
    ValidateMembers
    
    If IsNull(mValue) Then   ' vbNull
        sValue = "null"
    ElseIf IsEmpty(mValue) Then ' Empty
        sValue = "null"
    ElseIf IsObject(mValue) Then
        If mValue Is Nothing Then   ' Nothing
            sValue = "null" ' null
        Else
            Set filterDomain = mValue
            sValue = filterDomain.ToJson()
        End If
    Else
        sValue = WebHelpers.ConvertToJson(mValue)
    End If

    ToJson = "[" & WebHelpers.ConvertToJson(mFieldExpr) & ", " & WebHelpers.ConvertToJson(mOperator) & ", " & sValue & "]"
End Function

''
' Sets the comparison value for the criterion.
'
' @param aValueExpr Comparison value
''
Public Sub SetValue(aValueExpr As Variant)
    If IsObject(aValueExpr) Then
        Set mValue = aValueExpr
    Else
        mValue = aValueExpr
    End If
End Sub

''
' Sets the full criteria parameters (field, operator, and value).
'
' @param aFieldExpr Field expression string
' @param aOperator Logical operator string
' @param aValueExpr Comparison value
' @return OdFilterCriteria Current instance for method chaining
''
Public Function SetCriteria(aFieldExpr As String, aOperator As String, aValueExpr As Variant) As OdFilterCriteria
    Set SetCriteria = Me
    mFieldExpr = aFieldExpr
    mOperator = aOperator
    SetValue aValueExpr
End Function

''
' Sets the field expression and clears previous operator and value.
'
' @param aFieldExpr Field expression string
' @return OdFilterCriteria Current instance for method chaining
''
Public Function SetFieldExpr(aFieldExpr As String) As OdFilterCriteria
    Set SetFieldExpr = Me
    mFieldExpr = aFieldExpr
    mOperator = ""
    mValue = Empty  ' Clear
End Function

''
' Helper method to create a new criteria with a specific operator and value.
'
' @param aOperator Logical operator string
' @param aValueExpr Comparison value
' @return OdFilterCriteria New criteria instance
''
Private Function SetOpeValue(aOperator As String, aValueExpr As Variant) As OdFilterCriteria
    With New OdFilterCriteria
        Set SetOpeValue = .SetCriteria(mFieldExpr, aOperator, aValueExpr)
    End With
End Function

''
' Creates an equality (=) criterion.
'
' @param aValueExpr Comparison value
' @return OdFilterCriteria New criteria instance
''
Public Function Eq(aValueExpr As Variant) As OdFilterCriteria
    Set Eq = SetOpeValue("=", aValueExpr)
End Function

''
' Creates a non-equality (!=) criterion.
'
' @param aValueExpr Comparison value
' @return OdFilterCriteria New criteria instance
''
Public Function NotEq(aValueExpr As Variant) As OdFilterCriteria
    Set NotEq = SetOpeValue("!=", aValueExpr)
End Function

''
' Creates a greater-than (>) criterion.
'
' @param aValueExpr Comparison value
' @return OdFilterCriteria New criteria instance
''
Public Function Gt(aValueExpr As Variant) As OdFilterCriteria
    Set Gt = SetOpeValue(">", aValueExpr)
End Function

''
' Creates a greater-than-or-equal (>=) criterion.
'
' @param aValueExpr Comparison value
' @return OdFilterCriteria New criteria instance
''
Public Function Ge(aValueExpr As Variant) As OdFilterCriteria
    Set Ge = SetOpeValue(">=", aValueExpr)
End Function

''
' Creates a less-than (<) criterion.
'
' @param aValueExpr Comparison value
' @return OdFilterCriteria New criteria instance
''
Public Function Lt(aValueExpr As Variant) As OdFilterCriteria
    Set Lt = SetOpeValue("<", aValueExpr)
End Function

''
' Creates a less-than-or-equal (<=) criterion.
'
' @param aValueExpr Comparison value
' @return OdFilterCriteria New criteria instance
''
Public Function Le(aValueExpr As Variant) As OdFilterCriteria
    Set Le = SetOpeValue("<=", aValueExpr)
End Function

''
' Creates an unset-or-equal (=?) criterion.
'
' @param aValueExpr Comparison value
' @return OdFilterCriteria New criteria instance
''
Public Function UnsetOrEq(aValueExpr As Variant) As OdFilterCriteria
    Set UnsetOrEq = SetOpeValue("=?", aValueExpr)
End Function

''
' Creates an explicit-like (=like) criterion.
'
' @param aValueExpr Comparison value
' @return OdFilterCriteria New criteria instance
''
Public Function EqLike(aValueExpr As Variant) As OdFilterCriteria
    Set EqLike = SetOpeValue("=like", aValueExpr)
End Function

''
' Creates a case-sensitive like criterion.
'
' @param aValueExpr Comparison value
' @return OdFilterCriteria New criteria instance
''
Public Function IsLike(aValueExpr As Variant) As OdFilterCriteria
    Set IsLike = SetOpeValue("like", aValueExpr)
End Function

''
' Creates a case-sensitive not-like criterion.
'
' @param aValueExpr Comparison value
' @return OdFilterCriteria New criteria instance
''
Public Function NotLike(aValueExpr As Variant) As OdFilterCriteria
    Set NotLike = SetOpeValue("not like", aValueExpr)
End Function

''
' Creates a case-insensitive ilike criterion.
'
' @param aValueExpr Comparison value
' @return OdFilterCriteria New criteria instance
''
Public Function IsILike(aValueExpr As Variant) As OdFilterCriteria
    Set IsILike = SetOpeValue("ilike", aValueExpr)
End Function

''
' Creates a case-insensitive not-ilike criterion.
'
' @param aValueExpr Comparison value
' @return OdFilterCriteria New criteria instance
''
Public Function NotILike(aValueExpr As Variant) As OdFilterCriteria
    Set NotILike = SetOpeValue("not ilike", aValueExpr)
End Function

''
' Creates an explicit case-insensitive =ilike criterion.
'
' @param aValueExpr Comparison value
' @return OdFilterCriteria New criteria instance
''
Public Function EqILike(aValueExpr As Variant) As OdFilterCriteria
    Set EqILike = SetOpeValue("=ilike", aValueExpr)
End Function

''
' Creates an inclusion (in) criterion.
'
' @param aValueExpr Comparison collection or domain
' @return OdFilterCriteria New criteria instance
''
Public Function IsIn(aValueExpr As Variant) As OdFilterCriteria
    Set IsIn = SetOpeValue("in", aValueExpr)
End Function

''
' Creates a non-inclusion (not in) criterion.
'
' @param aValueExpr Comparison collection or domain
' @return OdFilterCriteria New criteria instance
''
Public Function NotIn(aValueExpr As Variant) As OdFilterCriteria
    Set NotIn = SetOpeValue("not in", aValueExpr)
End Function

''
' Creates a child-of relation criterion.
'
' @param aValueExpr Comparison value
' @return OdFilterCriteria New criteria instance
''
Public Function ChildOf(aValueExpr As Variant) As OdFilterCriteria
    Set ChildOf = SetOpeValue("child_of", aValueExpr)
End Function

''
' Creates a parent-of relation criterion.
'
' @param aValueExpr Comparison value
' @return OdFilterCriteria New criteria instance
''
Public Function ParentOf(aValueExpr As Variant) As OdFilterCriteria
    Set ParentOf = SetOpeValue("parent_of", aValueExpr)
End Function

''
' Creates an any relation criterion.
'
' @param aValueExpr Comparison domain or collection
' @return OdFilterCriteria New criteria instance
''
Public Function IsAny(aValueExpr As Variant) As OdFilterCriteria
    Set IsAny = SetOpeValue("any", aValueExpr)
End Function

''
' Creates a not-any relation criterion.
'
' @param aValueExpr Comparison domain or collection
' @return OdFilterCriteria New criteria instance
''
Public Function NotAny(aValueExpr As Variant) As OdFilterCriteria
    Set NotAny = SetOpeValue("not any", aValueExpr)
End Function

''
' Creates an equality-to-null (= null) criterion.
'
' @return OdFilterCriteria New criteria instance
''
Public Function EqNull() As OdFilterCriteria
    Set EqNull = SetOpeValue("=", Null)
End Function

''
' Creates a non-equality-to-null (!= null) criterion.
'
' @return OdFilterCriteria New criteria instance
''
Public Function NotNull() As OdFilterCriteria
    Set NotNull = SetOpeValue("!=", Null)
End Function
