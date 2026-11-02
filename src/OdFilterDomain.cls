VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdFilterDomain"
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
' Class: OdFilterDomain
' ==============================================================================
' Description:
'   Represents an Odoo search domain collection (first-order logical predicates),
'   allowing conditions and combinations to be added, serialized to JSON, and built into collections.
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

Private mList As New Collection ' list of criteria or combinations

''
' Adds an arity condition or logical combination to the domain.
'
' @param aArity Criteria, combination, or value to add
' @return OdFilterDomain Current instance for method chaining
''
Public Function AddArity(aArity As Variant) As OdFilterDomain
    Set AddArity = Me
    mList.Add aArity
End Function

''
' Converts a variant item within the domain into its JSON string representation.
'
' @param v Variant item (criteria, combination, or value)
' @return String JSON representation
''
Private Function VarToJson(v As Variant) As String
    If Not IsObject(v) Then
        Dim sJson As String
        Dim colParsed As Collection
        If VarType(v) = vbString Then
            VarToJson = """" & v & """"
        Else
            VarToJson = CStr(v)
        End If
        Exit Function
    End If

    If TypeOf v Is OdFilterCriteria Then
        Dim oCriteria As OdFilterCriteria
        Set oCriteria = v
        VarToJson = oCriteria.ToJson()
        Exit Function
    End If
    
    If TypeOf v Is OdFilterCombi Then
        Dim oCombi As OdFilterCombi
        Set oCombi = v
        VarToJson = oCombi.ToJson()
        Exit Function
    End If
    
    Debug.Assert TypeOf v Is OdFilterDomain
    Dim oDomain As OdFilterDomain
    Set oDomain = v
    VarToJson = oDomain.ToJson()
End Function

''
' Serializes the entire domain into a JSON string format.
'
' @return String JSON representation of the domain
''
Public Function ToJson() As String
    Dim s() As String
    Dim i As Integer
    
    If mList.Count = 0 Then
        ToJson = "[]"
        GoTo ExitProc
    End If
    
    ReDim s(1 To mList.Count)
    For i = LBound(s) To UBound(s)
        s(i) = VarToJson(mList(i))
    Next i
    
    ToJson = "[" & Join(s, ", ") & "]"
    
ExitProc:
    Exit Function
End Function

''
' Builds and parses the domain JSON into a VBA Collection structure.
'
' @return Collection Parsed domain collection
''
Public Function Build() As Collection
    Set Build = WebHelpers.ParseJson(Me.ToJson())
End Function

''
' Builds the domain and appends it to an existing parameter collection.
'
' @param aList Target collection to append to
' @return Collection Built domain collection
''
Public Function BuildAndAppend(aList As Collection) As Collection
    Set BuildAndAppend = Build
    aList.Add BuildAndAppend
End Function
