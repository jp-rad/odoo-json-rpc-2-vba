VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdFilterCombi"
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
' Class: OdFilterCombi
' ==============================================================================
' Description:
'   Represents logical combinations (AND, OR, NOT) for Odoo search domain criteria,
'   handling operator prefixing and recursive JSON serialization of conditions.
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

Private Const CLOGIC_NOT As String = "'!'"
Private Const CLOGIC_AND As String = "'&'"
Private Const CLOGIC_OR As String = "'|'"

Private mLogic As String
Private mArityList As Collection

''
' Serializes the logical combination and its nested arities into a JSON string format.
'
' @return String JSON representation of the filter combination
''
Public Function ToJson() As String
    Dim s() As String
    Dim cnt As Long
    Dim i As Long
    Dim v As Variant
    
    If mLogic = CLOGIC_NOT Then
        ReDim s(0 To 1)
        s(0) = CLOGIC_NOT
        s(1) = mArityList(1).ToJson()
    Else
        cnt = mArityList.Count
        ReDim s(0 To (cnt - 1) * 2)
        For i = 0 To cnt - 2
            s(i) = mLogic
        Next i
        For Each v In mArityList
            s(i) = v.ToJson()
            i = i + 1
        Next v
    End If
    ToJson = Join(s, ", ")
End Function

''
' Sets the logical operator and initializes the arity collection.
'
' @param aLogic Logical operator string ('&', '|', or '!')
' @param aArity1 First criteria or combination object
' @param aArity2 Second criteria or combination object (optional for NOT)
' @return OdFilterCombi Current instance for method chaining
''
Private Function SetLogic(aLogic As String, aArity1 As Object, aArity2 As Object) As OdFilterCombi
    Set SetLogic = Me
    mLogic = aLogic
    Set mArityList = New Collection
    mArityList.Add aArity1
    If mLogic <> CLOGIC_NOT Then
        mArityList.Add aArity2
    End If
End Function

''
' Initializes a logical AND combination.
'
' @param aArity1 First condition
' @param aArity2 Second condition
' @return OdFilterCombi Initialized combination instance
''
Public Function SetAndLogic(aArity1 As Object, aArity2 As Object) As OdFilterCombi
    Set SetAndLogic = SetLogic(CLOGIC_AND, aArity1, aArity2)
End Function

''
' Initializes a logical OR combination.
'
' @param aArity1 First condition
' @param aArity2 Second condition
' @return OdFilterCombi Initialized combination instance
''
Public Function SetOrLogic(aArity1 As Object, aArity2 As Object) As OdFilterCombi
    Set SetOrLogic = SetLogic(CLOGIC_OR, aArity1, aArity2)
End Function

''
' Initializes a logical NOT combination.
'
' @param aArity Condition to negate
' @return OdFilterCombi Initialized combination instance
''
Public Function SetNotLogic(aArity As Object) As OdFilterCombi
    Set SetNotLogic = SetLogic(CLOGIC_NOT, aArity, Nothing)
End Function

''
' Adds an additional arity condition to the logical combination.
'
' @param aArity Criteria or combination object to add
' @return OdFilterCombi Current instance for method chaining
''
Public Function Add(aArity As Object) As OdFilterCombi
    Set Add = Me
    mArityList.Add aArity
End Function
