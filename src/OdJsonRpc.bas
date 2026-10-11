Attribute VB_Name = "OdJsonRpc"
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
' Module: OdJsonRpc
' ==============================================================================
' Description:
'   Provides general-purpose utility helper functions (date formatting, ISO conversion,
'   value checks) and factory functions to conveniently instantiate clients,
'   repositories, models, and domain search filters.
' ==============================================================================
''

''
' ------------------------------------------------------------------------------
' Define common error constants for Odoo VBA Client project
' ------------------------------------------------------------------------------
Public Const ERR_UNINITIALIZED         As Long = vbObjectError + 1001
Public Const ERR_AUTHENTICATION_FAILED As Long = vbObjectError + 1002
Public Const ERR_HTTP_ERROR            As Long = vbObjectError + 1003
Public Const ERR_INVALID_RESPONSE      As Long = vbObjectError + 1004
Public Const ERR_EMPTY_RESPONSE        As Long = vbObjectError + 1005
Public Const ERR_ODOO_RPC_ERROR        As Long = vbObjectError + 1006
Public Const ERR_API_EXECUTION         As Long = vbObjectError + 1007

''
' ------------------------------------------------------------------------------
' Helpers
' ------------------------------------------------------------------------------
''

''
' Formats a Date value to Odoo standard date string format (yyyy-mm-dd).
'
' @param aDate Date value to format
' @return String Formatted date string
''
Public Function FormatDate(aDate As Date) As String
    FormatDate = Format(aDate, "yyyy-mm-dd")
End Function

''
' Parses an Odoo date string into a Date type.
'
' @param aDateString Date string to parse
' @return Date Parsed date value
''
Public Function ParseDate(aDateString As String) As Date
    ParseDate = CDate(aDateString)
End Function

''
' Converts a Date value to an ISO datetime string format.
'
' @param aDatetime Date value to convert
' @return String ISO datetime string
''
Public Function ConvertToIsoDatetime(aDatetime As Date) As String
    ConvertToIsoDatetime = WebHelpers.ConvertToIso(aDatetime)
End Function

''
' Parses an ISO datetime string into a Date type.
'
' @param aIsoString ISO datetime string
' @return Date Parsed datetime value
''
Public Function ParseIsoDatetime(ByVal aIsoString As String) As Date
    If InStr(aIsoString, "T") = 0 Then
        aIsoString = Replace(aIsoString, " ", "T") & "Z"
    End If
    ParseIsoDatetime = WebHelpers.ParseIso(aIsoString)
End Function

''
' Checks whether a given value is set (non-empty, non-null, and not False).
'
' @param vValue Variant value to check
' @return Boolean True if value is set
''
Function IsValueSet(ByVal vValue As Variant) As Boolean
    If VarType(vValue) = vbBoolean Then
        IsValueSet = (vValue <> False)
    ElseIf IsNull(vValue) Then
        IsValueSet = False
    ElseIf IsObject(vValue) Then
        If vValue Is Nothing Then
            IsValueSet = False
        Else
            IsValueSet = True
        End If
    Else
        IsValueSet = True
    End If
End Function

''
' Returns the value if set, otherwise returns the specified default value.
'
' @param vValue Target value to evaluate
' @param aDefaultValue Default fallback value
' @return Variant Evaluated value or default
''
Function ValueOrDefault(ByVal vValue As Variant, Optional ByVal aDefaultValue As Variant = "") As Variant
    If IsValueSet(vValue) Then
        If IsObject(vValue) Then
            Set ValueOrDefault = vValue
        Else
            ValueOrDefault = vValue
        End If
    Else
        If IsObject(aDefaultValue) Then
            Set ValueOrDefault = aDefaultValue
        Else
            ValueOrDefault = aDefaultValue
        End If
    End If
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
' ------------------------------------------------------------------------------
' Factory
' ------------------------------------------------------------------------------
''

''
' Creates a new OdClient instance.
'
' @return OdClient New client instance
''
Public Function NewOdClient() As OdClient
    Set NewOdClient = New OdClient
End Function

''
' Creates a new OdRepository instance.
'
' @return OdRepository New repository instance
''
Public Function NewOdRepository() As OdRepository
    Set NewOdRepository = New OdRepository
End Function

''
' Creates a new Od2Client instance.
'
' @return Od2Client New client instance
''
Public Function NewOd2Client() As Od2Client
    Set NewOd2Client = New Od2Client
End Function

''
' Creates a new Od2Repository instance.
'
' @return Od2Repository New repository instance
''
Public Function NewOd2Repository() As Od2Repository
    Set NewOd2Repository = New Od2Repository
End Function

''
' Creates a new OdxRepository facade instance.
'
' @return OdxRepository New facade instance
''
Public Function NewOdxRepository() As OdxRepository
    Set NewOdxRepository = New OdxRepository
End Function

''
' Creates and initializes a new OdxModel instance.
'
' @param aRepository OdxRepository instance
' @param aModelName Target Odoo model name
' @return OdxModel Initialized model view instance
''
Public Function NewOdxModel(aRepository As OdxRepository, aModelName As String) As OdxModel
    Set NewOdxModel = New OdxModel
    NewOdxModel.InitBase aRepository, aModelName
End Function

''
' Creates a new OdFilterDomain instance.
'
' @return OdFilterDomain New filter domain instance
''
Public Function NewDomain() As OdFilterDomain
    Set NewDomain = New OdFilterDomain
End Function

''
' Creates a new OdFilterCriteria instance with a field expression.
'
' @param aFieldExpr Field expression string
' @return OdFilterCriteria New filter criteria instance
''
Public Function NewField(aFieldExpr As String) As OdFilterCriteria
    With New OdFilterCriteria
        Set NewField = .SetFieldExpr(aFieldExpr)
    End With
End Function

''
' Creates a new OdFilterCriteria instance with full criteria parameters.
'
' @param aFieldExpr Field expression string
' @param aOperator Logical operator string
' @param aValueExpr Comparison value
' @return OdFilterCriteria New filter criteria instance
''
Public Function NewCriteria(aFieldExpr As String, aOperator As String, aValueExpr As Variant) As OdFilterCriteria
    With New OdFilterCriteria
        Set NewCriteria = .SetCriteria(aFieldExpr, aOperator, aValueExpr)
    End With
End Function

''
' Creates a new logical AND combination (OdFilterCombi).
'
' @param aArity1 First criteria or combination
' @param aArity2 Second criteria or combination
' @return OdFilterCombi New combination instance
''
Public Function NewAnd(aArity1 As Object, aArity2 As Object) As OdFilterCombi
    With New OdFilterCombi
        Set NewAnd = .SetAndLogic(aArity1, aArity2)
    End With
End Function

''
' Creates a new logical OR combination (OdFilterCombi).
'
' @param aArity1 First criteria or combination
' @param aArity2 Second criteria or combination
' @return OdFilterCombi New combination instance
''
Public Function NewOr(aArity1 As Object, aArity2 As Object) As OdFilterCombi
    With New OdFilterCombi
        Set NewOr = .SetOrLogic(aArity1, aArity2)
    End With
End Function

''
' Creates a new logical NOT combination (OdFilterCombi).
'
' @param aArity Criteria or combination to negate
' @return OdFilterCombi New combination instance
''
Public Function NewNot(aArity As Object) As OdFilterCombi
    With New OdFilterCombi
        Set NewNot = .SetNotLogic(aArity)
    End With
End Function

''
' Creates a new Collection instance.
'
' @return Collection New collection instance
''
Public Function NewList() As Collection
    Set NewList = New Collection
End Function

''
' Creates a new Dictionary instance.
'
' @return Dictionary New dictionary instance
''
Public Function NewDict() As Dictionary
    Set NewDict = New Dictionary
End Function
