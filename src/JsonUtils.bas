Attribute VB_Name = "JsonUtils"
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
' Module: JsonUtils.bas
' Purpose: JSON Lookup & Count Utilities
' ------------------------------------------------------------------------------

Private Const CESCAPED_SLASH As String = vbBack

''
' Counts the number of items in a JSON array string.
'
' @param aJsonArray JSON array string to parse
' @return Variant The number of items in the array, or #VALUE! on error
''
Public Function JSONCOUNT(ByRef aJsonArray As Variant) As Variant
On Error GoTo ErrHandler
    Dim oColl As Collection

    Set oColl = WebHelpers.ParseJson(aJsonArray)
    JSONCOUNT = oColl.Count

ExitProc:
    Exit Function
ErrHandler:
    JSONCOUNT = CVErr(xlErrValue)
    Resume ExitProc
End Function

''
' Extracts a value from a JSON string using a path expression.
'
' @param aJsonInput JSON string to parse
' @param aJsonPath Path to the target value in the JSON structure using '/' as a separator.
'                  Array elements can be accessed with square brackets, e.g., "items[0]".
' @param aReturnNAForFalse (Optional, default True) If True, returns #N/A for False boolean values.
' @return Variant The value found at the specified path, or #N/A / #REF! / #VALUE! on error
''
Public Function JSONLOOKUP(ByRef aJsonInput As Variant, Optional ByVal aJsonPath As String = "", Optional ByVal aReturnNAForFalse As Boolean = True) As Variant
On Error GoTo ErrHandler
    Dim oJsonObject As Object

    ' Parse JSON string into an object
    Set oJsonObject = WebHelpers.ParseJson(aJsonInput)

    ' Escape "\/"
    aJsonPath = Replace(aJsonPath, "\/", CESCAPED_SLASH)

    ' Normalize jsonPath
    aJsonPath = Replace(aJsonPath, "][", "]/[")

    ' Extract value using the specified JSON path
    ExtractJsonValue oJsonObject, aJsonPath, JSONLOOKUP
    If IsNull(JSONLOOKUP) Then
        JSONLOOKUP = CVErr(xlErrNA)
        GoTo ExitProc
    End If
    If IsObject(JSONLOOKUP) Then
        JSONLOOKUP = WebHelpers.ConvertToJson(JSONLOOKUP)
        GoTo ExitProc
    End If
    If aReturnNAForFalse Then
        If VarType(JSONLOOKUP) = vbBoolean Then
            If Not JSONLOOKUP Then
                JSONLOOKUP = CVErr(xlErrNA)
                GoTo ExitProc
            End If
        End If
    End If

ExitProc:
    Exit Function
ErrHandler:
    Select Case Err.Number
        Case 9 ' Index out of bounds, Dictionary or Collection
            JSONLOOKUP = CVErr(xlErrRef)
        Case Else ' General fallback
            JSONLOOKUP = CVErr(xlErrValue)
    End Select
    Resume ExitProc
End Function

''
' Recursively extracts a value from a JSON object or array using an escaped path expression.
'
' @param aJsonObject Root JSON object or collection
' @param aEscapedJsonPath Escaped path expression
' @param aResult [Out] Target variant to store the extracted value
' @return Boolean True on successful extraction
''
Private Function ExtractJsonValue(ByRef aJsonObject As Object, ByVal aEscapedJsonPath As String, ByRef aResult As Variant) As Boolean
    Dim sRemainingPath As String
    Dim oNestedJsonObject As Object
    Dim arrPathSegments() As String
    Dim sKey As String
    Dim nIndex As Long
    Dim bHasIndex As Boolean
    Dim vTempItem As Variant
    Dim oDict As Object
    Dim oColl As Object

    ' Split the escaped JSON path by '/' to extract the first key or index
    arrPathSegments = Split(aEscapedJsonPath, "/", 2)
    If UBound(arrPathSegments) >= 0 Then
        sKey = Trim(arrPathSegments(0))
    Else
        sKey = ""
    End If
    If UBound(arrPathSegments) > 0 Then
        sRemainingPath = arrPathSegments(1)
    Else
        sRemainingPath = ""
    End If

    ' Handle key with index notation, e.g., item[0]
    If InStr(sKey, "[") > 0 Then
        If Right(sKey, 1) <> "]" Then
            Err.Raise 9 ' Index out of bounds, returns CVErr(xlErrRef)
        End If
        sKey = Replace(sKey, "[", "/")
        sKey = Replace(sKey, "]", "")
        arrPathSegments = Split(sKey, "/")
        sKey = Trim(arrPathSegments(0))
        If Not IsNumeric(arrPathSegments(1)) Then
            Err.Raise 9 ' Index out of bounds, returns CVErr(xlErrRef)
        End If
        nIndex = CLng(arrPathSegments(1)) + 1 ' Collections in VBA are 1-based
        bHasIndex = True
    Else
        bHasIndex = False
    End If

    ' Retrieve the value from the dictionary
    If sKey = "" Then
        Set vTempItem = aJsonObject
    Else
        sKey = Replace(sKey, CESCAPED_SLASH, "/") ' Unescape slash (restore escaped "/")
        Set oDict = aJsonObject
        If Not oDict.Exists(sKey) Then
            Err.Raise 9 ' Index out of bounds, returns CVErr(xlErrRef)
        End If
        If IsObject(oDict.Item(sKey)) Then
            Set vTempItem = oDict.Item(sKey)
        Else
            vTempItem = oDict.Item(sKey)
        End If
    End If
    If IsNull(vTempItem) Then
        GoTo ExitProc
    End If

    ' Retrieve value from collection if indexed
    If bHasIndex Then
        If TypeName(vTempItem) <> "Collection" Then
            Err.Raise 9 ' Index out of bounds, returns CVErr(xlErrRef)
        End If
        Set oColl = vTempItem
        If IsObject(oColl.Item(nIndex)) Then
            Set vTempItem = oColl.Item(nIndex)
        Else
            vTempItem = oColl.Item(nIndex)
        End If
    End If
    If IsNull(vTempItem) Then
        GoTo ExitProc
    End If

    ' Recursively resolve the remaining path
    If sRemainingPath <> "" Then
        Set oNestedJsonObject = vTempItem
        ExtractJsonValue oNestedJsonObject, sRemainingPath, vTempItem
    End If

ExitProc:
    ' Assign result
    If IsObject(vTempItem) Then
        Set aResult = vTempItem
    Else
        aResult = vTempItem
    End If

End Function
