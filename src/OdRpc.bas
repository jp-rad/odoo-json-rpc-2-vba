Attribute VB_Name = "OdRpc"
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
' Module: OdRpc.bas
' ------------------------------------------------------------------------------
''
' Description:
'   Provides general-purpose utility helper public functions (date formatting, ISO conversion,
'   value checks) and factory public functions to conveniently instantiate clients,
'   repositories, models, and domain search filters.
' ==============================================================================
''

''
' ------------------------------------------------------------------------------
' Define common error constants for Odoo VBA Client project
' ------------------------------------------------------------------------------
''
Public Const CERR_UNINITIALIZED        As Long = vbObjectError + 1001
Public Const CERR_AUTHENTICATION_FAILED As Long = vbObjectError + 1002
Public Const CERR_HTTP_ERROR            As Long = vbObjectError + 1003
Public Const CERR_INVALID_RESPONSE      As Long = vbObjectError + 1004
Public Const CERR_EMPTY_RESPONSE        As Long = vbObjectError + 1005
Public Const CERR_ODOO_RPC_ERROR        As Long = vbObjectError + 1006
Public Const CERR_API_EXECUTION         As Long = vbObjectError + 1007

''
' ------------------------------------------------------------------------------
' Define default module constant for external ID handling
' ------------------------------------------------------------------------------
''
Private Const CDEFAULT_MODULE As String = "__rpc_vba__"

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
Public Function FormatDate(ByRef aDate As Date) As String
    FormatDate = Format(aDate, "yyyy-mm-dd")
End Function

''
' Parses an Odoo date string into a Date type.
'
' @param aDateString Date string to parse
' @return Date Parsed date value
''
Public Function ParseDate(ByRef aDateString As String) As Date
    ParseDate = CDate(aDateString)
End Function

''
' Formats a local Date value into an Odoo UTC datetime string format.
'
' @param aDatetime Local Date value to convert
' @param aConvertToUtc Optional. Whether to convert to UTC (default: True)
' @return String UTC datetime string
''
Public Function FormatDatetime(ByRef aDatetime As Date, Optional ByVal aConvertToUtc As Boolean = True) As String
    Dim targetDate As Date
    If aConvertToUtc Then
        targetDate = WebHelpers.convertToUtc(aDatetime)
    Else
        targetDate = aDatetime
    End If
    FormatDatetime = Format(targetDate, "yyyy-mm-dd hh:nn:ss")
End Function

''
' Parses an Odoo UTC datetime string into a local Date type.
'
' @param aUtcString UTC datetime string
' @param aConvertToLocal Optional. Whether to convert to local time (default: True)
' @return Date Parsed local datetime value
''
Public Function ParseDatetime(ByRef aUtcString As String, Optional ByVal aConvertToLocal As Boolean = True) As Date
    Dim parsedDate As Date
    parsedDate = CDate(aUtcString)
    
    If aConvertToLocal Then
        ParseDatetime = WebHelpers.ParseUtc(parsedDate)
    Else
        ParseDatetime = parsedDate
    End If
End Function

''
' Assigns a value to a variable, using Set if it is an object, or Let otherwise.
'
' @param vVal1 Source value to assign
' @param vVal2 Target variable to receive the value
' @return Boolean True if assigned with Set, False if assigned with Let
''
Public Function LetOrSetValue(ByRef vVal1 As Variant, ByRef vVal2 As Variant) As Boolean
    If IsObject(vVal1) Then
        Set vVal2 = vVal1
        LetOrSetValue = True
    Else
        vVal2 = vVal1
        LetOrSetValue = False
    End If
End Function

''
' Checks whether a given value is assigned (non-empty, non-null, and not False).
'
' @param vValue Variant value to check
' @return Boolean True if the value is assigned
''
Private Function IsAssigned(ByRef vValue As Variant) As Boolean
    If IsNull(vValue) Then
        IsAssigned = False
    ElseIf IsObject(vValue) Then
        IsAssigned = Not (vValue Is Nothing)
    ElseIf VarType(vValue) = vbBoolean Then
        IsAssigned = CBool(vValue)
    Else
        IsAssigned = True
    End If
End Function

''
' Returns the value if it is assigned, otherwise returns the specified default value.
'
' @param vValue Target value to evaluate
' @param aDefaultValue Default fallback value
' @return Variant Evaluated value or default
''
Public Function ValueOrDefault(ByRef vValue As Variant, ByRef aDefaultValue As Variant) As Variant
    If IsAssigned(vValue) Then
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
' Normalizes and converts line breaks in the target text between LF and CRLF.
'
' @param aTargetText Target string to process
' @param aToCrLf Boolean True to convert line breaks to CRLF, False to keep them as LF
' @return String Processed string with normalized line breaks
''
Public Function NormalizeLineBreaks(ByRef aTargetText As String, ByRef aToCrLf As Boolean) As String
    NormalizeLineBreaks = Replace(Replace(aTargetText, vbCrLf, vbLf), vbCr, vbLf)
    If aToCrLf Then
        NormalizeLineBreaks = Replace(NormalizeLineBreaks, vbLf, vbCrLf)
    End If
End Function

''
' Converts a byte array into a Base64 encoded string.
' - Removes carriage returns and line feeds from the resulting text.
'
' @param aByteArray Byte() Raw binary data to be encoded
' @return String Clean Base64 encoded string
' @throws Error if MSXML2.DOMDocument instantiation or encoding fails
''
Public Function BytesToBase64(ByRef aByteArray As Variant) As String
    Dim oDom As Object
    Dim oElem As Object
    Dim sB64Text As String
    
    Set oDom = CreateObject("MSXML2.DOMDocument")
    Set oElem = oDom.createElement("b64")
    oElem.DataType = "bin.base64"
    oElem.nodeTypedValue = aByteArray
    
    sB64Text = oElem.Text
    BytesToBase64 = Replace(Replace(sB64Text, vbCr, ""), vbLf, "")
End Function

''
' Converts a Base64 encoded string back into a byte array.
'
' @param aBase64Str String Base64 encoded string to decode
' @return Byte() Decoded raw binary data as a byte array
' @throws Error if MSXML2.DOMDocument instantiation or decoding fails
''
Public Function Base64ToBytes(ByRef aBase64Str As String) As Byte()
    Dim oDom As Object
    Dim oElem As Object
    
    Set oDom = CreateObject("MSXML2.DOMDocument")
    Set oElem = oDom.createElement("b64")
    oElem.DataType = "bin.base64"
    oElem.Text = aBase64Str
    
    Base64ToBytes = oElem.nodeTypedValue
End Function

''
' Detects the image file format from the beginning of a byte array.
' - Analyzes magic numbers (file signatures) to identify common image types.
'
' @param aFileBytes Byte() Raw binary data of the image file
' @return String Image format/extension ("png", "jpg", "gif", "bmp", "webp", or "unknown")
''
Public Function GetImageFormat(ByRef aFileBytes() As Byte) As String
    Dim nLength As Long
    
    On Error GoTo ErrorHandler
    nLength = UBound(aFileBytes) - LBound(aFileBytes) + 1
    
    ' Need at least 12 bytes for reliable detection (especially WebP)
    If nLength < 12 Then
        GetImageFormat = "unknown"
        Exit Function
    End If
    
    ' 1. PNG: 89 50 4E 47 0D 0A 1A 0A
    If aFileBytes(0) = &H89 And aFileBytes(1) = &H50 And _
       aFileBytes(2) = &H4E And aFileBytes(3) = &H47 And _
       aFileBytes(4) = &HD And aFileBytes(5) = &HA And _
       aFileBytes(6) = &H1A And aFileBytes(7) = &HA Then
        GetImageFormat = "png"
        Exit Function
    End If
    
    ' 2. JPEG / JPG: FF D8 FF
    If aFileBytes(0) = &HFF And aFileBytes(1) = &HD8 And aFileBytes(2) = &HFF Then
        GetImageFormat = "jpg"
        Exit Function
    End If
    
    ' 3. GIF: 47 49 46 38 ("GIF8")
    If aFileBytes(0) = &H47 And aFileBytes(1) = &H49 And _
       aFileBytes(2) = &H46 And aFileBytes(3) = &H38 Then
        GetImageFormat = "gif"
        Exit Function
    End If
    
    ' 4. BMP: 42 4D ("BM")
    If aFileBytes(0) = &H42 And aFileBytes(1) = &H4D Then
        GetImageFormat = "bmp"
        Exit Function
    End If
    
    ' 5. WebP: RIFF....WEBP (Bytes 0-3: "RIFF", Bytes 8-11: "WEBP")
    If aFileBytes(0) = &H52 And aFileBytes(1) = &H49 And _
       aFileBytes(2) = &H46 And aFileBytes(3) = &H46 And _
       aFileBytes(8) = &H57 And aFileBytes(9) = &H45 And _
       aFileBytes(10) = &H42 And aFileBytes(11) = &H50 Then
        GetImageFormat = "webp"
        Exit Function
    End If
    
    GetImageFormat = "unknown"
    Exit Function

ErrorHandler:
    GetImageFormat = "unknown"
End Function

''
' Reads a file from local disk and returns its contents as a byte array.
'
' @param aFilePath String Absolute or relative path to the source file
' @return Byte() Binary contents of the loaded file
' @throws Error if file does not exist or ADODB.Stream fails
''
Public Function LoadBinaryFile(ByRef aFilePath As String) As Byte()
    Dim oStream As Object
    
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Type = 1 ' adTypeBinary
    oStream.Open
    oStream.LoadFromFile aFilePath
    
    LoadBinaryFile = oStream.Read()
    oStream.Close
End Function

''
' Writes a byte array into a file on local disk. Overwrites if file exists.
'
' @param aFilePath String Absolute or relative path where the file should be saved
' @param aFileBytes Byte() Binary data to write to disk
' @param aSaveCreateOverWrite Optional Boolean True to overwrite, False to fail if exists (Default: True)
' @throws Error if ADODB.Stream fails or disk write is unsuccessful
''
Public Sub SaveBinaryFile(ByRef aFilePath As String, ByRef aFileBytes() As Byte, Optional ByVal aSaveCreateOverWrite As Boolean = True)
    Dim oStream As Object
    
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Type = 1 ' adTypeBinary
    oStream.Open
    oStream.Write aFileBytes
    oStream.SaveToFile aFilePath, IIf(aSaveCreateOverWrite, 2, 1)   ' 1 = adSaveCreateNotExist, 2 = adSaveCreateOverWrite
    oStream.Close
End Sub

''
' Extracts only the filename and extension from a given file path.
'
' @param aFilePath String Full path to the file
' @return String Filename component
''
Public Function ExtractFileName(ByRef aFilePath As String) As String
    Dim nPos1 As Long, nPos2 As Long, nLastPos As Long
    nPos1 = InStrRev(aFilePath, "\")
    nPos2 = InStrRev(aFilePath, "/")
    If nPos1 > nPos2 Then nLastPos = nPos1 Else nLastPos = nPos2
    
    If nLastPos > 0 Then
        ExtractFileName = Mid$(aFilePath, nLastPos + 1)
    Else
        ExtractFileName = aFilePath
    End If
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
' @return OdxModel Initialized model instance
''
Public Function NewOdxModel(ByRef aRepository As OdxRepository, ByRef aModelName As String) As OdxModel
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
Public Function NewField(ByRef aFieldExpr As String) As OdFilterCriteria
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
Public Function NewCriteria(ByRef aFieldExpr As String, ByRef aOperator As String, ByRef aValueExpr As Variant) As OdFilterCriteria
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
Public Function NewAnd(ByRef aArity1 As Object, ByRef aArity2 As Object) As OdFilterCombi
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
Public Function NewOr(ByRef aArity1 As Object, ByRef aArity2 As Object) As OdFilterCombi
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
Public Function NewNot(ByRef aArity As Object) As OdFilterCombi
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

''
' ------------------------------------------------------------------------------
' Odoo External ID Management
' ------------------------------------------------------------------------------
''

''
' Splits an External ID into module and object name and returns them as a String array.
'
' @param aExternalId External ID string (e.g., "base_module.partner_root_1")
' @return String array where index 0 is "module" and index 1 is "name"
' @throws ERR_API_EXECUTION if the external ID is empty, invalid, or does not contain a dot.
''
Public Function SplitExternalId(ByRef aExternalId As String) As String()
    Dim arrParts() As String
    
    aExternalId = Trim(aExternalId)
    
    If aExternalId = "" Then
        Err.Raise CERR_API_EXECUTION, "SplitExternalId", "External ID cannot be empty."
    End If
    
    If InStr(aExternalId, ".") = 0 Then
        Err.Raise CERR_API_EXECUTION, "SplitExternalId", "External ID must contain a dot (module.name)."
    End If
    
    arrParts = Split(aExternalId, ".", 2)
            
    arrParts(0) = Trim(arrParts(0))
    arrParts(1) = Trim(arrParts(1))
    
    SplitExternalId = arrParts
End Function

''
' Normalizes an External ID string.
' If the dot is missing, it supplements the default module name before calling SplitExternalId.
'
' @param aExternalId Raw External ID string
' @param aDefaultModule Optional default module name when dot is missing (Default: CDEFAULT_MODULE)
' @return Normalized External ID string
' @throws ERR_API_EXECUTION if the external ID is empty, or if the object name is missing (e.g., "base.").
''
Public Function NormalizeExternalId(ByRef aExternalId As String, Optional ByVal aDefaultModule As String = CDEFAULT_MODULE) As String
    Dim arrParts() As String
    Dim sModuleName As String
    Dim sObjName As String
    
    aExternalId = Trim(aExternalId)
    
    ' Check if the input is empty
    If aExternalId = "" Then
        Err.Raise CERR_API_EXECUTION, "NormalizeExternalId", "External ID cannot be empty."
    End If
        
    ' Supplement the module name with default module if dot is missing
    If InStr(aExternalId, ".") = 0 Then
        aExternalId = "." & aExternalId
    End If
    
    ' Split the External ID using the existing public function (returns String array)
    arrParts = SplitExternalId(aExternalId)
    
    sModuleName = Trim(arrParts(0))
    sObjName = Trim(arrParts(1))
    
    If sModuleName = "" Then
        If Trim(aDefaultModule) = "" Then aDefaultModule = CDEFAULT_MODULE
        sModuleName = aDefaultModule
    End If
    
    ' Guard against empty object names (e.g., "base." or just a dot)
    If sObjName = "" Then
        Err.Raise CERR_API_EXECUTION, "NormalizeExternalId", "Object name cannot be empty in External ID (module.name)."
    End If
        
    ' Clean up spaces and replace extra dots in object name with underscores
    sModuleName = Replace(sModuleName, " ", "_")
    sObjName = Replace(sObjName, " ", "_")
    sObjName = Replace(sObjName, ".", "_")
    
    ' Rebuild and return the normalized External ID
    NormalizeExternalId = sModuleName & "." & sObjName
End Function
