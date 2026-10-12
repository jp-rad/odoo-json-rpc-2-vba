VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Od2Client"
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

' ==============================================================================
' Class: Od2Client
' Purpose: Odoo External JSON-2 API Client
' Architecture Note: Encapsulates HTTP POST requests to present a Remote Procedure
'                    Call (RPC) interface for invoking Odoo model methods.
' Endpoint Pattern: /json/2/<model>/<method>
' ==============================================================================

Private m_baseUrl As String
Private m_apiKey As String
Private m_database As String
Private m_userAgent As String
Private m_initialized As Boolean
Private m_webClient As WebClient

' ------------------------------------------------------------------------------
' Class Initialization & Termination
' ------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set m_webClient = New WebClient
    m_initialized = False
    m_userAgent = "VBA-Od2Client"
End Sub

Private Sub Class_Terminate()
    Set m_webClient = Nothing
End Sub

' ------------------------------------------------------------------------------
' Public Properties
' ------------------------------------------------------------------------------

Public Property Get RefWebClient() As WebClient
    Set RefWebClient = m_webClient
End Property

Public Property Get UserAgent() As String
    UserAgent = m_userAgent
End Property

Public Property Let UserAgent(AValue As String)
    m_userAgent = AValue
End Property

' ------------------------------------------------------------------------------
' Public Methods
' ------------------------------------------------------------------------------

''
' Initialize the Odoo Client with Base URL, API Key, and Database Name
'
' @param aUrl Base URL of the Odoo instance
' @param aDb Database name
' @param aApiKey Odoo API Key
'
' @return Od2Client Initialized Od2Client instance
''
Public Function Initialize(aUrl As String, aDb As String, aApiKey As String) As Od2Client
    ' Remove trailing slash if present
    If Right$(aUrl, 1) = "/" Then
        m_baseUrl = Left$(aUrl, Len(aUrl) - 1)
    Else
        m_baseUrl = aUrl
    End If
    m_database = aDb
    m_apiKey = aApiKey
    
    m_webClient.BaseUrl = m_baseUrl
    m_initialized = True
    
    Set Initialize = Me
End Function

''
' Retrieve Odoo Server Version (GET /web/version)
'
' @return Dictionary {"result": Version Info payload}
''
Public Function GetVersion() As Dictionary
    CheckInitialization
    
    ' Header options
    Dim opts As New Dictionary
    opts.Add "Headers", CreateAuthHeaders()
    
    Dim resp As WebResponse
    Set resp = m_webClient.GetJson("/web/version", opts)
    
    If resp.StatusCode = WebStatusCode.Ok Then
        Set GetVersion = WrapResult(resp.Data)
        Exit Function
    Else
        Dim errDesc As String
        errDesc = "HTTP " & resp.StatusCode & ": " & resp.StatusDescription
        If Not resp.Data Is Nothing Then
            If TypeOf resp.Data Is Dictionary Then
                Dim errDict As Dictionary
                Set errDict = resp.Data
                If errDict.Exists("message") Then
                    errDesc = errDesc & " - " & errDict("message")
                End If
            End If
        End If
        
        Err.Raise Number:=ERR_HTTP_ERROR, _
                  Source:="Od2Client", _
                  Description:=errDesc
    End If
End Function

''
' Execute a Remote Procedure Call (RPC) to an Odoo model method via External JSON-2 API.
' Transmits parameters as JSON payload over HTTP POST to /json/2/<model>/<method>.
'
' @param aModel Target Odoo Model (e.g., "res.partner")
' @param aMethod Target Remote Method to invoke (e.g., "search_read", "create", "unlink")
' @param aParams Optional RPC keyword arguments (positional/keyword args, context, domain)
' @param aExpectPrimitive Optional Boolean specifying whether a primitive/single value is expected
' @return Dictionary {"result": Result payload}
''
Public Function ExecuteModelMethod(aModel As String, aMethod As String, Optional aParams As Dictionary, Optional aExpectPrimitive As Boolean = False) As Dictionary
    CheckInitialization
    
    ' RPC Endpoint syntax: /json/2/<model>/<method>
    Dim reqUrl As String
    reqUrl = "/json/2/" & aModel & "/" & aMethod
    
    ' Use a local variable to prevent modifying external references passed ByRef
    Dim payloadParams As Dictionary
    If aParams Is Nothing Then
        Set payloadParams = New Dictionary
    Else
        Set payloadParams = aParams
    End If
    
    ' Header options
    Dim opts As New Dictionary
    opts.Add "Headers", CreateAuthHeaders()
    
    ' Execute HTTP POST request (RPC call invocation)
    Dim resp As WebResponse
    Set resp = m_webClient.PostJson(reqUrl, payloadParams, opts)
    
    If resp.StatusCode = WebStatusCode.Ok Then
        If Not resp.Data Is Nothing Then
            Set ExecuteModelMethod = WrapResult(resp.Data, aExpectPrimitive)
            Exit Function
        Else
            Err.Raise Number:=ERR_EMPTY_RESPONSE, _
                      Source:="Od2Client", _
                      Description:="Response data is empty."
        End If
    Else
        ' Parse JSON-2 RPC error payload structure (name, message, arguments, debug)
        Dim errorMsg As String
        errorMsg = "HTTP " & resp.StatusCode & ": " & resp.StatusDescription
        
        If Not resp.Data Is Nothing Then
            If TypeOf resp.Data Is Dictionary Then
                Dim jsonErr As Dictionary
                Set jsonErr = resp.Data
                If jsonErr.Exists("message") Then
                    errorMsg = errorMsg & " | Odoo Error: " & jsonErr("message")
                End If
            End If
        End If
        
        Err.Raise Number:=ERR_ODOO_RPC_ERROR, _
                  Source:="Od2Client", _
                  Description:=errorMsg
    End If
End Function

' ------------------------------------------------------------------------------
' Private Helper Methods
' ------------------------------------------------------------------------------

Private Sub CheckInitialization()
    If Not m_initialized Then
        Err.Raise Number:=ERR_UNINITIALIZED, _
                  Source:="Od2Client", _
                  Description:="Od2Client is not initialized. Call Initialize() before executing requests."
    End If
End Sub

''
' Wraps a payload into a Dictionary with "result" key, with optional primitive validation/extraction.
'
' @param aPayload Value or object to encapsulate
' @param aExpectPrimitive Optional Boolean specifying whether to validate and extract single element from Collection
' @return Dictionary {"result": aPayload}
''
Private Function WrapResult(aPayload As Variant, Optional aExpectPrimitive As Boolean = False) As Dictionary
    Dim dataPayload As Variant
    
    ' Handle appropriate assignment separation for Object and Primitive (Let / Set)
    If IsObject(aPayload) Then
        Set dataPayload = aPayload
    Else
        dataPayload = aPayload
    End If
    
    If aExpectPrimitive Then
        If IsEmpty(dataPayload) Or IsNull(dataPayload) Then
            Err.Raise Number:=ERR_API_EXECUTION, _
                      Source:="Od2Client", _
                      Description:="Expected primitive value, but got Null or Empty."
        End If
        
        If TypeOf dataPayload Is Collection Then
            If dataPayload.Count = 1 Then
                Dim extractedVal As Variant
                extractedVal = dataPayload(1)
                
                ' Switch assignment based on whether the extracted value is an object
                If IsObject(extractedVal) Then
                    Set dataPayload = extractedVal
                Else
                    dataPayload = extractedVal
                End If
            Else
                Err.Raise Number:=ERR_API_EXECUTION, _
                          Source:="Od2Client", _
                          Description:="Expected single element collection, but got Count = " & dataPayload.Count
            End If
        Else
            Err.Raise Number:=ERR_API_EXECUTION, _
                      Source:="Od2Client", _
                      Description:="Expected Collection type for primitive extraction, but got a different type."
        End If
    End If

    Dim resDict As New Dictionary
    If IsObject(dataPayload) Then
        Set resDict("result") = dataPayload
    Else
        resDict("result") = dataPayload
    End If
    Set WrapResult = resDict
End Function

''
' Build Authentication & Routing Headers for Odoo External JSON-2 API
'
' @return Collection of KeyValue pairs representing HTTP headers
''
Private Function CreateAuthHeaders() As Collection
    Dim hdrs As New Collection
    
    hdrs.Add WebHelpers.CreateKeyValue("Authorization", "Bearer " & m_apiKey)
    hdrs.Add WebHelpers.CreateKeyValue("Content-Type", "application/json; charset=utf-8")
    
    If Len(m_userAgent) > 0 Then
        hdrs.Add WebHelpers.CreateKeyValue("User-Agent", m_userAgent)
    End If
    
    ' Optional database connection header
    If Len(m_database) > 0 Then
        hdrs.Add WebHelpers.CreateKeyValue("X-Odoo-Database", m_database)
    End If
    
    Set CreateAuthHeaders = hdrs
End Function
