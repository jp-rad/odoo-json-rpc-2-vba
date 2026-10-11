VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdClient"
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
' Class: OdClient
' Purpose: Odoo External API Client
' Architecture Note: Encapsulates JSON-RPC requests to common and object services.
' Endpoint Pattern: /jsonrpc
' ==============================================================================

Private m_baseUrl As String
Private m_database As String
Private m_username As String
Private m_apiKey As String
Private m_uid As Long
Private m_userAgent As String
Private m_initialized As Boolean
Private m_webClient As WebClient

' ------------------------------------------------------------------------------
' Class Initialization & Termination
' ------------------------------------------------------------------------------
Private Sub Class_Initialize()
    Set m_webClient = New WebClient
    m_initialized = False
    m_uid = 0
    m_userAgent = "VBA-OdClient"
End Sub

Private Sub Class_Terminate()
    Set m_webClient = Nothing
End Sub

' ------------------------------------------------------------------------------
' Public Properties
' ------------------------------------------------------------------------------

Public Property Get UserAgent() As String
    UserAgent = m_userAgent
End Property

Public Property Let UserAgent(ByVal AValue As String)
    m_userAgent = AValue
End Property

Public Property Get UserId() As Long
    UserId = m_uid
End Property

' ------------------------------------------------------------------------------
' Public Methods
' ------------------------------------------------------------------------------

''
' Initialize the Odoo Client with Base URL, Database Name, Username, and API Key
'
' @param aUrl Base URL of the Odoo instance (e.g., "https://your-instance.odoo.com")
' @param aDb Database name
' @param aUsername User email / login name
' @param aApiKey Odoo API Key
'
' @return OdClient Initialized OdClient instance
''
Public Function Initialize(aUrl As String, aDb As String, aUserName As String, aApiKey As String) As OdClient
    ' Remove trailing slash if present
    If Right$(aUrl, 1) = "/" Then
        m_baseUrl = Left$(aUrl, Len(aUrl) - 1)
    Else
        m_baseUrl = aUrl
    End If
    m_database = aDb
    m_username = aUserName
    m_apiKey = aApiKey
    
    m_webClient.BaseUrl = m_baseUrl
    m_initialized = True
    
    Set Initialize = Me
End Function

''
' Retrieve Odoo Server Version (common service -> version)
'
' @return Version Info Dictionary
''
Public Function GetVersion() As Dictionary
    CheckInitialization
    
    ' The version method typically takes an empty argument collection
    Dim argsList As New Collection
    
    Set GetVersion = CallJsonRpc("common", "version", argsList)
End Function

''
' Authenticate with Odoo (common service -> authenticate)
' Automatically invoked if needed, but can be called explicitly to verify credentials.
'
' @return Authenticated UID (Long)
''
Public Function Authenticate() As Long
    CheckInitialization
    
    Dim argsList As New Collection
    argsList.Add m_database
    argsList.Add m_username
    argsList.Add m_apiKey
    argsList.Add New Dictionary ' Empty struct required by Odoo authenticate
    
    Dim res As Dictionary
    Set res = CallJsonRpc("common", "authenticate", argsList)
    
    If Not IsNull(res("result")) And IsNumeric(res("result")) Then
        m_uid = CLng(res("result"))
        If m_uid > 0 Then
            Authenticate = m_uid
            Exit Function
        End If
    End If
    
    Err.Raise Number:=ERR_AUTHENTICATION_FAILED, _
              Source:="OdClient", _
              Description:="Authentication failed. Invalid credentials or database."
End Function

''
' Execute a Remote Procedure Call (RPC) to an Odoo model method via JSON-RPC execute_kw.
'
' @param aModel Target Odoo Model (e.g., "res.partner")
' @param aMethod Target Remote Method to invoke (e.g., "search_read", "create")
' @param aArgs Positional arguments for the method (Collection)
' @param aKwargs Optional keyword arguments for the method (Dictionary)
' @return Result data from Odoo
''
Public Function ExecuteKw(aModel As String, aMethod As String, Optional aArgs As Collection = Nothing, Optional aKwargs As Dictionary = Nothing) As Variant
    CheckInitialization
    
    ' Automatically authenticate if not yet authenticated
    If m_uid = 0 Then
        Authenticate
    End If
    
    ' Construct base execute_kw arguments: [db, uid, password, model, method, args, kwargs]
    Dim rpcArgsList As New Collection
    rpcArgsList.Add m_database
    rpcArgsList.Add m_uid
    rpcArgsList.Add m_apiKey
    rpcArgsList.Add aModel
    rpcArgsList.Add aMethod
    
    If aArgs Is Nothing Then
        rpcArgsList.Add New Collection
    Else
        rpcArgsList.Add aArgs
    End If
    
    If Not aKwargs Is Nothing Then
        rpcArgsList.Add aKwargs
    End If
    
    Set ExecuteKw = CallJsonRpc("object", "execute_kw", rpcArgsList)
End Function

' ------------------------------------------------------------------------------
' Private Helper Methods
' ------------------------------------------------------------------------------

Private Sub CheckInitialization()
    If Not m_initialized Then
        Err.Raise Number:=ERR_UNINITIALIZED, _
                  Source:="OdClient", _
                  Description:="OdClient is not initialized. Call Initialize() before executing requests."
    End If
End Sub

''
' Build and send JSON-RPC HTTP POST request to /jsonrpc
'
''
Private Function CallJsonRpc(aService As String, aMethod As String, aArgs As Collection) As Dictionary
    ' Construct JSON-RPC payload
    Dim payload As New Dictionary
    payload.Add "jsonrpc", "2.0"
    payload.Add "method", "call"
    payload.Add "id", Int((100000 * Rnd) + 1)
    
    ' Odoo expects named parameters: service, method, args inside params
    Dim params As New Dictionary
    params.Add "service", aService
    params.Add "method", aMethod
    params.Add "args", aArgs
    payload.Add "params", params
    
    ' Configure headers
    Dim hdrs As New Collection
    hdrs.Add WebHelpers.CreateKeyValue("Content-Type", "application/json; charset=utf-8")
    If Len(m_userAgent) > 0 Then
        hdrs.Add WebHelpers.CreateKeyValue("User-Agent", m_userAgent)
    End If
    
    Dim opts As New Dictionary
    opts.Add "Headers", hdrs
    
    ' Execute HTTP POST request
    Dim resp As WebResponse
    Set resp = m_webClient.PostJson("/jsonrpc", payload, opts)
    
    If resp.StatusCode = WebStatusCode.Ok Then
        If Not resp.Data Is Nothing Then
            Dim jsonResp As Dictionary
            Set jsonResp = resp.Data
            
            ' Check for Odoo JSON-RPC errors
            If jsonResp.Exists("error") Then
                Dim errInfo As Variant
                Set errInfo = jsonResp("error")
                Err.Raise Number:=ERR_ODOO_RPC_ERROR, _
                          Source:="OdClient", _
                          Description:="Odoo RPC Error: " & errInfo("message")
            ElseIf jsonResp.Exists("result") Then
                Set CallJsonRpc = jsonResp
                Exit Function
            Else
                Err.Raise Number:=ERR_INVALID_RESPONSE, _
                          Source:="OdClient", _
                          Description:="Invalid JSON-RPC response structure: missing 'result' or 'error'."
            End If
        Else
            Err.Raise Number:=ERR_EMPTY_RESPONSE, _
                      Source:="OdClient", _
                      Description:="Response data is empty."
        End If
    Else
        Err.Raise Number:=ERR_HTTP_ERROR, _
                  Source:="OdClient", _
                  Description:="HTTP " & resp.StatusCode & ": " & resp.StatusDescription
    End If
End Function
