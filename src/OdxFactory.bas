Attribute VB_Name = "OdxFactory"
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
' Module: OdxFactory
' ==============================================================================
' Description:
'   Handles configuration loading from external text files and provides
'   factory methods to instantiate both `OdClient` and `Od2Client` instances.
' ==============================================================================
''

Private Const CONFIG_URL        As String = "URL"
Private Const CONFIG_DB         As String = "DB"
Private Const CONFIG_USERNAME   As String = "USERNAME"
Private Const CONFIG_APIKEY     As String = "APIKEY"
Private Const CONFIG_INSECURE   As String = "INSECURE"
Private Const CONFIG_TIMEOUTMS  As String = "TIMEOUTMS"

Private Const DEFAULT_URL       As String = "https://localhost"
Private Const DEFAULT_DB        As String = "dev_odoo"
Private Const DEFAULT_USERNAME  As String = "admin"
Private Const DEFAULT_APIKEY    As String = "admin"
Private Const DEFAULT_TIMEOUTMS As Long = 5000
Private Const DEFAULT_INSECURE  As Boolean = False

Private pOdUrl       As String
Private pOdDb        As String
Private pOdUserName  As String
Private pOdApiKey    As String
Private pOdTimeoutMs As String
Private pOdInsecure  As String

Private pLoadedConfigPath As String

''
' ------------------------------------------------------------------------------
' Configuration Getters
' ------------------------------------------------------------------------------
''

Public Property Get OdUrl() As String
    If Len(pOdUrl) > 0 Then
        OdUrl = pOdUrl
    Else
        OdUrl = DEFAULT_URL
    End If
End Property

Public Property Get OdDb() As String
    If Len(pOdDb) > 0 Then
        OdDb = pOdDb
    Else
        OdDb = DEFAULT_DB
    End If
End Property

Public Property Get OdUserName() As String
    If Len(pOdUserName) > 0 Then
        OdUserName = pOdUserName
    Else
        OdUserName = DEFAULT_USERNAME
    End If
End Property

Public Property Get OdApiKey() As String
    If Len(pOdApiKey) > 0 Then
        OdApiKey = pOdApiKey
    Else
        OdApiKey = DEFAULT_APIKEY
    End If
End Property

Public Property Get OdTimeoutMs() As Long
On Error Resume Next
    OdTimeoutMs = DEFAULT_TIMEOUTMS
    OdTimeoutMs = CLng(pOdTimeoutMs)
End Property

Public Property Get OdInsecure() As Boolean
On Error Resume Next
    OdInsecure = DEFAULT_INSECURE
    OdInsecure = CBool(pOdInsecure)
End Property


''
' ------------------------------------------------------------------------------
' Configuration Loader & Initializer
' ------------------------------------------------------------------------------
''

''
' Loads configuration parameters from a specified text file.
' Creates a default configuration file if it does not exist.
'
' @param configPath Path to the configuration text file
''
Private Sub LoadConfigFromTextFile(configPath As String)
    If Dir(configPath) = "" Then
        CreateDefaultConfigFile configPath
    End If
    
    Dim fileNum As Integer
    fileNum = FreeFile
    Open configPath For Input As #fileNum
    
    Do Until EOF(fileNum)
        Dim lineText As String
        Line Input #fileNum, lineText
        
        lineText = Trim$(lineText)
        If Len(lineText) > 0 And Left$(lineText, 1) <> "#" And InStr(lineText, "=") > 0 Then
            Dim keyName As String
            Dim valStr As String
            
            keyName = UCase$(Trim$(Split(lineText, "=")(0)))
            valStr = Trim$(Mid$(lineText, InStr(lineText, "=") + 1))
            
            If Len(valStr) > 0 Then
                Select Case keyName
                    Case CONFIG_URL:        pOdUrl = valStr
                    Case CONFIG_DB:         pOdDb = valStr
                    Case CONFIG_USERNAME:   pOdUserName = valStr
                    Case CONFIG_APIKEY:     pOdApiKey = valStr
                    Case CONFIG_INSECURE:   pOdInsecure = valStr
                    Case CONFIG_TIMEOUTMS:  pOdTimeoutMs = valStr
                End Select
            End If
        End If
    Loop
    Close #fileNum
End Sub

''
' Creates a default configuration file with template values.
'
' @param FilePath Path where the default configuration file should be created
''
Private Sub CreateDefaultConfigFile(FilePath As String)
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open FilePath For Output As #fileNum
    Print #fileNum, "--------------------"
    Print #fileNum, " External API"
    Print #fileNum, "--------------------"
    Print #fileNum, CONFIG_URL & "=" & DEFAULT_URL
    Print #fileNum, CONFIG_DB & "=" & DEFAULT_DB
    Print #fileNum, CONFIG_USERNAME & "=" & DEFAULT_USERNAME
    Print #fileNum, CONFIG_APIKEY & "=" & DEFAULT_APIKEY
    Print #fileNum, CONFIG_TIMEOUTMS & "=" & DEFAULT_TIMEOUTMS
    Print #fileNum, CONFIG_INSECURE & "=" & DEFAULT_INSECURE
    Close #fileNum
End Sub

''
' ------------------------------------------------------------------------------
' Client Factory Methods
' ------------------------------------------------------------------------------
''

''
' Creates and initializes an OdClient instance.
'
' @param aUrl Optional Odoo server URL
' @param aDb Optional database name
' @param aUserName Optional username
' @param aKey Optional API key or password
' @param aConfigPath Optional path to configuration file
' @return OdClient Initialized OdClient instance
''
Public Function CreateOdClient( _
    Optional aUrl As String = "", Optional aDb As String = "", _
    Optional aUserName As String = "", Optional aKey As String = "", _
    Optional aConfigPath As String = "" _
) As OdClient
    Dim targetUrl As String: targetUrl = aUrl
    Dim targetDb As String: targetDb = aDb
    Dim targetUserName As String: targetUserName = aUserName
    Dim targetKey As String: targetKey = aKey
    
    If Len(aConfigPath) > 0 Then
        If pLoadedConfigPath <> aConfigPath Then
            LoadConfigFromTextFile aConfigPath
            pLoadedConfigPath = aConfigPath
        End If
    End If
    
    If Len(targetUrl) = 0 Then targetUrl = OdUrl
    If Len(targetDb) = 0 Then targetDb = OdDb
    If Len(targetUserName) = 0 Then targetUserName = OdUserName
    If Len(targetKey) = 0 Then targetKey = OdApiKey
    
    Set CreateOdClient = NewOdClient
    With CreateOdClient
        .Initialize targetUrl, targetDb, targetUserName, targetKey
        With .RefWebClient
            .TimeoutMs = OdTimeoutMs
            .Insecure = OdInsecure
        End With
    End With
End Function

''
' Creates and initializes an Od2Client instance (External JSON-2 API).
'
' Note:
'   External API v2 does not require or use a username for authentication;
'   it authenticates using the database and API key (or token) only.
'
' @param aUrl Optional Odoo server URL
' @param aDb Optional database name
' @param aKey Optional API key
' @param aConfigPath Optional path to configuration file
' @return Od2Client Initialized Od2Client instance
''
Public Function CreateOd2Client( _
    Optional aUrl As String = "", Optional aDb As String = "", Optional aKey As String = "", _
    Optional aConfigPath As String = "" _
) As Od2Client
    Dim targetUrl As String: targetUrl = aUrl
    Dim targetDb As String: targetDb = aDb
    Dim targetKey As String: targetKey = aKey
    
    If Len(aConfigPath) > 0 Then
        If pLoadedConfigPath <> aConfigPath Then
            LoadConfigFromTextFile aConfigPath
            pLoadedConfigPath = aConfigPath
        End If
    End If
    
    If Len(targetUrl) = 0 Then targetUrl = OdUrl
    If Len(targetDb) = 0 Then targetDb = OdDb
    If Len(targetKey) = 0 Then targetKey = OdApiKey
    
    ' Note: External API v2 authentication is handled without a username.
    Set CreateOd2Client = NewOd2Client
    With CreateOd2Client
        .Initialize targetUrl, targetDb, targetKey
        With .RefWebClient
            .TimeoutMs = OdTimeoutMs
            .Insecure = OdInsecure
        End With
    End With
End Function
