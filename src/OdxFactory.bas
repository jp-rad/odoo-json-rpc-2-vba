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

Private Const CONFIG_URL      As String = "URL"
Private Const CONFIG_DB       As String = "DB"
Private Const CONFIG_USERNAME As String = "USERNAME"
Private Const CONFIG_KEY      As String = "KEY"

Private Const DEFAULT_URL      As String = "https://server"
Private Const DEFAULT_DB       As String = "db_name"
Private Const DEFAULT_USERNAME As String = "user_name"
Private Const DEFAULT_KEY      As String = "api_key"

Private pOdUrl       As String
Private pOdDb        As String
Private pOdUserName  As String
Private pOdKey       As String

Private pLoadedConfigPath As String

''
' ------------------------------------------------------------------------------
' Configuration Getters
' ------------------------------------------------------------------------------
''

Public Property Get OdUrl() As String
    OdUrl = pOdUrl
End Property

Public Property Get OdDb() As String
    OdDb = pOdDb
End Property

Public Property Get OdUserName() As String
    OdUserName = pOdUserName
End Property

Public Property Get OdKey() As String
    OdKey = pOdKey
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
                    Case CONFIG_URL:      pOdUrl = valStr
                    Case CONFIG_DB:       pOdDb = valStr
                    Case CONFIG_USERNAME: pOdUserName = valStr
                    Case CONFIG_KEY:      pOdKey = valStr
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
    Print #fileNum, CONFIG_KEY & "=" & DEFAULT_KEY
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
    
    If Len(targetUrl) = 0 Then targetUrl = pOdUrl
    If Len(targetDb) = 0 Then targetDb = pOdDb
    If Len(targetUserName) = 0 Then targetUserName = pOdUserName
    If Len(targetKey) = 0 Then targetKey = pOdKey
    
    Set CreateOdClient = NewOdClient
    CreateOdClient.Initialize targetUrl, targetDb, targetUserName, targetKey
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
    
    If Len(targetUrl) = 0 Then targetUrl = pOdUrl
    If Len(targetDb) = 0 Then targetDb = pOdDb
    If Len(targetKey) = 0 Then targetKey = pOdKey
    
    ' Note: External API v2 authentication is handled without a username.
    Set CreateOd2Client = NewOd2Client
    CreateOd2Client.Initialize targetUrl, targetDb, targetKey
End Function
