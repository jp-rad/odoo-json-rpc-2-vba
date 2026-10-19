Attribute VB_Name = "SpecRunner"
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
' Module: SpecRunner
' Purpose: Entry point and central configuration provider for unit tests
' ------------------------------------------------------------------------------

''
' Main Test Runner Entrypoint for all suites.
''
Public Sub DoAll()
    DoAllType1
    DoAllType2
End Sub

Public Sub DoAllType1()
    DoTutorialExternalApi1
    DoTutrialOdxModel
    RunSpecsType1
End Sub

Public Sub DoAllType2()
    DoTutorialExternalApi2
    DoTutrialOdxModel2
    RunSpecsType2
End Sub

''
' Main Test Runner Entrypoint for Type 1 suite.
''
Public Sub RunSpecsType1()
    Dim Suite As New TestSuite
    Suite.Description = "OdClient Test Suite"
        
    ' Register specs from external spec modules
    SpecOdFilter.RegisterSpecs Suite
    SpecJsonUtils.RegisterSpecs Suite
    SpecOdJsonRpc.RegisterSpecs Suite

    SpecOdClient.RegisterSpecs Suite
    SpecOdRepository.RegisterSpecs Suite
    
    SpecOdxRepository.RegisterSpecs Suite, False
    SpecOdxModelRemoteSync.RegisterSpecs Suite, False
    
    PrintSummary Suite
End Sub

''
' Main Test Runner Entrypoint for Type 2 suite.
''
Public Sub RunSpecsType2()
    Dim Suite As New TestSuite
    Suite.Description = "Od2Client Test Suite"
        
    ' Register specs from external spec modules
    SpecOdFilter.RegisterSpecs Suite
    SpecJsonUtils.RegisterSpecs Suite
    SpecOdJsonRpc.RegisterSpecs Suite

    SpecOd2Client.RegisterSpecs Suite
    SpecOd2Repository.RegisterSpecs Suite
    
    SpecOdxRepository.RegisterSpecs Suite, True
    SpecOdxModelRemoteSync.RegisterSpecs Suite, True
    
    PrintSummary Suite
End Sub

''
' ------------------------------------------------------------------------------
' Helper
' ------------------------------------------------------------------------------
''

''
' Prints the test summary and failed test details to the debug window.
'
' @param Suite TestSuite instance
''
Private Sub PrintSummary(Suite As TestSuite)
    Debug.Print "========================================"
    Debug.Print " Test Suite: " & Suite.Description
    Debug.Print "----------------------------------------"
    Debug.Print " Passed : " & Suite.PassedTests.Count
    Debug.Print " Failed : " & Suite.FailedTests.Count
    Debug.Print " Pending: " & Suite.PendingTests.Count
    Debug.Print " Total  : " & Suite.Tests.Count
    Debug.Print "========================================"
    
    If Suite.FailedTests.Count > 0 Then
        Debug.Print ""
        Debug.Print "[FAILED TESTS SUMMARY]"
        
        Dim testObj As Object
        For Each testObj In Suite.FailedTests
            Dim v As Variant
            For Each v In testObj.Failures
                Debug.Print " - " & v
            Next v
        Next testObj
        
        Debug.Print "========================================"
    End If
End Sub

''
' Validates whether a given string is valid JSON format.
'
' @param aJson JSON string to validate
' @return Boolean True if valid JSON
''
Public Function IsValidJson(aJson As String) As Boolean
On Error GoTo ErrHandler
    WebHelpers.ParseJson aJson
    IsValidJson = True
ExitProc:
    Exit Function
ErrHandler:
    Debug.Print "Invalid JSON:", aJson
    Debug.Print "----- Err.Description -----"
    Debug.Print Err.Description
    Debug.Print "---------------------------"
    Resume ExitProc
End Function

''
' Asserts that two JSON strings are equivalent after parsing and re-serializing.
'
' @param oTest TestCase instance
' @param aExcept Expected JSON string
' @param aActual Actual JSON string
''
Public Sub TestIsEqualJson(oTest As TestCase, aExcept As String, aActual As String)
On Error Resume Next
    Debug.Assert IsValidJson(aExcept)
    Debug.Assert IsValidJson(aActual)
    oTest.IsEqual WebHelpers.ConvertToJson(WebHelpers.ParseJson(aExcept)), _
        WebHelpers.ConvertToJson(WebHelpers.ParseJson(aActual))
End Sub

''
' Asserts that two JSON comma-separated strings are equivalent.
'
' @param oTest TestCase instance
' @param aExcept Expected JSON CSV string
' @param aActual Actual JSON CSV string
''
Public Sub TestIsEqualJsonCsv(oTest As TestCase, aExcept As String, aActual As String)
    TestIsEqualJson oTest, "[" & aExcept & "]", "[" & aActual & "]"
End Sub

''
' ------------------------------------------------------------------------------
' Factory
' ------------------------------------------------------------------------------
''

''
' Creates and configures an OdClient instance for testing.
'
' @return OdClient Initialized OdClient instance
''
Public Function CreateOdClient() As OdClient
    Dim configPath As String
    
    configPath = ThisWorkbook.FullName & ".secret.txt"
    Set CreateOdClient = OdxFactory.CreateOdClient(aConfigPath:=configPath)

End Function

''
' Creates and configures an Od2Client instance for testing.
'
' @return Od2Client Initialized Od2Client instance
''
Public Function CreateOd2Client() As Od2Client
    Dim configPath As String
    
    configPath = ThisWorkbook.FullName & ".secret.txt"
    Set CreateOd2Client = OdxFactory.CreateOd2Client(aConfigPath:=configPath)

End Function
