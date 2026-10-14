Attribute VB_Name = "SpecOdJsonRpc"
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
' Module: SpecOdJsonRpc
' Purpose: Contains unit test specifications for Odoo External ID utility functions
' Standard:
'   - Validates module and object name separation via SplitExternalId.
'   - Validates formatting, normalization, and sanitization via NormalizeExternalId.
' ------------------------------------------------------------------------------

''
' Register and run all External ID utility test specifications.
'
' @param Suite TestSuite instance
''
Public Sub RegisterSpecs(Suite As TestSuite)
    SpecSplitExternalId Suite
    SpecNormalizeExternalId Suite
End Sub

' ------------------------------------------------------------------------------
' Test Specifications
' ------------------------------------------------------------------------------

''
' Tests for SplitExternalId function.
'
' @param Suite TestSuite instance
''
Private Sub SpecSplitExternalId(Suite As TestSuite)
    Dim Tests As New TestSuite
    Dim Test As TestCase
    Dim splitRes() As String
    
    With Suite.Test("TestSplitExternalId - Module and Object Separation")
        Set Test = Tests.Test("should pass")
        
        ' 1. Normal case with explicit module prefix
        splitRes = SplitExternalId("base.partner_root")
        Test.NotUndefined splitRes
        Test.IsEqual splitRes(0), "base"
        Test.IsEqual splitRes(1), "partner_root"
        
        ' 2. Whitespace trimming inside parts
        splitRes = SplitExternalId("  base  .  partner_root  ")
        Test.IsEqual splitRes(0), "base"
        Test.IsEqual splitRes(1), "partner_root"
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    ' Error handling tests for SplitExternalId (must contain a dot)
    With Suite.Test("TestSplitExternalId - Error Handling")
        Set Test = Tests.Test("should raise error for missing dot or empty")
        
        ' Note: Depending on your TestSuite framework, you would wrap these in On Error Resume Next / Assert Err.Number
        ' Example check for missing dot (should throw ERR_API_EXECUTION)
        On Error Resume Next
        SplitExternalId "partner_123"
        Test.IsEqual Err.Number, ERR_API_EXECUTION
        
        SplitExternalId ""
        Test.IsEqual Err.Number, ERR_API_EXECUTION
        
        SplitExternalId "base."
        Test.IsEqual Err.Number, ERR_API_EXECUTION
        On Error GoTo 0
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
End Sub

''
' Tests for NormalizeExternalId function.
'
' @param Suite TestSuite instance
''
Private Sub SpecNormalizeExternalId(Suite As TestSuite)
    Dim Tests As New TestSuite
    Dim Test As TestCase
    
    With Suite.Test("TestNormalizeExternalId - Formatting and Sanitization")
        Set Test = Tests.Test("should pass")
        
        ' Standard external ID remains unchanged
        Test.IsEqual NormalizeExternalId("base.partner_root"), "base.partner_root"
        
        ' Missing module prepends default module namespace (DEFAULT_MODULE_IMPORT)
        Test.IsEqual NormalizeExternalId("res_partner_123"), "__import__" & ".res_partner_123"
        
        ' Multiple dots and spaces in object name are sanitized (e.g., "base.partner.root" -> "base.partner_root")
        Test.IsEqual NormalizeExternalId("base.partner.root"), "base.partner_root"
        
        ' Whitespace trimming and space-to-underscore replacement in module and object
        Test.IsEqual NormalizeExternalId("  my module . item # 1   "), "my_module.item_#_1"
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
    
    ' Error handling tests for NormalizeExternalId
    With Suite.Test("TestNormalizeExternalId - Error Handling")
        Set Test = Tests.Test("should raise error for empty or invalid format")
        
        On Error Resume Next
        NormalizeExternalId ""
        Test.IsEqual Err.Number, ERR_API_EXECUTION
        
        NormalizeExternalId "base."
        Test.IsEqual Err.Number, ERR_API_EXECUTION
        On Error GoTo 0
        
        .IsEqual Test.Result, TestResultType.Pass
    End With
End Sub
