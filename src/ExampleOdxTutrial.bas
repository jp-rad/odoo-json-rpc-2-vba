Attribute VB_Name = "ExampleOdxTutrial"
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

Private Const CODOO_ATTR_NAME     As String = "name"
Private Const CODOO_ATTR_STRING   As String = "string"
Private Const CODOO_ATTR_TYPE     As String = "type"
Private Const CODOO_ATTR_RELATION As String = "relation"
Private Const CODOO_ATTR_FKEY     As String = "relation_field"   ' foreign key
Private Const CODOO_ATTR_REQUIRED As String = "required"

''
' Runs tutorial examples for OdxModel using standard JSON-RPC API.
''
Public Sub DoTutrialOdxModel()
    ExecuteTutrial CreateOdxRepository(False)
End Sub

''
' Runs tutorial examples for OdxModel using External JSON-2 API.
''
Public Sub DoTutrialOdxModel2()
    ExecuteTutrial CreateOdxRepository(True)
End Sub

''
' Executes the OdxModel tutorial workflow, generating schemas and data sheets.
'
' @param oRepo Initialized OdxRepository instance
''
Private Sub ExecuteTutrial(oRepo As OdxRepository)
    Dim oModelView As OdxModel
    
    Dim oDomain As OdFilterDomain
    Dim oNames As OdFilterDomain
    
    Dim wb As Workbook
    Dim sht As Worksheet
    Dim rng As Range
    
    Dim colNameList As Collection
    Dim colTagList As Collection
    
    Dim v As Variant
    Dim s As String
        
    Set wb = Workbooks.Add
    Set sht = wb.Sheets(1)
    SetWorksheetName sht, "OdxModel Example"
    Set rng = sht.Range("A1")
    
    ' Target names definition
    Set colTagList = New Collection
    Set colNameList = New Collection
    s = "Mitchell"
    colTagList.Add s
    colNameList.Add "Mitchell Admin", s
    s = "Marc"
    colTagList.Add s
    colNameList.Add "Marc Demo", s
    s = "Gemini"
    colTagList.Add s
    colNameList.Add "Gemini Furniture", s
            
    DebugRange rng, "--------------"
    DebugRange rng, " OdxModel"
    DebugRange rng, "--------------"
    
    ' ==================
    ' New OdxModelView
    ' ==================
    Set oModelView = NewOdxModel(oRepo, "res.partner")
    With oModelView
        ' -------------
        ' res.partner
        ' -------------
        DebugRange rng, .ModelName
        DebugRangeSchema rng, .ModelName, oRepo.ModelSchema(.ModelName)
        
        .AddExternalIdField ' (option) External ID
        
        .AddField "name"
        .AddField "is_company"
        .AddField "is_public"
        
        With .AddField("company_id")    ' many2one
            ' --------------------------
            ' res.company
            ' --------------------------
            DebugRange rng, "company_id: many2one --> " & .ModelName
            Debug.Assert .IsMany2One
            DebugRangeSchema rng, .ModelName, oRepo.ModelSchema(.ModelName)
            
            .AddField "name"
            .AddField "city"
            
            With .AddField("currency_id")   ' many2one
                ' --------------------------
                ' res.currency
                ' --------------------------
                DebugRange rng, "company_id: many2one --> currency_id: many2one --> " & .ModelName
                Debug.Assert .IsMany2One
                DebugRangeSchema rng, .ModelName, oRepo.ModelSchema(.ModelName)
                
                .AddField "name"
                .AddField "symbol"

            End With
            
            .AddField "layout_background"

            .AddExternalIdField ' (option) External ID
        End With
        
        With .AddField("child_ids")     ' one2many
            ' --------------------------
            ' res.partner
            ' --------------------------
            DebugRange rng, "child_ids: one2many --> " & .ModelName
            Debug.Assert .IsOne2Many
            DebugRangeSchema rng, .ModelName, oRepo.ModelSchema(.ModelName)
            
            .AddExternalIdField "custom_external_id" ' (option) External ID
        
            .AddField "name"
        End With
                
        With .AddField("category_id")   ' many2many
            ' --------------------------
            ' res.partner.category
            ' --------------------------
            DebugRange rng, "category_id: many2many --> " & .ModelName
            Debug.Assert .IsMany2Many
            DebugRangeSchema rng, .ModelName, oRepo.ModelSchema(.ModelName)
            
            .AddField "name"
        End With
    End With

    ' ==================
    ' Fetch data
    ' ==================
    Set oNames = NewDomain
    For Each v In colNameList
        oNames.AddArity v
    Next v
    Set oDomain = NewDomain
    oDomain.AddArity NewField("name").IsIn(oNames)
    oModelView.FetchSearchAndReadRemote oDomain.Build
    
    ' ==================
    ' All records view
    ' ==================
    oModelView.ClearFilter   ' Unfiltered
    With oModelView
        DebugPrintModelView wb, .RefMe
        With .GetRelatedModel("company_id", True)
            DebugPrintModelView wb, .RefMe
            With .GetRelatedModel("currency_id", True)
                DebugPrintModelView wb, .RefMe
            End With
        End With
        With .GetRelatedModel("child_ids", True)
            DebugPrintModelView wb, .RefMe
        End With
        With .GetRelatedModel("category_id", True)
            DebugPrintModelView wb, .RefMe
        End With
    End With
    
    ' ==================
    ' Filtered records view
    ' ==================
    For Each v In colTagList
        s = colNameList(CStr(v))
        
        oModelView.SetFilter "name = '" & s & "'"   ' Filtered
        With oModelView
            DebugPrintModelView wb, .RefMe, CStr(v)
            With .GetRelatedModel("company_id")
                DebugPrintModelView wb, .RefMe, v & "(company_id)"
                With .GetRelatedModel("currency_id")
                    DebugPrintModelView wb, .RefMe, v & "(currency_id)"
                End With
            End With
            With .GetRelatedModel("child_ids")
                DebugPrintModelView wb, .RefMe, v & "(child_ids)"
            End With
            With .GetRelatedModel("category_id")
                DebugPrintModelView wb, .RefMe, v & "(category_id)"
            End With
        End With
    Next v
    
    sht.Select
    wb.Saved = True
    wb.Activate
End Sub

''
' Outputs a value to the debug window and writes it to the worksheet range.
'
' @param rng Worksheet Range reference
' @param v Value to output
''
Private Sub DebugRange(ByRef rng As Range, v As Variant)
    Debug.Print v
    rng = v
    Set rng = rng.Offset(1, 0)
End Sub

''
' Writes model schema definitions to the worksheet.
'
' @param rng Worksheet Range reference
' @param aModelName Target model name
' @param dicModelSchema Dictionary of field schema definitions
''
Private Sub DebugRangeSchema(ByRef rng As Range, aModelName As String, dicModelSchema As Dictionary)
    ' Debug.Print WebHelpers.ConvertToJson(dicModelSchema)
    
    Dim tag As String
    Dim i As Long
    Dim dic As Dictionary
    
    tag = aModelName
    
    ' Header
    rng.Offset(0, 0) = tag
    rng.Offset(0, 1) = "#"
    rng.Offset(0, 2) = CODOO_ATTR_NAME
    rng.Offset(0, 3) = CODOO_ATTR_REQUIRED
    rng.Offset(0, 4) = CODOO_ATTR_TYPE
    rng.Offset(0, 5) = CODOO_ATTR_RELATION
    rng.Offset(0, 6) = CODOO_ATTR_FKEY
    rng.Offset(0, 7) = CODOO_ATTR_STRING
    ' Next row
    Set rng = rng.Offset(1, 0)
    
    ' Fields
    For i = 1 To dicModelSchema.Count
        Set dic = dicModelSchema.Items(i - 1)
        rng.Offset(0, 0) = tag
        rng.Offset(0, 1) = i
        rng.Offset(0, 2) = dic(CODOO_ATTR_NAME)
        rng.Offset(0, 3) = dic(CODOO_ATTR_REQUIRED)
        rng.Offset(0, 4) = dic(CODOO_ATTR_TYPE)
        rng.Offset(0, 5) = dic(CODOO_ATTR_RELATION)
        rng.Offset(0, 6) = dic(CODOO_ATTR_FKEY)
        rng.Offset(0, 7) = dic(CODOO_ATTR_STRING)
        ' Next row
        Set rng = rng.Offset(1, 0)
    Next i
End Sub

''
' Attempts to rename a worksheet safely, ignoring errors.
'
' @param aWs Worksheet instance
' @param aName Desired worksheet name
' @return Boolean True if successful
''
Private Function TrySetWorksheetName(aWs As Worksheet, aName As String) As Boolean
On Error GoTo ErrHandler
    aWs.Name = aName
    TrySetWorksheetName = True
ExitProc:
    Exit Function
ErrHandler:
    Resume ExitProc
End Function

''
' Sets a worksheet name uniquely, appending an incremental suffix if a conflict occurs.
'
' @param aWs Worksheet instance
' @param aName Base worksheet name
''
Private Function SetWorksheetName(aWs As Worksheet, aName As String)
    Dim i As Long
    If Not TrySetWorksheetName(aWs, aName) Then
        i = 2
        Do Until TrySetWorksheetName(aWs, aName & " (" & i & ")")
            i = i + 1
        Loop
    End If
End Function

''
' Dumps recordset contents of an OdxModel view into a new worksheet.
'
' @param aWb Workbook instance
' @param aMv OdxModel instance
' @param optSheetName Optional custom sheet name
''
Private Sub DebugPrintModelView(aWb As Workbook, aMv As OdxModel, Optional optSheetName As String = "")
    Dim rs As ADODB.Recordset
    Dim sht As Worksheet
    Dim rng As Range
    Dim i As Long
    
    Dim sSheetName As String
    
    ' Clone and Filtered
    Set rs = aMv.Recordset.Clone
    rs.Filter = aMv.Recordset.Filter
    
    If optSheetName = "" Then
        sSheetName = aMv.ModelName
    Else
        sSheetName = optSheetName
    End If
    Set sht = aWb.Sheets.Add(After:=aWb.Worksheets(aWb.Worksheets.Count))
    SetWorksheetName sht, sSheetName
    
    Set rng = sht.Range("A1")
    For i = 0 To rs.Fields.Count - 1
        rng.Offset(0, i) = rs.Fields(i).Name
    Next i
    
    If Not rs.BOF Then
        rs.MoveFirst
    End If
    rng.Offset(1, 0).CopyFromRecordset rs
    
End Sub
