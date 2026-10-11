Attribute VB_Name = "ExampleOd2Tutrial"
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

'
' External API - odoo docs
'
' Odoo is usually extended internally via modules, but many of its features and all of its data
' are also available from the outside for external analysis or integration with various tools.
' Part of the Models API is easily available over XML-RPC and accessible from a variety of languages.
'
' https://www.odoo.com/documentation/master/developer/reference/external_api.html
'

''
' Runs tutorial examples for Odoo External JSON-2 API operations.
' Reference: https://www.odoo.com/documentation/master/developer/reference/external_api.html#calling-methods
''
Public Sub DoTutorialExternalApi2()
    Dim objClient As Od2Client
    Dim objRepo As Od2Repository
    Dim retDict As Dictionary
    
    Dim paramsDict As Dictionary
    Dim domainList As Collection
    Dim fieldsList As Collection
    Dim attrList As Collection
    Dim valuesDict As Dictionary
    Dim ids As Collection
    
    Dim nId As Long

    ' --------------------------------------------------
    ' Direct instantiation and initialization of Repository
    ' --------------------------------------------------
    Dim configPath As String
    configPath = ThisWorkbook.FullName & ".secret.txt"
    Set objClient = OdxFactory.CreateOd2Client(aConfigPath:=configPath)
        
    Set objRepo = NewOd2Repository
    objRepo.Initialize objClient

    ' --------------------------------------------------
    ' Get Version
    ' --------------------------------------------------
    Debug.Print "---------"
    Debug.Print " version"
    Debug.Print "---------"
    
    Set retDict = objClient.GetVersion
    
    Debug.Print WebHelpers.ConvertToJson(retDict, 4)
    Debug.Print

    ' --------------------------------------------------
    ' execute_kw
    ' python: models.execute_kw(db, uid, password, 'res.partner', 'name_search', ['foo'], {'limit': 10})
    ' --------------------------------------------------
    Debug.Print "------------"
    Debug.Print " execute_kw"
    Debug.Print "------------"
    
    Set paramsDict = NewDict
    paramsDict.Add "name", "Azure"   ' ['foo']
    paramsDict.Add "limit", 1        ' {'limit': 10}
    Set retDict = objRepo.ExecuteModelMethod("res.partner", "name_search", paramsDict)
    
    Debug.Print WebHelpers.ConvertToJson(retDict, 4)
    Debug.Print
    
    ' --------------------------------------------------
    ' List records - search
    ' python: models.execute_kw(db, uid, password, 'res.partner', 'search', [[['is_company', '=', True]]])
    ' --------------------------------------------------
    Debug.Print "--------------"
    Debug.Print " List records"
    Debug.Print "--------------"
    
    With NewDomain
        .AddArity NewField("is_company").Eq(True)   ' ['is_company', '=', True]
        Set domainList = .Build
    End With
    Set retDict = objRepo.SearchRecord("res.partner", aDomain:=domainList)
    
    Debug.Print WebHelpers.ConvertToJson(retDict, 4)
    Debug.Print
    
    ' --------------------------------------------------
    ' Pagination - search
    ' python: models.execute_kw(db, uid, password, 'res.partner', 'search', [[['is_company', '=', True]]], {'offset': 10, 'limit': 5})
    ' --------------------------------------------------
    Debug.Print "------------"
    Debug.Print " Pagination"
    Debug.Print "------------"
    
    With NewDomain
        .AddArity NewField("is_company").Eq(True)   ' ['is_company', '=', True]
        Set domainList = .Build
    End With
    ' {'offset': 10, 'limit': 5}
    Set retDict = objRepo.SearchRecord("res.partner", aDomain:=domainList, aOffset:=3, aLimit:=5)
    
    Debug.Print WebHelpers.ConvertToJson(retDict, 4)
    Debug.Print
    
    ' --------------------------------------------------
    ' Count records - search_count
    ' python: models.execute_kw(db, uid, password, 'res.partner', 'search_count', [[['is_company', '=', True]]])
    ' --------------------------------------------------
    Debug.Print "---------------"
    Debug.Print " Count records"
    Debug.Print "---------------"
    
    Set paramsDict = NewDict
    With NewDomain
        .AddArity NewField("is_company").Eq(True)   ' ['is_company', '=', True]
        Set domainList = .Build
    End With
    paramsDict.Add "domain", domainList
    Set retDict = objRepo.ExecuteModelMethod("res.partner", "search_count", paramsDict)
    
    Debug.Print WebHelpers.ConvertToJson(retDict, 4)
    Debug.Print
    
    ' --------------------------------------------------
    ' Read records - search, read
    ' python: ids = models.execute_kw(db, uid, password, 'res.partner', 'search', [[['is_company', '=', True]]], {'limit': 1})
    '         [record] = models.execute_kw(db, uid, password, 'res.partner', 'read', [ids])
    '         # count the number of fields fetched by default
    '         len(record)
    '         models.execute_kw(db, uid, password, 'res.partner', 'read', [ids], {'fields': ['name', 'country_id', 'comment']})
    ' --------------------------------------------------
    Debug.Print "--------------"
    Debug.Print " Read records"
    Debug.Print "--------------"
    
    With NewDomain
        .AddArity NewField("is_company").Eq(True)   ' ['is_company', '=', True]
        Set domainList = .Build
    End With
    Set retDict = objRepo.SearchRecord("res.partner", aDomain:=domainList, aLimit:=1)
    nId = retDict("result")(1)
    
    Set retDict = objRepo.ReadRecord("res.partner", nId)    ' [ids]
    
    Debug.Print "ids:", nId
    Debug.Print WebHelpers.ConvertToJson(retDict, 4)
    Debug.Print
    
    Set fieldsList = NewList ' {'fields': ['name', 'country_id', 'comment']}
    With fieldsList
        .Add "name"
        .Add "country_id"
        .Add "comment"
    End With
    Set retDict = objRepo.ReadRecord("res.partner", nId, aFields:=fieldsList)
    Debug.Print "----- fields -----"
    Debug.Print WebHelpers.ConvertToJson(retDict, 4)
    Debug.Print
    
    ' --------------------------------------------------
    ' List record fields - fields_get
    ' python: models.execute_kw(db, uid, password, 'res.partner', 'fields_get', [], {'attributes': ['string', 'help', 'type']})
    ' --------------------------------------------------
    Debug.Print "--------------------"
    Debug.Print " List record fields"
    Debug.Print "--------------------"
    
    Set attrList = NewList ' {'attributes': ['string', 'help', 'type']}
    With attrList
        .Add "string"
        .Add "help"
        .Add "type"
    End With
    Set retDict = objRepo.GetModelFields("res.partner", aAttributes:=attrList)
    
    Debug.Print WebHelpers.ConvertToJson(retDict, 4)
    Debug.Print
    
    ' --------------------------------------------------
    ' Search and read - search_read
    ' python: models.execute_kw(db, uid, password, 'res.partner', 'search_read', [[['is_company', '=', True]]], {'fields': ['name', 'country_id', 'comment'], 'limit': 5})
    ' --------------------------------------------------
    Debug.Print "-----------------"
    Debug.Print " Search and read"
    Debug.Print "-----------------"
    
    Set fieldsList = NewList ' {'fields': ['name', 'country_id', 'comment'], 'limit': 5}
    With fieldsList
        .Add "name"
        .Add "country_id"
        .Add "comment"
    End With
    Set retDict = objRepo.SearchReadRecord("res.partner", aFields:=fieldsList, aLimit:=5)
    
    Debug.Print WebHelpers.ConvertToJson(retDict, 4)
    Debug.Print
    
    ' --------------------------------------------------
    ' Create records - create
    ' python: id = models.execute_kw(db, uid, password, 'res.partner', 'create', [{'name': "New Partner"}])
    ' --------------------------------------------------
    Debug.Print "----------------"
    Debug.Print " Create records"
    Debug.Print "----------------"
    
    ' [{'name': "New Partner"}]
    Set valuesDict = NewDict
    valuesDict.Add "name", "New Partner"
    Set retDict = objRepo.CreateRecord("res.partner", aValues:=valuesDict)
    nId = retDict("result")(1)
    
    Debug.Print "id:", nId
    Debug.Print
    
    ' --------------------------------------------------
    ' Update records - write
    ' python: models.execute_kw(db, uid, password, 'res.partner', 'write', [[id], {'name': "Newer partner"}])
    '         # get record name after having changed it
    '         models.execute_kw(db, uid, password, 'res.partner', 'read', [[id], ['display_name']])
    ' --------------------------------------------------
    Debug.Print "----------------"
    Debug.Print " Update records"
    Debug.Print "----------------"
    
    ' [[id], {'name': "Newer partner"}]
    Set ids = NewList
    ids.Add nId
    
    Set valuesDict = NewDict
    valuesDict.Add "name", "Newer parther"
    
    Set retDict = objRepo.WriteRecord("res.partner", ids, valuesDict)
    
    Debug.Print WebHelpers.ConvertToJson(retDict, 4)
    Debug.Print
    
    ' # get record name after having changed it
    ' [[id], ['display_name']]
    Set ids = NewList
    ids.Add nId
    
    Set fieldsList = NewList
    fieldsList.Add "display_name"
    
    Set retDict = objRepo.ReadRecord("res.partner", ids, aFields:=fieldsList)
    
    Debug.Print WebHelpers.ConvertToJson(retDict, 4)
    Debug.Print
    
    ' --------------------------------------------------
    ' Delete records - unlink
    ' python: models.execute_kw(db, uid, password, 'res.partner', 'unlink', [[id]])
    '         # check if the deleted record is still in the database
    '         models.execute_kw(db, uid, password, 'res.partner', 'search', [[['id', '=', id]]])
    ' --------------------------------------------------
    Debug.Print "----------------"
    Debug.Print " Delete records"
    Debug.Print "----------------"
    
    ' [[id]]
    Set retDict = objRepo.UnlinkRecord("res.partner", nId)
    Debug.Print WebHelpers.ConvertToJson(retDict, 4)
    Debug.Print
    
    ' # check if the deleted record is still in the database
    Set retDict = objRepo.ReadRecord("res.partner", ids)
    
    Debug.Print WebHelpers.ConvertToJson(retDict, 4)
    Debug.Print
            
End Sub
