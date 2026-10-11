# Odoo External API Client and Data Access Library for Excel VBA

Seamlessly connect Excel and VBA applications to Odoo ERP with advanced ADODB Recordset mapping, schema caching, and hierarchical relation traversing.

`odoo-json-rpc-2-vba` is a robust, production-grade library designed to bridge Microsoft Excel (via VBA) and Odoo ERP systems. Traditionally, interacting with enterprise ERP endpoints from a desktop spreadsheet application involves handling complex network requests, manual JSON parsing, and tedious data structuring. This library solves these challenges by providing a clean, modular, and object-oriented architecture tailored specifically for VBA developers.

By abstracting low-level network communication and mapping remote Odoo models directly to local ADODB Recordsets, the library empowers developers to perform complex searches, reads, writes, and relational data traversals with minimal boilerplate code.

- [Odoo Docs - External JSON-2 API](https://www.odoo.com/documentation/master/developer/reference/external_api.html)
- [Odoo Docs - External RPC API](https://www.odoo.com/documentation/20.0/developer/reference/external_rpc_api.html)
- odoo-json-rpc-2-vba (GitHub repository: [https://github.com/jp-rad/odoo-json-rpc-2-vba](https://github.com/jp-rad/odoo-json-rpc-2-vba))


**License:** MIT License  
**Copyright:** (c) 2022-2027 jp-rad  


## Module Architecture

`odoo-json-rpc-2-vba` is structured into architectural layers separating low-level API communication, repository abstraction, and high-level data view management.

### 1. Low-Level Clients
- **`OdClient`**: Manages communication with Odoo's traditional JSON-RPC External API (authentication via username and API key, calling model methods via `execute_kw`).
- **`Od2Client`**: Manages communication with Odoo's new External JSON-2 API (optimized endpoint using database and API key/token authentication without requiring a username).

### 2. Repositories (Data Access Layer)
- **`OdRepository`**: Wraps `OdClient` to provide a standard repository interface for traditional API models.
- **`Od2Repository`**: Wraps `Od2Client` to provide a repository interface tailored for the JSON-2 API.

### 3. Facade & Model View Layer
- **`OdxRepository`**: A unified facade interface that abstracts away the underlying differences between `OdRepository` and `Od2Repository`, allowing your application code to interact with Odoo uniformly regardless of the API version used. It dynamically delegates operations and includes built-in model schema caching.
- **`OdxModel`**: High-level model view controller that integrates with ADODB Recordsets. It handles model schemas, filters, data fetching, and hierarchical relation traversing seamlessly within Excel VBA.


## OdxRepository Core Methods

`OdxRepository` provides a unified set of methods to interact with Odoo models regardless of whether the backend uses `OdRepository` or `Od2Repository`:

| Method Name | Description |
| :--- | :--- |
| `Initialize(aRepository)` | Initializes the facade wrapper with either an `OdRepository` or `Od2Repository` instance. |
| `IsInitialized()` | Returns a Boolean indicating whether the repository has been properly initialized. |
| `GetModelFields(aModel, [aFields], [aAttributes])` | Retrieves raw field definitions and attributes for the specified model via Odoo's `fields_get`. |
| `GetModelSchema(aModelName)` | Retrieves a deep-copied schema definition for a specified model, leveraging internal memory caching (`FetchModelFields`). |
| `SearchRecord(aModel, [aDomain], [aOffset], [aLimit], [aOrder])` | Searches records matching domain criteria and returns a collection of matching record IDs. |
| `ReadRecord(aModel, aRecordId, [aFields])` | Reads specific field values for a given record ID or collection of record IDs. |
| `SearchReadRecord(aModel, [aDomain], [aFields], [aOffset], [aLimit], [aOrder])` | Combines search and read in a single optimized RPC call to fetch record field dictionaries. |
| `CreateRecord(aModel, aValues)` | Creates a new record on the specified model using provided field-value pairs and returns the new ID(s). |
| `WriteRecord(aModel, aRecordId, aValues)` | Updates an existing record (or collection of records) by ID with new key-value pairs. |
| `UnlinkRecord(aModel, aRecordId)` | Deletes specified record(s) by ID. |
| `GetCurrentUserId()` | Retrieves the currently authenticated user's ID. |


## API Version Support & Repository Creation

This library supports both APIs for communicating with Odoo:
1. **Traditional External API** (`OdClient` / `OdRepository`)
2. **New External JSON-2 API** (`Od2Client` / `Od2Repository`)

You can instantiate and initialize the repository by passing the connection arguments directly.

### Example 1: Creating Repository for Traditional External API (`OdClient`)
*Note: Pass your Odoo API key to the key parameter.*

```vb
Private Function CreateTraditionalRepository() As OdxRepository
    Dim oClient As OdClient
    Dim oRepo As OdRepository
    
    ' Create and initialize OdClient with direct arguments: (Url, Db, UserName, ApiKey)
    Set oClient = NewOdClient
    oClient.Initialize "https://localhost", "dev_odoo", "admin", "your_api_key"
    
    Set oRepo = NewOdRepository
    oRepo.Initialize oClient
    
    ' Initialize and return OdxRepository facade
    Set CreateTraditionalRepository = NewOdxRepository
    CreateTraditionalRepository.Initialize oRepo
End Function
```

### Example 2: Creating Repository for New External JSON-2 API (`Od2Client`)

*Note: External API v2 authenticates using the database and API key only, without requiring a username.*

```vb
Private Function CreateJson2Repository() As OdxRepository
    Dim oClient2 As Od2Client
    Dim oRepo2 As Od2Repository
    
    ' Create and initialize Od2Client with direct arguments: (Url, Db, ApiKey)
    Set oClient2 = NewOd2Client
    oClient.Initialize "https://localhost", "dev_odoo", "your_api_key"
    
    Set oRepo2 = NewOd2Repository
    oRepo2.Initialize oClient2
    
    ' Initialize and return OdxRepository facade
    Set CreateJson2Repository = NewOdxRepository
    CreateJson2Repository.Initialize oRepo2
End Function
```


### Excel VBA Examples (`OdxModel`)

#### 1. Simple Search and Read (`res.partner`)

Using `OdxModel` with an initialized repository to fetch specific records:

```vb
Sub DoSimpleSearchReadOdxModel()
    Dim oRepo As OdxRepository
    Dim oModel As OdxModel
    Dim oDomain As OdFilterDomain
    
    ' Use your repository creation function here
    ' Set oRepo = CreateTraditionalRepository()
    Set oRepo = CreateJson2Repository()
    
    ' Create OdxModel for 'res.partner'
    Set oModel = NewOdxModel(oRepo, "res.partner")
    With oModel
        .AddField "name"
        .AddField "is_company"
        .AddField "is_public"
    End With
    
    ' Fetch data with a domain filter
    Set oDomain = NewDomain
    oDomain.AddArity NewField("name").IsIn(Array("Mitchell Admin", "Gemini Furniture"))
    oModel.FetchSearchAndReadRemote oDomain.Build
    
    Debug.Print "Fetched record count: " & oModel.Recordset.RecordCount
End Sub
```

#### 2. Advanced: Traversing Relations (`res.partner` and related fields)

Traversing **Many2one**, **One2many**, and **Many2many** relations hierarchically:

```vb
Sub DoRelationsTraversalOdxModel()
    Dim oRepo As OdxRepository
    Dim oModel As OdxModel
    Dim oDomain As OdFilterDomain
    
    ' Set oRepo = CreateTraditionalRepository()
    Set oRepo = CreateJson2Repository()
    
    ' Define model and nested relations
    Set oModel = NewOdxModel(oRepo, "res.partner")
    With oModel
        .AddField "name"
        .AddField "is_company"
        
        ' --- Many2one: company_id -> res.company ---
        With .AddField("company_id")
            Debug.Assert .IsMany2One
            .AddField "name"
            .AddField "city"
            
            ' --- Nested Many2one: currency_id -> res.currency ---
            With .AddField("currency_id")
                Debug.Assert .IsMany2One
                .AddField "name"
                .AddField "symbol"
            End With
        End With
        
        ' --- One2many: child_ids -> res.partner ---
        With .AddField("child_ids")
            Debug.Assert .IsOne2Many
            .AddField "name"
        End With
        
        ' --- Many2many: category_id -> res.partner.category ---
        With .AddField("category_id")
            Debug.Assert .IsMany2Many
            .AddField "name"
        End With
    End With
    
    ' Fetch records matching criteria
    Set oDomain = NewDomain
    oDomain.AddArity NewField("name").IsIn(Array("Mitchell Admin", "Gemini Furniture"))
    oModel.FetchSearchAndReadRemote oDomain.Build
    
    ' Access data and related models
    Do Until oModel.Recordset.EOF
        Debug.Print "Partner Name: " & oModel.Recordset.Fields("name").Value
        
        Dim oCompanyModel As OdxModel
        Set oCompanyModel = oModel.GetRelatedModel("company_id")
        If Not oCompanyModel Is Nothing Then
            If oCompanyModel.Recordset.RecordCount = 1 Then
                Debug.Print "  -> Company Name: " & oCompanyModel.Recordset.Fields("name").Value
            Else
                Debug.Print "  -> Company Count: " & oCompanyModel.Recordset.RecordCount
            End If
        End If
        
        Dim oChildModel As OdxModel
        Set oChildModel = oModel.GetRelatedModel("child_ids")
        If Not oChildModel Is Nothing Then
            Debug.Print "  -> Child Partner Count: " & oChildModel.Recordset.RecordCount
        End If
        
        oModel.Recordset.MoveNext
    Loop
End Sub
```


## Installation, Setup & Example Workbook

### 1. Clone the Repository

Clone the repository **with the** `--recursive` **option** (this is strictly required to include all required submodule dependencies):

```bash
git clone --recursive https://github.com/jp-rad/odoo-json-rpc-2-vba
```

### 2. Generate and Set Up the Example Workbook

Run the batch file to create the example workbook:

```bash
cd odoo-json-rpc-2-vba
./create_workbook.bat
```

**Open the Excel Files:**

* Open both `odoo-json-rpc-2-vba example.xlsm` and `odoo-json-rpc-2-vba.xlam`.

**Configure References in Visual Basic Editor (VBE):**

* Open the Visual Basic Editor (VBE).
* In VBE, select the project `odoo-json-rpc-2-vba example.xlsm`.
* Go to **Tools** > **References**.
* In the References dialog, select `OdooJsonRpc2VBA`.
* **Additionally, select the following reference in the add‑in project `odoo-json-rpc-2-vba.xlam`:**
* **Microsoft ActiveX Data Objects 2.8 Library**
*(or another available version such as 6.1, depending on your system)*
This reference is required for ADODB Recordset operations used by the library.


**Run the Tutorial Methods:**

* Open the **Immediate Window** in VBE.
* Run either of the following commands depending on which API version you wish to test:
* For Traditional External API: `DoTutrialOdxModel`
* For New External JSON-2 API: `DoTutrialOdxModel2`


Refer to the following document for step-by-step details:

* [Odoo Docs - External API: Calling methods](https://www.odoo.com/documentation/20.0/developer/reference/external_rpc_api.html#calling-methods)


## Notes

### VBS Tools Runtime Error

Programmatic access to the Office VBA project may be denied. If this occurs, refer to:

* [Programmatic access to Office VBA project is denied](https://support.microsoft.com/en-us/topic/programmatic-access-to-office-vba-project-is-denied-960d5265-6592-9400-31bc-b2ddfb94b445)

### Date and Datetime Fields

When working with Date or Datetime fields in Odoo from VBA, note:

* Odoo expects dates as strings in specific formats:
* Date: `YYYY-MM-DD`
* Datetime: `YYYY-MM-DDTHH:MM:SS`


* Use `OdJsonRpc.FormatDate` / `OdJsonRpc.ConvertToIsoDatetime` when sending values, and `OdJsonRpc.ParseDate` / `OdJsonRpc.ParseIsoDatetime` when reading values.


## Credits

This project utilizes the following open-source tools and libraries from **VBA-tools**:

1. **vba-dictionary** (GitHub: [https://github.com/VBA-tools/VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary))
2. **vba-json** (GitHub: [https://github.com/VBA-tools/VBA-JSON](https://github.com/VBA-tools/VBA-JSON)) - *Note: This project utilizes `WebHelpers` provided within vba-web.*
3. **vba-web** (GitHub: [https://github.com/VBA-tools/VBA-Web](https://github.com/VBA-tools/VBA-Web))
    * Includes a modified version of `WebResponse` (v4.1.6 Custom), updated to support parsing primitive JSON responses by automatically wrapping non-object/non-array content into a JSON array structure.

4. **vba-test** (GitHub: [https://github.com/VBA-tools/VBA-Test](https://github.com/VBA-tools/VBA-Test))

- License: MIT License (MIT)  
Copyright (c) Tim Hall
