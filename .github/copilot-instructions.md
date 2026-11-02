# Excel VBA Coding Conventions and Standards

This document defines the coding conventions, naming rules, structural standards, and unit testing guidelines for Excel VBA projects (such as `odoo-json-rpc-2-vba`). It is designed to be easily readable and parsable by both LLMs and human developers.

---

## 1. File Structure and Metadata

* **License Notice**: Every source file (modules, class modules) must begin with the standard MIT License comment block at the very top.


* **Option Explicit**: Must be declared immediately after the license header in every module to enforce explicit variable declaration.


* **Attribute Directives**: Do not manually add or manage Attribute declarations (such as `VB_Name` or `VERSION 1.0 CLASS`) in code blocks, as these are managed internally by the VBA IDE. Modules are added simply as `.bas` or `.cls` files.


* **Module Header Format**: The descriptive comment block following `Option Explicit` must use the standard file-based naming representation:
* For standard modules: `' ------------------------------------------------------------------------------` followed by `' Module: ModuleName.bas`
* For class modules: `' ------------------------------------------------------------------------------` followed by `' Module: ClassName.cls`



---

## 2. Naming Conventions

### Variable Prefixes (Hungarian Notation for Local Variables)

All local variables must use a lowercase prefix indicating their data type or scope to avoid collisions with properties and keywords.

| Data Type / Category | Prefix | Example |
| --- | --- | --- |
| Object (General) | `o` | `oRepo`, `oModel`, `oRs`, `oSuite` |
| String (`String`) | `s` / `str` | `sFieldName`, `sFieldType`, `externalId` |
| Long Integer (`Long`) | `n` | `nId`, `nCount`, `nLimit` |
| Boolean (`Boolean`) | `b` | `bType2`, `bSuccess`, `bIncludeSynced` |
| Variant (`Variant`) | `v` | `vKey`, `vFieldValue` |
| Byte Array (`Byte()`) | `ab` | `abA`, `abB`, `abBytes` |
| Collection (`Collection`) | `col` | `colFields`, `colDomain` |
| Dictionary (`Dictionary`) | `dic` | `dicRecord`, `dicConfig` |
| Array (`Array`) | `arr` | `arrParts`, `arrValues` |


### Member Variable Naming (Module-Level Variables)

* **Scope Prefix (`m`)**: All module-level variables (member variables) must start with the scope prefix `m`.
* **No Type Prefixes**: Unlike local variables, type-specific Hungarian notation prefixes (such as `o`, `s`, or `n`) are **not** applied to member variables. They use `m` followed immediately by a PascalCase name.
* **Examples**:
* Object: `mRs` (`m` + PascalCase)
* String: `mModuleName` (`m` + PascalCase)
* Boolean: `mIsInitialized` (`m` + PascalCase)





### Parameter Naming

* **All procedure parameters must start with the prefix `a**` (e.g., `aSuite`, `aFieldName`, `aDataType`, `aIncludeSynced`).



### Constants Naming

* **Module-level or Global Constants**: Must start with `C`, followed by uppercase letters with underscore separators (e.g., `CBIT_MODEL`, `CATTR_PK_ID`, `CFIELD_STATUS`).



### Classes, Methods, and Properties

* **Class Names, Public Methods, and Properties**: Use PascalCase (e.g., `OdxModel`, `OdxRepository`, `CreateNewRecordSet`, `RegisterSpecs`).


* **Private Methods**: Use PascalCase or camelCase for internal helper routines (e.g., `IsBinaryField`, `AreByteArraysEqual`, `AddField`).



---

## 3. Parameter Passing and Performance Rules (ByRef vs ByVal)

To balance high execution performance with code safety, follow these rules strictly:

1. **Default / Required Parameters (`ByRef`)**:
* For maximum execution speed, all required parameters must use pass-by-reference (`ByRef`, which is the default in VBA, or explicitly omitted).


2. **Optional Parameters (`ByVal`)**:
* All `Optional` parameters **must** be explicitly declared with `ByVal` and assigned a clear default value (e.g., `Optional ByVal aLimit As Long = -1`).


3. **Internal Value Modification (`ByVal`)**:
* If a parameter's value is modified or reassigned inside the procedure body, it **must** be declared with `ByVal` to prevent unintended side effects on the caller and to avoid declaring extra local variables.



---

## 4. Unit Testing Framework (`vba-test`)

All unit testing within the project must use the **`vba-test`** framework.

1. **Framework Core**:
* Unit tests must be structured using `TestSuite` and reported via `ImmediateReporter` (or appropriate listener).




2. **Test Registration Structure**:
* Test modules (e.g., `SpecMemRs.bas`, `SpecOdxRepository.bas`) must expose an entry point procedure, typically `RegisterSpecs(aSuite As TestSuite)`, accepting the `TestSuite` instance with the `a` parameter prefix.




3. **Assertions**:
* Use built-in assertions provided by `vba-test` (e.g., `.IsEqual`, `.IsTrue`, `.IsFalse`, `.IsUndefined`) to assert expected outputs against actual outcomes.




4. **Example Test Pattern**:

```vba
''
' Register and run all test specifications for MemRs module.
'
' @param aSuite TestSuite instance
Public Sub RegisterSpecs(aSuite As TestSuite)
    Dim oSuite As TestSuite
    Set oSuite = aSuite

    With oSuite.Test("should create a valid recordset")
        Dim oRs As Object
        Set oRs = MemRs.CreateNewRecordSet()
        .IsTrue Not (oRs Is Nothing)
    End With
End Sub

```

---

## 5. Commenting Rules

* **Strict ASCII Only**: All source code, including comments, docstrings, and literals, must be written using **ASCII characters only**. Do not use multi-byte characters, Japanese text, emojis, or special symbols.


* **File Header**: Include the full MIT License text in ASCII.


* **Section Dividers**: Use ASCII banner/border lines to separate logical sections within a module.


* **Documentation Comments**: Use double apostrophes (`''`) for procedure-level and test-level documentation. Explicitly document parameters using `@param` and return values using `@return`.


* **Inline Comments**: Use single apostrophes (`'`) for concise implementation notes, branch logic explanations, or step descriptions.
