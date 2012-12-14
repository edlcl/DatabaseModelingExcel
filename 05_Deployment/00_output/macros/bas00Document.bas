Attribute VB_Name = "bas00Document"
Option Explicit

'------------------------------------------------------------------------
'-- Bug fixed list
'-  <3.2.3>
'*  Bugs fix:
'*      [SQL Server] Cannot get FK when the reversing table is in a schema rather than dbo.
'
'-  <3.2.2>
'*  Bugs fix:
'*      A mistake in sample.bat
'
'-  <3.2.0>
'*  New feature:.
'*      Support only Create Table SQL script
'
'-  <3.1.0>
'*  New feature:.
'*      Support ignore some worksheets when generating SQL scripts
'
'-  <3.0.0 RC1>
'*  Support Oracle.
'*  Add a vb scripts sample which generate scripts
'
'-  <2.0.0 RC1>
'*  Support MySQL.
'*  Add a vb scripts sample which generate scripts
'*  More details Help
'*  Add denote function
'*  Per document for one database type
'
'-  <Bug Fixed>
'*  Generate Foreign Key Error.
'*  Generate Index bug: non-unique index is as unique index.
'*  Reverse Index bug: error when judgment an index is unique or not.
'*  Reverse index bug: error when there are more than one column in an index.
'
'-  <1.6.2006.0814>
'   Clear note when reverse
'   Default (1) change to -1 when reverse
'------------------------------------------------------------------------

'------------------------------------------------------------------------
'-- Current Features
'   < Database Type >
'   * Support SQL Server
'   * Support MySQL
'
'------------------------------------------------------------------------

'------------------------------------------------------------------------
'-- Future Features
'   < Database Type >
'   No other database want to supported.
'
'   < Script Capability>
'   * Support user define type
'
'   < Other>
'   * Generate unified XML
'   * multiple lines for FK, Index
'   * Non-macro, should be a msoffice add-ins utility
'------------------------------------------------------------------------

'------------------------------------------------------------------------
'-- How to define version
'   * MajorVersion is changed for big feature be added.
'   * MinorVersion is changed for normal feature be added.
'   * Revision number is changed for bug fix.
'   * Release type definition. a#: arlfa release; b#: beta release; rc#:realease cadidate; <empty>: product release
'------------------------------------------------------------------------

'------------------------------------------------------------------------
'-- How to support a new database
'   * Add constant for new databsae in basAppSetting File
'       Like: Public Const DBName_Oracle                      As String = "Oracle"
'   * Add menu item for new database in basToolbar
'   * Add below code files
'       basSQL_<NewDatabase>
'       basReverse_<NewDatabase>
'       frmReverse_<NewDatabase>
'   * Add a template excel file in 05_Deployment\Resources
'       DatabaseModeling_Template_<NewDatabase>.xls
'   * Update build script for update macro for the new template excel file.
'
'   [Technical points]
'       Create Table
'       Create PK(Cluster, Unique, Non-unique), Index(Cluster, Unique, Non-unique), FK,
'       Create Table and Field comment
'       Drop Table
'       Drop FK
'       Get Table(s) schema information
'------------------------------------------------------------------------

