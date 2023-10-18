# SQL on Excel Data using VBA (dbConn) Wiki!

This scripts in the README file are lift and shift codes which can be directly pasted in the `Visual Basic Editor of MS Excel`. This script will help in filtering the summarizing the data by treating it as a SQL database and hence, SQL queries can be used upon the Excel data.

`SQL queries` are more user friendly as compared to front-end Excel filtering coded in VBA.

## Pre-requisites

- MS Excel
- OfficeSetup.exe*

<span style="color: aqua">* It is recently observed that the MS Office pack does not come loaded with the required libraries (DLLs). It was installed separately using the executable file [`OfficeSetup.exe`](https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=AccessRuntimeRetail&language=en-us)<br>
OS Architecture (32 bit \ 64 bit) plays an important role in installation of both MS pack and the DLLs and you will need a MS Account to download.
</span>


## Code Setup

<b><i>Step 1:</b></i> From `Tools` -> `References`, add `Microsoft ActiveX Data Objects X.X Library` and `Microsoft ActiveX Data Objects Recordset X.X Library`, latest versions


<b><i>Step 2:</b></i> Insert a `Class Module` and name it `dbConn`

<b><i>Step 3:</b></i> Paste the below code in the class module

<code>

    Option Explicit
    Dim fcount As Long, rcount As Long
    Private cn As ADODB.Connection, rs As ADODB.Recordset

    Private Sub Class_Initialize()
        Call OpenConnection
    End Sub

    Private Sub Class_Terminate()
        Call CloseConnection
    End Sub


    Private Sub OpenConnection()
        Dim dbPath As String

        dbPath = ThisWorkbook.FullName
        Set cn = New ADODB.Connection
        Set rs = New ADODB.Recordset
        
        cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & dbPath & ";" & _
        "Extended Properties=""Excel 12.0 Macro;" & "HDR=YES;"";"
    End Sub

    Private Sub CloseConnection()
        cn.Close
        Set cn = Nothing
        Set rs = Nothing
    End Sub

    Public Sub ExecuteQuery(ByVal strSql As String)
        Set rs = cn.Execute(strSql)
    End Sub

    Public Function GetRecords(ByVal strSql As String) As Variant
        Dim var, i, j
        
        If rs.State = adStateOpen Then rs.Close
        
        rs.Open strSql, cn, adOpenStatic, adLockOptimistic
        rcount = rs.RecordCount
        fcount = rs.Fields.Count
        
        If rcount = 0 Then
            GetRecords = var
            Exit Function
        End If
        
        ReDim var(rcount - 1, fcount - 1)
        For i = LBound(var, 1) To UBound(var, 1)
            For j = LBound(var, 2) To UBound(var, 2)
                If Not IsNull(rs.Fields(j)) Then
                    var(i, j) = rs.Fields(j)
                End If
            Next
            rs.MoveNext
        Next
        GetRecords = var
        rs.Close
    End Function

    Private Sub GetCounts()
        If rs.State = adStateOpen Then rs.Close
        
        rs.Open strSql, cn, adOpenStatic, adLockOptimistic
        rcount = rs.RecordCount
        fcount = rs.Fields.Count
        rs.Close
    End Sub

    Public Function RecCount(ByVal strSql As String) As Long
        If rs.State = adStateOpen Then rs.Close
        
        rs.Open strSql, cn, adOpenStatic, adLockOptimistic
        RecCount = rs.RecordCount
        rs.Close
    End Function

    Public Function FieldCount(ByVal strSql As String) As Long
        If rs.State = adStateOpen Then rs.Close
        
        rs.Open strSql, cn, adOpenForwardOnly, adLockReadOnly
        FieldCount = rs.Fields.Count
        rs.Close
    End Function

</code>

<span style="color: aqua">Note that the `Provider=...` connection string can change with the change in MS Office version. The current setting is compatible with MS Office 365</span>

<b><i>Step 4:</b></i> In the `Module`, along with other functions add the below code with SQL queries to be called from the front-end
```

<code>

    Sub test2()

    Dim cn As dbConn
    Dim ArrRaw As Variant
    Dim strSql As String

    strSql = "Select Col1, sum(Col2), sum(Col3) from [Sheet1$] where Col1 = 'Group2' and Col2 > 40 group by Col1"

    Set cn = New dbConn
    ArrRaw = cn.GetRecords(strSql)

    'For pasting the values
    Worksheets("Sheet3").Range("A2").Resize(UBound(ArrRaw, 1) + 1, UBound(ArrRaw, 2) + 1) = ArrRaw

    End Sub

</code>

## Sample Workbook

A sample workbook `sample_data.xlsm` is provided utilizing the above scripts to achive the SQL functionalities within Excel