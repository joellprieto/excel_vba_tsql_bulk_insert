# excel_vba_tsql_bulk_insert
excel_vba_tsql_bulk_insert

### connection string - note the SQL Server running here is docker thus path is using non-windows convention
```vba
Sub copy_file(source as string, target_folder as String)
    With New FileSystemObject
        If .FileExists(source_file) And .FileExists(target_folder) Then
            .CopyFile source_file, target_folder & "\", True
        End If
    End With
End Sub

Sub call_bulk_insert()
    Set conn = New ADODB.Connection
    User_ID = "set user id here if odbc is setup"
    Password = "set password here if odbc is setup"
    conn.ConnectionString = "dsn=DOCKER_SQL_SERVER;User ID=" & User_ID & ";Password=" & Password
    conn.Open
    bulk_insert_string = "if object_id('data_upload.dbo.people', 'U') is not null truncate table data_upload.dbo.people; " & _
                         "if object_id('data_upload.dbo.people', 'U') is not null drop table data_upload.dbo.people; " & _
                         "select top 0 * into data_upload.dbo.people from data_upload.dbo.people_clone_format; " & _
                         "bulk insert data_upload.dbo.people " & _
                         "from '/var/opt/mssql/bulk_insert/sample_data/sample_data.csv' " & _
                         "with (FIELDTERMINATOR=',',FIRSTROW=2,FORMAT='CSV',FORMATFILE='/var/opt/mssql/bulk_insert/sample_data/sample_data.format'); "

    conn.Execute bulk_insert_string
    conn.Close
End Sub

Sub delete_target_file(target_file as String)
    With New FileSystemObject
        If .FileExists(target_file) Then
            .DeleteFile target_file
        End If
    End With
End Sub

```

### since TSQL insert treats as characters we will append the corresponding file format for transformation
```xml
<?xml version="1.0"?>
<BCPFORMAT xmlns="http://schemas.microsoft.com/sqlserver/2004/bulkload/format" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<RECORD>
  <FIELD ID="1" xsi:type="CharTerm" TERMINATOR=","/>
  <FIELD ID="2" xsi:type="CharTerm" TERMINATOR=","/>
  <FIELD ID="3" xsi:type="CharTerm" TERMINATOR=","/>
  <FIELD ID="4" xsi:type="CharTerm" TERMINATOR="\r\n"/>
</RECORD>
<ROW>
  <COLUMN SOURCE="1" NAME="name" xsi:type="SQLNVARCHAR"/>
  <COLUMN SOURCE="2" NAME="age" xsi:type="SQLINT"/>
  <COLUMN SOURCE="3" NAME="country" xsi:type="SQLNVARCHAR"/>
  <COLUMN SOURCE="4" NAME="dob" xsi:type="SQLDATE"/>
</ROW>
</BCPFORMAT>
```
### below is the sample data file in CSV format
```csv
name,age,country,dob
john,20,AUS,20241201
peter,20,SIN,20241103
```


