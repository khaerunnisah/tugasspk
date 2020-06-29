Attribute VB_Name = "Module1"
Option Explicit
Public koneksidb As New ADODB.Connection
Sub bukadb()
Set koneksidb = New ADODB.Connection
koneksidb.CursorLocation = adUseClient
koneksidb.ConnectionString = "driver={mysql odbc 3.51 driver};server=localhost;database=db_saham;uid=root;option="
If koneksidb.State = adStateClosed Then koneksidb.Open
End Sub
