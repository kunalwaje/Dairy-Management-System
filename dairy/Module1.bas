Attribute VB_Name = "Module1"
Option Explicit
Public rss As New Recordset
Public con As New Connection
Public rs As New Recordset
Public nm As Integer

Public Sub connect()
If con.State = 1 Then
con.Close
End If
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\PROJECT\Dairy Management\dairy\dairy\dairy managment.mdb;Persist Security Info=False;Persist Security Info=False"
con.Open
'MsgBox ("Connected to database")

End Sub

Public Sub retdata(ByVal q As String)
If rs.State = 1 Then
rs.Close
End If
rs.CursorLocation = adUseClient
rs.Open q, con, adOpenDynamic, adLockPessimistic
End Sub

Public Sub setdata(ByVal q As String)
If rss.State = 1 Then
rs.Close
End If
rss.CursorLocation = adUseClient
rss.Open q, con, adOpenDynamic, adLockPessimistic
End Sub

Public Sub numquery(ByVal q As String)
If rs.State = 1 Then
rs.Close
End If
rs.CursorLocation = adUseClient

rs.Open q, con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Or Not rs.BOF Then
nm = rs.Fields(0).Value
End If

End Sub


