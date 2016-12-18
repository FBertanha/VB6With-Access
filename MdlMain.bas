Attribute VB_Name = "MdlMain"
Public access As New ADODB.Connection
Public rs As New ADODB.Recordset

Public Sub openBase()
    access.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\data.mdb;")
End Sub

Public Function executeQuery(query As String, Optional opt As String = "") As Recordset
    Set executeQuery = access.Execute(query)
End Function

