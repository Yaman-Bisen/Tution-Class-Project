Attribute VB_Name = "Module1"
Public cnn As New ADODB.Connection
Public Sub connect()
    If cnn.State = 1 Then cnn.Close
    cnn.Provider = "Microsoft.Jet.OLEDB.4.0"
    'cnn.ConnectionString = "data source=" & App.Path & "\dsoft.mdb"
    cnn.ConnectionString = "data source=" & App.Path & "\dsoft1.mdb"
    cnn.Open
End Sub
