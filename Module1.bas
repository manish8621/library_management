Attribute VB_Name = "Module1"
Public db As New ADODB.Connection
Public rs As New ADODB.Recordset
Public constr As String

Public Sub loadcon()
db.Provider = "Microsoft.Jet.OLEDB.4.0"
db.ConnectionString = "Data Source=D:\Documents\VisualBasic\LibraryProject\library.mdb"
db.Open constr

End Sub
