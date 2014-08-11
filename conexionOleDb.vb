'=-------
' Copyright Pedro Santana
' http://www.pecesama.net/
'
' Liberado tal cual, sin garantías ni responsabilidades, etc.
' Se da permiso de uso, copia y modificaciones,
' siempre y cuando se me de crédito, sólo que no
' me molestes si no te funciona.
'=------- 

Public Class conexionOleDb
	' Metodos
	Public Sub New(ByVal rutaBase As String)
		Me.strCon = ""
		Me.Error = ""
		Me.strCon = ("Provider=Microsoft.Jet.OLEDB.4.0 ;Data Source=" & rutaBase & ";")
	End Sub

	Public Sub New(ByVal rutaBase As String, ByVal password As String)
		Me.strCon = ""
		Me.Error = ""
		Me.strCon = String.Concat(New String() { "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=", rutaBase, ";Jet OLEDB:Database Password=", password, ";" })
	End Sub

	Private Function ActualizaBase(ByVal conjuntoDeDatos As DataTable) As Integer
		Try
			If (Not conjuntoDeDatos Is Nothing) Then
				Dim errors As DataRow() = conjuntoDeDatos.GetErrors
				If (errors.Length = 0) Then
					Dim num As Integer = Me.miAdaptador.Update(conjuntoDeDatos)
					conjuntoDeDatos.AcceptChanges
					Me.Error = ""
					Me.misDatos = conjuntoDeDatos
					Return num
				End If
				Me.Error = ""
				Dim row As DataRow
				For Each row In errors
					Dim column As DataColumn
					For Each column In row.GetColumnsInError
						Me.Error = (Me.Error & row.GetColumnError(column) & ChrW(10))
					Next
				Next
				conjuntoDeDatos.RejectChanges
				Me.misDatos = conjuntoDeDatos
			End If
			Return -1
		Catch exception As Exception
			Me.Error = exception.Message
			Return -1
		End Try
	End Function

	Public Sub cerrarConexion()
		If (Not Me.miCon Is Nothing) Then
			Me.miCon.Close
		End If
	End Sub

	Public Function conectar() As Boolean
		If (Not Me.miCon Is Nothing) Then
			Me.miCon.Close
		End If
		Try
			Me.miCon = New OleDbConnection(Me.strCon)
			Me.miCon.Open
			Me.Error = ""
			Return True
		Catch exception As Exception
			Me.Error = exception.Message
			Return False
		End Try
	End Function

	Public Function ejecutaSql(ByVal CadenaSql As String) As DataTable
		Dim selectCommandText As String = CadenaSql
		Me.misDatos = New DataTable
		Me.miAdaptador = New OleDbDataAdapter(selectCommandText, Me.miCon)
		Try
			Me.miAdaptador.Fill(Me.misDatos)
			Me.Error = ""
		Catch exception As Exception
			Me.misDatos = Nothing
			Me.Error = exception.Message
		End Try
		Return Me.misDatos
	End Function

	Public Function estaConectado() As Boolean
		Return (Not Me.miCon Is Nothing)
	End Function

	' Propiedades
	Public Property error As String
		Get
			Return Me.Error
		End Get
		Set(ByVal value As String)
			Me.Error = value
		End Set
	End Property

	' Campos
	Private Error As String
	Private miAdaptador As OleDbDataAdapter
	Private miCon As OleDbConnection
	Private misDatos As DataTable
	Private strCon As String
End Class