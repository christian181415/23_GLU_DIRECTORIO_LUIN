Imports System.Configuration
Imports System.Data.OleDb

Public Class ClassDirectorioData
    Dim CadenaConexion As String = ConfigurationManager.ConnectionStrings("CadenaConexion").ConnectionString

    Public Function SELECTDB_MINIMO(LBox As ListBox, TxtNombre As TextBox)
        Try
            Using ConexionDB As New OleDbConnection(CadenaConexion)
                Dim Consulta As String = "SELECT Nombre FROM Empleados" &
                                        " WHERE Nombre LIKE '[" & TxtNombre.Text & "]*' UNION" &
                                        " SELECT Nombre FROM Sucursales" &
                                        " WHERE Nombre LIKE '[" & TxtNombre.Text & "]*' UNION" &
                                        " SELECT Nombre FROM Puestos" &
                                        " WHERE Nombre LIKE '[" & TxtNombre.Text & "]*';"
                Dim Comando As OleDbCommand = New OleDbCommand(Consulta)
                Comando.Connection = ConexionDB
                ConexionDB.Open()
                Dim Reader As OleDbDataReader = Comando.ExecuteReader
                MsgBox(Consulta)

                While Reader.Read
                    MsgBox(Reader.GetString(0))
                    If Reader.GetString(0) <> String.Empty Then
                        LBox.Items.Add(Reader.GetString(0))
                    Else
                        ClassMsgBoxCustom.Show("No se encontro ningun usuario con esas credenciales", "Error | Corporativo LUIN", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    End If
                End While
                ConexionDB.Close()
            End Using
        Catch ex As Exception
            ClassMsgBoxCustom.Show(ex.Message, "Error | Corporativo LUIN", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function


    Public Function SELECTDB_BUSCADOR_PROCEDURE(TxtNombre As TextBox, ProcAlm As String, LBox As ListBox)
        Try
            Using ConexionDB As New OleDbConnection(CadenaConexion)
                Dim Comando As New OleDbCommand
                Comando.Connection = ConexionDB
                Comando.CommandType = CommandType.StoredProcedure
                Comando.CommandText = ProcAlm
                Comando.Parameters.AddWithValue("@Criterio", TxtNombre.Text)

                ConexionDB.Open()
                Dim Reader As OleDbDataReader = Comando.ExecuteReader
                LBox.Items.Clear()
                LBox.DataSource = Nothing
                While Reader.Read
                    If Reader.GetString(0) <> "" Then
                        LBox.Items.Add(Reader.GetString(0))
                    Else
                        ClassMsgBoxCustom.Show("No se encontro ningun usuario con esas credenciales", "Error | Corporativo LUIN", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    End If
                End While
                ConexionDB.Close()
            End Using
        Catch ex As Exception
            ClassMsgBoxCustom.Show(ex.Message, "Error | Corporativo LUIN", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function


    Public Function SELECTDB_DATA_PROCEDURE(LBox As ListBox, ProcAlm As String, lbNombre As Label, lbPuesto As Label, lbCorreo As Label, lbTelF As Label, lbTelM As Label, lbExt As Label)
        Try
            Using ConexionDB As New OleDbConnection(CadenaConexion)
                Dim Comando As New OleDbCommand
                Comando.Connection = ConexionDB
                Comando.CommandType = CommandType.StoredProcedure
                Comando.CommandText = ProcAlm
                Comando.Parameters.AddWithValue("@Criterio", LBox.SelectedItem)

                ConexionDB.Open()
                Dim Reader As OleDbDataReader = Comando.ExecuteReader
                While Reader.Read
                    lbNombre.Text = Reader.GetString(0)
                    lbPuesto.Text = Reader.GetString(1)
                    lbCorreo.Text = Reader.GetString(2)
                    lbTelF.Text = Reader.GetString(3)
                    lbTelM.Text = Reader.GetString(4)
                    lbExt.Text = Reader.GetString(5)
                End While
                ConexionDB.Close()
            End Using
        Catch ex As Exception
            ClassMsgBoxCustom.Show(ex.Message, "Error | Corporativo LUIN", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Public Function INSERTDB_TICKET_PROCEDURE()

    End Function
End Class
