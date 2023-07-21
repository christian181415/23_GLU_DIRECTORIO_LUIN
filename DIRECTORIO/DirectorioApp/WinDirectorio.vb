Public Class WinDirectorio
    Dim NewConfig As New ClassDirectorioConfig
    Dim NewData As New ClassDirectorioData

    Private Sub WinDirectorio_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            NewConfig.CircleSize(Me, P_Logo, Lb_Hora.Text, Lb_Fecha.Text, Timer1.Interval)
            Timer1.Start()

            Dim r As Rectangle = My.Computer.Screen.WorkingArea
            Dim Largo = r.Width - Width
            Dim Alto = r.Height - Height
            Location = New Point(r.Width - Width - 10, 10)

            'Dim Empresa = "INEXISTENTE"
            'Dim Empresa = "FGI"
            'Dim Empresa = "MEDICINE"
            'Dim Empresa = "MOLDAPTA"
            'Dim Empresa = "PLASTIKISIMO"
            'Dim Empresa = "PLASTIKA"
            Dim Empresa = System.Environment.GetEnvironmentVariable("COMPUTERNAME")
            NewConfig.SeleccionarColor(Empresa, Footer1, Footer2, Footer3, Footer4, Footer5, Footer6, Footer7, btnLogo,
                                      FondoEmpresa, btnBackSearch, btnBackList, btnBackResult, btnBackTicket,
                                      btnBackResultAviso, btnBackAviso, btnCrearTicket, btnGuardarTicket, NotifyIcon, Me, btnMostrarAvisos)
        Catch ex As Exception

        End Try
    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Lb_Hora.Text = String.Format("{0:hh:mm tt}", DateTime.Now)

        'If btnMostrarAvisos.Visible = True Then
        '    If btnMostrarAvisos.Width = 26 Then
        '        For i As Integer = 26 To 29
        '            btnMostrarAvisos.Width = i
        '            btnMostrarAvisos.Height = i
        '        Next
        '    ElseIf btnMostrarAvisos.Width = 29 Then
        '        btnMostrarAvisos.Width = 26
        '        btnMostrarAvisos.Height = 26
        '    End If
        'End If

    End Sub
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        NewConfig.TimerAction(Timer2, P_Logo, P_Busqueda, P_ResultadoPerfil, P_Perfil, P_Soporte, P_Ticket, P_ResultadoAviso, P_Aviso, Me.NotifyIcon)
    End Sub
    Private Sub MenuCerrar_Click(sender As Object, e As EventArgs) Handles CerrarDirectorio.Click
        Application.Exit()
    End Sub


    Private Sub btnLogo_Click(sender As Object, e As EventArgs) Handles btnLogo.Click
        NewConfig.CambiarVentana(P_Logo, 309, 0, P_Busqueda, 0, 0)
        NewConfig.SesionSistema(Timer2)
    End Sub
    Private Sub btnMostrarAvisos_Click(sender As Object, e As EventArgs) Handles btnMostrarAvisos.Click
        NewConfig.CambiarVentana(P_Logo, 309, 0, P_ResultadoAviso, 0, 0)
        NewConfig.SesionSistema(Timer2)
    End Sub
    Private Sub btnCrearTicket_Click(sender As Object, e As EventArgs) Handles btnCrearTicket.Click
        NewConfig.CambiarVentana(P_Logo, 309, 0, P_Ticket, 0, 0)
        NewConfig.SesionSistema(Timer2)
    End Sub


    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        If txtBuscar.Text <> String.Empty Then
            NewData.SELECTDB_BUSCADOR_PROCEDURE(txtBuscar, "SELECT_BUSCADOR", lbResultado)
            NewConfig.CambiarVentana(P_Busqueda, 618, 0, P_ResultadoPerfil, 0, 0)
            txtBuscar.Text = ""
        Else
            ClassMsgBoxCustom.Show("Rellene los campos necesarios", "Error | Corporativo LUIN", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub


    Private Sub btnBackSearch_Click(sender As Object, e As EventArgs) Handles btnBackSearch.Click
        NewConfig.CambiarVentana(P_Busqueda, 618, 0, P_Logo, 0, 0)
        txtBuscar.Text = ""
    End Sub


    Private Sub btnBackList_Click(sender As Object, e As EventArgs) Handles btnBackList.Click
        NewConfig.CambiarVentana(P_ResultadoPerfil, 927, 0, P_Busqueda, 0, 0)
    End Sub
    Private Sub lbResultado_DoubleClick(sender As Object, e As EventArgs) Handles lbResultado.DoubleClick
        NewData.SELECTDB_DATA_PROCEDURE(lbResultado, "SELECT_DATA", lbNombreData, lbPuestoData, lbCorreoData, lbTelefonoData, lbCelularData, lbExtensionData)
        NewConfig.CambiarVentana(P_ResultadoPerfil, 927, 0, P_Perfil, 0, 0)
    End Sub


    Private Sub btnBackResult_Click(sender As Object, e As EventArgs) Handles btnBackResult.Click
        NewConfig.CambiarVentana(P_Perfil, 1236, 0, P_ResultadoPerfil, 0, 0)
    End Sub
    Private Sub btnCorreo_Click(sender As Object, e As EventArgs) Handles btnCorreo.Click
        'System.Diagnostics.Process.Start("http://corporativoluin.mx/")
        System.Diagnostics.Process.Start("mailto:" & lbCorreoData.Text)
        'Dim RutaOutlook As New ProcessStartInfo
        'RutaOutlook.FileName = "C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.exe"
        'RutaOutlook.Arguments = "Mail.To.Add(" & lbCorreoData.Text & ")"
        'Process.Start(RutaOutlook)
    End Sub
    Private Sub PBoxUbicacion_Click(sender As Object, e As EventArgs) Handles PBoxUbicacion.Click
        System.Diagnostics.Process.Start("https://goo.gl/maps/gyueKXvgZWXuMnjW6")
    End Sub


    Private Sub BtnBackFormato_Click(sender As Object, e As EventArgs) Handles BtnBackFormato.Click

    End Sub


    Private Sub btnBackTicket_Click(sender As Object, e As EventArgs) Handles btnBackTicket.Click
        NewConfig.CambiarVentana(P_Ticket, 618, 309, P_Logo, 0, 0)
        txtTitulo.Text = ""
        txtProblema.Text = ""
    End Sub


    Private Sub lbResultadoAviso_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbResultadoAviso.SelectedIndexChanged
        If lbResultadoAviso.SelectedIndex = 0 Then
            NewConfig.CambiarVentana(P_ResultadoAviso, 927, 309, P_Aviso, 0, 0)
        End If
    End Sub
    Private Sub btnBackResultAviso_Click(sender As Object, e As EventArgs) Handles btnBackResultAviso.Click
        NewConfig.CambiarVentana(P_ResultadoAviso, 927, 309, P_Logo, 0, 0)
    End Sub


    Private Sub btnBackAviso_Click(sender As Object, e As EventArgs) Handles btnBackAviso.Click
        NewConfig.CambiarVentana(P_Aviso, 1236, 309, P_ResultadoAviso, 0, 0)
    End Sub


    Private Sub MostrarDirectorio_Click(sender As Object, e As EventArgs) Handles MostrarDirectorio.Click
        Me.WindowState = FormWindowState.Minimized
        Me.WindowState = FormWindowState.Normal
    End Sub
    Private Sub NotifyIcon_MouseClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon.MouseClick
        If e.Button = MouseButtons.Left Then
            Me.WindowState = FormWindowState.Minimized
            Me.WindowState = FormWindowState.Normal
        End If
    End Sub

End Class
