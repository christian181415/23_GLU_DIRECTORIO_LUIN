Imports System.Runtime.InteropServices
Imports System.Threading

Public Class ClassConfig
    Private Declare Function GetLastInputInfo Lib "user32" (ByRef plii As UltimoClick) As Boolean
#Region "ClassConfig | FUNCION ESPECIFICA PARA CAMBIAR EL COLOR"
    Public Function SeleccionarColor(Empresa As String, P_Logo As Panel, P_Busqueda As Panel, P_ResultadoPerfil As Panel, P_Perfil As Panel, P_Ticket As Panel,
                                 P_ResultadoAviso As Panel, P_Aviso As Panel, Btn_Logo As PictureBox, Fondo_Empresa As PictureBox, Btn_BackSearch As Button,
                                 Btn_BackList As Button, Btn_BackResult As Button, Btn_BackTicket As Button, Btn_BackResultAviso As Button, Btn_BackAviso As Button,
                                 Btn_CrearTicket As PictureBox, Btn_GuardarTicket As Button, NotifyIcon As NotifyIcon, BloqueoForm As Object, Btn_MostrarAvisos As PictureBox)

        '------------------------------------------------------- COLORES EMPRESARIALES
        Dim CORP = "#8A2A2B"
        Dim OverBackCORP = "#FF0003" 'FB0101
        Dim MEDICINE = "#16A6D9"
        Dim OverBackMDN = "#00BCFF"  '00BCFF
        Dim PLASTIKISIMO = "#00577F"
        Dim OverBackPSKI = "#00AFFF"  '0087C4
        Dim PLASTIKA = "#001871"
        Dim OverBackPSK = "#0036FF"  '0036FF

        If Empresa = "SSIT02" Then
            'USUARIOS CORP
            CambiarColor(P_Logo, P_Busqueda, P_ResultadoPerfil, P_Perfil, P_Ticket, P_ResultadoAviso, P_Aviso, Btn_CrearTicket,
                        Btn_BackSearch, Btn_GuardarTicket, Btn_BackList, Btn_BackResult, Btn_BackTicket, Btn_BackResultAviso,
                        Btn_BackAviso, Btn_Logo, Fondo_Empresa, CORP, OverBackCORP, "Logo_CORP.jpg", "Logo_CORP_Hoja.png")
        ElseIf Empresa = "FGI" Then
            'USUARIOS FGI
            CambiarColor(P_Logo, P_Busqueda, P_ResultadoPerfil, P_Perfil, P_Ticket, P_ResultadoAviso, P_Aviso, Btn_CrearTicket,
                        Btn_BackSearch, Btn_GuardarTicket, Btn_BackList, Btn_BackResult, Btn_BackTicket, Btn_BackResultAviso,
                        Btn_BackAviso, Btn_Logo, Fondo_Empresa, MEDICINE, OverBackMDN, "Logo_FGI.jpg", "Logo_FGI.jpg")
        ElseIf Empresa.Substring(0, 4) = "MD" Then
            'USUARIOS MEDICINE
            CambiarColor(P_Logo, P_Busqueda, P_ResultadoPerfil, P_Perfil, P_Ticket, P_ResultadoAviso, P_Aviso, Btn_CrearTicket,
                        Btn_BackSearch, Btn_GuardarTicket, Btn_BackList, Btn_BackResult, Btn_BackTicket, Btn_BackResultAviso,
                        Btn_BackAviso, Btn_Logo, Fondo_Empresa, MEDICINE, OverBackMDN, "Logo_Medicine.jpg", "Logo_Medicine.jpg")
        ElseIf Empresa.Substring(0, 4) = "MDTP" Then
            'USUARIOS MOLDAPTA
            CambiarColor(P_Logo, P_Busqueda, P_ResultadoPerfil, P_Perfil, P_Ticket, P_ResultadoAviso, P_Aviso, Btn_CrearTicket,
                        Btn_BackSearch, Btn_GuardarTicket, Btn_BackList, Btn_BackResult, Btn_BackTicket, Btn_BackResultAviso,
                        Btn_BackAviso, Btn_Logo, Fondo_Empresa, PLASTIKA, OverBackPSK, "Logo_Moldapta.jpg", "Logo_Moldapta.jpg")
        ElseIf Empresa.Substring(0, 4) = "PSKI" Then
            'USUARIOS PLASTIKISIMO
            CambiarColor(P_Logo, P_Busqueda, P_ResultadoPerfil, P_Perfil, P_Ticket, P_ResultadoAviso, P_Aviso, Btn_CrearTicket,
                        Btn_BackSearch, Btn_GuardarTicket, Btn_BackList, Btn_BackResult, Btn_BackTicket, Btn_BackResultAviso,
                        Btn_BackAviso, Btn_Logo, Fondo_Empresa, PLASTIKISIMO, OverBackPSKI, "Logo_Plastikisimo.jpg", "Logo_Plastikisimo.jpg")
        ElseIf Empresa.Substring(0, 4) = "PSK" Then
            'USUARIOS PLASTIKA
            CambiarColor(P_Logo, P_Busqueda, P_ResultadoPerfil, P_Perfil, P_Ticket, P_ResultadoAviso, P_Aviso, Btn_CrearTicket,
                        Btn_BackSearch, Btn_GuardarTicket, Btn_BackList, Btn_BackResult, Btn_BackTicket, Btn_BackResultAviso,
                        Btn_BackAviso, Btn_Logo, Fondo_Empresa, PLASTIKA, OverBackPSK, "Logo_PSK.png", "Logo_PSK.png")
        Else
            NotifyIcon.Visible = False
            NotifyIcon.Visible = True
            NotifyIcon.BalloonTipTitle = "Equipo no registrado"
            NotifyIcon.BalloonTipText = "Comunicate con sistemas para resolver tu problema."
            NotifyIcon.ShowBalloonTip(5000)
            BloqueoForm.Enabled = False
            Btn_Logo.Enabled = False
            Btn_MostrarAvisos.Enabled = False
            Btn_CrearTicket.Enabled = False
            P_Logo.BackColor = ColorTranslator.FromHtml(CORP)
        End If
    End Function

#End Region


#Region "ClassConfig | FUNCION GENERAL PARA CAMBIAR EL COLOR"
    Public Function CambiarColor(P_Logo As Panel, P_Busqueda As Panel, P_ResultadoPerfil As Panel, P_Perfil As Panel, P_Ticket As Panel, P_ResultadoAviso As Panel,
                                 P_Aviso As Panel, Btn_CrearTicket As PictureBox, Btn_BackSearch As Button, Btn_GuardarTicket As Button, Btn_BackList As Button,
                                 Btn_BackResult As Button, Btn_BackTicket As Button, Btn_BackResultAviso As Button, Btn_BackAviso As Button, Btn_Logo As PictureBox,
                                 Fondo_Empresa As PictureBox, EMPRESA As String, OverBackEMPRESA As String, Logo As String, Icono As String)
        P_Logo.BackColor = ColorTranslator.FromHtml(EMPRESA)
        P_Busqueda.BackColor = ColorTranslator.FromHtml(EMPRESA)
        P_ResultadoPerfil.BackColor = ColorTranslator.FromHtml(EMPRESA)
        P_Perfil.BackColor = ColorTranslator.FromHtml(EMPRESA)
        P_Ticket.BackColor = ColorTranslator.FromHtml(EMPRESA)
        P_ResultadoAviso.BackColor = ColorTranslator.FromHtml(EMPRESA)
        P_Aviso.BackColor = ColorTranslator.FromHtml(EMPRESA)
        Btn_CrearTicket.BackColor = ColorTranslator.FromHtml(EMPRESA)
        Btn_GuardarTicket.BackColor = ColorTranslator.FromHtml(EMPRESA)
        Btn_BackSearch.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackEMPRESA)
        Btn_BackList.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackEMPRESA)
        Btn_BackResult.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackEMPRESA)
        Btn_BackTicket.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackEMPRESA)
        Btn_BackResultAviso.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackEMPRESA)
        Btn_BackAviso.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackEMPRESA)
        Btn_GuardarTicket.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml(OverBackEMPRESA)
        Btn_Logo.Image = Image.FromFile(System.IO.Directory.GetCurrentDirectory + "\Resources\Empresas\Logos\" & Logo)
        Fondo_Empresa.BackgroundImage = Image.FromFile(System.IO.Directory.GetCurrentDirectory + "\Resources\Empresas\Iconos\" & Icono)
    End Function

#End Region


#Region "ClassConfig | FUNCION PARA REDIMENCIONAR EL SISTEMA"
    Public Function CircleSize(Medida As Object, Logo As Panel, Hora As String, Fecha As String, Tiempo As Integer)
        'Prepara formulario 
        Medida.Size = New Size(300, 300)
        Logo.Location = New Point(0, 0)
        'P_Perfil.Location = New Point(0, 0)
        Hora = String.Format("{0:hh:mm tt}", DateTime.Now)
        Fecha = DateTime.Now.ToLongDateString()
        Tiempo = 1000  ' Un segundo

        'La propiedad GraphicsPath representa 
        'una serie de lineas y curvas conectadas.
        Dim miPath As New System.Drawing.Drawing2D.GraphicsPath

        'Esta linea agrega un elipse al grafico
        'usando las propiedades de ancho y alto del From
        miPath.AddEllipse(0, 0, Medida.Width, Medida.Height)

        'Agregamos el path a una nueva propiedad de Region
        Dim miRegion As New Region(miPath)

        'Por ultimo asignamos nuestra region a la del Formulario
        Medida.Region = miRegion
    End Function
#End Region


#Region "ClassConfig | FUNCION PARA MONITOREAR EL MOVIMIENTO DEL USUARIO"
    Structure UltimoClick
        Public cbSize As Integer
        Public dwTime As Integer
    End Structure

    Dim Click As New UltimoClick()
    Public Function SesionSistema(Tiempo As Object)
        Click.cbSize = Marshal.SizeOf(Click)
        Tiempo.Interval = 1000
        Tiempo.Start()
    End Function
    Public Function TimerAction(Tiempo As Object, P_Logo As Panel, P_Buscar As Panel, P_ResultP As Panel, P_Perfil As Panel, P_Soporte As Panel, P_Ticket As Panel, P_ResultA As Panel, P_Aviso As Panel, NotifyIcon As NotifyIcon)
        GetLastInputInfo(Click)

        Dim Total As Integer = Environment.TickCount
        Dim Ultimo As Integer = Click.dwTime
        Dim Intervalo As Integer = (Total - Ultimo) / 1000

        Dim Segundos As Integer = 3
        If Intervalo >= Segundos Then
            Tiempo.Stop()
            'MsgBox(Convert.ToString(Segundos) + " SEGUNDOS DE INACTIVIDAD")
            NotifyIcon.Visible = False
            NotifyIcon.Visible = True
            NotifyIcon.BalloonTipTitle = "Inactividad en el sistema"
            NotifyIcon.BalloonTipText = "Será devuelto a la pantalla principal."
            NotifyIcon.ShowBalloonTip(5000)
            P_Logo.Location = New Point(0, 0)
            P_Buscar.Location = New Point(618, 0)
            P_ResultP.Location = New Point(927, 0)
            P_Perfil.Location = New Point(1236, 0)
            P_Soporte.Location = New Point(309, 309)
            P_Ticket.Location = New Point(618, 309)
            P_ResultA.Location = New Point(927, 309)
            P_Aviso.Location = New Point(1236, 309)
        End If
    End Function
#End Region


#Region "ClassConfig | FUNCION PARA MOVER PLANTILLAS"
    Public Function CambiarVentana(P1 As Panel, P1X As Integer, P1Y As Integer, P2 As Panel, P2X As Integer, P2Y As Integer)
        P1.Location = New Point(P1X, P1Y)
        P2.Location = New Point(P2X, P2Y)
    End Function
#End Region
End Class
