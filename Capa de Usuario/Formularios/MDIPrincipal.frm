VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H00807000&
   Caption         =   "SISTEMA DE CONTROL DE PERSONAL - Versión 3.0 - Menú Principal - COMPUSOFT"
   ClientHeight    =   6825
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9060
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIPrincipal.frx":08CA
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar SbPrincipal 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6450
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
            Picture         =   "MDIPrincipal.frx":1E18D
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1E887
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   2822
            MinWidth        =   2822
            Picture         =   "MDIPrincipal.frx":1F161
            TextSave        =   "07:38 p.m."
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMantenimiento 
      Caption         =   "&Mantenimientos"
      Begin VB.Menu mnuArea 
         Caption         =   "Area"
      End
      Begin VB.Menu mnuCargo 
         Caption         =   "Cargo"
      End
      Begin VB.Menu mnuCategoria 
         Caption         =   "Categoría"
      End
      Begin VB.Menu mnuLinea11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTipoHora 
         Caption         =   "Tipo de Hora"
      End
      Begin VB.Menu mnuMotivoJustificacion 
         Caption         =   "Motivo de Justificación"
      End
      Begin VB.Menu mnuFeriado 
         Caption         =   "Feriados"
      End
      Begin VB.Menu mnuLinea12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGrupoTrabajador 
         Caption         =   "Grupo de Trabajador"
      End
      Begin VB.Menu mnuLinea14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrabajador 
         Caption         =   "Trabajador"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuLinea13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCliente 
         Caption         =   "Cliente"
      End
      Begin VB.Menu mnuUsuarios 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu mnuDefinirAccesoUsuario 
         Caption         =   "Accesos de Usuario"
      End
      Begin VB.Menu mnuLinea15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuMovimientos 
      Caption         =   "M&ovimientos"
      Begin VB.Menu mnuDefinirHorarios 
         Caption         =   "Definir Horarios"
         Begin VB.Menu mnuHorarioxTrabajador 
            Caption         =   "Horario por Trabajador"
         End
         Begin VB.Menu mnuHorarioxGrupos 
            Caption         =   "Horario por Grupos"
         End
         Begin VB.Menu mnuHorarioxDias 
            Caption         =   "Horario por Días"
         End
      End
      Begin VB.Menu mnuModificarHorarios 
         Caption         =   "Modificar Horarios"
      End
      Begin VB.Menu mnuLinea21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuJustificaciones 
         Caption         =   "Justificaciones"
      End
      Begin VB.Menu mnuSuspensiones 
         Caption         =   "Suspensiones"
      End
      Begin VB.Menu mnuPermisos 
         Caption         =   "Permisos"
      End
      Begin VB.Menu mnuLinea22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenerarCompromisosMarcacion 
         Caption         =   "Generar Compromisos de Marcación"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuLinea23 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMarcacion 
         Caption         =   "Marcación"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuConsAsistMens 
         Caption         =   "Consolidado de Asistencia Mensual"
      End
   End
   Begin VB.Menu mnuReportes 
      Caption         =   "&Reportes"
      Begin VB.Menu mnuListaTrab 
         Caption         =   "Listado de Trabajadores"
      End
      Begin VB.Menu mnuLinea31 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListadoHor 
         Caption         =   "Listado de Horario"
      End
      Begin VB.Menu mnuLinea32 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRegMarc 
         Caption         =   "Registro de Marcaciones"
      End
      Begin VB.Menu mnuComparacion 
         Caption         =   "Comparación horas trabajadas"
      End
      Begin VB.Menu mnuLinea33 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListaSusp 
         Caption         =   "Listado de Suspensiones"
      End
      Begin VB.Menu mnuListJust 
         Caption         =   "Listado de Justificaciones"
      End
      Begin VB.Menu mnuRegNoMarc 
         Caption         =   "Registro de No Marcaciones"
      End
      Begin VB.Menu mnuRecordTard 
         Caption         =   "Record de Tardanzas"
      End
      Begin VB.Menu mnuRecordDeInasist 
         Caption         =   "Record de Inasistencias"
      End
      Begin VB.Menu mnuRecordDeSobretiempo 
         Caption         =   "Récord de Sobretiempos"
      End
      Begin VB.Menu mnuLinea34 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDetAsisTrab 
         Caption         =   "Detalle de Asistencia por Trabajador"
      End
      Begin VB.Menu mnuLinea35 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConsMens 
         Caption         =   "Consolidado Mensual de Asistencia"
      End
      Begin VB.Menu mnuConsolidadoTrabajador 
         Caption         =   "Consolidado de Asistencia por Trabajador"
      End
      Begin VB.Menu mnuLinea36 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBitacora 
         Caption         =   "Bitácora"
      End
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Activate()
   Me.SbPrincipal.Panels(1).Text = "Usuario : " + gstrNombreUsuario
   Me.SbPrincipal.Panels(1).MinWidth = Len(Me.SbPrincipal.Panels(1).Text) * 100 + 500
   Me.SbPrincipal.Panels(2).Text = "Cliente: " + gstrNameClient
   Me.SbPrincipal.Panels(2).MinWidth = Len(Me.SbPrincipal.Panels(2).Text) * 100 + 500
   Me.SbPrincipal.Panels(3).MinWidth = MDIPrincipal.Width - Me.SbPrincipal.Panels(1).MinWidth - Me.SbPrincipal.Panels(4).MinWidth - Me.SbPrincipal.Panels(2).MinWidth
   'mnuMarcacion_Click
End Sub

Private Sub MDIForm_Load()
    'ConfiguraNombreCliente
    gaParametrosReport(2, 1) = "Pliego"
    gaParametrosReport(2, 2) = gstrNameClient
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuArea_Click()
    FrmConsArea.ShowForm
End Sub

Private Sub mnuBitacora_Click()
    FrmRepBitacora.ShowForm
End Sub

Private Sub mnuCargo_Click()
    FrmConsCargo.ShowForm
End Sub

Private Sub mnuCategoria_Click()
    FrmConsCategoria.ShowForm
End Sub

Private Sub mnuCliente_Click()
    FrmMantClientes.ShowForm
End Sub

Private Sub mnuComparacion_Click()
    FrmRepComparacionHoras.ShowForm
End Sub

Private Sub mnuConsAsistMens_Click()
    FrmConsConsolidadoMensual.ShowForm
End Sub

Private Sub mnuConsMens_Click()
    FrmRepConsolidado.ShowForm
End Sub

Private Sub mnuConsolidadoTrabajador_Click()
    FrmRepConsolidadoxTrabajador.ShowForm
End Sub

Private Sub mnuDefinirAccesoUsuario_Click()
    FrmBusqUsuario.strModo = Query
    FrmBusqUsuario.ShowForm
End Sub

Private Sub mnuDetAsisTrab_Click()
    FrmRepDetAsisTrab.ShowForm
End Sub

Private Sub mnuFeriado_Click()
    FrmConsFeriado.ShowForm
End Sub

Private Sub mnuGenerarCompromisosMarcacion_Click()
    FrmBusqCompromisosMarcacion.ShowForm
End Sub

Private Sub mnuGrupoTrabajador_Click()
    FrmConsGrupo.ShowForm
End Sub

Private Sub mnuHorario_Click()
    FrmConsHorario.ShowForm
End Sub

Private Sub mnuHolguras_Click()
    FrmMantHolguras.ShowForm
End Sub

Private Sub mnuHorarioxDias_Click()
    FrmBusqHorarioxDia.ShowForm
End Sub

Private Sub mnuHorarioxGrupos_Click()
    FrmBusqHorarioxGrupo.ShowForm
End Sub

Private Sub mnuHorarioxTrabajador_Click()
    FrmBusqHorarioxTrabajador.ShowForm
End Sub

Private Sub mnuJustificaciones_Click()
   FrmBusqJustificacion.ShowForm
End Sub

Private Sub mnuListadoHor_Click()
    frmRepListadoHorario.ShowForm
End Sub

Private Sub mnuListaSusp_Click()
    frmRepListSusp.ShowForm
End Sub

Private Sub mnuListaTrab_Click()
    frmRepTrabajadores.ShowForm
End Sub

Private Sub mnuListJust_Click()
    FrmRepListJust.ShowForm
End Sub

Private Sub mnuMarcacion_Click()
    FrmMantMarcacion.intAction = enumActions.AddNew
    FrmMantMarcacion.ShowForm
End Sub

Private Sub mnuModifMarc_Click()
    FrmBusqEliminarHorario.ShowForm

End Sub

Private Sub mnuModificarHorarios_Click()
    FrmMantHorarioxTrabajadorBK.intAction = 1
    FrmMantHorarioxTrabajadorBK.ShowForm
End Sub

Private Sub mnuMotivoJustificacion_Click()
    FrmConsMotivoJustificacion.ShowForm
End Sub

Private Sub mnuPermisos_Click()
    FrmBusqPermiso.ShowForm
End Sub

Private Sub mnuRecordDeInasist_Click()
    FrmRepRecordInasist.ShowForm
End Sub

Private Sub mnuRecordDeSobretiempo_Click()
    FrmRepRecordSobretiempo.ShowForm
End Sub

Private Sub mnuRecordTard_Click()
    frmRepRecordTard.ShowForm
End Sub

Private Sub mnuRegMarc_Click()
    FrmRepRegistroMarc.ShowForm
End Sub

Private Sub mnuRegNoMarc_Click()
    FrmRepRegistroNoMarc.ShowForm
End Sub

Private Sub mnuSalir_Click()
Unload Me
End Sub

Private Sub mnuSuspensiones_Click()
    FrmConsSuspension.ShowForm
End Sub

Private Sub mnuTipoHora_Click()
    FrmConsTipoHora.ShowForm
End Sub

Private Sub mnuTrabajador_Click()
    FrmBusqTrabajador.strModo = Query
    FrmBusqTrabajador.ShowForm
End Sub

Private Sub mnuUsuarios_Click()
    FrmConsUsuario.ShowForm
End Sub
