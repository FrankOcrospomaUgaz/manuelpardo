Attribute VB_Name = "ModErrHandler"
'Constantes para operaciones de Mantenimiento
Public Const MSG_ACTION_ADDNEW As String = "Crear"
Public Const MSG_ACTION_UPDATE As String = "Modificar"
Public Const MSG_ACTION_DELETE As String = "Eliminar"
Public Const MSG_ACTION_QUERY As String = "Consultar"
Public Const MSG_ACTION_ACEPTAR As String = "Aceptar"
Public Const MSG_ACTION_FINISH As String = "Finalizar"
Public Const MSG_ACTION_IMPRIMIR As String = "Imprimir"
Public Const MSG_ACTION_AGREGAR As String = "Agregar"
Public Const MSG_ACTION_QUITAR As String = "Quitar"
Public Const MSG_ACTION_SETTING As String = "Seleccionar"

'Constantes para mensajes
Public Const MSG_GENERAR As String = "GENERANDO"

Public Const MSG_RAISE_ERROR As String = "Se produjo un error inesperado."
Public Const MSG_SELECTION_EMPTY As String = "Debe seleccionar un item."
Public Const MSG_OPERATION_FAIL As String = "Falló la operación."
Public Const MSG_OPERATION_SUCESS As String = "La operación se realizó con éxito."
Public Const MSG_DUPLICATE_KEY As String = "El código ya existe."
Public Const MSG_INVALIDATE_KEY As String = "El dato ingresado no es valido !"
Public Const MSG_DUPLICATE_REGISTER As String = "El Registro ya existe."
Public Const MSG_DATE_CONSOLIDADO As String = "Imposible poner Feriado a una fecha donde ya se generó consolidado"
Public Const MSG_DATE_CONSOLIDADO_MARCACION As String = "Imposible Modificar Marcación, ya se generó consolidado"
Public Const MSG_RELATION_CHILD As String = "Este Código tiene información relacionada"
Public Const MSG_REQUERY_FILE As String = "Archivo requerido"
Public Const MSG_REQUERY_FIELD As String = "Campo requerido"
Public Const MSG_INVALIDATE_FILE As String = "El valor ingresado es invalido"
Public Const MSG_INVALIDATE_DATE As String = "La fecha de inicio debe ser menor"
Public Const MSG_INVALIDATE_DATEFIRST As String = "La Fecha de Inicio debe ser Mayor a la Fecha Actual"
Public Const MSG_QUESTION_DELETE As String = "Desea eliminar este registro ?"
Public Const MSG_QUESTION_UPDATE As String = "Desea modificar este registro ?"
Public Const MSG_EMPTY As String = "El detalle relacionado está vacío"
Public Const MSG_EMPTY_TIPOHORA As String = "No ha indicado Tipo de Hora"
Public Const MSG_QUESTION_CONTINUE As String = "Desea Continuar ?"

Public Const MSG_QUESTION_SAVE As String = "Desea Grabar los Cambios ?"
Public Const MSG_INVALIDATE_PASSWORD As String = "Error de Confirmación, vuelva a escribir su contraseña"
Public Const MSG_QUESTION_EXIT As String = "Desea abandonar la aplicación?"
Public Const MSG_QUESTION_CLOSE_SESSION As String = "Desea cerrar la Sesión e inciar con un usuario distinto?"

Public Const MSG_INFORM_NOTEXISTREGISTEREPORT As String = "No Existe Registros Para el Reporte"
Public Const MSG_INFORM_GENERATE As String = " ha sido  Generado"
Public Const MSG_INFORM_CIERREOPERATION As String = " Esta operacion No se le asignado una fecha de cierre"
Public Const MSG_INFORM_NOINCLUDE As String = " No se pueden Ingresar mas registros, El mes ha sido cerrado "

Public Const MSG_FRACASO_USUARIO_SISTEMA As String = "No es un Usuario autorizado del Sistema, vuelva a ingresar su Nombre de Usuario"
Public Const MSG_FRACASO_CONTRASENA_SISTEMA As String = "Contraseña incorrecta, vuelva a ingresar su clave"
Public Const MSG_EXIST_USER As String = "El Nombre de Usuario ya existe"


'Constantes de error por formulario
Public Const MGeneral_OrdenaGrilla As Integer = 500
Public Const MGeneral_SettingControls As Integer = 501
Public Const MDIPrincipal_Load As Integer = 502
Public Const MGeneral_InsertDataRS As Integer = 503
Public Const FrmSegAcceso_LoadData As Integer = 504
Public Const FrmSegAcceso_CmdAceptar_Click   As Integer = 505

Public Const FrmConsArea_LoadData As Integer = 1000
Public Const FrmConsArea_AddNew As Integer = 1001
Public Const FrmConsArea_Update As Integer = 1002
Public Const FrmConsArea_Delete As Integer = 1003

Public Const FrmMantArea_ShowForm As Integer = 1004
Public Const FrmMantArea_LoadData As Integer = 1005
Public Const FrmMantArea_AddNew As Integer = 1006
Public Const FrmMantArea_Update As Integer = 1007

Public Const FrmConsCargo_LoadData As Integer = 1008
Public Const FrmConsCargo_AddNew As Integer = 1009
Public Const FrmConsCargo_Update As Integer = 1010
Public Const FrmConsCargo_Delete As Integer = 1011

Public Const FrmMantCargo_ShowForm As Integer = 1012
Public Const FrmMantCargo_LoadData As Integer = 1013
Public Const FrmMantCargo_AddNew As Integer = 1014
Public Const FrmMantCargo_Update As Integer = 1015

Public Const FrmConsCategoria_LoadData As Integer = 1016
Public Const FrmConsCategoria_AddNew As Integer = 1017
Public Const FrmConsCategoria_Update As Integer = 1018
Public Const FrmConsCategoria_Delete As Integer = 1019

Public Const FrmMantCategoria_ShowForm As Integer = 1020
Public Const FrmMantCategoria_LoadData As Integer = 1021
Public Const FrmMantCategoria_AddNew As Integer = 1022
Public Const FrmMantCategoria_Update As Integer = 1023

Public Const FrmConsTipoHora_LoadData As Integer = 1024
Public Const FrmConsTipoHora_AddNew As Integer = 1025
Public Const FrmConsTipoHora_Update As Integer = 1026
Public Const FrmConsTipoHora_Delete As Integer = 1027

Public Const FrmMantTipoHora_ShowForm As Integer = 1028
Public Const FrmMantTipoHora_LoadData As Integer = 1029
Public Const FrmMantTipoHora_AddNew As Integer = 1030
Public Const FrmMantTipoHora_Update As Integer = 1031

Public Const FrmConsMotivoJustificacion_LoadData As Integer = 1032
Public Const FrmConsMotivoJustificacion_AddNew As Integer = 1033
Public Const FrmConsMotivoJustificacion_Update As Integer = 1034
Public Const FrmConsMotivoJustificacion_Delete As Integer = 1035

Public Const FrmMantMotivoJustificacion_ShowForm As Integer = 1036
Public Const FrmMantMotivoJustificacion_LoadData As Integer = 1037
Public Const FrmMantMotivoJustificacion_AddNew As Integer = 1038
Public Const FrmMantMotivoJustificacion_Update As Integer = 1039

Public Const FrmBusqTrabajador_ShowForm As Integer = 1040
Public Const FrmBusqTrabajador_LoadData As Integer = 1041
Public Const FrmBusqTrabajador_Queries As Integer = 1042

Public Const FrmConsTrabajador_LoadData As Integer = 1043
Public Const FrmConsTrabajador_AddNew As Integer = 1044
Public Const FrmConsTrabajador_Update As Integer = 1045
Public Const FrmConsTrabajador_Delete As Integer = 1046
Public Const FrmConsTrabajador_Suspender As Integer = 1047
Public Const FrmConsTrabajador_Print As Integer = 1048
Public Const FrmConsTrabajador_ShowForm As Integer = 1049

Public Const FrmMantTrabajador_ShowForm As Integer = 1050
Public Const FrmMantTrabajador_LoadData As Integer = 1051
Public Const FrmMantTrabajador_AddNew As Integer = 1052
Public Const FrmMantTrabajador_Update As Integer = 1053

Public Const FrmConsGrupo_LoadData As Integer = 1054
Public Const FrmConsGrupo_AddNew As Integer = 1055
Public Const FrmConsGrupo_Update As Integer = 1056
Public Const FrmConsGrupo_Delete As Integer = 1057

Public Const FrmMantGrupo_ShowForm As Integer = 1058
Public Const FrmMantGrupo_LoadData As Integer = 1059
Public Const FrmMantGrupo_AddNew As Integer = 1060
Public Const FrmMantGrupo_Update As Integer = 1061

Public Const FrmConsFeriado_LoadData As Integer = 1062
Public Const FrmConsFeriado_AddNew As Integer = 1063
Public Const FrmConsFeriado_ShowForm As Integer = 1064
Public Const FrmConsFeriado_Update As Integer = 1065
Public Const FrmConsFeriado_Delete As Integer = 1066

Public Const FrmMantFeriado_ShowForm As Integer = 1067
Public Const FrmMantFeriado_LoadData As Integer = 1068
Public Const FrmMantFeriado_AddNew As Integer = 1069
Public Const FrmMantFeriado_Update As Integer = 1070
Public Const FrmMantFeriado_ExisteConsolidado As Integer = 1071

Public Const FrmConsHorario_LoadData As Integer = 2000
Public Const FrmConsHorario_AddNew As Integer = 2001
Public Const FrmConsHorario_Update As Integer = 2002
Public Const FrmConsHorario_Delete As Integer = 2003

Public Const FrmMantHorario_ShowForm As Integer = 2004
Public Const FrmMantHorario_LoadData As Integer = 2005
Public Const FrmMantHorario_AddNew As Integer = 2006
Public Const FrmMantHorario_Update As Integer = 2007

Public Const FrmBusqHorarioxTrabajador_ShowForm As Integer = 2008
Public Const FrmBusqHorarioxTrabajador_LoadData As Integer = 2009
Public Const FrmBusqHorarioxTrabajador_Queries As Integer = 2010

Public Const FrmConsHorarioxTrabajador_LoadData As Integer = 2011
Public Const FrmConsHorarioxTrabajador_AddNew As Integer = 2012
Public Const FrmConsHorarioxTrabajador_Update As Integer = 2013
Public Const FrmConsHorarioxTrabajador_Delete As Integer = 2014

Public Const FrmMantHorarioxTrabajador_LoadData As Integer = 2015
Public Const FrmMantHorarioxTrabajador_ShowForm As Integer = 2016
Public Const FrmMantHorarioxTrabajador_AddNew As Integer = 2017
Public Const FrmMantHorarioxTrabajador_ConfiguraArray As Integer = 2018

Public Const FrmBusqCompromisoMarcacion_ShowForm As Integer = 2019
Public Const FrmBusqCompromisoMarcacion_LoadData As Integer = 2020
Public Const FrmBusqCompromisoMarcacion_Queries As Integer = 2021

Public Const FrmConsCompromisoMarcacion_LoadData As Integer = 2022
Public Const FrmConsCompromisoMarcacion_AddNew As Integer = 2023
Public Const FrmConsCompromisoMarcacion_Update As Integer = 2024
Public Const FrmConsCompromisoMarcacion_Delete As Integer = 2025
Public Const FrmConsCompromisoMarcacion_Imprimir As Integer = 2026

Public Const FrmMantCompromisoMarcacion_ShowForm As Integer = 2030
Public Const FrmMantCompromisoMarcacion_LoadData As Integer = 2031
Public Const FrmMantCompromisoMarcacion_AddNew As Integer = 2032
Public Const FrmMantCompromisoMarcacion_ConfiguraArray As Integer = 2033
Public Const FrmMantCompromisoMarcacion_LlenarRstPersonal As Integer = 2034

Public Const FrmMantHorarioxTrabajadorUpdate_ShowForm As Integer = 2035
Public Const FrmMantHorarioxTrabajadorUpdate_LoadData As Integer = 2036
Public Const FrmMantHorarioxTrabajadorUpdate_Update As Integer = 2037

Public Const FrmMantCompromisoMarcacionUpdate_ShowForm As Integer = 2060
Public Const FrmMantCompromisoMarcacionUpdate_LoadData As Integer = 2061
Public Const FrmMantCompromisoMarcacionUpdate_Update As Integer = 2062
Public Const FrmMantCompromisoMarcacionUpdate_ExisteConsolidado As Integer = 2063

'****************  agregado  FrmMantCompromisoMarcacionDelete **********
Public Const FrmMantCompromisoMarcacionDelete_ShowForm As Integer = 2064
Public Const FrmMantCompromisoMarcacionDelete_LoadData As Integer = 2065
Public Const FrmMantCompromisoMarcacionDelete_Update As Integer = 2066
Public Const FrmMantCompromisoMarcacionDelete_ValidateConsolidado As Integer = 2067


Public Const FrmConsSuspension_LoadData As Integer = 2070
Public Const FrmConsSuspension_AddNew As Integer = 2071
Public Const FrmConsSuspension_Update As Integer = 2072
Public Const FrmConsSuspension_Delete As Integer = 2073

Public Const FrmMantSuspension_ShowForm As Integer = 2080
Public Const FrmMantSuspension_LoadData As Integer = 2081
Public Const FrmMantSuspension_AddNew As Integer = 2082
Public Const FrmMantSuspension_Update As Integer = 2083
Public Const FrmMantSuspension_ValidateConsolidado As Integer = 2084

Public Const FrmMantMarcacion_ShowForm As Integer = 2090
Public Const FrmMantMarcacion_LoadData As Integer = 2091
Public Const FrmMantMarcacion_AddNew As Integer = 2092
Public Const FrmMantMarcacion_Load As Integer = 2093

Public Const FrmMantDetalleMarcacion_ShowForm As Integer = 2100
Public Const FrmMantDetalleMarcacion_LoadData As Integer = 2101
Public Const FrmMantDetalleMarcacion_AddNew As Integer = 2102

Public Const FrmBusqHorarioxGrupo_ShowForm As Integer = 2110
Public Const FrmBusqHorarioxGrupo_LoadData As Integer = 2111
Public Const FrmBusqHorarioxGrupo_Queries As Integer = 2112

Public Const FrmConsHorarioxGrupo_LoadData As Integer = 2120
Public Const FrmConsHorarioxGrupo_AddNew As Integer = 2121
Public Const FrmConsHorarioxGrupo_Update As Integer = 2122
Public Const FrmConsHorarioxGrupo_Delete As Integer = 2123

'nuevo deletegrupo debia estar en el grupo de FrmConsHorarioxTrabajador
Public Const FrmConsHorarioxTrabajador_DeleteGrupo As Integer = 2124
'FrmConsHorarioxTrabajador_DeleteGrupo
'FrmConsCompromisoMarcacion_DeleteGrupo


Public Const FrmMantHorarioxGrupo_LoadData As Integer = 2220
Public Const FrmMantHorarioxGrupo_ShowForm As Integer = 2221
Public Const FrmMantHorarioxGrupo_AddNew As Integer = 2222
Public Const FrmMantHorarioxGrupo_ConfiguraArray As Integer = 2223

Public Const FrmBusqJustificacion_ShowForm As Integer = 2230
Public Const FrmBusqJustificacion_LoadData As Integer = 2231
Public Const FrmBusqJustificacion_Queries As Integer = 2232

Public Const FrmBusqPermiso_ShowForm As Integer = 2233
Public Const FrmBusqPermiso_LoadData As Integer = 2234
Public Const FrmBusqPermiso_Queries As Integer = 2235

Public Const FrmConsJustificacion_LoadData As Integer = 2240
Public Const FrmConsJustificacion_AddNew As Integer = 2241
Public Const FrmConsJustificacion_Update As Integer = 2242
Public Const FrmConsJustificacion_Delete As Integer = 2243

Public Const FrmConsPermiso_LoadData As Integer = 2244
Public Const FrmConsPermiso_AddNew As Integer = 2245
Public Const FrmConsPermiso_Update As Integer = 2246
Public Const FrmConsPermiso_Delete As Integer = 2247

Public Const FrmMantJustificacion_LoadData As Integer = 2250
Public Const FrmMantJustificacion_ShowForm As Integer = 2251
Public Const FrmMantJustificacion_AddNew As Integer = 2252
Public Const FrmMantJustificacion_ConfiguraArray As Integer = 2253
Public Const FrmMantJustificacion_ValidateConsolidado As Integer = 2254

Public Const FrmMantPermiso_LoadData As Integer = 2262
Public Const FrmMantPermiso_ShowForm As Integer = 2263
Public Const FrmMantPermiso_AddNew As Integer = 2264
Public Const FrmMantPermiso_ConfiguraArray As Integer = 2264
Public Const FrmMantPermiso_ValidateConsolidado As Integer = 2265

'***FrmMantJustificacionGrupo
Public Const FrmMantJustificacionGrupo_LoadData As Integer = 2255
Public Const FrmMantJustificacionGrupo_ShowForm As Integer = 2256
Public Const FrmMantJustificacionGrupo_AddNew As Integer = 2257
Public Const FrmMantJustificacionGrupo_ConfiguraArray As Integer = 2258
Public Const FrmMantJustificacionGrupo_ValidateConsolidado As Integer = 2259
'FrmMantJustificacionGrupo_AddNewGrupo
Public Const FrmMantJustificacionGrupoDia_LoadData As Integer = 2262
Public Const FrmMantJustificacionGrupoDia_ShowForm As Integer = 2263
Public Const FrmMantJustificacionGrupoDia_AddNew As Integer = 2264
Public Const FrmMantJustificacionGrupoDia_ConfiguraArray As Integer = 2264
Public Const FrmMantJustificacionGrupoDia_ValidateConsolidado As Integer = 2265

Public Const FrmMantHolgura_CmdAceptar_Click As Integer = 2260
Public Const FrmMantHolgura_ShowForm As Integer = 2261

Public Const FrmConsMarcaciones_LoadData As Integer = 2270
Public Const FrmConsMarcaciones_Update As Integer = 2271
Public Const FrmConsMarcaciones_AddNew As Integer = 2272

Public Const FrmMantMarcaciones_LoadData As Integer = 2280
Public Const FrmMantMarcaciones_Update As Integer = 2281
Public Const FrmMantMarcaciones_ShowForm As Integer = 2282
Public Const FrmMantMarcaciones_CmdAceptar_Click As Integer = 2283
Public Const FrmRepTrabajadores_Load As Integer = 2284
'Muestra mensajes de error

Public Const FrmConsUsuario_LoadData As Integer = 2290
Public Const FrmConsUsuario_AddNew As Integer = 2291
Public Const FrmConsUsuario_Update As Integer = 2292
Public Const FrmConsUsuario_Delete As Integer = 2293

Public Const FrmMantUsuario_ShowForm As Integer = 2300
Public Const FrmMantUsuario_LoadData As Integer = 2301
Public Const FrmMantUsuario_AddNew As Integer = 2302
Public Const FrmMantUsuario_Update As Integer = 2303

Public Const FrmBusqUsuario_ShowForm As Integer = 2310
Public Const FrmBusqUsuario_LoadData As Integer = 2311
Public Const FrmBusqUsuario_Queries As Integer = 2312

Public Const FrmConsAccesoUsuario_LoadData As Integer = 2320
Public Const FrmConsAccesoUsuario_Update As Integer = 2321

Public Const FrmMantAccesoUsuario_cmdAceptar As Integer = 2330
Public Const FrmMantAccesoUsuario_LoadData As Integer = 2331
Public Const FrmMantAccesoUsuario_ConfiguraArray As Integer = 2334
Public Const FrmMantAccesoUsuario_Update As Integer = 2335

Public Const FrmConsEliminarHorario_LoadData As Integer = 2340
Public Const FrmConsEliminarHorario_ConfiguraArray As Integer = 2341
Public Const FrmConsEliminarHorario_Delete As Integer = 2342


Public Const FrmRepRecordInasistencias_Imprimir As Integer = 2350
Public Const FrmRepRecordInasistencias_Load As Integer = 2351
Public Const FrmRepListadoHorario_Imprimir As Integer = 2352
Public Const FrmRepListadoHorario_Load As Integer = 2353
Public Const FrmRepTrabajadores_Imprimir As Integer = 2354
Public Const FrmRegistroNoMarc_Imprimir As Integer = 2355
Public Const FrmRepRegistroNoMarc_Load As Integer = 2356
Public Const FrmRepListJust_Imprimir As Integer = 2357
Public Const FrmRepListJust_Load As Integer = 2358
Public Const FrmRepListSusp_Imprimir As Integer = 2359
Public Const FrmRepListSusp_Load As Integer = 2360
Public Const FrmRepConsolidado_Imprimir As Integer = 2361
Public Const FrmMantConsolidado_Imprimir As Integer = 2362
Public Const FrmMantConsolidado_CmdAceptar_Click As Integer = 2363
Public Const FrmRepConsolidado_Toolbar As Integer = 2364
Public Const FrmRepRecordTardanzas_Load As Integer = 2365
Public Const FrmRepConsolidadoxTrabajador_Imprimir As Integer = 2366
Public Const FrmRepDetAsisTrab_Imprimir As Integer = 2367
Public Const FrmRepRecordSobretiempos_Load As Integer = 2368
Public Const FrmRepRecordSobretiempos_Imprimir As Integer = 2369
Public Const FrmRepRecordTardanzas_Imprimir As Integer = 2370
Public Const FrmRepBitacora_Load As Integer = 2371
Public Const FrmRepBitacora_Imprimir As Integer = 2372
Public Const FrmRepComparacionHoras_Imprimir As Integer = 2373
Public Const FrmRepComparacionHoras_Load As Integer = 2374
Public Const FrmRepRegistroMarc_Load As Integer = 2375
Public Const FrmRepRegistroMarc_Imprimir As Integer = 2376

Public Const FrmMantClientes_ShowForm As Integer = 2377
Public Const FrmMantClientes_LoadData As Integer = 2378
Public Const FrmMantClientes_AddNew As Integer = 2379
Public Const FrmMantClientes_Update As Integer = 2376


Public Const FrmConsJustificacion_AddNewGrupo As Integer = 2380 'para FrmConsJustificacion_AddNewGrupo
Public Const FrmConsJustificacion_AddNewGrupoDia As Integer = 2381 'justificar grupo dias
Public Const FrmMantModJustificacion_LoadData As Integer = 2382
Public Const FrmMantModJustificacion_ShowForm As Integer = 2383
Public Const FrmMantModJustificacion_UpDate As Integer = 2384

Public Sub ErrorMessage(ByVal ilngNumber As Long, ByVal istrSource As String, ByVal istrDescription As String)
Dim strErrorText As String
MsgBox MSG_RAISE_ERROR & Chr(13) & Chr(13) & _
       "Codigo Error de SIAPER  : " & CStr(ilngNumber) & Chr(13) & Chr(13) & _
       "Origen                  : " & istrSource & Chr(13) & _
       "Descripcion             : " & istrDescription & Chr(13) & _
       "Texto del Error         : " & strErrorText & Chr(13), vbInformation, App.Title
End Sub
