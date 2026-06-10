Attribute VB_Name = "modManError"
'*** Codificacion de Errores ************************
'*****************************************************
Public Const cData_RunSQLSentence As Integer = 90
Public Const cData_GetDataFromSQLSentence As Integer = 91
Public Const cData_RunSQLSoloDetalle As Integer = 92
Public Const cData_ConfiguraConexion As Integer = 93
Public Const cData_GetDataFromStoredProcedure As Integer = 94
Public Const cData_RunStoredProcedure As Integer = 95
Public Const cData_RunStoredProcedureSoloDetalle As Integer = 96

Public Const clsGeneral_GetCampos As Integer = 100
Public Const clsGeneral_AddNewBitacora As Integer = 101

'***************************************************
'                       Personal

Public Const clsPersonal_AddNewArea As Integer = 150
Public Const clsPersonal_UpdateArea As Integer = 151
Public Const clsPersonal_DeleteArea As Integer = 152
Public Const clsPersonal_GetArea As Integer = 153

Public Const clsPersonal_AddNewCargo As Integer = 154
Public Const clsPersonal_UpdateCargo As Integer = 155
Public Const clsPersonal_DeleteCargo As Integer = 156
Public Const clsPersonal_GetCargo As Integer = 157

Public Const clsPersonal_AddNewCategoria As Integer = 158
Public Const clsPersonal_UpdateCategoria As Integer = 159
Public Const clsPersonal_DeleteCategoria As Integer = 160
Public Const clsPersonal_GetCategoria As Integer = 161

Public Const clsPersonal_AddNewTipoHora As Integer = 162
Public Const clsPersonal_UpdateTipoHora As Integer = 163
Public Const clsPersonal_DeleteTipoHora As Integer = 164
Public Const clsPersonal_GetTipoHora As Integer = 165

Public Const clsPersonal_AddNewMotivoJustificacion As Integer = 166
Public Const clsPersonal_UpdateMotivoJustificacion As Integer = 167
Public Const clsPersonal_DeleteMotivoJustificacion As Integer = 168
Public Const clsPersonal_GetMotivoJustificacion As Integer = 169

Public Const clsPersonal_AddNewTrabajador As Integer = 170
Public Const clsPersonal_UpdateTrabajador As Integer = 171
Public Const clsPersonal_DeleteTrabajador As Integer = 172
Public Const clsPersonal_GetTrabajador As Integer = 173

Public Const clsPersonal_AddNewGrupo As Integer = 174
Public Const clsPersonal_UpdateGrupo As Integer = 175
Public Const clsPersonal_DeleteGrupo As Integer = 176
Public Const clsPersonal_GetGrupo As Integer = 177

Public Const clsPersonal_GetConsolidadoAsistencia As Integer = 178
Public Const clsPersonal_UpdateHolgura As Integer = 179
Public Const clsPersonal_GetHolgura As Integer = 180

Public Const clsPersonal_JustificarFeriadosRemunerados As Integer = 181
Public Const clsPersonal_AddNewFeriado As Integer = 182
Public Const clsPersonal_UpdateFeriado As Integer = 183
Public Const clsPersonal_DeleteFeriado As Integer = 184
Public Const clsPersonal_GetFeriado As Integer = 185

Public Const clsPersonal_GetUnicoTrabajador As Integer = 186
Public Const clsPersonal_GetCliente As Integer = 187
Public Const clsPersonal_AddNewCliente As Integer = 188
Public Const clsPersonal_UpdateCliente As Integer = 189

Public Const clsPersonal_AddNewMotivoJustificacionGrupo As Integer = 190
Public Const clsPersonal_UpdateMotivoJustificacionGrupo As Integer = 191
Public Const clsPersonal_DeleteMotivoJustificacionGrupo As Integer = 192
Public Const clsPersonal_GetMotivoJustificacionGrupo As Integer = 193
'**** Agregado para tolerancia
Public Const clsPersonal_UpdateTolerancia As Integer = 194





'***************************************************
'                       Horario

Public Const clsHorario_AddNewHorario As Integer = 1001
Public Const clsHorario_UpdateHorario As Integer = 1002
Public Const clsHorario_DeleteHorario As Integer = 1003
Public Const clsHorario_GetHorario As Integer = 1004

Public Const clsHorario_GetHorarioxTrabajador As Integer = 1005
Public Const clsHorario_DeleteHorarioxTrabajador As Integer = 1006
Public Const clsHorario_AddNewHorarioxTrabajador As Integer = 1007

Public Const clsHorario_UpdateMarcacion As Integer = 1008
Public Const clsHorario_GetCompromisoMarcacion As Integer = 1009

Public Const clsHorario_GetTrabajadorSinCompromiso As Integer = 1010
Public Const clsHorario_AddNewCompromisoMarcacionxTrabajador As Integer = 1011
Public Const clsHorario_DeleteMarcacion As Integer = 1012
Public Const clsHorario_GetHorarioxTrabajadorSinCompromiso As Integer = 1013
Public Const clsHorario_ValidateBussines As Integer = 1014
Public Const clsHorario_DeleteHorarioYMarcacion As Integer = 1015
Public Const clsHorario_GetHorarioGrupo As Integer = 1016
Public Const clsHorario_GetCompromisoHorarioxTrabajador As Integer = 1017
Public Const clsHorario_UpdateHorarioxTrabajador As Integer = 1018
Public Const clsHorario_UpdateHorarioxGrupo As Integer = 1019 'agregado para eliminar horarioxgrupo
Public Const clsHorario_Transaccion_Marcacion As Integer = 1020

'***************************************************
'                       Movimientos
Public Const clsMovimientos_UpdateHoraEntradaYSalidaMarcacion As Integer = 1999
Public Const clsMovimientos_GetMarcacionTrabajador As Integer = 2000
Public Const clsMovimientos_AddNewSuspension As Integer = 2001
Public Const clsMovimientos_UpdateSuspension As Integer = 2002
Public Const clsMovimientos_DeleteSuspension As Integer = 2003
Public Const clsMovimientos_GetSuspension As Integer = 2004

Public Const clsMovimientos_UpdateMarcacion As Integer = 2005
Public Const clsMovimientos_GetMarcacion As Integer = 2006
Public Const clsMovimientos_GetMarcacionVarios As Integer = 2007

Public Const clsMovimientos_GetJustificacion As Integer = 2008
Public Const clsMovimientos_GetFaltas As Integer = 2009
Public Const clsMovimientos_AddNewJustificacion As Integer = 2010
Public Const clsMovimientos_ExisteMesYAño As Integer = 2011
Public Const clsMovimientos_GetConsolidadosxTrabajador As Integer = 2012
Public Const clsMovimientos_GetRecordsetDesconectado As Integer = 2013
Public Const clsMovimientos_GetMotivoJustificacion As Integer = 2014
Public Const clsMovimientos_GetMarcacionesTrabajador As Integer = 2015
Public Const clsMovimientos_GetMinutoSuspensionTrabajador As Integer = 2016
Public Const clsMovimientos_UpdateEstadoSuspendidoToNormal As Integer = 2017
Public Const clsMovimientos_GetExisteConsolidado As Integer = 2018
Public Const clsMovimientos_ValidateBussines As Integer = 2019
Public Const clsMovimientos_GetTardanzas As Integer = 2020
Public Const clsMovimientos_GetJornadasIncompletas As Integer = 2021
Public Const clsMovimientos_DeleteJustificacion As Integer = 2022
' ******* Agregado para Consultar las marcaciones a eliminar
Public Const clsMovimientos_GetConsxEliminar As Integer = 2023
Public Const clsMovimientos_NewEliminar As Integer = 2024
Public Const clsMovimientos_GetConsHorarioxEliminar As Integer = 2025
Public Const clsMovimientos_AddNewPermiso As Integer = 2026
Public Const clsMovimientos_AddNewJustificacionGrupoDia As Integer = 2027

Public Const clsMovimientos_UpdateJustificacion As Integer = 2028

'***************************************************
'                       Reportes
Public Const clsReportes_GetReporteTrabajadores As Integer = 3011
Public Const clsReportes_GetReporteHorTrabajadores As Integer = 3012
Public Const clsReportes_GetReporteSuspTrabajadores As Integer = 3013
Public Const clsReportes_GetReporteJustificacionTrabajadores As Integer = 3014
Public Const clsReportes_GetRegistroNoMarcacionesD As Integer = 3015
Public Const clsReportes_GetRegistroNoMarcacionesL As Integer = 3016
Public Const clsReportes_GetRecordTardanzasD As Integer = 3017
Public Const clsReportes_GetRecordInasistenciasD As Integer = 3018
Public Const clsReportes_GetRecordInasistenciasL As Integer = 3019
Public Const clsReportes_GetReportConsolidados As Integer = 3029
Public Const clsReportes_GetDetalleAsistenciaxTrabajador As Integer = 3021
Public Const clsReportes_GetDetalleAsistenciaxTrabajadorL As Integer = 3020
Public Const clsReportes_GetBitacora As Integer = 3022
Public Const clsReportes_GetRecordSobretiemposD As Integer = 3023
Public Const clsReportes_GetRecordSobretiemposL As Integer = 3024
Public Const clsReportes_GetRecordTardanzasL As Integer = 3025
Public Const clsReportes_GetRegistroMarcacionesDia As Integer = 3026
Public Const clsReportes_GetRegistroMarcacionesFecha As Integer = 3027
Public Const clsReportes_GetToleranciaTardanza As Integer = 3028
Public Const clsReportes_GetToleranciaJornadaIncompleta As Integer = 3029
Public Const clsReportes_GetComparacionHoraTrab As Integer = 3030
Public Const clsReportes_GetRegistroNoMarcacionesAmbasD As Integer = 3031
Public Const clsReportes_GetRegistroNoMarcacionesAmbasL As Integer = 3032
Public Const clsReportes_GetRegistroNoMarcacionesESD As Integer = 3033
Public Const clsReportes_GetRegistroNoMarcacionesESL As Integer = 3034
Public Const clsReportes_GetDetalleAsistenciaxTrabajadorJust As Integer = 3035
Public Const clsReportes_GetRecordInasistenciasSoliDiaD As Integer = 3036
Public Const clsReportes_GetRecordInasistenciasSoliDiaL As Integer = 3037
Public Const clsReportes_GetComparacionHoraTrabL As Integer = 3038
Public Const clsReportes_GetReporteJustificacionTrabL As Integer = 3039
Public Const clsReportes_GetComparacionHoraTrabAsistencia As Integer = 3040
Public Const clsReportes_GetRegistroMarcacionesDiaTipo As Integer = 3041
Public Const clsReportes_GetRegistroMarcacionesFechaTipo As Integer = 3042

'***************************************************
'                       Consolidado

Public Const clsConsolidado_AddNewConsolidado As Integer = 4019
Public Const ClsConsolidado_AddNewDetalleConsolidado As Integer = 4020
Public Const clsConsolidado_GetIncidYMinutoSalidaAntRemunerada As Integer = 4021

Public Const clsPersonal_GetTotalHoraDebioTrabajar As Integer = 2022
Public Const clsPersonal_GetTotalHoraTrabajada As Integer = 2023
Public Const clsPersonal_GetTotalMinSuspension As Integer = 2024
Public Const clsPersonal_DeleteConsolidado As Integer = 2025
Public Const clsPersonal_DeleteDetalleConsolidado As Integer = 2026
Public Const clsPersonal_UpdateConsolidadoAsistencia As Integer = 2027
Public Const clsPersonal_SetCodigoConsolidadoEnMarcacion As Integer = 2028
Public Const clsConsolidado_GetIncidYMinutoSalidaAntNoRemunerada As Integer = 2029
Public Const clsConsolidado_GetIncidYMinutoInasNoRemunerada As Integer = 2030
Public Const clsConsolidado_GetIncidYMinutoTardRemunerado As Integer = 2031
Public Const clsConsolidado_GetIncidYMinutoTardNoRemunerado As Integer = 2032
Public Const clsConsolidado_GetIncidYMinutoInasRemunerada As Integer = 2033
Public Const clsConsolidado_JustificarFeriadosNoRemunerados As Integer = 2034
Public Const clsConsolidado_GetTrabajadorNoEliminado As Integer = 2035
Public Const clsConsolidado_GetTotalHoraExtra As Integer = 2036

'*********************************************
'                       Seguridad
Public Const clsSeguridad_GetUsuario As Integer = 3000
Public Const clsSeguridad_ValidateBussines As Integer = 3001
Public Const clsSeguridad_DeleteUsuario As Integer = 3002
Public Const clsSeguridad_UpdateUsuario As Integer = 3003
Public Const clsSeguridad_AddNewUsuario As Integer = 3004
Public Const clsSeguridad_GetOpcionMenuDeUsuario As Integer = 3005
Public Const clsSeguridad_GetValidarUsuario As Integer = 3006
Public Const clsSeguridad_GetAccesoUnicoUsuario As Integer = 3007
Public Const clsSeguridad_DeleteAccesoUnicoUsuario As Integer = 3008
Public Const clsSeguridad_UpdateAccesoUsuario As Integer = 3009
Public Const clsSeguridad_GetNombreCliente As Integer = 3010
Public Const clsSeguridad_GetOpcionesModulo As Integer = 3011

'*********************************************
'                       Crear Base
Public Const clsCreaBase_CreaBase As Integer = 4001
