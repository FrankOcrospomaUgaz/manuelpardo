Attribute VB_Name = "modConstantes"
'****** Constantes del sistema
Public Const gstrHoraRemunerada As String = "S"
Public Const gstrHoraNoRemunerada As String = "N"
Public Const gstrFeriadoRemunerado As String = "S"
Public Const gstrFeriadoNoRemunerado As String = "N"
Public Const gdatFecha_Nula As Date = 1 / 1 / 2999
Public Const gdatFecha_Nula_RangoInicio As Date = 1 / 1 / 1900
Public Const gdatFecha_Nula_RangoFin As Date = 31 / 12 / 2999
Public Const gstrEstadoTrabajadorNormal As String = "N"
Public Const gstrEstadoTrabajadorSuspendido As String = "S"
Public Const gstrEstadoTrabajadorEliminado As String = "E"
Public Const gstrSuspensionEstaActiva As String = "A"
Public Const gstrSuspensionEstaInactiva As String = "I"
Public Const gintMinutosHolgura As Integer = 60 'minutos
Public Const gstrNoMarcoEntrada As String = "NME"
Public Const gstrNoMarcoSalida As String = "NMS"
Public Const gstrEntrada As String = "ENTRADA"
Public Const gstrSalida As String = "SALIDA"
Public Const gintGrupoxDefecto As Integer = 0
Public Const gintMotivoFeriadoRemunerado As Integer = 1
Public Const gintMotivoFeriadoNoRemunerado As Integer = 2
Public Const gintMotivoPermiso As Integer = 3
Public Const gintHoraExtra As Integer = 2
Public Const gstrCompromisoNoGenerado As String = 0
Public Const gstrCompromisoGenerado As String = 1
Public Const gintUsuarioAdministrador As Integer = 1
'**** Rango para la Hora
Public Const gdatHora_Nula_RangoInicio As Date = "00:00:00"
Public Const gdatHora_Nula_RangoFin As Date = "23:59:59"

'********** Permiso **********
Public Const gintPermiso As Integer = 3


'******* Constantes para mensajes **********
Public Const MSG_RELATION_CHILD As String = "Este Elemento tiene información relacionada"

'******* enumerados para validaciones ******

Public Enum EnumActionsHorario
    DeleteHorarioxTrabajador = 1
    DeleteMarcacion
    UpdateHorarioxTrabajador
    UpdateHorarioxGrupo
    DeleteHorarioxGrupo
End Enum

Public Enum EnumActionsMovimientos
    DeleteJustificacion = 1
    DeleteSuspension
    DeleteFeriado
    AddNewFeriado
    UpdateFeriado
End Enum

Public Enum EnumActionsSeguridad
    Delete_Usuario = 1
    Delete_Sistema
    Delete_Opcion
End Enum

Public Enum EnumActionsPersonal
    Delete_Area = 1
    Delete_Cargo
    Delete_Grupo
    Delete_Categoria
    Delete_MotivoJustificacion
    Delete_TipoHora
    Delete_Trabajador
End Enum

