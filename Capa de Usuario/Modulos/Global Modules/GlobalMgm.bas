Attribute VB_Name = "GlobalMgm"
Global cnxSQL As New ADODB.Connection
Global gstrUsuario As String
Global gintCursorLocation As Integer
Global Const gLinkColor = &H7D6546
Global Const gHoverColor = &H270B0
Global gstrCodigoUsusario As String
Global gstrNombreFormulario As String
Public Const gstrConexion = "Provider=SQLOLEDB.1;Password=12JULIOS;Persist Security Info=True;User ID=sa;Initial Catalog=PERSONAL;Data Source=PERSONAL"
Public Const Auth_ID = "9920114198077182"
Public Const gdatFechaInicioNull = "01/01/1900"
Public Const gdatFechaFinNull = "31/12/3000"
Global FRegTemplate As String

Public Const gstrHoraRemunerada As String = "S"
Public Const gstrHoraNoRemunerada As String = "N"
Public Const gstrFeriadoRemunerado As String = "S"
Public Const gstrFeriadoNoRemunerado As String = "N"
Public Const gdatFecha_Nula As Date = 1 / 1 / 2999
Public Const gstrEstadoTrabajadorNormal As String = "N"
'Public Const gstrEstadoTrabajadorSuspendido As String = "S"
'Public Const gstrEstadoTrabajadorEliminado As String = "E"
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


Public gobjReporte As Object

Sub Conectar()
If cnxSQL.State Then cnxSQL.Close
cnxSQL.CursorLocation = gintCursorLocation
cnxSQL.ConnectionString = GlobalMgm.gstrConexion
cnxSQL.Open
End Sub

Sub AccesLogin()
Dim Cadena As String
FrmPresentacion.Show
'If getKeyForm(App.Title, "Loaded") = "" Then
'    frmWelMGM.Show
'Else
'If getKeyForm(App.Title, "Agree") = "" Then
'    frmKeyMGM.Show
'Else
'    'nrodias >>> ViewTable
'    'activa >>> TimeOut
'    'nrousos >>> TimeConexion
'    'usosdia >>> System
'    'desactiva >>> ReadProxy
'    'ultidia >>> ReadOnly
'    'primerdia >>> SysBat
'    Dim dblNroDias As Double
'    Dim dblNroUsos As Double
'    Dim dblNroUsosDia As Double
'    Dim datUltimoDia As Date
'    Dim datDiaInstalacion As Date
'
'    Dim strActiva As String
'    Dim strDesactiva As String
'    Dim clave_ok As String
'
'    dblNroDias = Val(getKeyForm("System", "ViewTable"))
'    dblNroUsos = Val(getKeyForm("System", "TimeConexion"))
'    dblNroUsosDia = Val(getKeyForm("System", "System"))
'    clave_ok = getKeyForm(App.Title, "Identifier")
'    If clave_ok = "2817708914110299" Then
'        'SaveKeyForm App.Title, "Identifier", "2376498424535"
'        SaveKeyForm App.Title, "Identifier", "2817708914110299"
'        APIMgm.GrabarClave "SOFTWARE\" & "Root" & "\" & "System", "Identifier", "2817708914110299", cHKEY_LOCAL_MACHINE
'    End If
'    If clave_ok = "2376498424535" And clave_ok = APIMgm.LeerClave("SOFTWARE\" & "Root" & "\" & "System", "Identifier", cHKEY_LOCAL_MACHINE) Then
'        If getKeyForm("System", "TimeOut") = "" Then
'            SaveKeyForm "System", "TimeOut", "si"
'        End If
'        strActiva = getKeyForm("System", "TimeOut")
'
'        If getKeyForm("System", "ReadProxy") = "" Then
'            SaveKeyForm "System", "ReadProxy", "yes"
'        End If
'        strDesactiva = getKeyForm("System", "ReadProxy")
'
'        If strActiva = "no" Or strDesactiva = "not" Then
'            SaveKeyForm "System", "ViewTable", Str(31)
'            SaveKeyForm "System", "TimeConexion", Str(31)
'            SaveKeyForm "System", "System", Str(21)
'            SaveKeyForm "System", "TimeOut", "no"
'            SaveKeyForm "System", "ReadProxy", "not"
'            SaveKeyForm "System", "Identifier", "34234235345346"
'            APIMgm.GrabarClave "SOFTWARE\" & "Root" & "\" & "System", "Identifier", "34234235345346", cHKEY_LOCAL_MACHINE
'        End If
'
'        If getKeyForm("System", "ReadOnly") = "" Then
'            SaveKeyForm "System", "ReadOnly", CStr(Date)
'        End If
'        datUltimoDia = CDate(getKeyForm("System", "ReadOnly"))
'
'        If getKeyForm("System", "SysBat") = "" Then
'            SaveKeyForm "System", "SysBat", CStr(Date)
'        End If
'        datDiaInstalacion = CDate(getKeyForm("System", "SysBat"))
'
'        NroDiasInstalacion = DateDiff("d", datDiaInstalacion, Date)
'        If NroDiasInstalacion > 30 Then
'            SaveKeyForm "System", "ViewTable", Str(31)
'            SaveKeyForm "System", "TimeConexion", Str(31)
'            SaveKeyForm "System", "System", Str(21)
'            SaveKeyForm "System", "TimeOut", "no"
'            SaveKeyForm "System", "ReadProxy", "not"
'            SaveKeyForm "System", "Identifier", "34234235345346"
'            APIMgm.GrabarClave "SOFTWARE\" & "Root" & "\" & "System", "Identifier", "34234235345346", cHKEY_LOCAL_MACHINE
'        End If
'
'        If datUltimoDia = Date Then
'            If dblNroUsosDia > 20 Then
'                SaveKeyForm "System", "ViewTable", Str(31)
'                SaveKeyForm "System", "TimeConexion", Str(31)
'                SaveKeyForm "System", "System", Str(21)
'                SaveKeyForm "System", "TimeOut", "no"
'                SaveKeyForm "System", "ReadProxy", "not"
'                SaveKeyForm "System", "Identifier", "34234235345346"
'                APIMgm.GrabarClave "SOFTWARE\" & "Root" & "\" & "System", "Identifier", "34234235345346", cHKEY_LOCAL_MACHINE
'            Else
'                SaveKeyForm "System", "System", Str(dblNroUsosDia + 1)
'            End If
'        ElseIf datUltimoDia < Date Then
'            If NroDiasInstalacion > dblNroDias Then
'                SaveKeyForm "System", "ViewTable", Str(NroDiasInstalacion + 1)
'            Else
'                SaveKeyForm "System", "ViewTable", Str(dblNroDias + 1)
'            End If
'            SaveKeyForm "System", "TimeConexion", Str(dblNroUsos + 1)
'            SaveKeyForm "System", "ReadOnly", CStr(Date)
'            SaveKeyForm "System", "System", 0
'        Else
'            SaveKeyForm "System", "ViewTable", Str(31)
'            SaveKeyForm "System", "TimeConexion", Str(31)
'            SaveKeyForm "System", "TimeOut", "no"
'            SaveKeyForm "System", "ReadProxy", "not"
'            SaveKeyForm "System", "Identifier", "34234235345346"
'            APIMgm.GrabarClave "SOFTWARE\" & "Root" & "\" & "System", "Identifier", "34234235345346", cHKEY_LOCAL_MACHINE
'        End If
'
'        dblNroDias = Val(getKeyForm("System", "ViewTable"))
'        dblNroUsos = Val(getKeyForm("System", "TimeConexion"))
'        dblNroUsosDia = Val(getKeyForm("System", "System"))
'
'        If dblNroDias > 30 And dblNroUsos > 30 Then
'            'MsgBox "El Sistema a caducado por Favor contactese con Compusoft", , "Sistema Caducado"
'            FrmPruebaEvaluacion.Show
'            FrmPruebaEvaluacion.TxtUsuario = FrmPresentacion.TxtUsuario & Chr(13) & "El Sistema a caducado por Favor contactese con Compusoft"
'        Else
'            'MsgBox "Le quedan " & 30 - dblNroDias & " dias, como Prueba de Evaluación" & Chr(13) & "Y un total de " & 30 - dblNroUsosDia & " usos para este Dia", , "Sistema de Asistencia - Prueba de Evaluación"
'            FrmPresentacion.Show
'            FrmPresentacion.TxtUsuario = FrmPresentacion.TxtUsuario & Chr(13) & "Le quedan " & 30 - dblNroDias & " dias, como Prueba de Evaluación" & Chr(13) & "Y un total de " & 30 - dblNroUsosDia & " usos para este Dia"
'        End If
'    Else
'        FrmPruebaEvaluacion.Show
'        FrmPruebaEvaluacion.TxtUsuario = FrmPresentacion.TxtUsuario & Chr(13) & "La fecha del Sistema ha sido modificado y no se ejecutará" & Chr(13) & "Por Favor contactese con Compusoft"
'        'MsgBox "Faltan Archivos del Sistema, El Sistema ha sido modificado y no se ejecutará por Favor contactese con Compusoft", , "Sistema Caducado"
'        'End
'    End If
'End If
End Sub
Sub Main()
AccesLogin
End Sub
Function BuscarControl(ByRef Frm As Form, ByVal ctrl_name As String) As Long
Dim ctrl As Control
Dim i As Long
For i = 0 To Frm.Controls.Count - 1
    Set ctrl = Frm.Controls(i)
    If LCase(ctrl_name) = LCase(ctrl.Name) Then
        BuscarControl = i
        Exit Function
    End If
Next
BuscarControl = -1
End Function
Sub Subrayar(ByRef lbl As Label, ByVal Valor As Boolean)
If lbl.Font.Underline <> Valor Then
    lbl.Font.Underline = Valor
End If
End Sub
Sub Centrar(ByRef Frm As Form)
If Frm.WindowState = 0 Then
Frm.Top = (Screen.Height * 0.85 - Frm.Height) / 2
Frm.Left = (Screen.Width * 0.99 - Frm.Width) / 2
End If
End Sub
Sub Centrar_Ajustar(ByRef Frm As Form)
If Frm.WindowState = 0 Then
Frm.Width = Screen.Width * 0.99
Frm.Height = Screen.Height * 0.85
Centrar Frm
End If
End Sub

Function WeekDay(ByVal FECHA As String) As String
Dim dia_semana As Long
dia_semana = DatePart("w", FECHA)
Select Case dia_semana
'    Case 1: WeekDay = "Domingo  "
'    Case 2: WeekDay = "Lunes    "
'    Case 3: WeekDay = "Martes   "
'    Case 4: WeekDay = "Miércoles"
'    Case 5: WeekDay = "Jueves   "
'    Case 6: WeekDay = "Viernes   "
'    Case 7: WeekDay = "Sábado   "
    Case 1: WeekDay = 1
    Case 2: WeekDay = 2
    Case 3: WeekDay = 3
    Case 4: WeekDay = 4
    Case 5: WeekDay = 5
    Case 6: WeekDay = 6
    Case 7: WeekDay = 7
End Select
End Function

Function getNumberPart(ByVal texto As String) As Long
Dim longitud As Long
Dim i As Long
Dim letra As String
If IsNumeric(texto) Then
    getNumberPart = Val(texto)
    Exit Function
End If
longitud = Len(texto)
i = longitud
Do While i > 1
    letra = Mid(texto, i, 1)
    If letra Like "." Or letra Like "[A-Z]" Or letra Like "[a-z]" Then
        Exit Do
    Else
        i = i - 1
    End If
Loop
If i < longitud Then
    getNumberPart = Val(Mid(texto, i + 1))
Else
    getNumberPart = 0
End If
End Function
Function DaysOfMonth(ByVal Mes As Long, ByVal anio As Long) As Long
Select Case Mes
Case 1, 3, 5, 7, 8, 10, 12
    DaysOfMonth = 31
Case 2
    If (anio Mod 400 = 0) Or (anio Mod 4 = 0 And Not anio Mod 100 = 0) Then
        DaysOfMonth = 29
    Else
        DaysOfMonth = 28
    End If

Case 4, 6, 9, 11
    DaysOfMonth = 30
End Select
End Function
Sub ClearMaskEdit(ByRef txtMsk As Object)
    Dim msk As String
    msk = txtMsk.Mask
    txtMsk.Mask = ""
    txtMsk = ""
    txtMsk.Mask = msk
End Sub
Function NumberOfMonth(ByVal des_mes As String) As Long
NumberOfMonth = -1
Select Case LCase(des_mes)
    Case "enero"
        NumberOfMonth = 1
    Case "febrero"
        NumberOfMonth = 2
    Case "marzo"
        NumberOfMonth = 3
    Case "abril"
        NumberOfMonth = 4
    Case "mayo"
        NumberOfMonth = 5
    Case "junio"
        NumberOfMonth = 6
    Case "julio"
        NumberOfMonth = 7
    Case "agosto"
        NumberOfMonth = 8
    Case "setiembre"
        NumberOfMonth = 9
    Case "octubre"
        NumberOfMonth = 10
    Case "noviembre"
        NumberOfMonth = 11
    Case "diciembre"
        NumberOfMonth = 12
End Select
End Function
Sub SaveGridColumnWidth(ByRef dgrGrid As DataGrid)
Dim colGrid As Column
With dgrGrid
    For Each colGrid In dgrGrid.Columns
        APIMgm.SaveKeyForm dgrGrid.Parent.Name & "\" & .Name, colGrid.Caption, colGrid.Width
    Next
End With
End Sub
Sub GetGridColumnWidth(ByRef dgrGrid As DataGrid)
Dim colWidth As Integer
With dgrGrid
    For i = 0 To .Columns.Count - 1
        colWidth = Val(APIMgm.getKeyForm(dgrGrid.Parent.Name & "\" & .Name, .Columns(i).Caption))
        If colWidth > 0 Then .Columns(i).Width = colWidth
    Next
End With
End Sub
Sub SaveListColumnWidth(ByRef lstReport As ListView)
Dim colHeader As ColumnHeader
With lstReport
    For Each colHeader In lstReport.ColumnHeaders
        APIMgm.SaveKeyForm lstReport.Parent.Name & "\" & .Name, colHeader.Text, colHeader.Width
    Next
End With
End Sub
Sub GetListColumnWidth(ByRef lstReport As ListView)
Dim colWidth As Integer
With lstReport
    For i = 1 To .ColumnHeaders.Count
        colWidth = Val(APIMgm.getKeyForm(lstReport.Parent.Name & "\" & .Name, .ColumnHeaders(i).Text))
        If colWidth > 0 Then .ColumnHeaders(i).Width = colWidth
    Next
End With
End Sub
Sub SaveFormSettings(ByRef Frm As Form)
With Frm
    If .WindowState = vbNormal Then
        APIMgm.SaveKeyForm .Name, "Left", .Left
        APIMgm.SaveKeyForm .Name, "Top", .Top
        APIMgm.SaveKeyForm .Name, "Height", .Height
        APIMgm.SaveKeyForm .Name, "Width", .Width
    End If
End With
End Sub
Sub GetFormSettings(ByRef Frm As Form)
Dim getValue As String
With Frm
    If Not .WindowState = vbNormal Then Exit Sub
    getValue = APIMgm.getKeyForm(.Name, "Left")
    If IsNumeric(getValue) Then .Left = getValue
    getValue = APIMgm.getKeyForm(.Name, "Top")
    If IsNumeric(getValue) Then .Top = getValue
    getValue = APIMgm.getKeyForm(.Name, "Height")
    If Not .BorderStyle = vbSizable And Not .BorderStyle = vbSizableToolWindow Then Exit Sub
    If IsNumeric(getValue) Then .Height = getValue
    getValue = APIMgm.getKeyForm(.Name, "Width")
    If IsNumeric(getValue) Then .Width = getValue
End With
End Sub
Sub GetControlSettings(ByRef ctrl As Control)
Dim vGetValue As Variant
vGetValue = APIMgm.getKeyForm(ctrl.Parent.Name, ctrl.Name)
If TypeOf ctrl Is DataList Then
    ctrl.BoundText = vGetValue
ElseIf TypeOf ctrl Is DataCombo Then
    ctrl.BoundText = vGetValue
ElseIf TypeOf ctrl Is DTPicker Then
    If IsDate(vGetValue) Then ctrl.Value = vGetValue
ElseIf TypeOf ctrl Is TextBox Then
    If vGetValue <> "" Then ctrl.Text = vGetValue
ElseIf TypeOf ctrl Is CheckBox Then
    If IsNumeric(vGetValue) Then ctrl.Value = vGetValue
ElseIf TypeOf ctrl Is OptionButton Then
    If vGetValue <> "" Then ctrl.Value = CBool(vGetValue)
ElseIf TypeOf ctrl Is ComboBox Then
    If vGetValue <> "" Then ctrl.Text = vGetValue
End If
End Sub
Sub SaveControlSettings(ByRef ctrl As Control)
If TypeOf ctrl Is DataList Then
    Call APIMgm.SaveKeyForm(ctrl.Parent.Name, ctrl.Name, ctrl.BoundText)
ElseIf TypeOf ctrl Is DataCombo Then
    Call APIMgm.SaveKeyForm(ctrl.Parent.Name, ctrl.Name, ctrl.BoundText)
ElseIf TypeOf ctrl Is DTPicker Then
    Call APIMgm.SaveKeyForm(ctrl.Parent.Name, ctrl.Name, ctrl.Value)
ElseIf TypeOf ctrl Is TextBox Then
    Call APIMgm.SaveKeyForm(ctrl.Parent.Name, ctrl.Name, ctrl.Text)
ElseIf TypeOf ctrl Is CheckBox Then
    Call APIMgm.SaveKeyForm(ctrl.Parent.Name, ctrl.Name, ctrl.Value)
ElseIf TypeOf ctrl Is OptionButton Then
    Call APIMgm.SaveKeyForm(ctrl.Parent.Name, ctrl.Name, ctrl.Value)
ElseIf TypeOf ctrl Is ComboBox Then
    Call APIMgm.SaveKeyForm(ctrl.Parent.Name, ctrl.Name, ctrl.Text)
End If

End Sub
Sub SelTextBox(ByRef txt As TextBox)
Dim lenText As Integer
lenText = Len(txt)
If lenText > 0 Then
    txt.SelStart = 0
    txt.SelLength = lenText
End If
End Sub
Sub UpProgressBar(ByRef pb As ProgressBar, ByVal vValue As Integer)
If pb.Value + vValue <= pb.Max Then
    pb.Value = pb.Value + vValue
Else
    pb.Value = pb.Min
End If
End Sub
Sub LimpiarLista(ByRef lsvData As ListView)
Dim i As Integer
Dim J As Integer
For i = 1 To lsvData.ListItems.Count
    lsvData.ListItems(i).Text = " "
    For J = 1 To lsvData.ColumnHeaders.Count - 1
        lsvData.ListItems(i).ListSubItems(J).Text = " "
    Next
Next
End Sub
Function BuscarIzq(ByVal Cadena As String, ByVal Caracter As String) As Integer
For i = Len(Cadena) To 1 Step -1
    If Mid(Cadena, i, 1) = Caracter Then
        BuscarIzq = i
        Exit Function
    End If
Next
End Function
Function SeleccionarFila(ByVal CODIGO As Variant, ByRef lsvLista As ListView)
Dim i As Integer
Dim FlagSelect As Boolean
With lsvLista.ListItems
    For i = 1 To .Count
        If IsNumeric(CODIGO) Then
            If Val(CODIGO) = Val(.Item(i).Text) Then
                .Item(i).Selected = True
                .Item(i).EnsureVisible
                FlagSelect = True
            Else
                .Item(i).Selected = False
            End If
        Else
            If CODIGO = .Item(i).Text Then
                .Item(i).Selected = True
                FlagSelect = True
                .Item(i).EnsureVisible
                Exit For
            Else
                .Item(i).Selected = False

            End If
        End If
    Next
If Not FlagSelect And .Count > 0 Then
    .Item(1).Selected = True
End If
End With

End Function

Function ValidarCantidad(ByVal vDescripcion_Unidad As String, ByVal vTamano_Unidad As Double, ByVal vCantidad As String) As Boolean
ValidarCantidad = True
If vDescripcion_Unidad <> "METRO" And vDescripcion_Unidad <> "KILO" And vDescripcion_Unidad <> "ROLLO" And vCantidad Like "*.*" Then
    vCantidad_Verificar = Val(vCantidad) * vTamano_Unidad
    If vCantidad_Verificar Like "*.*" Then
        ValidarCantidad = False
    End If
End If
End Function

Public Function EjecutarProcedimientosAlmacenadosComplejo(ByVal ArraySP As Variant) As Boolean
Dim cmd As ADODB.Command

Dim i As Integer
'On Error GoTo ErrHandler
    Set cmd = New ADODB.Command
    cnxSQL.BeginTrans
    Set cmd.ActiveConnection = cnxSQL
    cmd.CommandType = adCmdStoredProc
    For i = LBound(ArraySP) To UBound(ArraySP)
        If ArraySP(i) = "" Then
        Exit For
        End If
        cmd.CommandText = ArraySP(i)
        cmd.Execute , , ADODB.adExecuteNoRecords
    Next
    
    Set cmd.ActiveConnection = Nothing
    cnxSQL.CommitTrans
    Set cmd = Nothing
    EjecutarProcedimientosAlmacenadosComplejo = True
'Exit Function
'ErrHandler:
'    cnxSQL.RollbackTrans
'    EjecutarProcedimientosAlmacenadosComplejo = False
'    Set cmd = Nothing
'    Err.Raise Err.Number, "ProDatos - RunStoredProcedureOnlyDetail", Err.Description
End Function


