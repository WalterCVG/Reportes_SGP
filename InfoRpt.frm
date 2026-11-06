VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Begin VB.Form InfoRpt 
   Caption         =   "Informes"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12750
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "InfoRpt.frx":0000
   ScaleHeight     =   9090
   ScaleWidth      =   12750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExportaCVS 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Exportar a Excell"
      Height          =   495
      Left            =   8520
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   8895
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      lastProp        =   600
      _cx             =   20981
      _cy             =   15690
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "InfoRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private crxApp As New CRAXDRT.Application
Private crxRpt As CRAXDRT.Report
Dim strSelectionfrm As String
Public Rst As ADODB.Recordset ' Cambiado a Public para asegurar que no se destruya prematuramente
Public NameDestino As String
Private Sub Form_Load()
    Dim crxParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crxParamDef As CRAXDRT.ParameterFieldDefinition
    Dim sql As String
    Dim WDETALLE As String
    Dim tabla As String
    Dim rutaRpt As String
    
    On Error GoTo ErrHandler
    
    ' 1. Definir tabla, ruta del reporte y título según la opción seleccionada
    Select Case RptMenu.Txtpoquery.Text
        Case "21"
            tabla = "TmpReporteEjecucionCitas"
            rutaRpt = App.Path & "\RptCitasEjecutadas.rpt"
            NameDestino = "CitasEjecutadas"
            WDETALLE = "Ejecución Citas programadas " & Format(ReporteCitas.TxFechaIni.Text, "dd/mm/yyyy") & " al : " & Format(ReporteCitas.TxFechaFin.Text, "dd/mm/yyyy")
        Case "22"
            tabla = "TmpReporteEjecucionCitasConsolidado"
            rutaRpt = App.Path & "\RptConsolidadoEjecucionDiaria.rpt"
            NameDestino = "ConsolidadoEjecucionDiaria"
            WDETALLE = "Consolidado Diario de Ejecución de Citas " & Format(ReporteCitas.TxFechaIni.Text, "dd/mm/yyyy") & " al : " & Format(ReporteCitas.TxFechaFin.Text, "dd/mm/yyyy")
        Case "23"
            tabla = "TmpReporteEjecucionCitasCrystal"
            rutaRpt = App.Path & "\RptEjecucionAutoevaluacion.rpt"
            NameDestino = "EjecucionAutoevaluacion"
            WDETALLE = "Seguimiento de Citas Médicas periodo del " & Format(ReporteCitas.TxFechaIni.Text, "dd/mm/yyyy") & " al : " & Format(ReporteCitas.TxFechaFin.Text, "dd/mm/yyyy")
        Case Else
            MsgBox "Tipo de reporte no válido.", vbExclamation
            Exit Sub
    End Select

    ' 2. Verificar la conexión global (cnx8) antes de hacer cualquier cosa
    On Error Resume Next
    If cnx8 Is Nothing Then
        MsgBox "La conexión a la base de datos no ha sido inicializada.", vbCritical
        Exit Sub
    End If
    If cnx8.State <> adStateOpen Then
        MsgBox "La conexión a la base de datos no está activa. Por favor, reinicie el proceso.", vbCritical
        Exit Sub
    End If
    On Error GoTo ErrHandler

    ' 3. Construir el SQL unificado (Ahora TODAS filtran por IdSesion)
    sql = "SELECT * FROM " & tabla & " WHERE IdSesion = '" & gIdSesion & "'"

    ' 4. Cargar el Recordset
    Set Rst = New ADODB.Recordset
    Rst.Open sql, cnx8, adOpenStatic, adLockReadOnly

    ' 5. Abrir el reporte de Crystal y enlazar los datos
    Set crxApp = New CRAXDRT.Application
    Set crxRpt = crxApp.OpenReport(rutaRpt, 1)
    
    crxRpt.DiscardSavedData
    crxRpt.EnableParameterPrompting = False
    crxRpt.Database.SetDataSource Rst

    ' 6. Asignar Parámetros
    Set crxParamDefs = crxRpt.ParameterFields
    For Each crxParamDef In crxParamDefs
        crxParamDef.ClearCurrentValueAndRange
        
        Select Case UCase(crxParamDef.ParameterFieldName)
            Case "RAZONSOCIAL"
               crxParamDef.AddCurrentValue wrazonsocial
            Case "SISTEMA"
                crxParamDef.AddCurrentValue wsistema
            Case "DETALLE"
                crxParamDef.AddCurrentValue WDETALLE
        End Select
    Next

    ' 7. Configurar el Visor y mostrar
    CRViewer1.DisplayGroupTree = False
    crxRpt.PrinterSetup (0)
    crxRpt.PaperSize = crPaperA4
    
    CRViewer1.ReportSource = crxRpt
    CRViewer1.ViewReport
    CRViewer1.Zoom (100)
    
    Screen.MousePointer = vbDefault
    
    Set crxParamDefs = Nothing
    Set crxParamDef = Nothing
    
    ' NOTA: El Recordset Rst NO se cierra aquí. Se cierra en Form_Unload.
    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
      Beep 500, 300
        Beep 500, 300
    If Err.Number = -2147206461 Then
        MsgBox "El archivo de reporte no se encuentra en la ruta especificada:" & vbCrLf & rutaRpt, vbCritical
    Else
        MsgBox "Error al cargar el reporte:" & vbCrLf & Err.Description, vbCritical
    End If
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub CmdExportaCVS_Click()

    Dim f As Integer
    Dim i As Integer
    Dim ruta As String
    Dim linea As String
    
     CreaDir ("C:\RPT")
     
    '----ruta = App.Path & "\ReporteCitas_" & Format(Now, "yyyymmdd_hhmmss") & ".csv"
    ruta = "C:\RPT\" & NameDestino & "_" & Format(Now, "yyyymmdd_hhmmss") & ".csv"
    f = FreeFile
    
    If Rst Is Nothing Or Rst.EOF Then
        MsgBox "No hay datos para exportar.", vbExclamation
        Exit Sub
    End If
    
    Open ruta For Output As #f
    
    ' Escribir encabezados
    For i = 0 To Rst.Fields.Count - 1
        linea = linea & """" & Rst.Fields(i).Name & """;"
    Next i
    Print #f, linea
    
    ' Escribir datos
    Rst.MoveFirst
    Do While Not Rst.EOF
        linea = ""
        For i = 0 To Rst.Fields.Count - 1
            If IsNull(Rst.Fields(i).Value) Then
                linea = linea & """"""
            Else
                linea = linea & """" & CStr(Rst.Fields(i).Value) & """"
            End If
            If i < Rst.Fields.Count - 1 Then linea = linea & ";"
        Next i
        Print #f, linea
        Rst.MoveNext
    Loop
    
    Close #f
    Rst.MoveFirst
    Beep 500, 300
    MsgBox "Archivo exportado exitosamente en:" & vbCrLf & ruta, vbInformation
End Sub



' ? EL RECORDSET SE CIERRA Y DESTRUYE SOLO CUANDO SE CIERRA LA VENTANA DEL REPORTE
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    If Not Rst Is Nothing Then
        If Rst.State = adStateOpen Then Rst.Close
        Set Rst = Nothing
    End If
    
    Set crxRpt = Nothing
    Set crxApp = Nothing
End Sub

#If False Then
Private Sub CmdExportar_Click()
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim i As Integer
    Dim fila As Long
    Dim ruta As String
    Dim creadoPorVB As Boolean
    
    On Error GoTo ErrHandler
    
    If Rst Is Nothing Then
        MsgBox "No hay datos para exportar.", vbExclamation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    ' 1. Intentar obtener una instancia de Excel ya abierta
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        ' No hay instancia, crear una nueva
        Set xlApp = CreateObject("Excel.Application")
        creadoPorVB = True
    End If
    On Error GoTo ErrHandler
    
    ' 2. Configurar Excel
    xlApp.Visible = True ' Lo hacemos visible para que el usuario vea el progreso
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)
    
    ' 3. Escribir encabezados
    For i = 0 To Rst.Fields.Count - 1
        xlSheet.Cells(1, i + 1).Value = Rst.Fields(i).Name
        xlSheet.Cells(1, i + 1).Font.Bold = True
        xlSheet.Cells(1, i + 1).Interior.Color = RGB(200, 200, 200)
    Next i
    
    ' 4. Escribir datos
    Rst.MoveFirst
    fila = 2
    Do While Not Rst.EOF
        For i = 0 To Rst.Fields.Count - 1
            If IsNull(Rst.Fields(i).Value) Then
                xlSheet.Cells(fila, i + 1).Value = ""
            Else
                xlSheet.Cells(fila, i + 1).Value = Rst.Fields(i).Value
            End If
        Next i
        Rst.MoveNext
        fila = fila + 1
    Loop
    
    ' 5. Ajustar y guardar
    xlSheet.Columns.AutoFit
    
    
    CreaDir ("C:\RPT")
    
    '--ruta = App.Path & "\ReporteCitas_" & Format(Now, "yyyymmdd_hhmmss") & ".xls"
    
    ruta = "C:\RPT\ReporteCitas_" & Format(Now, "yyyymmdd_hhmmss") & ".xls"
    xlBook.SaveAs ruta
    
    MsgBox "Archivo exportado exitosamente en:" & vbCrLf & ruta, vbInformation
    
    ' 6. Liberar memoria (el orden es crucial)
    Set xlSheet = Nothing
    Set xlBook = Nothing
    ' Solo cerramos Excel si fuimos nosotros quienes lo creamos
    If creadoPorVB Then xlApp.Quit
    Set xlApp = Nothing
    
    Rst.MoveFirst ' Regresar el cursor para no romper el reporte de Crystal
    
    Screen.MousePointer = vbDefault
    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    MsgBox "Error al exportar a Excel: " & Err.Description, vbCritical
    ' Asegurarse de liberar Excel incluso si hay error
    On Error Resume Next
    If Not xlApp Is Nothing Then
        If creadoPorVB Then xlApp.Quit
    End If
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
End Sub
#End If

