VERSION 5.00
Begin VB.Form ReporteCitas 
   Caption         =   "Reporte Citas SGP"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7800
   Icon            =   "ReporteCitas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "ReporteCitas.frx":030A
   ScaleHeight     =   4695
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSalir 
      Caption         =   "SALIR"
      Height          =   735
      Left            =   6600
      TabIndex        =   6
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton CmdGenerar 
      Caption         =   "Generar"
      Height          =   1095
      Left            =   2640
      TabIndex        =   5
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox TxFechaFin 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox TxFechaIni 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   7215
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   135
      Left            =   6240
      TabIndex        =   3
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Al:"
      Height          =   255
      Left            =   4800
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo  del:"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "ReporteCitas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cNamedbf As String

Private Sub Form_Load()
  
  cDSource8 = "192.168.1.8"     '"HPVALQUI"        '
  cICatal8 = "SIGH"
  cUsr8 = "sa"
  cPwd8 = "P4ssw0rd"
   
   CString8 = "Provider=SQLOLEDB; Data Source='" + cDSource8 + "';" & _
       "initial catalog=" + cICatal8 + "; user id='" + cUsr8 + "'; password='" + cPwd8 + "';"
   Set cnx8 = New ADODB.Connection
   With cnx8
      .ConnectionString = CString8
      .Open
   End With
   cnx8.CommandTimeout = 300

   Select Case RptMenu.Txtpoquery.Text
    Case "21"
      Label1.Caption = "Ejecución Citas del periodo"
      cNamedbf = "C:\RPT\EjecucionCitas.XLS"
    Case "22"
      Label1.Caption = "Consolidado Diario Ejecución Citas"
      cNamedbf = "C:\RPT\ConsolidadoCitas.XLS"
    Case "23"
      Label1.Caption = "Citas Consolidado de Ejecución - Autoevaluación"
      cNamedbf = "C:\RPT\CitasAutoevaluacion.XLS"
  End Select
  
  
End Sub

Private Sub cmdGenerar_Click()
    Dim FechaIni As String
    Dim FechaFin As String
    Dim IdServ   As Integer
    On Error GoTo ErrHandler
    FechaIni = Format(CDate(TxFechaIni.Text), "yyyy-mm-dd")
    FechaFin = Format(CDate(TxFechaFin.Text), "yyyy-mm-dd")
    IdServ = 0

    Screen.MousePointer = vbHourglass

    Select Case RptMenu.Txtpoquery.Text
        Case "21"
            gIdSesion = EjecutarReporte( _
                            "dbo.usp_select_ReporteEjecucionCitas_20261101", _
                            FechaIni, FechaFin, IdServ, 0, 0)
        Case "22"
            gIdSesion = EjecutarReporte( _
                            "dbo.usp_select_ReporteEjecucionCitasConsolidado_20261101", _
                            FechaIni, FechaFin, IdServ, 0, 0)
        Case "23"
            gIdSesion = EjecutarReporte( _
                            "dbo.usp_select_ReporteEjecucionCitasCrystal_20261101", _
                            FechaIni, FechaFin, IdServ, 0, 0)
    End Select

    Screen.MousePointer = vbDefault
    Call cmdImprimir_Click
    Exit Sub

ErrHandler:
    Screen.MousePointer = vbDefault
    MsgBox "Error al generar el reporte: " & vbCrLf & Err.Description, vbCritical
End Sub

Private Function EjecutarReporte( _
        ByVal NombreUSP As String, _
        ByVal FechaIni As String, _
        ByVal FechaFin As String, _
        ByVal IdServ As Integer, _
        ByVal IdEsp As Integer, _
        ByVal IdMed As Integer) As String

    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cnx8
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = NombreUSP
    cmd.CommandTimeout = 0

    ' 1. Parámetros comunes para TODOS los reportes
    cmd.Parameters.Append cmd.CreateParameter("@FechaInicio", adDate, adParamInput, , CDate(FechaIni))
    cmd.Parameters.Append cmd.CreateParameter("@FechaFin", adDate, adParamInput, , CDate(FechaFin))
    cmd.Parameters.Append cmd.CreateParameter("@IdServicio", adInteger, adParamInput, , IdServ)

    ' 2. Evaluamos si es el reporte "21" (que solo lleva 4 parámetros)
    ' Si no es el 21 (es decir, es el 22 o 23), agregamos Especialidad y Médico
    If InStr(1, NombreUSP, "Consolidado", vbTextCompare) > 0 Or _
       InStr(1, NombreUSP, "Crystal", vbTextCompare) > 0 Then
       
        cmd.Parameters.Append cmd.CreateParameter("@IdEspecialidad", adInteger, adParamInput, , IdEsp)
        cmd.Parameters.Append cmd.CreateParameter("@IdMedico", adInteger, adParamInput, , IdMed)
        
    End If

    ' 3. El parámetro de salida va al final
    cmd.Parameters.Append cmd.CreateParameter("@IdSesion", adGUID, adParamOutput, 16)

    cmd.Execute
    EjecutarReporte = CStr(cmd.Parameters("@IdSesion").Value)
    Set cmd = Nothing
End Function


Private Sub cmdImprimir_Click()
    Dim rsCheck As ADODB.Recordset
    Dim sSQL    As String
    On Error GoTo ErrHandler

    Select Case RptMenu.Txtpoquery.Text
        Case "21": sSQL = "SELECT COUNT(*) AS Total FROM TmpReporteEjecucionCitas WHERE IdSesion = '" & gIdSesion & "'"
        Case "22": sSQL = "SELECT COUNT(*) AS Total FROM TmpReporteEjecucionCitasConsolidado WHERE IdSesion = '" & gIdSesion & "'"
        Case "23": sSQL = "SELECT COUNT(*) AS Total FROM TmpReporteEjecucionCitasCrystal WHERE IdSesion = '" & gIdSesion & "'"
    End Select

    Set rsCheck = cnx8.Execute(sSQL)
    If rsCheck!Total = 0 Then
        MsgBox "No hay datos para imprimir. Genere el reporte primero.", vbExclamation
        rsCheck.Close
        Exit Sub
    End If
    rsCheck.Close

    InfoRpt.Left = 0: InfoRpt.Top = 0
    InfoRpt.Show
    Exit Sub

ErrHandler:
    MsgBox "Error al imprimir: " & vbCrLf & Err.Description, vbCritical
End Sub
Private Sub cmdsalir_Click()
   #If False Then
   For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        ' Form Abierto
        If Forms(cont).Name <> "RptMenu" Then
           Unload Forms(cont)
        End If
      End If
   Next
   #End If
   RptMenu.Txtpoquery.Text = ""
   Me.Hide
   Unload Me
End Sub

