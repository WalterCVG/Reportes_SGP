VERSION 5.00
Begin VB.Form RptMenu 
   Caption         =   "Menu Principal"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   9390
   LockControls    =   -1  'True
   Picture         =   "RptMenu.frx":0000
   ScaleHeight     =   6930
   ScaleWidth      =   9390
   Begin VB.TextBox Txtpoquery 
      Height          =   375
      Left            =   8520
      MaxLength       =   2
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   1080
      Top             =   6000
   End
   Begin VB.TextBox TxAño 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   7800
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   3
      Top             =   5640
      Width           =   795
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   5040
   End
   Begin VB.Label lblinicio 
      BackStyle       =   0  'Transparent
      Caption         =   "Lblinicio"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Lblping 
      BackStyle       =   0  'Transparent
      Caption         =   "ping"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo Activo"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   7560
      TabIndex        =   4
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label LblDS 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LblDS"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   6960
      TabIndex        =   1
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label Lblhora 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "18:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Menu opcion_1 
      Caption         =   "&Historias"
      Begin VB.Menu opc11 
         Caption         =   "rpt11"
      End
      Begin VB.Menu opc12 
         Caption         =   "rpt12"
      End
      Begin VB.Menu raya11 
         Caption         =   "-"
      End
   End
   Begin VB.Menu Opcion_2 
      Caption         =   "&Citas"
      Begin VB.Menu opc21 
         Caption         =   "&Ejecución de Citas"
      End
      Begin VB.Menu Opc22 
         Caption         =   "&Consolidado Diario"
      End
      Begin VB.Menu opc23 
         Caption         =   "&Citas AUTOEVALUACIÓN"
      End
   End
   Begin VB.Menu opcion_3 
      Caption         =   "&Oficina 1"
   End
   Begin VB.Menu Opcion_4 
      Caption         =   "&Oficina 2"
   End
   Begin VB.Menu Opcion_5 
      Caption         =   "&Utilitarios"
      Begin VB.Menu opc51 
         Caption         =   "&USUARIOS"
      End
      Begin VB.Menu opc52 
         Caption         =   "File configurador"
      End
      Begin VB.Menu opc59 
         Caption         =   "Varios"
      End
   End
   Begin VB.Menu opcout 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "RptMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim nIntervPing As Integer
Public wtpoquery As String
Private Sub Form_Load()
   lblinicio.Caption = Format(Time, "hh:mm:ss")
   Lblping.Caption = ""
   nmintPing = 0    'control mints  para lanzar PING
   Me.Caption = "Menu Principal - " & genaño & "   -   Sesión: " + wUsuNombre
   LblDS.Caption = cICatal + "-" + UCase(IIf(Val(Right(cDSource, 2)) > 0, Right(cDSource, 2), Trim(cDSource)))
   TxAño = genaño
   If wUsuario <> "AMASTER" Then
      ' Opc29.Visible = False
   End If
   Dim wtpoquery As String
End Sub

Private Sub opc11_Click()
   acceso ("11")
   If wacceso Then
      'Load tCadena
      'tCadena.Left = 0: tCadena.Top = 0
      'tCadena.Refresh
      'tCadena.Show
   Else
      MsgBox "Usuario No Autorizado", vbExclamation
      Exit Sub
   End If
End Sub
Private Sub opc12_Click()
   'acceso ("12")
   If wacceso Then
      'Load ccosto
      'ccosto.Left = 0: ccosto.Top = 0
      'ccosto.Refresh
      'ccosto.Show
   Else
      MsgBox "Usuario No Autorizado", vbExclamation
      Exit Sub
   End If
End Sub
Private Sub opc13_Click()
   'acceso ("13")
   If wacceso Then
   Else
      MsgBox "Usuario No Autorizado", vbExclamation
      Exit Sub
   End If
End Sub
Private Sub opc14_Click()
  ' acceso ("14")
   If wacceso Then
   Else
      MsgBox "Usuario No Autorizado", vbExclamation
      Exit Sub
   End If
End Sub

Private Sub opc21_Click()
 
 acceso ("21")

  ' If wacceso Then
      '-EjecucionCitas
      Txtpoquery = "21"
      Load ReporteCitas
      ReporteCitas.Left = 0: ReporteCitas.Top = 0
      ReporteCitas.Refresh
      ReporteCitas.Show
 '  Else
  '    MsgBox "Usuario No Autorizado", vbExclamation
 '     Exit Sub
  ' End If
End Sub
Private Sub opc22_Click()
   acceso ("22")
  
   'If wacceso Then
      Txtpoquery = "22"
      '-EjecucionCitasConsolidado
      Load ReporteCitas
      ReporteCitas.Left = 0: ReporteCitas.Top = 0
      ReporteCitas.Refresh
      ReporteCitas.Show
 '  Else
  '    MsgBox "Usuario No Autorizado", vbExclamation
 '     Exit Sub
  ' End If
End Sub
Private Sub opc23_Click()
 
   acceso ("23")
   'If wacceso Then
      '-CONSOLIDADO_CitasAutoevaluado
      Txtpoquery = "23"
      Load ReporteCitas
      ReporteCitas.Left = 0: ReporteCitas.Top = 0
      ReporteCitas.Refresh
      ReporteCitas.Show
  ' Else
  '    MsgBox "Usuario No Autorizado", vbExclamation
  '    Exit Sub
  ' End If
End Sub

Private Sub opc24_Click()
   acceso ("24")
   If wacceso Then
   Else
      MsgBox "Usuario No Autorizado", vbExclamation
      Exit Sub
   End If
End Sub
Private Sub opc25_Click()
   acceso ("25")
   If wacceso Or weditor = "S" Then
   Else
      MsgBox "Usuario No Autorizado ", vbExclamation
      Exit Sub
   End If
End Sub

Private Sub opc31_Click()
   acceso ("31")
   If wacceso Then
   Else
      MsgBox "Usuario No Autorizado", vbExclamation
      Exit Sub
   End If
End Sub
Private Sub opc32_Click()
   acceso ("32")
   If wacceso Then
   Else
      MsgBox "Usuario No Autorizado", vbExclamation
      Exit Sub
   End If
End Sub
Private Sub opc33_Click()
   acceso ("33")
   If wacceso Then
   Else
      MsgBox "Usuario No Autorizado", vbExclamation
      Exit Sub
   End If
End Sub

Private Sub opc34_Click()
   acceso ("34")
   If wacceso Then
   Else
      MsgBox "Usuario No Autorizado", vbExclamation
      Exit Sub
   End If
End Sub

Private Sub opc41_Click()
    acceso ("41")
    If wacceso Then
    Else
      MsgBox "Usuario No Autorizado", vbExclamation
      Exit Sub
    End If
End Sub

Private Sub opc42_Click()
    acceso ("42")
    If wacceso Then
    Else
      MsgBox "Usuario No Autorizado", vbExclamation
      Exit Sub
    End If
End Sub
Private Sub opc43_Click()
    acceso ("43")
    If wacceso Then
    Else
      MsgBox "Usuario No Autorizado", vbExclamation
      Exit Sub
    End If
End Sub
Private Sub opc44_Click()
    acceso ("44")
    If wacceso Then
    Else
      MsgBox "Usuario No Autorizado", vbExclamation
      Exit Sub
    End If
End Sub

Private Sub opc51_Click()
    acceso ("51") Or wrights = "MAS"
    If wacceso Then
       Load RptUsuario
       RptUsuario.Left = 0: RptUsuario.Top = 0
       RptUsuario.Refresh
       RptUsuario.Show
    Else
      MsgBox "Usuario No Autorizado", vbExclamation
      Exit Sub
    End If
End Sub
Private Sub opc52_Click()
    acceso ("52")
    If wacceso Or wrights = "MAS" Then
    Else
      MsgBox "Usuario No Autorizado", vbExclamation
      Exit Sub
    End If
End Sub
Private Sub opc53_Click()
    acceso ("53")
    If wacceso Or wrights = "MAS" Then
    
    Else
      MsgBox "Usuario No Autorizado", vbExclamation
      Exit Sub
    End If
End Sub
Private Sub opc54_Click()
    acceso ("54")
    If wacceso Or wrights = "MAS" Then
    Else
      MsgBox "Usuario No Autorizado", vbExclamation
      Exit Sub
    End If
End Sub
Private Sub opc59_Click()
    If wrights = "MAS" Then
    
       '--Call Ver_xls
       '''Call IniciaClave
       Load Form1
       Form1.Left = 0: Form1.Top = 0
       Form1.Refresh
       Form1.Show
    Else
      MsgBox "Acceso denegado..", vbExclamation
      Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
   On Error Resume Next
   Dim ws As Workspace
   Dim db As Database
   Dim Rs As Recordset
   For Each ws In Workspaces
       For Each db In ws.Databases
           'For Each Rs In db.Recordsets
           '    Rs.Close
           '    Set Rs = Nothing
           'Next
           'db.Close
           Set db = Nothing
       Next
       ws.Close
       Set ws = Nothing
   Next
   End
End Sub

#If False Then
Private Sub CmdManual_Click()
   cObjeto = wmanualPdf
   Call CmdVerPDF_Click
End Sub
Private Sub CmdVerPDF_Click()
   ShellExecute hwnd, "Open", Trim(cObjeto), vbNullString, vbNullString, SW_NORMAL
End Sub
#End If
   
   
Private Sub opcout_Click()
   Unload Me
   On Error Resume Next
   Dim ws As Workspace
   Dim db As Database
   Dim Rs As Recordset
   For Each ws In Workspaces
       For Each db In ws.Databases
           'For Each Rs In db.Recordsets
           '   Rs.Close
           '   Set Rs = Nothing
           'Next
           'db.Close
           Set db = Nothing
       Next
       ws.Close
       Set ws = Nothing
   Next
   If cnx.State = adStateOpen Then cnx.Close
   Set cnx = Nothing
   
   End
End Sub

Private Sub Timer1_Timer()
   Lblhora.Caption = Format(Time, "hh:mm:ss")
End Sub
Private Sub Timer2_Timer()
   '-lanza ping cada 7 minutos (intervalo 60000milisegundos)
   nIntervPing = nIntervPing + 1
   If nIntervPing >= 7 Then
      nIntervPing = 0
      Lblping.Caption = Format(Time, "hh:mm:ss")
      Dim oShell As Object
      Set oShell = CreateObject("WScript.Shell")
      oShell.Run "ping -n 1 192.168.1.12", 0, True
      Set oShell = Nothing
   End If
End Sub
'Public Function CloseAllRecordsets() As Integer
'On Error Resume Next
'-Be sure to Close all Data Objects upon Exit - and release memory
'-Put it on close of the Main Menu - (make sure you have no bound forms open)
'Dim wsCurr As Workspace
'Dim dbCurr As Database
'Dim rs As Recordset
'Dim Frm As Form
'For Each wsCurr In Workspaces
'   For Each dbCurr In wsCurr.Databases
'      For Each rs In dbCurr.Recordsets
'         MsgBox "The recordset " & vbCrLf & rs.Name & vbCrLf & _
'         rs.RecordCount & " record(s) was left open - now closing it.", vbCritical, "Validation"
'         rs.Close
'         Set rs = Nothing
'      Next
'      dbCurr.Close
'      Set dbCurr = Nothing
'    Next
'    wsCurr.Close
'    Set wsCurr = Nothing
'Next
'End Function

Private Sub UpDown1_DownClick()
    If TxAño = 2012 Then
       Exit Sub
    End If
    For cont = Forms.Count - 1 To 0 Step -1
       If Forms(cont).Visible = False Then
       Else
          ' Form Abierto
          If Forms(cont).Name <> "CpMenu" Then
             Unload Forms(cont)
          End If
       End If
    Next
    
    TxAño = TxAño - 1
    'wflg = "NO"
    'w_ano = TxAño
    'Call CompruebaTabla
    'If wflg = "NO" Then
    '   TxAño = TxAño + 1
    '   Exit Sub
    'End If
    
    genaño = Trim(TxAño)
    Me.Caption = "Menu Principal - " & genaño & "   -   Sesión: " + wUsuNombre
End Sub
Private Sub UpDown1_UpClick()
    For cont = Forms.Count - 1 To 0 Step -1
       If Forms(cont).Visible = False Then
       Else
          ' Form Abierto
          If Forms(cont).Name <> "CpMenu" Then
             Unload Forms(cont)
          End If
       End If
    Next
       
    TxAño = TxAño + 1
    'wflg = "NO"
    'w_ano = TxAño
    'Call CompruebaTabla
    'If wflg = "NO" Then
    '   TxAño = TxAño - 1
    '   Exit Sub
    'End If
    
    genaño = Trim(TxAño)
    Me.Caption = "Menu Principal - " & genaño & "   -   Sesión: " + wUsuNombre
End Sub
Public Sub acceso(x)
   wacceso = False
   Select Case x
   Case "11"
     If Mid(wacceso1, 1, 1) = "X" Then wacceso = True
    Case "12"
     If Mid(wacceso1, 2, 1) = "X" Then wacceso = True
    Case "13"
     If Mid(wacceso1, 3, 1) = "X" Then wacceso = True
   Case "14"
     If Mid(wacceso1, 4, 1) = "X" Then wacceso = True
   Case "15"
     If Mid(wacceso1, 5, 1) = "X" Then wacceso = True
   Case "16"
     If Mid(wacceso1, 6, 1) = "X" Then wacceso = True
   
   Case "21"
     If Mid(wacceso2, 1, 1) = "X" Then wacceso = True
    Case "22"
     If Mid(wacceso2, 2, 1) = "X" Then wacceso = True
    Case "23"
     If Mid(wacceso2, 3, 1) = "X" Then wacceso = True
    Case "24"
     If Mid(wacceso2, 4, 1) = "X" Then wacceso = True
    Case "25"
     If Mid(wacceso2, 5, 1) = "X" Then wacceso = True
    Case "26"
     If Mid(wacceso2, 6, 1) = "X" Then wacceso = True
    
    Case "31"
     If Mid(wacceso3, 1, 1) = "X" Then wacceso = True
   Case "32"
     If Mid(wacceso3, 2, 1) = "X" Then wacceso = True
   Case "33"
     If Mid(wacceso3, 3, 1) = "X" Then wacceso = True
   Case "34"
     If Mid(wacceso3, 4, 1) = "X" Then wacceso = True
   Case "35"
     If Mid(wacceso3, 5, 1) = "X" Then wacceso = True
     
   Case "41"
     If Mid(wacceso4, 1, 1) = "X" Then wacceso = True
   Case "42"
     If Mid(wacceso4, 2, 1) = "X" Then wacceso = True
   Case "43"
     If Mid(wacceso4, 3, 1) = "X" Then wacceso = True
   Case "44"
     If Mid(wacceso4, 4, 1) = "X" Then wacceso = True
   Case "45"
     If Mid(wacceso4, 5, 1) = "X" Then wacceso = True
   End Select
End Sub

