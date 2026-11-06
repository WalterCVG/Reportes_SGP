VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form inicio 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Ingreso"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "INICIO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "INICIO.frx":0442
   ScaleHeight     =   5865
   ScaleMode       =   0  'User
   ScaleWidth      =   10012.19
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   8400
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton CmdChkPwd 
      Height          =   495
      Left            =   7080
      Picture         =   "INICIO.frx":4F2F
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Cambiar clave"
      Top             =   1800
      Width           =   615
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
      Left            =   6600
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   7
      Top             =   3240
      Width           =   795
   End
   Begin VB.TextBox Txusuario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   7800
      MaxLength       =   10
      PasswordChar    =   "&"
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton CmdIN 
      Caption         =   "Ingreso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   2
      Top             =   3240
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5520
      Top             =   4440
   End
   Begin VB.Label LblClienteIp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label LblNick 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nick"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label LblDia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LblDia"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5880
      TabIndex        =   13
      Top             =   5040
      Width           =   3255
   End
   Begin VB.Label LblDS 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LblDs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   6960
      TabIndex        =   12
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Lblhora 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   6120
      TabIndex        =   9
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODO"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   8
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Lblrazon 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   5760
      TabIndex        =   6
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Lblnombre 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6120
      TabIndex        =   3
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   8400
      Picture         =   "INICIO.frx":57F9
      ToolTipText     =   "Salir"
      Top             =   4080
      Width           =   480
   End
   Begin VB.Label Lblsistema 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Walter C. Valqui Guarniz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   5760
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sistemas H.E.P. Informatica "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   1455
   End
End
Attribute VB_Name = "inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dr As Scripting.Drive                  'objeto drive
Dim FSO As New Scripting.FileSystemObject  'objeto Filesystemobject
Dim unafecha As Date

'- Saber Ip cliente/host, agregar componente Microsoft Winsock control 6.0
'- agrega control winsock y usar sus propiedades ha criterio
Private Sub Form_Load()
  wflg = "NO"
  If App.PrevInstance Then
       MsgBox "Aplicación Ya esta  en Ejecución", vbExclamation, "Ya está siendo ejecutada"
       End
  End If
 
  Winsock1.RemoteHost = Winsock1.LocalHostName
  w_clienName = Winsock1.LocalHostName  ' captura Nombre Pc cliente de ejecucion
  w_clienteIP = Winsock1.LocalIP        ' captura IP cliente de ejecucion
  '-valida nombre estación max. 10posc.
  If Len(w_clienName) > 10 Then
     Beep 500, 300
     If MsgBox("Nombre del equipo EXCEDE lo permitido, continua ?  ", vbQuestion + vbYesNo, "[" + w_clienName + "]") = vbNo Then
        End
     End If
     w_clienName = Left(w_clienName, 10)
  End If
  LblDS.Caption = cICatal + " - " + IIf(Val(Right(cDSource, 2)) > 0, Right(cDSource, 2), Trim(cDSource))
  
  'Razon social
  wrazonsocial = "HOSPITAL DE EMERGENCIAS PEDIATRICAS"
  wsistema = "REPORTES adicionales SGP"
  
  Lblrazon.Caption = wrazonsocial
  Lblsistema.Caption = wsistema
  Txusuario = ""
  Txclave = ""
  CmdIN.Enabled = False
  Image2.Enabled = True
  genaño = Year(Date)
  TxAño = genaño:  w_ano = genaño
  genmes = Month(Date)
           wMeses = "ENERO    FEBRERO  MARZO    ABRIL    MAYO     JUNIO    "
  wMeses = wMeses + "JULIO    AGOSTO   SETIEMBREOCTUBRE  NOVIEMBREDICIEMBRE"
  '                    XXXXXXXX1XXXXXXXX2XXXXXXXX3XXXXXXXX4XXXXXXXX5XXXXXXXX6
  gennombmes = ""
  For x = 1 To 12
    If genmes = x Then
       gennombmes = Trim(Mid(wMeses, (x - 1) * 9 + 1, 9))
    End If
  Next x
  
  gendia = Day(Date)
  unafecha = Right("0" + Trim(Str(gendia)), 2) & "/" + Right("0" + Trim(Str(genmes)), 2) + "/" + Trim(Str(genaño))
  LblDia.Caption = WeekdayName(Weekday(unafecha)) + " " & gendia & " de " & gennombmes & " del " & genaño
  
  '- 02/Ene/2000  Domingo
  ''MsgBox (WeekdayName(Weekday("02/01/2000")) + " " + Format("02/01/2000", "dd/mm/yyyy"))
  If Weekday("02/01/2000") <> 1 Or UCase(Left(WeekdayName(Weekday("02/01/2000")), 3)) <> "DOM" Then  'No es Domingo dia Inicio semana
     MsgBox "[" + WeekdayName(Weekday(unafecha)) + "], Error, Inicio de Semana no es Domingo, Comunique a Sistemas.", vbCritical, "Configuración Calendario"
     End
  End If
  
   ' Dim cDSource, cICatal, CString, cUsr, cPwd As String
  
  cDSource = "192.168.2.10"     '"HPVALQUI"         '
  cICatal = "HepControlP"
  cUsr = "sa"
  cPwd = "P4ssw0rd"
   
   CString = "Provider=SQLOLEDB; Data Source='" + cDSource + "';" & _
       "initial catalog=" + cICatal + "; user id='" + cUsr + "'; password='" + cPwd + "';"
   Set cnx = New ADODB.Connection
   With cnx
      .ConnectionString = CString
      .Open
   End With
   cnx.CommandTimeout = 300
  
  Exit Sub
TrataError:
    Beep 500, 300
    For Each adoErr In cnx.Errors
        MsgBox _
        Prompt:=adoErr.Description, _
        Title:="Error Nro. " & CStr(adoErr.Number)
    Next
    End
End Sub

Private Sub Label4_Click()
   Load RegSerial
   RegSerial.Left = 0: RegSerial.Top = 0
   RegSerial.Refresh
   RegSerial.Show
End Sub

Private Sub Txusuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      '  Call chkNroserial
      If Trim(Txusuario) = "" Then
         Beep 500, 300
         MsgBox "No indico usuario..", vbCritical
         End
      End If
      Txusuario = UCase(Txusuario)
      Call Cmdfirmar
      
      Exit Sub
   End If
End Sub

Private Sub CmdChkPwd_Click()
    If wcambia = "SI" Or wProgUpr = "S" Then
       Load ChkPwd
       ChkPwd.Left = 0: ChkPwd.Top = 0
       ChkPwd.Refresh
       ChkPwd.Show
    End If
End Sub

Private Sub Cmdfirmar()
  Sw = False
  
  wrights = "OPE":  wDerEspec = "N": wTerminal = ""
  wccosto = "XX"
  If Trim(Txusuario) = "8E9P4SS" Or Trim(Txusuario) = "8E9OEI" Then
     '-Administrador
     wUsuario = "AMASTER"
     wNick = "AMaster"
     wUsuNombre = "AMASTER"
     wrights = "MAS"
     wDerEspec = "S"
     Lblnombre.Caption = wUsuNombre
     Sw = True
     wacceso1 = "XXXXXXXXXXXXXXX": wacceso2 = "XXXXXXXXXXXXXXX": wacceso3 = "XXXXXXXXXXXXXXX": wacceso4 = "XXXXXXXXXXXXXXX"
     wacceso5 = "XXXXXXXXXXXXXXX": wacceso6 = "XXXXXXXXXXXXXXX"
     wccosto = "90"
 
  Else
       '-Generales del servidor
       Set Rs = New ADODB.Recordset
       sql = "SELECT * FROM CPgeneral WHERE (baja is NULL) and documento='" & Trim(Txusuario) & "'"    'Dni
       '--  MsgBox "2.1 " + sql
       Rs.Open sql, cnx, adOpenStatic, adLockReadOnly, adCmdText
       If Not Rs.EOF Then
          wccosto = IIf(IsNull(Rs!ccosto), wccosto, Rs!ccosto)
          wNick = Left(Rs!nombres, 1) + Left(Rs!apelpat, 9)
          Lblnombre.Caption = Left(Trim(Rs!apelpat) + " " + Trim(Rs!apelmat) + " " + Trim(Rs!nombres), 35)
          wUsuNombre = Left(Trim(Rs!apelpat) + " " + Trim(Rs!apelmat) + " " + Trim(Rs!nombres), 35) + " (D)"
          Sw = True
       Else
            wccosto = "88"
            wUsuNombre = "INVITADO"
            Sw = True
            Lblnombre.Caption = wUsuNombre
       End If
       Rs.Close
        wacceso1 = "000000000000000": wacceso2 = "XXX000000000000": wacceso3 = "000000000000000"
        wacceso4 = "000000000000000":        wacceso5 = "000000000000000": wacceso6 = "000000000000000"
  
       #If False Then
       '-valida Usuario
       Set Rs = New ADODB.Recordset  ' <--- Agrega "ADODB." aquí
       sql = "SELECT * FROM RPTUsuarios WHERE Usuario='" & Txusuario & "'"
       Rs.Open sql, cnx, adOpenStatic, adLockReadOnly, adCmdText
       If Not Rs.EOF Then
          wUsuario = Rs!usuario
          wrights = Rs!derechos
          wDerEspec = IIf(IsNull(Rs!DerEspec), "", Rs!DerEspec)
          wBaja = IIf(IsNull(Rs!BAJA), "", Rs!BAJA)
          Lblnombre.Caption = Rs!Nombre
          wUsuNombre = Rs!Nombre
          wccosto = Rs!ccosto
          npos = InStr(Rs!Nombre, " ")
          wNick = Left(wUsuNombre, 1) + Trim(Mid(wUsuNombre, npos + 1, 9))
          wUsuNombre = Rs!Nombre + " (U)"
          wacceso1 = IIf(IsNull(Rs!acceso1), "000000000000000", Rs!acceso1): wacceso2 = IIf(IsNull(Rs!acceso2), "000000000000000", Rs!acceso2): wacceso3 = IIf(IsNull(Rs!acceso3), "000000000000000", Rs!acceso3): wacceso4 = IIf(IsNull(Rs!acceso4), "000000000000000", Rs!acceso4)
          wacceso5 = IIf(IsNull(Rs!acceso5), "000000000000000", Rs!acceso5): wacceso6 = IIf(IsNull(Rs!acceso6), "000000000000000", Rs!acceso6): wacceso7 = IIf(IsNull(Rs!acceso7), "000000000000000", Rs!acceso7)
          Sw = True
          '''MsgBox "1.usr"
       End If
       Rs.Close
       #End If
     
  End If
  If wccosto = "XX" Then
     Lblnombre.Caption = "<E> Servidor No autorizado..!"
     Txusuario = ""
     Txusuario.SetFocus
     Exit Sub
  End If
  
  LblNick.Caption = wNick
  LblClienteIp.Caption = w_clienteIP
  If w_clienteIP <> "" Then
      wTerminal = Right(w_clienteIP, 3)
      If Left(wTerminal, 1) = "." Then
         wTerminal = Right(w_clienteIP, 2)    '.19
      ElseIf Mid(wTerminal, 2, 1) = "." Then
            wTerminal = Right(w_clienteIP, 1)  '.7
      End If
  End If
     
  CmdIN.Enabled = True
  CmdIN.SetFocus
End Sub


Private Sub UpDown1_DownClick()
    If TxAño <= 1980 Then
       TxAño = 2090
       Exit Sub
    End If
    TxAño = TxAño - 1
    genaño = Trim(TxAño)
End Sub
Private Sub UpDown1_UpClick()
    If TxAño >= 2090 Then
       TxAño = 1980
       Exit Sub
    End If
    TxAño = TxAño + 1
    genaño = Trim(TxAño)
End Sub

Private Sub cmdIN_Click()
   If Trim(Txusuario) = "" Then
      Beep 500, 300
      MsgBox "No indico usuario..", vbCritical
      End
   End If
   wflg = "NO"
   w_ano = TxAño
   
   genaño = TxAño
   Unload Me
   ''Me.Hide
   Load RptMenu
   RptMenu.Left = 0: RptMenu.Top = 0
   RptMenu.Show
End Sub
Private Sub Image2_Click()
   cnx.Close
   cnx8.Close
   Set cnx = Nothing
   Set cnx8 = Nothing
   End
End Sub
Private Sub Timer1_Timer()
   Lblhora.Caption = Format(Time, "hh:mm:ss")
End Sub

#If False Then
Private Sub chkNroserial()
     For Each dr In FSO.Drives
         If UCase$(dr.DriveLetter) = "C" Then
            NumerodeSerie = CStr(dr.SerialNumber)
            Exit For
         End If
      Next
      encclave = ""
      WVencripta (NumerodeSerie)
      Set Rs = New Recordset
      sql = "SELECT * FROM CPregserial WHERE regserial='" & encclave & "'"
      Rs.Open sql, cnx, adOpenStatic, adLockReadOnly, adCmdText
      If Rs.EOF Then
         MsgBox ("EQUIPO NO Autorizado..!")
         Rs.Close
         Call Image2_Click
      End If
      Rs.Close
End Sub
#End If

Public Sub VALMAC()
   'Dim colNetAdapters, objWMIService As Object
   Dim strComputer As String
   '-strComputer = "."  direccion IP Servidor ha Scanear
   strComputer = cDSource
   'Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
   'Set colNetAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
   'For Each objitem In colNetAdapters
   '    If objitem.MACAddress = cMAC Then
          wacceso = True
   '    End If
   'Next
End Sub



'Nota:
'====
'Si quieres usar un icono para el ejecutable diferente al creado por el VB6, puedes hacer lo siguiente:
'-Añade un formulario al proyecto.
'-Cambia el icono del formulario por el que quieras usar.
'-En el cuado de propiedades de las propiedades del proyecto (menú Proyecto>Propiedades de Link...), selecciona la ficha Generar, y selecciona ese formulario del combo que hay junto a Icono.
'-El Objeto inicial (de la ficha General) debe seguir siendo Sub Main.
'-Compilar el proyecto y el ejecutable tendrá el icono indicado
