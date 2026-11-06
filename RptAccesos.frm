VERSION 5.00
Begin VB.Form RptAccesos 
   Caption         =   "Accesos"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "RptAccesos.frx":0000
   ScaleHeight     =   3045
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Chk2 
      Caption         =   "Citas Autoevaluación"
      Height          =   350
      Index           =   2
      Left            =   1920
      TabIndex        =   18
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "Consolidado Diario"
      Height          =   350
      Index           =   1
      Left            =   1920
      TabIndex        =   17
      Top             =   840
      Width           =   1815
   End
   Begin VB.CheckBox Chk3 
      Caption         =   "Query 2"
      Height          =   350
      Index           =   1
      Left            =   4080
      TabIndex        =   15
      Top             =   840
      Width           =   1575
   End
   Begin VB.CheckBox Chk5 
      Caption         =   "Configurador"
      Height          =   350
      Index           =   3
      Left            =   8040
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "Rpt2"
      Height          =   350
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   1575
   End
   Begin VB.CheckBox Chk5 
      Height          =   350
      Index           =   2
      Left            =   8040
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.CheckBox Chk5 
      Caption         =   "Usuarios"
      Height          =   350
      Index           =   1
      Left            =   8040
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.CheckBox Chk3 
      Caption         =   "Query Libre"
      Height          =   350
      Index           =   0
      Left            =   4080
      TabIndex        =   12
      Top             =   480
      Width           =   1575
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "Ejecución Citas"
      Height          =   350
      Index           =   0
      Left            =   1920
      TabIndex        =   10
      Top             =   480
      Width           =   1575
   End
   Begin VB.CheckBox Chk4 
      Caption         =   "Listado"
      Height          =   350
      Index           =   1
      Left            =   6120
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "Rpt1"
      Height          =   350
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   7320
      Picture         =   "RptAccesos.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cerrar"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   4920
      TabIndex        =   7
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CheckBox Chk4 
      Caption         =   "Reporte"
      Height          =   350
      Index           =   0
      Left            =   6120
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Oficina2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   16
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Oficina1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "UTILITARIOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8040
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Citas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Historias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "RptAccesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   RptAccesos.Caption = "USUARIO  : " & RptUsuarios.TxUsuario.Text
   If RptUsuarios.Txacceso1.Text = "" Then
   Else
      For i = 1 To 2
          Chk1(i - 1).Value = IIf(Mid(RptUsuarios.Txacceso1.Text, i, 1) = "X", 1, 0)
      Next
   End If
   If RptUsuarios.Txacceso2.Text = "" Then
   Else
      For i = 1 To 3
         Chk2(i - 1).Value = IIf(Mid(RptUsuarios.Txacceso2.Text, i, 1) = "X", 1, 0)
      Next
   End If
   If RptUsuarios.Txacceso3.Text = "" Then
   Else
      For i = 1 To 2
          Chk3(i - 1).Value = IIf(Mid(RptUsuarios.Txacceso3.Text, i, 1) = "X", 1, 0)
      Next
   End If
   If RptUsuarios.Txacceso4.Text = "" Then
   Else
      For i = 1 To 2
          Chk4(i - 1).Value = IIf(Mid(RptUsuarios.Txacceso4.Text, i, 1) = "X", 1, 0)
      Next
   End If
   If RptUsuarios.Txacceso5.Text = "" Then
   Else
      For i = 1 To 3
          Chk5(i - 1).Value = IIf(Mid(RptUsuarios.Txacceso5.Text, i, 1) = "X", 1, 0)
      Next
   End If

   
   
End Sub

Private Sub cmdOK_Click()
   If IsNull(RptUsuarios.Txacceso1.Text) Then
   Else
     For i = 1 To 2
         RptUsuarios.Txacceso1.Text = IIf(Chk1(i - 1).Value = 1, Left(RptUsuarios.Txacceso1.Text, i - 1) + "X" + Right(RptUsuarios.Txacceso1.Text, 15 - i), Left(RptUsuarios.Txacceso1.Text, i - 1) + "0" + Right(RptUsuarios.Txacceso1.Text, 15 - i))
     Next
   End If
   If RptUsuarios.Txacceso2.Text = "" Then
   Else
      For i = 1 To 1
          RptUsuarios.Txacceso2.Text = IIf(Chk2(i - 1).Value = 1, Left(RptUsuarios.Txacceso2.Text, i - 1) + "X" + Right(RptUsuarios.Txacceso2.Text, 15 - i), Left(RptUsuarios.Txacceso2.Text, i - 1) + "0" + Right(RptUsuarios.Txacceso2.Text, 15 - i))
      Next
   End If
   If RptUsuarios.Txacceso3.Text = "" Then
   Else
      For i = 1 To 2
          RptUsuarios.Txacceso3.Text = IIf(Chk3(i - 1).Value = 1, Left(RptUsuarios.Txacceso3.Text, i - 1) + "X" + Right(RptUsuarios.Txacceso3.Text, 15 - i), Left(RptUsuarios.Txacceso3.Text, i - 1) + "0" + Right(RptUsuarios.Txacceso3.Text, 15 - i))
      Next
   End If
   If RptUsuarios.Txacceso4.Text = "" Then
   Else
      For i = 1 To 5
          RptUsuarios.Txacceso4.Text = IIf(Chk4(i - 1).Value = 1, Left(RptUsuarios.Txacceso4.Text, i - 1) + "X" + Right(RptUsuarios.Txacceso4.Text, 15 - i), Left(RptUsuarios.Txacceso4.Text, i - 1) + "0" + Right(RptUsuarios.Txacceso4.Text, 15 - i))
      Next
   End If
   SwOk = True
   Unload Me
   Me.Hide
End Sub

Private Sub cmdsalir_Click()
  Me.Hide     '1er esconde
  Unload Me   '2do descarga
End Sub

