VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form RptUsuario 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   Caption         =   "Mantenimiento de Usuarios"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "RptUsuarios.frx":0000
   ScaleHeight     =   6135
   ScaleMode       =   0  'User
   ScaleWidth      =   11384.93
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4560
      TabIndex        =   17
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton Cmdsalir 
      Height          =   495
      Left            =   9960
      Picture         =   "RptUsuarios.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Cerrar"
      Top             =   5400
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3135
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   5530
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "USUARIOS"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   9360
      TabIndex        =   9
      Top             =   3360
      Width           =   1575
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eli&minar"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Cmdnuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   480
      TabIndex        =   0
      Top             =   3360
      Width           =   8415
      Begin VB.TextBox Txderespr 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   15
         Top             =   1800
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox ChkDerEspR 
         Alignment       =   1  'Right Justify
         Caption         =   "Der. Especiales"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox Txcambia 
         Enabled         =   0   'False
         Height          =   195
         Left            =   6840
         MaxLength       =   1
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox CboCambia 
         Height          =   315
         Left            =   7440
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdAccesoS 
         Caption         =   "ACCESOS"
         Height          =   375
         Left            =   7320
         TabIndex        =   24
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Txacceso1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5760
         TabIndex        =   18
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Txacceso2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5760
         TabIndex        =   19
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Txacceso3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5760
         TabIndex        =   20
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Txacceso4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5760
         TabIndex        =   21
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Txacceso5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5760
         TabIndex        =   22
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox TxOKPwd 
         BackColor       =   &H80000000&
         ForeColor       =   &H80000000&
         Height          =   285
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TxUsuario 
         Height          =   285
         Left            =   960
         MaxLength       =   10
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Txnombre 
         Height          =   285
         Left            =   960
         MaxLength       =   40
         TabIndex        =   10
         Top             =   840
         Width           =   4335
      End
      Begin VB.TextBox Txclave 
         BackColor       =   &H80000000&
         ForeColor       =   &H80000000&
         Height          =   285
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Cambia"
         Height          =   255
         Left            =   6720
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Usuario"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Clave"
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "RptUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DD As New ADODB.Recordset
Dim Sw, SwOk As Boolean
Dim marcador As Variant
Dim Rs As New ADODB.Recordset

Private Sub Form_Load()
  If wacceso5 = "XXXXXXXXXXXXXXX" Or wrights = "MAS" Then
     Txclave.ForeColor = &H80000002
     TxOKPwd.ForeColor = &H80000002
  End If
    Set DD = New ADODB.Recordset
    With DD
      .Fields.Append "usuario", adVarChar, 10
      .Fields.Append "clave", adVarChar, 10
      .Fields.Append "cambia", adVarChar, 1
      .Fields.Append "nombre", adVarChar, 40
      .Fields.Append "DerEsp", adVarChar, 1
      .Open
    End With
    Set DataGrid1.DataSource = DD
    For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
        Select Case i
         Case 0
          .Caption = "Codigo"
          .Width = Me.TextWidth(String(10, "S"))
         Case 1
          .Caption = "CLave"
          .Width = Me.TextWidth(String(10, "S"))
         Case 2
          '.Caption = "Cambia"
          .Width = Me.TextWidth(String(6, "S"))
         Case 3
          .Caption = "Descripcion"
          .Width = Me.TextWidth(String(35, "S"))
         Case 4
          .Caption = "DerEsp"
          .Width = Me.TextWidth(String(6, "S"))
         End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
    Next
    Set Rs = New ADODB.Recordset
    sql = "SELECT * FROM RptUsuarios "
    Rs.Open sql, cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not Rs.EOF And Not Rs.BOF Then
    Rs.MoveFirst
    Do While Not Rs.EOF
       DD.AddNew
       DD!usuario = Rs!usuario
       DD!clave = Rs!clave
       DD!cambia = IIf(IsNull(Rs!cambia), "", Rs!cambia)
       DD!Nombre = Rs!Nombre
       DD!derespr = IIf(IsNull(Rs!derespr), "N", Rs!derespr)
       DD.Update
       Rs.MoveNext
    Loop
    End If
    Rs.Close
    'LLena el Combo Modulo
    Call Limpiacampos
End Sub

Private Sub CmdAccesos_Click()
   If Trim(TxUsuario.Text) = "" Then Exit Sub
   SwOk = False
   Load RptAccesos
   RptAccesos.Left = 1500: RptAccesos.Top = 1300
   RptAccesos.Show
End Sub
Private Sub DataGrid1_Click()
   If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
   TxUsuario.Text = DataGrid1.Columns(0)
   Txclave.Text = DataGrid1.Columns(1)
   Txcambia.Text = DataGrid1.Columns(2)
   Txnombre.Text = DataGrid1.Columns(3)
   Txderespr.Text = DataGrid1.Columns(4)
   Check1.Enabled = False
   Select Case Txcambia.Text
    Case "S"
      CboCambia.ListIndex = 0
    Case "N"
      CboCambia.ListIndex = 1
   End Select
  
   If Txderesp.Text = "N" Then
      ChkDerEspR.Value = 0
   Else
      ChkDerEspR.Value = 1
   End If
   Set Rs = New ADODB.Recordset
   sql = "Select * From RptUsuarios where  usuario='" & RptUsuarios.TxUsuario.Text & "'"
   Rs.Open sql, cnx, adOpenStatic, adLockReadOnly, adCmdText
   If Rs.EOF Then
      Txacceso1.Text = "000000000000000"
      Txacceso2.Text = "000000000000000"
      Txacceso3.Text = "000000000000000"
      Txacceso4.Text = "000000000000000"
      Txacceso5.Text = "000000000000000"
   Else
      Txacceso1.Text = IIf(IsNull(Rs!acceso1), "" + "000000000000000", Left(Rs!acceso1 + "000000000000000", 15))
      Txacceso2.Text = IIf(IsNull(Rs!acceso2), "" + "000000000000000", Left(Rs!acceso2 + "000000000000000", 15))
      Txacceso3.Text = IIf(IsNull(Rs!acceso3), "" + "000000000000000", Left(Rs!acceso3 + "000000000000000", 15))
      Txacceso4.Text = IIf(IsNull(Rs!acceso4), "" + "000000000000000", Left(Rs!acceso4 + "000000000000000", 15))
      Txacceso5.Text = IIf(IsNull(Rs!acceso5), "" + "000000000000000", Left(Rs!acceso5 + "000000000000000", 15))
   End If
   encclave = ""
   WVdesencripta (Rs!clave)
   Rs.Close
   Txclave.Text = encclave
   TxOKPwd.Text = encclave
   desabilitatextos
   Cmdnuevo.Caption = "&Nuevo"
   CmdEditar.Caption = "&Editar"
   CmdEditar.Enabled = True
   CmdEliminar.Enabled = True
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
  With DD      ' Adodc1.Recordset
    If (.Sort = .Fields(ColIndex).[Name] & " Asc") Then
       .Sort = .Fields(ColIndex).[Name] & " Desc"
    Else
       .Sort = .Fields(ColIndex).[Name] & " Asc"
    End If
  End With
  Call Limpiacampos
End Sub

Private Sub Cmdnuevo_Click()
  If wacceso Or wDerEsp = "MAS" Then
  Else
     MsgBox "Usuario No Autorizado", vbExclamation
     Exit Sub
   End If
  If Cmdnuevo.Caption = "&Nuevo" Then
     Limpiacampos
     habilitatextos
     TxUsuario.SetFocus
     CmdEliminar.Enabled = False
     Sw = True
     Cmdnuevo.Caption = "&Grabar"
     CmdEditar.Caption = "&Cancelar"
     CmdEditar.Enabled = True
  Else                  'Grabar
     If Txclave.Text = UCase(TxOKPwd) Then
        encclave = ""
        WVencripta (Txclave.Text)
        Text1.Text = encclave
     Else
        MsgBox "Claves No Coinciden..  ", vbExclamation
        TxOKPwd = ""
        Exit Sub
     End If
     If TxModulo.Text = "" Then
        MsgBox "Defina modulo...  ", vbExclamation
        Exit Sub
     End If
     If Txderechos.Text = "" Then
        MsgBox "Defina derechos...  ", vbExclamation
        Exit Sub
     End If
     On Error GoTo err_trans
     cnx.BeginTrans
     Dim sql As String
     If Check1.Value = 1 Then
        Txestacion.Text = ""
     End If
     If Sw Then
        sql = "INSERT INTO RptUsuarios(usuario,clave,cambia,cmp,nombre,modulo,derechos,DerEspR,acceso1,acceso2,acceso3,acceso4,acceso5,acceso6,acceso7) " & _
        "VALUES('" + TxUsuario.Text + "','" + Text1.Text + "','" + Txcambia.Text + "','" + TxCmp.Text + "','" + Txnombre.Text + "','" + TxModulo.Text + "','" + Txderechos.Text + "','" + Txderespr.Text + "','" & _
        Txacceso1.Text + "','" + Txacceso2.Text + "','" + Txacceso3.Text + "','" + Txacceso4.Text + "','" + Txacceso5.Text + "','" + Txacceso6.Text + "','" + Txacceso7.Text + "')"
     Else
        sql = "UPDATE RptUsuarios SET clave='" + Text1.Text + "',cambia='" + Txcambia.Text + "',cmp='" + TxCmp.Text + "', nombre='" + Txnombre.Text + "',modulo='" + TxModulo.Text + "',derechos='" + Txderechos.Text & _
        "',derespr='" + Txderespr.Text + "',estacion='" + Txestacion.Text + "',acceso1='" + Txacceso1.Text + "',acceso2='" + Txacceso2.Text + "',acceso3='" + Txacceso3.Text + "',acceso4='" + Txacceso4.Text + "',acceso5='" & _
        Txacceso5.Text + "',acceso6='" + Txacceso6.Text + "',acceso7='" + Txacceso7.Text + "' where usuario='" + TxUsuario.Text + "'"
    End If
    cnx.Execute sql
    If Sw Then
       DD.AddNew
       DD!usuario = TxUsuario.Text
    End If
    DD!clave = Text1.Text
    DD!cambia = Txcambia.Text
    DD!Nombre = Txnombre.Text
    DD!modulo = TxModulo.Text
    DD!derechos = Txderechos.Text
    DD!derespr = Txderespr.Text
    DD!CMP = TxCmp.Text
    DD!estacion = Txestacion.Text
    DD.Update
    cnx.CommitTrans
    If TxCmp = "S" Then
       '-Actualiza Tabla Medicos
       sql = "SELECT * FROM medico WHERE cmp='" & Trim(TxUsuario.Text) & "'"
       Set Rs = New ADODB.Recordset
       Rs.Open sql, Cnx0, adOpenStatic, adLockOptimistic, adCmdText
       If Rs.EOF Then
          Rs.AddNew
          Rs!CMP = Trim(TxCmp.Text)
          Rs!descripcion = TxDescripcion.Text
       End If
       'tipdoc,numero,telefono,codespe
       'Txtipdoc.Text, TxNumero.Text, TxTelefono.Text, TxCodEspe.Text
       Rs!clave = Txclave.Text
       Rs!derespr = Txderespr.Text
       Rs.Update
       Rs.Close
    End If
    
    desabilitatextos
    Limpiacampos
    CmdEliminar.Enabled = True
    Cmdnuevo.Caption = "&Nuevo"
    CmdEditar.Caption = "&Editar"
    DataGrid1.Refresh
    CmdEditar.Enabled = False
    CmdEliminar.Enabled = False
    DataGrid1.SetFocus
    Exit Sub
err_trans:
    MsgBox Err.Description
    cnx.RollbackTrans
  End If
End Sub
Private Sub Cmdeditar_Click()
  If wacceso Or wrights = "MAS" Then
  Else
     MsgBox "Usuario No Autorizado", vbExclamation
     Exit Sub
  End If
  If CmdEditar.Caption = "&Editar" Then
     habilitatextos
     TxUsuario.Locked = True
     Txclave.SetFocus
     CmdEliminar.Enabled = False
     Sw = False
     Cmdnuevo.Caption = "&Grabar"
     CmdEditar.Caption = "&Cancelar"
  Else                   'Cancelar
     desabilitatextos
     CmdEliminar.Enabled = True
     Cmdnuevo.Caption = "&Nuevo"
     CmdEditar.Caption = "&Editar"
     CmdEditar.Enabled = False
     CmdEliminar.Enabled = False
     Limpiacampos
     DataGrid1.SetFocus
  End If
End Sub
Private Sub CmdEliminar_Click()
  If wacceso Or wrights = "MAS" Then
  Else
     MsgBox "Usuario No Autorizado", vbExclamation
     Exit Sub
  End If
  If MsgBox("Elimina registro ", vbOKCancel + vbQuestion, "Confirme") <> vbCancel Then
     'Borra en tabla
     cnx.Execute "DELETE FROM Rptusuarios WHERE Usuario='" & TxUsuario.Text & "'"
     'Borra en Temporal
     DD.Delete
     'DataGrid1.Refresh
     Limpiacampos
  End If
End Sub
Private Sub cmdsalir_Click()
   Set Rs = Nothing
   Unload Me
   Me.Hide
End Sub
Sub habilitatextos()
    TxUsuario.Locked = False
    Txclave.Locked = False
    Txnombre.Locked = False
    CboCambia.Enabled = True
    CboModulo.Enabled = True
    Cboderechos.Enabled = True
    ChkDerEspR.Enabled = True
    If Trim(Txestacion) <> "" Then
       Check1.Enabled = True
    End If
End Sub
Sub desabilitatextos()
    TxUsuario.Locked = True
    Txclave.Locked = True
    Txnombre.Locked = True
    CboCambia.Enabled = False
    CboModulo.Enabled = False
    Cboderechos.Enabled = False
    ChkDerEspR.Enabled = False
    Check1.Enabled = False
End Sub
Sub Limpiacampos()
    TxUsuario.Text = ""
    Txclave.Text = ""
    TxOKPwd.Text = ""
    Text1.Text = ""
    Txnombre.Text = ""
    Txcambia.Text = "":    CboCambia.Text = "": TxCmp.Text = ""
    CboModulo.Text = "":    TxModulo.Text = ""
    Cboderechos.Text = "":    Txderechos.Text = ""
    ChkDerEspR.Value = 0:    Txderespr.Text = ""
    Txestacion.Text = "": Check1.Value = 0: Check1.Enabled = False
    Txacceso1.Text = "000000000000000"
    Txacceso2.Text = "000000000000000"
    Txacceso3.Text = "000000000000000"
    Txacceso4.Text = "000000000000000"
    Txacceso5.Text = "000000000000000"
    Txacceso6.Text = "000000000000000"
    Txacceso7.Text = "000000000000000"
End Sub
Private Sub Txclave_LostFocus()
    Txclave.Text = UCase(Txclave.Text)
End Sub

Private Sub Txusuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TxUsuario.Text = UCase(TxUsuario.Text)
      Txclave.SetFocus
      Exit Sub
   End If
End Sub
Private Sub Txclave_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Txclave.Text = UCase(Txclave.Text)
      encclave = ""
      WVencripta (Txclave.Text)
      Text1.Text = encclave
      TxOKPwd.SetFocus
      Exit Sub
   End If
End Sub
Private Sub TxOkPwd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     If Txclave.Text = UCase(TxOKPwd) Then
        Txnombre.SetFocus
        Exit Sub
     Else
        MsgBox "Claves No Coinciden..  ", vbExclamation
        TxOKPwd = ""
     End If
   End If
End Sub
Private Sub Txnombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Txnombre.Text = UCase(Txnombre.Text)
      CboModulo.SetFocus
      Exit Sub
   End If
End Sub
Private Sub Cbocambia_Click()
   Txcambia.Text = Left(CboCambia.Text, 1)
End Sub
Private Sub Cbomodulo_Click()
   TxModulo.Text = Left(CboModulo.Text, 3)
End Sub
Private Sub Cboderechos_Click()
   Txderechos.Text = Left(Cboderechos.Text, 3)
End Sub
Private Sub ChkDerEspR_Click()
   If ChkDerEspR.Value = 1 Then
      Txderespr.Text = "S"
   Else
      Txderespr.Text = "N"
   End If
End Sub



