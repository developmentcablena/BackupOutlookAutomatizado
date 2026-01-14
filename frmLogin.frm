VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   3495
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSenha 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Left            =   7200
      Top             =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Entrar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.ComboBox cboEmpresa 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   2295
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   3120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema de Backup Outlook"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmLogin.frx":10CA
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
vUsuario = Me.txtUsuario.Text
vSenha = Me.txtSenha.Text


If Trim(Me.txtUsuario.Text) = "" Then
    MsgBox "Digite o  usuario!", vbExclamation, "Sistemas"
    Exit Sub
End If

If Trim(Me.txtSenha.Text) = "" Then
    MsgBox "Digite a senha!", vbExclamation, "Sistemas"
    Exit Sub
End If

If Trim(Me.cboEmpresa.Text) = "" Then
    MsgBox "Favor selecione a Base de Dados!", vbExclamation, "Sistemas"
    Exit Sub
End If

PermissaoAD = False
Call GetUserFullName(vUsuario)
On Error GoTo ErrHandler
    Dim ldapObject As Object ' Usar Object para evitar problemas com interface
    Dim ldapConnection As Object ' Usar Object para a conexão LDAP
    Dim userDN As String
    Dim password As String
    Dim domain As String
    Dim adsPath As String
    Dim userObject As Object

    ' Pegando valores de login e senha
    domain = "LDAP://elesrv027/cn=Users,dc=cablenabr,dc=local" ' Ajuste para o seu domínio AD
    userDN = "CABLENABR\" & vUsuario ' Ajuste conforme o formato necessário
    password = txtSenha.Text
    
    ' Conectar ao AD usando as credenciais fornecidas
    Set ldapObject = GetObject("LDAP:")
    Set ldapConnection = ldapObject.OpenDSObject(domain, userDN, password, 0)
    
    
    If Not ldapConnection Is Nothing Then
        
        PermissaoAD = True
        Me.Label3.Caption = "Carregando a Base de Dados..."
        vEmpresa = Me.cboEmpresa.Text
        ProgressBar1.Min = 0
        ProgressBar1.Max = 100
        ProgressBar1.Value = 0
        Timer1.Interval = 5
        Timer1.Enabled = True
        Call Unload(Me)
        Call frmPrincipal.Show
        Exit Sub
    End If

ErrHandler:
    ' Em caso de erro, capturar o erro e mostrar mensagem ao usuário
    'MsgBox "Falha no login. Codigo de erro: " & Err.Number & ". Verifique suas credenciais.", vbExclamation, "Erro"
    If Len(Trim(Me.txtUsuario.Text)) = 0 Then
        MsgBox "Digite o usuário!", vbExclamation, "Suporte Simac"
        Me.txtUsuario.SetFocus
    Else
        MsgBox "Usuário ou senha inválidos.", vbCritical, "Erro de Login"
        Me.txtSenha.Text = ""
        Me.txtSenha.SetFocus
    End If

End Sub

Private Sub GetUserFullName(ByVal sUserName As String)
    Dim oRootDSE As Object
    Dim oConnection As Object
    Dim oCommand As Object
    Dim oRecordSet As Object
    Dim sDomain As String
    Dim sQuery As String
   
    ' Conectar ao AD
    Set oRootDSE = GetObject("LDAP://RootDSE")
    sDomain = "LDAP://elesrv027.cablenabr.local/DC=cablenabr,DC=local"
    

    Set oConnection = CreateObject("ADODB.Connection")
    oConnection.Provider = "ADsDSOObject"
    oConnection.Open "Active Directory Provider"
    
    Set oCommand = CreateObject("ADODB.Command")
    oCommand.ActiveConnection = oConnection
    
    ' Criar a consulta LDAP para buscar o usuário pelo sAMAccountName
    ' Se necessário, você pode alterar a parte do domínio ou a OU
    sQuery = "<" & sDomain & ">;(sAMAccountName=" & sUserName & ");cn;subtree"  '"<LDAP://" & sDomain & ">;(sAMAccountName=" & sUserName & ");cn;subtree"
    
    oCommand.CommandText = sQuery
    
    ' Executar a consulta
    Set oRecordSet = oCommand.Execute
    
    ' Verificar se encontrou resultados
    If Not oRecordSet.EOF Then
        vNomeAD = oRecordSet.Fields("cn").Value ' Obter o nome completo

    '    MsgBox "Nome completo do usuário: " & sFullName
    Else
    '    MsgBox "Usuário não encontrado no AD."
    End If
    
    ' Fechar conexão
    oRecordSet.Close
    oConnection.Close
End Sub




Private Sub Form_Load()

Me.cboEmpresa.AddItem "ELETRICOS"
Me.cboEmpresa.AddItem "TELECOM"




End Sub




Private Sub Timer1_Timer()
    On Error Resume Next

    If ProgressBar1.Value < ProgressBar1.Max Then
        ProgressBar1.Value = ProgressBar1.Value + 1   ' tamanho do passo
        DoEvents                                       ' mantém a UI responsiva
    Else
        Timer1.Enabled = False
        
        ' Fecha o splash
        Unload Me
        
        frmPrincipal.Show vbModal
    
    End If
End Sub

