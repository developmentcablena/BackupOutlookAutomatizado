VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   2985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5100
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   2985
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   3
      Top             =   1680
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
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      _Version        =   327680
      Appearance      =   1
      MouseIcon       =   "frmLogin.frx":10CA
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
      Left            =   0
      TabIndex        =   5
      Top             =   2400
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
      TabIndex        =   4
      Top             =   1080
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
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmLogin.frx":10E6
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
If Trim(Me.cboEmpresa.Text) = "" Then
    MsgBox "Favor selecione a Base de Dados!", vbExclamation, "Sistemas"
    Exit Sub
End If

Me.Label3.Caption = "Carregando a Base de Dados..."

vEmpresa = Me.cboEmpresa.Text
' Configura a barra
    ProgressBar1.Min = 0
    ProgressBar1.Max = 100
    ProgressBar1.Value = 0

    ' Inicia o timer para animar a barra
    Timer1.Interval = 5  ' 50 ms por passo -> ~5s no total (100 * 50ms)
    Timer1.Enabled = True
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

