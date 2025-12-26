VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.OCX"
Begin VB.Form frmProgramacao 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10980
   BeginProperty Font 
      Name            =   "Segoe UI Semibold"
      Size            =   8.25
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProgramacao.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExcluirMaquinalvw 
      Caption         =   "Excluir"
      Height          =   375
      Left            =   2520
      TabIndex        =   33
      Top             =   2040
      Width           =   975
   End
   Begin VB.CheckBox chkMaquinaAtiva 
      Caption         =   " Maquina Ativa"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   32
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtMensalDia 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5760
      TabIndex        =   7
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton cmdTestar 
      Caption         =   "Testar (Run)"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   29
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   28
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Novo"
      Height          =   375
      Left            =   2520
      TabIndex        =   27
      Top             =   1560
      Width           =   975
   End
   Begin MSMask.MaskEdBox mskDataInicio 
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   3480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483633
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskHoraInicio 
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   3480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   -2147483633
      AutoTab         =   -1  'True
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdCriar 
      Caption         =   "Criar"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox txtSaidaLog 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   4680
      Width           =   10815
   End
   Begin VB.TextBox txtData 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   24
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtHora 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   22
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox txtCaminhoScript 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   21
      Top             =   2880
      Width           =   5055
   End
   Begin VB.TextBox txtSenha 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   6600
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2160
      Width           =   4215
   End
   Begin VB.TextBox txtUsuario 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   1680
      Width           =   4215
   End
   Begin VB.TextBox txtMaquina 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   720
      Width           =   5175
   End
   Begin VB.TextBox txtNomeTarefa 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Maquinas"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3615
      Begin MSComctlLib.ListView lvwMaquinas 
         Height          =   2895
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   5106
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI Semibold"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.CommandButton cmdAtualizar 
         Caption         =   "Atualiar"
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtBuscarMaquina 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Filtro por nome:"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Prrogramação de tarefa do Outlook"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   6720
      Width           =   3255
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   735
      Left            =   0
      TabIndex        =   30
      Top             =   6600
      Width           =   11055
   End
   Begin VB.Label Label15 
      Caption         =   "Saída/Log"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Data inicio:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   23
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Maquina:"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   20
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Usuário p/ executar (RU):"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   19
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label8 
      Caption         =   "Senha (RP):"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Programa/Script:"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Hora inicio:"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   16
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Se Mensal:"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Dia:"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Nome da tarefa:"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmProgramacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' No topo do Form (declarar variáveis)
Option Explicit
Private shellObj As Object
Private fso As Object
Private Const CAMINHO_ARQ_TEL As String = "\\196.200.80.28\TempBackupOutlook\Base_De_Dados_Computadores_TELECOM.txt"
Private Const CAMINHO_ARQ_ELE As String = "\\196.200.80.28\TempBackupOutlook\Base_De_Dados_Computadores_ELETRICOS.txt"

    

Private Sub cmdAtualizar_Click()
     
Dim filtro As String
    filtro = Me.txtBuscarMaquina.Text  ' o campo "Filtro por nome"
    CarregarMaquinas filtro
End Sub



Private Function CarregarMaquinas(Optional ByVal filtro As String = "")


On Error GoTo TratarErro
     
     Dim ts As Object
     Dim linha As String
     Dim FiltroProc As String
     
     FiltroProc = UCase$(Trim$(filtro))
    
    Me.lvwMaquinas.ListItems.Clear
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(CAMINHO_ARQ) Then
        MsgBox "Arquivo não encontrado", vbExclamation, "Máquinas"
        Exit Function
    End If
    
    Set ts = fso.OpenTextFile(CAMINHO_ARQ, 1)
        
    Do While Not ts.AtEndOfStream
        linha = Trim$(ts.ReadLine)
        If Len(linha) > 0 Then
            ' Aplica filtro (case-insensitive). Troque a lógica conforme sua necessidade:
            ' 1) Contém:
            If FiltroProc = "" Or InStr(1, UCase$(linha), FiltroProc, vbTextCompare) > 0 Then
                Me.lvwMaquinas.ListItems.Add , , linha
            End If
        End If
    Loop
    
    ts.Close
    Exit Function
    
TratarErro:
    MsgBox "Erro ao carregar máquinas: " & Err.Description, vbCritical, "Máquinas"
End Function




Private Sub cmdTestar_Click()

If Trim(Me.txtNomeTarefa.Text) = "" Then
    MsgBox "Favor colocar o nome da Tarefa!", vbExclamation, "Sistema"
    Exit Sub
End If

If Trim(Me.txtMaquina.Text) = "" Then
    MsgBox "Favor colocar o nome da Maquina!", vbExclamation, "Sistema"
    Exit Sub
End If

If Trim(Me.txtUsuario.Text) = "" Then
    MsgBox "Favor colocar o seu usuário de rede!", vbExclamation, "Sistema"
    Me.txtUsuario.SetFocus
    Exit Sub
End If

If Not IsNumeric(Trim(Me.txtSenha.Text)) Then
    If Trim(Me.txtSenha.Text) = "" Then
        MsgBox "Favor colocar a senha de Usuário de Rede!", vbExclamation, "Sistema"
        Me.txtSenha.SetFocus
        Exit Sub
    End If
End If

If Trim(Me.txtCaminhoScript.Text) = "" Then
    MsgBox "Favor colocar o Caminho de Script!", vbExclamation, "Sistema"
    Me.txtCaminhoScript.SetFocus
    Exit Sub
End If

logFile = "C:\Logs\" & nomeTarefa & ".log"
nomeMaquina = Me.txtMaquina.Text
nomeTarefa = Me.txtNomeTarefa
usuario = "CABLENABR\" & Me.txtUsuario.Text
senha = Me.txtSenha.Text
script = Me.txtCaminhoScript.Text

comandoRun = "cmd /c SCHTASKS /Run /S " & nomeMaquina & _
                " /U """ & usuario & """ /P """ & senha & """ /TN """ & nomeTarefa & """  > """ & logFile & """ 2>&1"


Screen.MousePointer = 11

If VerificaMaquina(nomeMaquina) Then

    'Executa os comandos
    Set shellObj = CreateObject("WScript.Shell")
    shellObj.run comandoRun, 0, True   ' True para esperar terminar
    
    Dim CaminhoSaidaLog As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    CaminhoSaidaLog = logFile
    Me.txtSaidaLog = fso.OpenTextFile(CaminhoSaidaLog, 1).ReadAll

Else
    Kill logFile
    Open logFile For Append As #1
    Print #1, "Máquina não está acessível ou está desligada!"
    Close #1
    CaminhoSaidaLog = logFile
    
    Me.txtSaidaLog = fso.OpenTextFile(CaminhoSaidaLog, 1).ReadAll
End If

Screen.MousePointer = 0


End Sub


Private Sub cmdExcluir_Click()

If Trim(Me.txtNomeTarefa.Text) = "" Then
    MsgBox "Favor colocar o nome da Tarefa!", vbExclamation, "Sistema"
    Exit Sub
End If

If Trim(Me.txtMaquina.Text) = "" Then
    MsgBox "Favor colocar o nome da Maquina!", vbExclamation, "Sistema"
    Exit Sub
End If

If Trim(Me.txtUsuario.Text) = "" Then
    MsgBox "Favor colocar o seu usuário de rede!", vbExclamation, "Sistema"
    Me.txtUsuario.SetFocus
    Exit Sub
End If

If Not IsNumeric(Trim(Me.txtSenha.Text)) Then
    If Trim(Me.txtSenha.Text) = "" Then
        MsgBox "Favor colocar a senha de Usuário de Rede!", vbExclamation, "Sistema"
        Me.txtSenha.SetFocus
        Exit Sub
    End If
End If

If Trim(Me.txtCaminhoScript.Text) = "" Then
    MsgBox "Favor colocar o Caminho de Script!", vbExclamation, "Sistema"
    Me.txtCaminhoScript.SetFocus
    Exit Sub
End If

logFile = "C:\Logs\" & nomeTarefa & ".log"
nomeMaquina = Me.txtMaquina.Text
nomeTarefa = Me.txtNomeTarefa
usuario = "CABLENABR\" & Me.txtUsuario.Text
senha = Me.txtSenha.Text
script = Me.txtCaminhoScript.Text

comandoDelete = "cmd /c SCHTASKS /Delete /S " & nomeMaquina & _
                " /U """ & usuario & """ /P """ & senha & """ /TN """ & nomeTarefa & """ /F  > """ & logFile & """ 2>&1"

Screen.MousePointer = 11

If VerificaMaquina(nomeMaquina) Then

    'Executa os comandos
    Set shellObj = CreateObject("WScript.Shell")
    shellObj.run comandoDelete, 0, True   ' True para esperar terminar
    
    Dim CaminhoSaidaLog As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    CaminhoSaidaLog = logFile
    Me.txtSaidaLog = fso.OpenTextFile(CaminhoSaidaLog, 1).ReadAll
Else
    
    Open logFile For Append As #1
    Print #1, "Máquina não está acessível ou está desligada!"
    Close #1
    CaminhoSaidaLog = logFile
    
    Me.txtSaidaLog = fso.OpenTextFile(CaminhoSaidaLog, 1).ReadAll
End If

Screen.MousePointer = 0


End Sub

Private Sub cmdNovo_Click()
Dim fnum As Integer
fnum = FreeFile

Open CAMINHO_ARQ For Append As #fnum
    Print #fnum, Me.txtBuscarMaquina.Text
Close #fnum

CarregarMaquinas
End Sub

Private Sub cmdCriar_Click()

If Trim(Me.txtNomeTarefa.Text) = "" Then
    MsgBox "Favor colocar o nome da Tarefa!", vbExclamation, "Sistema"
    Exit Sub
End If

If Trim(Me.txtMaquina.Text) = "" Then
    MsgBox "Favor colocar o nome da Maquina!", vbExclamation, "Sistema"
    Exit Sub
End If

If Trim(Me.txtUsuario.Text) = "" Then
    MsgBox "Favor colocar o seu usuário de rede!", vbExclamation, "Sistema"
    Me.txtUsuario.SetFocus
    Exit Sub
End If

If Not IsNumeric(Trim(Me.txtSenha.Text)) Then
    If Trim(Me.txtSenha.Text) = "" Then
        MsgBox "Favor colocar a senha de Usuário de Rede!", vbExclamation, "Sistema"
        Me.txtSenha.SetFocus
        Exit Sub
    End If
End If

If Trim(Me.txtCaminhoScript.Text) = "" Then
    MsgBox "Favor colocar o Caminho de Script!", vbExclamation, "Sistema"
    Me.txtCaminhoScript.SetFocus
    Exit Sub
End If

If InStr(Me.mskHoraInicio.Text, "_") > 0 Then
    MsgBox "Favor colocar a Hora de Inicio!", vbExclamation, "Sistema"
    Me.mskHoraInicio.SetFocus
    Exit Sub
End If

If InStr(Me.mskDataInicio.Text, "_") > 0 Or Not IsDate(Me.mskDataInicio.Text) Then
    MsgBox "Data de Início está vazia ou contém uma data inválida!", vbExclamation, "Sistema"
    Me.mskDataInicio.SetFocus
    Exit Sub
End If

If Len(Trim(Me.txtMensalDia.Text)) = 0 Then
    MsgBox "Favor colocar o Dia!", vbExclamation, "Sistema"
    Me.txtMensalDia.SetFocus
    Exit Sub
End If

    
' Pega valores da tela
nomeMaquina = Me.txtMaquina.Text
nomeTarefa = Me.txtNomeTarefa
usuario = "CABLENABR\" & Me.txtUsuario.Text
senha = Me.txtSenha.Text
script = Me.txtCaminhoScript.Text
horario = Trim(Me.mskHoraInicio.Text)
dataInicio = Trim(Me.mskDataInicio.Text)
logFile = "C:\Logs\" & nomeTarefa & ".log"
diaMes = Me.txtMensalDia.Text
periodicidade = "/SC MONTHLY /D " & diaMes & ""
'OBSERVAÇÃOES se semanl /SC WEEKLY /D MON,TUE,WED /ST 21:00 ou se diario /SC DAIL


comandoCreate = "cmd /c SCHTASKS /Create /S " & nomeMaquina & _
                " /U """ & usuario & """ /P """ & senha & """ " & _
                "/TN """ & nomeTarefa & """ " & _
                "/TR """ & script & """ " & _
                periodicidade & " /ST " & horario & " /SD " & dataInicio & _
                " /RL HIGHEST /RU """ & usuario & """ /RP """ & senha & """ /F > """ & logFile & """ 2>&1"
                

Screen.MousePointer = 11

If VerificaMaquina(nomeMaquina) Then   'Verifica se a maquina está ligada.
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim CaminhoSaidaBacupPST As String
    Dim arq As Object
    
    Call ProcessarComputador(nomeMaquina)   'Faz a instalaçao dos arquivos necessario para o backup
    
    CaminhoSaidaBacupPST = "\\196.200.80.28\TempBackupOutlook\Log_PST\" & vEmpresa & "\backup_" & nomeMaquina & ".txt"
    'Se não existe a pasta onde vai ficar os dados do bkp feito, cria.
    If Not fso.FileExists(CaminhoSaidaBacupPST) Then
        Set arq = fso.createTextFile(CaminhoSaidaBacupPST, True)
    End If
    
    Set shellObj = CreateObject("WScript.Shell")
        shellObj.run comandoCreate, 0, True
    
    CaminhoSaidaLog = logFile
    Me.txtSaidaLog = fso.OpenTextFile(CaminhoSaidaLog, 1).ReadAll
Else
    Kill logFile
    Open logFile For Append As #1
    Print #1, "Máquina não está acessível ou está desligada!"
    Close #1
    CaminhoSaidaLog = logFile
    
    Me.txtSaidaLog = fso.OpenTextFile(CaminhoSaidaLog, 1).ReadAll
End If

Screen.MousePointer = 0
End Sub

Private Sub ProcessarComputador(ByVal vMaquina As String)
    Dim destino As String
    Dim destino2 As String
    Dim pastaOrigem As String
    Dim pastaOrigem2 As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    destino = "\\" & vMaquina & "\C$\Backup\" ' nome da pasta onde fica os arquivo para fazer backup
    destino2 = "\\" & vMaquina & "\C$\Windows\" ' destino onde vai ser copiado para executar o backup
    
    pastaOrigem = "\\196.200.80.28\TempBackupOutlook\Backup Outlook (arquivos)\" & vEmpresa & "\"
    
    pastaOrigem2 = "\\196.200.80.28\TempBackupOutlook\Backup Outlook (arquivos)\Arquivo\"

    On Error GoTo TrataErros

    ' Cria pasta se não existir
    If Not fso.FolderExists(destino) Then
        fso.CreateFolder destino
    End If
    
    ' Copia arquivos
    fso.CopyFile pastaOrigem & "*", destino, True
    fso.CopyFile pastaOrigem2 & "*", destino2, True

TrataErros:
' MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical
    Err.Clear
End Sub











Public Function VerificaMaquina(nomeMaquina As String) As Boolean
    Dim WshShell As Object
    Dim execObj As Object
    Dim saida As String
    Dim resposta As Boolean

    Set WshShell = CreateObject("WScript.Shell")

    ' Executa o ping oculto (sem abrir janela)
    Set execObj = WshShell.Exec("ping -n 1 " & nomeMaquina)

    ' Aguarda o término do comando
    Do While execObj.Status = 0
        DoEvents
    Loop

    ' Lê toda a saída do comando
    saida = LCase(execObj.StdOut.ReadAll)

    ' Faz verificação inteligente
    If (InStr(saida, "resposta de") > 0 Or InStr(saida, "reply from") > 0) _
        And InStr(saida, "inacess") = 0 _
        And InStr(saida, "tempo esgotado") = 0 _
        And InStr(saida, "unreachable") = 0 Then
        resposta = True
    Else
        resposta = False
    End If

    VerificaMaquina = resposta

    Set execObj = Nothing
    Set WshShell = Nothing
End Function







Private Sub Form_Load()
If vEmpresa = "ELETRICOS" Then
    CAMINHO_ARQ = CAMINHO_ARQ_ELE
ElseIf vEmpresa = "TELECOM" Then
    CAMINHO_ARQ = CAMINHO_ARQ_TEL
Else

End If

Me.cmdTestar.Visible = True
Call ListarMaquinasLVW
Me.txtNomeTarefa.Text = "BackupPST_Mensal"
nomeTarefa = Me.txtNomeTarefa.Text
Me.txtNomeTarefa.Enabled = False
Me.txtCaminhoScript.Text = "C:\backup\backup.bat"
Me.txtCaminhoScript.Enabled = False
Me.txtMaquina.Enabled = False
CarregarMaquinas ""
Me.Label1.BackColor = RGB(43, 87, 154)
Me.txtMensalDia.MaxLength = 2
Me.mskDataInicio = Date
Me.chkMaquinaAtiva.Value = 0
Me.chkMaquinaAtiva.ForeColor = RGB(250, 5, 5)
Me.txtSaidaLog.BackColor = RGB(224, 223, 237)
Me.txtNomeTarefa.BackColor = RGB(224, 223, 237)
Me.txtMaquina.BackColor = RGB(224, 223, 237)
Me.txtUsuario.BackColor = RGB(224, 223, 237)
Me.txtSenha.BackColor = RGB(224, 223, 237)
Me.txtCaminhoScript.BackColor = RGB(224, 223, 237)
Me.mskDataInicio.BackColor = RGB(224, 223, 237)
Me.mskHoraInicio.BackColor = RGB(224, 223, 237)
Me.txtMensalDia.BackColor = RGB(224, 223, 237)
Me.txtBuscarMaquina.BackColor = RGB(224, 223, 237)
Me.lvwMaquinas.BackColor = RGB(224, 223, 237)

End Sub

Private Sub ListarMaquinasLVW()
    Dim fso As Object
    Dim ts As Object
    Dim linha As String
    Dim caminho As String

    ' Configura o ListView para modo relatório (colunas)
    With Me.lvwMaquinas
        .View = lvwReport                 ' 3 - modo relatório
        .LabelEdit = lvwManual
        .ListItems.Clear
        .FullRowSelect = True
        .GridLines = True
        .HideColumnHeaders = False
        .ColumnHeaders.Clear
        ' ColumnHeaders.Add(Index, Key, Text, Width, Alignment)
        .ColumnHeaders.Add 1, "colMaquinas", "MÁQUINAS", 3000, lvwColumnLeft
    End With

    ' Instancia o FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Caminho do arquivo (UNC)
    'caminho = "\\196.200.80.28\TempBackupOutlook\Base_De_Dados_Computadores_" & vEmpresa & ".txt"
    'caminho = "\\196.200.80.28\TempBackupOutlook\Base_De_Dados_Computadores_TELECOM.txt"

    On Error GoTo TrataErro

    ' Abre o arquivo como texto (ForReading = 1)
    Set ts = fso.OpenTextFile(CAMINHO_ARQ, 1)

    ' Limpa itens antes de carregar
    Me.lvwMaquinas.ListItems.Clear

    ' Lê linha a linha e adiciona ao ListView
    Do While Not ts.AtEndOfStream
        linha = Trim$(ts.ReadLine)
        If Len(linha) > 0 Then
            ' Cada linha vira um item (primeira coluna)
            Me.lvwMaquinas.ListItems.Add , , linha
        End If
    Loop
    
    

    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    Exit Sub

TrataErro:
    MsgBox "Erro ao carregar máquinas: " & Err.Description, vbExclamation, "Leitura de arquivo"
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub



Private Sub lvwMaquinas_Click()
Dim ItemSelecionado As MSComctlLib.ListItem

Set ItemSelecionado = Me.lvwMaquinas.SelectedItem

If Not ItemSelecionado Is Nothing Then
    Me.txtMaquina = ItemSelecionado.Text
End If

If TaskExists(ItemSelecionado.Text, "\BackupPST_Mensal") Then
 Dim fso As Object
 Set fso = CreateObject("Scripting.FileSystemObject")
    Me.chkMaquinaAtiva.Value = 1
    Me.chkMaquinaAtiva.ForeColor = RGB(21, 5, 250)
Else
    Me.chkMaquinaAtiva.Value = 0
    Me.chkMaquinaAtiva.ForeColor = RGB(250, 5, 5)
End If

End Sub




Private Sub txtBuscarMaquina_Change()
Me.txtBuscarMaquina.Text = UCase$(Me.txtBuscarMaquina.Text)
Me.txtBuscarMaquina.SelStart = Len(Me.txtBuscarMaquina.Text)
CarregarMaquinas Me.txtBuscarMaquina.Text

End Sub


Public Function TaskExists(ByVal nomeMaquina As String, ByVal nomeTarefa As String) As Boolean
    Dim shell As Object
    Dim cmd As String
    Dim exitCode As Long
    
    cmd = "schtasks /Query /S " & nomeMaquina & " /TN " & Chr(34) & nomeTarefa & Chr(34)
    
    Set shell = CreateObject("WScript.Shell")
    ' Run retorna o código de saída (ERRORLEVEL)
    exitCode = shell.run(cmd, 0, True)
    
    ' Se exitCode = 0, tarefa existe; caso contrário, não existe
    TaskExists = (exitCode = 0)
End Function

