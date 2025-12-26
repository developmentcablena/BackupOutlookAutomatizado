VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Instalação"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7290
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo2 
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
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   360
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   2175
   End
   Begin VB.PictureBox ProgressBar1 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   7035
      TabIndex        =   4
      Top             =   5040
      Width           =   7095
   End
   Begin VB.TextBox txtVisualizadorGeral 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   960
      Width           =   7095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Falha"
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
      Left            =   5760
      TabIndex        =   2
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sucesso"
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
      Left            =   4200
      TabIndex        =   1
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Executar"
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
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Nome do computador:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fso As Object


Private Sub Combo1_Click()
If Trim(Me.Combo1.Text) = "Manual" Then
    Me.Combo2.Visible = True
    Me.Label2.Visible = True
    
    Me.Combo2.Clear
    Select Case vEmpresa
        Case "TELECOM"
            Call CarregarComputadores_TEL
        Case "ELETRICOS"
            Call CarregarComputadores_ELE
    End Select
ElseIf Trim(Me.Combo1.Text) = "Geral" Then
    Me.Combo2.Visible = False
    Me.Label2.Visible = False
Else
    MsgBox "Favor selecionar um status", vbInformation, "Informação"
    Me.Combo1.SetFocus
End If

End Sub


Private Sub CarregarComputadores_ELE()
    Dim linha As String
    Dim caminhoArquivo_ELE As String


    caminhoArquivo_ELE = "\\196.200.80.28\TempBackupOutlook\Base_De_Dados_Computadores_ELETRICOS.txt"
    Combo2.Clear ' Limpa a lista antes de carregar
    
    
    Open caminhoArquivo_ELE For Input As #1
    Do While Not EOF(1)
        Line Input #1, linha
        If Trim(linha) <> "" Then
            Combo2.AddItem linha
        End If
    Loop
    Close #1
End Sub
 
Private Sub CarregarComputadores_TEL()
    Dim linha As String
    Dim caminhoArquivo_TEL As String

    caminhoArquivo_TEL = "\\196.200.80.28\TempBackupOutlook\Base_De_Dados_Computadores_TELECOM.txt"
    Combo2.Clear ' Limpa a lista antes de carregar
    
    
    Open caminhoArquivo_TEL For Input As #1
    Do While Not EOF(1)
        Line Input #1, linha
        If Trim(linha) <> "" Then
            Combo2.AddItem linha
        End If
    Loop
    Close #1
End Sub



Private Sub Command1_Click()
If Trim(Me.Combo1.Text) = "Geral" Then
    '
ElseIf Trim(Me.Combo1.Text) = "Manual" Then
    If Trim(Me.Combo2.Text) = "" Then
        MsgBox "Favor selecione a maquina!", vbInformation, "Informação"
        Me.Combo2.SetFocus
        Exit Sub
        
    Else
    End If
Else
    MsgBox "Favor selecionar um status", vbInformation, "Informação"
    Me.Combo1.SetFocus
    Exit Sub
    
End If

Call CopiarArquivoParaComputador
End Sub

Private Sub Command2_Click()

Dim arquivo As String
Set fso = CreateObject("Scripting.FileSystemObject")

arquivo = "\\196.200.80.28\TempBackupOutlook\Log_Sucesso_" & vEmpresa & ".txt"
Me.txtVisualizadorGeral = fso.OpenTextFile(arquivo, 1).ReadAll

End Sub

Private Sub Command3_Click()
Dim arquivo As String
Set fso = CreateObject("Scripting.FileSystemObject")

arquivo = "\\196.200.80.28\TempBackupOutlook\Log_Falha_" & vEmpresa & ".txt"
Me.txtVisualizadorGeral = fso.OpenTextFile(arquivo, 1).ReadAll

End Sub


Private Sub CopiarArquivoParaComputador()
    Dim linha As String
    Dim caminhoArquivoLista As String
    Dim pastaOrigem As String
    Dim pastaOrigem2 As String
    Dim destino As String
    Dim fso As Object
    Dim totalLinhas As Long
    Dim contador As Long
    Dim vStatus As String

    ' Detecta modo
    If Trim(Me.Combo1.Text) = "Geral" Then
        vStatus = "Geral"
    ElseIf Trim(Me.Combo1.Text) = "Manual" Then
        vStatus = "Manual"
    Else
        MsgBox "Selecione um status válido!", vbExclamation
        Exit Sub
    End If

    Screen.MousePointer = 11
    Me.Enabled = False

    ' Caminhos fixos
    caminhoArquivoLista = "\\196.200.80.28\TempBackupOutlook\Base_De_Dados_Computadores_" & vEmpresa & ".txt"
    
    
    pastaOrigem = "\\196.200.80.28\TempBackupOutlook\Backup Outlook\" & vEmpresa & "\"
    
    pastaOrigem2 = "\\196.200.80.28\TempBackupOutlook\Backup Outlook\Arquivo\"

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Abre logs
    Open "\\196.200.80.28\TempBackupOutlook\Log_Sucesso_" & vEmpresa & ".txt" For Append As #10
    Open "\\196.200.80.28\TempBackupOutlook\Log_Falha_" & vEmpresa & ".txt" For Append As #11

    contador = 0

    ' =====================================================
    '   MODO GERAL ? lê TXT e conta linhas
    ' =====================================================
    If vStatus = "Geral" Then
        totalLinhas = 0
        Open caminhoArquivoLista For Input As #12
        Do While Not EOF(12)
            Line Input #12, linha
            If Trim(linha) <> "" Then
                totalLinhas = totalLinhas + 1
            End If
        Loop
        Close #12

        ' Configura barra de progresso
        ProgressBar1.Min = 0
        ProgressBar1.Max = totalLinhas
        ProgressBar1.Value = 0

        ' Agora processa os computadores
        Open caminhoArquivoLista For Input As #12
        Do While Not EOF(12)
            Line Input #12, linha
            linha = Trim(linha)
            If linha <> "" Then
                Call ProcessarComputador(linha, fso, pastaOrigem, pastaOrigem2, contador)
            End If
        Loop
        Close #12
    End If

    ' =====================================================
    '   MODO MANUAL ? usa Combo2
    ' =====================================================
    If vStatus = "Manual" Then
        linha = Trim(Me.Combo2.Text)
        If linha = "" Then
            MsgBox "Digite o nome da máquina no campo Manual!", vbExclamation
            GoTo Finalizar
        End If

        ' Configura barra para 1 item
        ProgressBar1.Min = 0
        ProgressBar1.Max = 1
        ProgressBar1.Value = 0

        Call ProcessarComputador(linha, fso, pastaOrigem, pastaOrigem2, contador)
    End If

Finalizar:
    Close #10
    Close #11

    Screen.MousePointer = 0
    Me.Enabled = True

    MsgBox "Processo concluído! Verifique os logs."
End Sub

' ========================================================================================
'   SUB PARA PROCESSAR UM ÚNICO COMPUTADOR
' ========================================================================================
Private Sub ProcessarComputador(ByVal linha As String, ByVal fso As Object, ByVal pastaOrigem As String, ByVal pastaOrigem2 As String, ByRef contador As Long)
    Dim destino As String
    destino = "\\" & linha & "\C$\Backup\" ' nome da pasta onde fica os arquivo para fazer backup
    destino2 = "\\" & linha & "\C$\Windows\" ' destino onde vai ser copiado para executar o backup

    On Error GoTo TrataErros

    ' Cria pasta se não existir
    If Not fso.FolderExists(destino) Then
        fso.CreateFolder destino
    End If
    

    ' Copia arquivos
    fso.CopyFile pastaOrigem & "*", destino, True
    fso.CopyFile pastaOrigem2 & "*", destino2, True

    ' Sucesso
    Print #10, linha & " - Copiado com sucesso - " & Now
    GoTo Continuar

TrataErros:
    Print #11, linha & " - FALHA - " & Err.Description & " - " & Now
    Err.Clear

Continuar:
    contador = contador + 1
    If contador <= ProgressBar1.Max Then
        ProgressBar1.Value = contador
    End If
    DoEvents
End Sub



Private Sub Form_Load()
Me.Label2.Visible = False
Me.Combo2.Visible = False
Me.Combo1.AddItem "Geral"
Me.Combo1.AddItem "Manual"
End Sub
