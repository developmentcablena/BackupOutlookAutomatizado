VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Backup de Outlook "
   ClientHeight    =   8355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12090
   ForeColor       =   &H8000000E&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Manual"
      Enabled         =   0   'False
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
      Left            =   5040
      TabIndex        =   19
      Top             =   1440
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H8000000F&
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
      ItemData        =   "Form1.frx":10CA
      Left            =   2640
      List            =   "Form1.frx":10CC
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Instalador "
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
      Left            =   4800
      TabIndex        =   17
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C00000&
      Caption         =   "Verificar PC"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      MaskColor       =   &H00C00000&
      MousePointer    =   1  'Arrow
      TabIndex        =   10
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      MousePointer    =   1  'Arrow
      TabIndex        =   9
      Top             =   7080
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000F&
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
      ItemData        =   "Form1.frx":10CE
      Left            =   240
      List            =   "Form1.frx":10D0
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   13680
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   2895
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "1"
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
         Left            =   2040
         TabIndex        =   8
         Top             =   360
         Width           =   1350
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "2"
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1240
      End
   End
   Begin VB.Timer Timer1 
      Left            =   15240
      Top             =   5400
   End
   Begin VB.CommandButton CommandButton2 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      TabIndex        =   4
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox TextBox2 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   4815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2040
      Width           =   11895
   End
   Begin VB.CommandButton cmdExecutar 
      BackColor       =   &H8000000D&
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   7080
      Width           =   1455
   End
   Begin VB.TextBox TextBox1 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema de Backup Outlook"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2880
      TabIndex        =   16
      Top             =   240
      Width           =   6375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Height          =   975
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   12135
   End
   Begin VB.Label Versao 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   10440
      TabIndex        =   14
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CABLENA DO BRASIL. LTDA "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   8040
      Width           =   4575
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   7920
      Width           =   12135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Computador:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private operacaoAtiva As Boolean
Private nomePC As String
Private arquivoSaida As String
Private arquivoSaidaLOG_ELE As String
Private arquivoSaidaLOG_TEL As String
Private shellObj As Object
Private backupConcluido As Boolean
Private comandoCreate As String
Private comandoRun As String
Private comandoDelete As String
Private comando As String

Private nomeMaquina As String
Private caminhoRede As String
Private fso2 As Object

Private Sub Check1_Click()
If Me.Check1.Value = 1 Then
    Me.TextBox1.Visible = True
    Me.Combo2.Visible = False
Else
    Me.TextBox1.Visible = False
    Me.Combo2.Visible = True
End If

End Sub

Private Sub cmdExecutar_Click()
If Trim(Me.Combo2.Text) = "" Then
    MsgBox "Favor colocar o nome da maquina desejavel!", vbExclamation, "Aviso"
    Me.Combo2.SetFocus
    Exit Sub
Else
    
    If Trim(Me.Combo1.Text) = "ELETRICOS" Then
        vEmpresa = "ELETRICOS"
    
        arquivoSaida = "\\196.200.80.28\TempBackupOutlook\backupELE.txt"
        arquivoSaidaLOG_ELE = "\\196.200.80.28\TempBackupOutlook\Log_ELETRICOS.txt"
        vLOG = arquivoSaidaLOG_ELE
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        If Not fso.FileExists(arquivoSaida) Then
            Set arq = fso.CreateTextFile(arquivoSaida, True)
        End If
        
        If Not fso.FileExists(arquivoSaidaLOG_ELE) Then
            Set arq = fso.CreateTextFile(arquivoSaidaLOG_ELE, True)
        End If
    
        operacaoAtiva = True
        backupConcluido = False
        nomePC = Combo2.Text
    
        Set shellObj = CreateObject("WScript.Shell")
        
        If Me.Option1.Value = True Then
    
            ' Comando para criar a tarefa
                comandoCreate = "cmd /c SCHTASKS /Create /S " & nomePC & _
                                " /U cablenabr\emperes /P emperes2025 /TN ""BackupPST"" " & _
                                "/TR ""C:\backup\backup.bat"" /SC ONCE /ST 23:59 /RL HIGHEST " & _
                                "/RU ""emperes"" /RP ""emperes2025"" /F > """ & arquivoSaidaLOG_ELE & """ 2>&1"
            
                ' Comando para executar a tarefa
                comandoRun = "cmd /c SCHTASKS /Run /S " & nomePC & _
                             " /U emperes /P emperes2025 /TN ""BackupPST"" >> """ & arquivoSaidaLOG_ELE & """ 2>&1"
                ' Comando para deletar a tarefa
                comandoDelete = "cmd /c SCHTASKS /Delete /S " & nomePC & _
                                " /U emperes /P emperes2025 /TN ""BackupPST"" /F  >> """ & arquivoSaidaLOG_ELE & """ 2>&1"
                            
                'Executa os comandos
                shellObj.Run comandoCreate, 0, True   ' True para esperar terminar
                shellObj.Run comandoRun, 0, True      ' Espera terminar
                shellObj.Run comandoDelete, 0, False  ' Não precisa esperar
                
                ' Iniciar o timer para atualizar a saída e a barra de progresso
                Timer1.Interval = 1000 ' Atualizar a cada segundo
                Timer1.Enabled = True
        Else
            comando = "cmd /c cd C:\SysInternals\2.20 && psexec -h -s \\" & nomePC & " c:\Backup\backup.bat"
        
            ' Executar o comando sem esperar finalizar
            shellObj.Run comando, 0, False
        
            ' Iniciar o timer para atualizar a saída e a barra de progresso
            Timer1.Interval = 1000 ' Atualizar a cada segundo
            Timer1.Enabled = True
        End If
    ElseIf Trim(Me.Combo1.Text) = "TELECOM" Then
        
        
        
        arquivoSaida = "\\196.200.80.28\TempBackupOutlook\backupTEL.txt"
        arquivoSaidaLOG_TEL = "\\196.200.80.28\TempBackupOutlook\Log_TELECOM.txt"
        vLOGTEL = arquivoSaidaLOG_TEL
        vEmpresa = "TELECOM"
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        If Not fso.FileExists(arquivoSaida) Then
            Set arq = fso.CreateTextFile(arquivoSaida, True)
        End If
        
        If Not fso.FileExists(arquivoSaidaLOG_TEL) Then
            Set arq = fso.CreateTextFile(arquivoSaidaLOG_TEL, True)
        End If
    
        operacaoAtiva = True
        backupConcluido = False
        nomePC = Combo2.Text
    
        Set shellObj = CreateObject("WScript.Shell")
        
        If Me.Option1.Value = True Then
    
            ' Comando para criar a tarefa
                comandoCreate = "cmd /c SCHTASKS /Create /S " & nomePC & _
                                " /U cablenabr\emperes /P emperes2025 /TN ""BackupPST"" " & _
                                "/TR ""C:\backup\backup.bat"" /SC ONCE /ST 23:59 /RL HIGHEST " & _
                                "/RU ""emperes"" /RP ""emperes2025"" /F > """ & arquivoSaidaLOG_TEL & """ 2>&1"
            
                ' Comando para executar a tarefa
                comandoRun = "cmd /c SCHTASKS /Run /S " & nomePC & _
                             " /U emperes /P emperes2025 /TN ""BackupPST"" >> """ & arquivoSaidaLOG_TEL & """ 2>&1"
                ' Comando para deletar a tarefa
                comandoDelete = "cmd /c SCHTASKS /Delete /S " & nomePC & _
                                " /U emperes /P emperes2025 /TN ""BackupPST"" /F  >> """ & arquivoSaidaLOG_TEL & """ 2>&1"
                            
                'Executa os comandos
                shellObj.Run comandoCreate, 0, True   ' True para esperar terminar
                shellObj.Run comandoRun, 0, True      ' Espera terminar
                shellObj.Run comandoDelete, 0, False  ' Não precisa esperar
                
                ' Iniciar o timer para atualizar a saída e a barra de progresso
                Timer1.Interval = 1000 ' Atualizar a cada segundo
                Timer1.Enabled = True

        Else
    
            ' Montar o comando psexec
           
            comando = "cmd /c cd C:\SysInternals\2.20 && psexec -h -s \\" & nomePC & " c:\Backup\backup.bat"
        
            ' Executar o comando sem esperar finalizar
            shellObj.Run comando, 0, False
        
            ' Iniciar o timer para atualizar a saída e a barra de progresso
            Timer1.Interval = 1000 ' Atualizar a cada segundo
            Timer1.Enabled = True
        End If
    Else
        MsgBox "Favor selecione uma empresa!", vbExclamation, "Aviso"
        Me.Combo1.SetFocus
    End If
End If
    
End Sub

Private Sub Combo1_Click()
Me.Combo2.Clear
Select Case Me.Combo1.Text
    Case "TELECOM"
        Call CarregarComputadores_TEL
    Case "ELETRICOS"
        Call CarregarComputadores_ELE
End Select

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

If Trim(Me.Combo1.Text) = "ELETRICOS" Then

    If Trim(vLOG) = "" Then
        MsgBox "Necessário executar o backup primeiro!", vbExclamation, "Aviso"
        Exit Sub
    Else
        Form2.Show vbModal
    End If
    
ElseIf Trim(Me.Combo1.Text) = "TELECOM" Then

    If Trim(vLOGTEL) = "" Then
        MsgBox "Necessário executar o backup primeiro!", vbExclamation, "Aviso"
        Exit Sub
    Else
        Form2.Show vbModal
    End If
Else
    MsgBox "Favor selecione uma empresa!", vbExclamation, "Aviso"
End If


End Sub

Private Sub Command2_Click()
nomePC = Me.Combo2.Text
If Trim(nomePC) = "" Then
    MsgBox "Favor colocar o nome do computador!", vbExclamation, "Aviso"
Else
    Screen.MousePointer = vbHourglass
    DoEvents  ' atualiza interface
    If VerificaMaquina(nomePC) Then
        MsgBox "Máquina está acessível. Pode iniciar o backup.", vbInformation, "Verificação OK"
    Else
        MsgBox "Máquina não está acessível ou está desligada!", vbCritical, "Erro de Conexão"
    End If
    Screen.MousePointer = vbDefault
End If

    

End Sub

Private Sub Command3_Click()
If Trim(Me.Combo1.Text) = "" Then
    MsgBox "Favor selecionar a Empresa!", vbExclamation, "Aviso"
    Me.Combo1.SetFocus
    Exit Sub
Else
    vEmpresa = Trim(Me.Combo1.Text)
    Form3.Show vbModal
End If

End Sub

Private Sub Form_Load()
Me.Versao.Caption = "Versão: " & App.Major & "." & App.Minor & "." & App.Revision
Me.TextBox1.Visible = False
Me.CommandButton2.Enabled = False
Me.Label3.BackColor = RGB(43, 87, 154)
Me.Option1.Value = True
Me.Label6.BackColor = RGB(43, 87, 154)
Me.Combo1.AddItem "ELETRICOS"
Me.Combo1.AddItem "TELECOM"
End Sub

Private Sub Label8_Click()
Unload Me
End Sub

Private Sub TextBox2_Change()
Me.BackColor = &HF0F0F0

End Sub

Private Sub Timer1_Timer()
    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(arquivoSaida) Then
        Dim texto As String
        texto = fso.OpenTextFile(arquivoSaida, 1).ReadAll

        TextBox2.Text = texto
        TextBox2.SelStart = Len(TextBox2.Text)
        TextBox2.SelLength = 0

        ' Verifica se o texto contém a frase final
        If InStr(1, texto, "Backup concluído com sucesso!", vbTextCompare) > 0 Then
            MsgBox "Backup finalizado com sucesso!", vbInformation

            ' Encerrar processos
            shellObj.Run "taskkill /im cmd.exe", 0, False
            shellObj.Run "taskkill /f /im PSEXESVC.exe", 0, False

            fso.DeleteFile arquivoSaida, True
            Timer1.Enabled = False
        End If
    Else
        TextBox2.Text = "Arquivo de saída não encontrado."
    End If

    Set fso = Nothing
End Sub


Private Sub CommandButton2_Click()


    ' Cancelar a operação
    operacaoAtiva = False
    Timer1.Enabled = False
    
    ' Tentar matar o processo (essa parte pode ser complicada no VB6)
    ' Uma abordagem é usar o comando taskkill no shell
    Dim comandoKill As String
    comandoKill = "taskkill /f /im cmd.exe"
    shellObj.Run comandoKill, 0, False
    
    comandoKill = "taskkill /f /im PSEXESVC.exe"
    shellObj.Run comandoKill, 0, False
    
    ' Limpar objetos
    Set shellObj = Nothing
    
    If Trim(Me.Combo1.Text) = "ELETRICOS" Then
        Kill "\\196.200.80.28\c$\Temp\backupELE.txt"
        MsgBox "Backup cancelado!!"
        Exit Sub
    Else
        Kill "\\196.200.80.28\c$\Temp\backupTEL.txt"
        MsgBox "Backup cancelado!!"
        Exit Sub
    End If
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





