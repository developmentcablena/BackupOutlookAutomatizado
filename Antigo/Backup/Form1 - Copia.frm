VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Backup de Outlook - v1.0.0.4"
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11205
   ForeColor       =   &H8000000E&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
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
      Left            =   9600
      TabIndex        =   3
      Top             =   240
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
      Height          =   5655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   840
      Width           =   11055
   End
   Begin VB.CommandButton CommandButton1 
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
      Left            =   7920
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   240
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2175
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
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private operacaoAtiva As Boolean
Private nomePC As String
Private arquivoSaida As String
Private shellObj As Object
Private backupConcluido As Boolean

Private Sub CommandButton1_Click()
MsgBox "Backup iniciando..."
    arquivoSaida = "C:\Temp\backup.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
' Verifica se o arquivo existe
    If Not fso.FileExists(arquivoSaida) Then
        ' Cria o arquivo se não existir
        Set arq = fso.CreateTextFile(arquivoSaida, True)
    End If
    

    ' Iniciar a operação
    operacaoAtiva = True
    backupConcluido = False
    nomePC = TextBox1.Text

    
    ' Criar objeto WScript.Shell
    Set shellObj = CreateObject("WScript.Shell")

    ' Montar o comando psexec
    Dim comando As String
    comando = "cmd /c cd C:\SysInternals\2.20 && psexec -h -s \\" & nomePC & " c:\Backup\backup.bat"

    ' Executar o comando sem esperar finalizar
    shellObj.Run comando, 0, False

    ' Iniciar o timer para atualizar a saída e a barra de progresso
    Timer1.Interval = 1000 ' Atualizar a cada segundo
    Timer1.Enabled = True

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
   
    Kill "C:\Temp\backup.txt"
    MsgBox "Backup cancelado!!"
    Exit Sub
    
End Sub







