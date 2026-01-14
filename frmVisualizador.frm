VERSION 5.00
Begin VB.Form frmVisualizador 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14460
   Icon            =   "frmVisualizador.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   14460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame z 
      Height          =   1215
      Left            =   5280
      TabIndex        =   7
      Top             =   0
      Width           =   3495
      Begin VB.OptionButton optBackupPST 
         Caption         =   "Backup PST"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   2895
      End
      Begin VB.OptionButton optBasico 
         Caption         =   "Informações Basico"
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
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton optInformacoesC 
         Caption         =   "Informações Complementares"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   3015
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1320
      Width           =   14295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Maquinas"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtxMaquina 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdAtualizar 
         Caption         =   "Atualizar"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   1
         Top             =   360
         Width           =   975
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
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Visualizador PST."
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
      Left            =   240
      TabIndex        =   5
      Top             =   7920
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
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   7800
      Width           =   14535
   End
End
Attribute VB_Name = "frmVisualizador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private shellObj As Object
Private Saida1 As String
Private pc As String

Private Sub cmdAtualizar_Click()

    Dim fso As Object
    Dim texto As String
    Dim arquivoPC

    pc = Me.txtxMaquina.Text
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Me.optBackupPST.Value = True Then
    
        Screen.MousePointer = 11

        vLOG = "\\196.200.80.28\TempBackupOutlook\Log_PST\" & vEmpresa & "\backup_" & pc & ".txt"

        If fso.FileExists(vLOG) Then
            
            Set ts = fso.OpenTextFile(vLOG, 1) ' ForReading
            If ts.AtEndOfStream Then
                texto = ""
                MsgBox "Arquivo está vazio!", vbInformation, "Sistemas"
            Else
                texto = ts.ReadAll
            End If
            ts.Close

            Me.Text1.Text = texto
            Me.Text1.SelStart = Len(Me.Text1.Text)
            Me.Text1.SelLength = 0
                
            Screen.MousePointer = 0
        Else
            Me.Text1.Text = "Arquivo de saída não encontrado."
            Screen.MousePointer = 0
        End If

    ElseIf Me.optBasico.Value = True Then
    
        Screen.MousePointer = 11
    
        Saida1 = "\\196.200.80.28\TempBackupOutlook\INFORMAOES_BASICO_ELE.txt"
        
        comandoQueryBasic = "cmd /c SCHTASKS /query /S " & pc & _
                " /U ""emperes"" /P ""emperes1nsid"" " & _
                "/TN ""BackupPST_Mensal"" > """ & Saida1 & """ 2>&1"
                                
        Set shellObj = CreateObject("WScript.Shell")
        shellObj.run comandoQueryBasic, 0, True
        
        Me.Text1.Text = fso.OpenTextFile(Saida1, 1).ReadAll
        Screen.MousePointer = 0
    ElseIf Me.optInformacoesC.Value = True Then
        
        Screen.MousePointer = 11
        
        Saida1 = "\\196.200.80.28\TempBackupOutlook\INFORMAOES_BASICO_" & vEmpresa & ".txt"
    
       comandoQueryComplementares = "cmd /c chcp 1252 > nul &  SCHTASKS /query /S " & pc & _
                " /U ""emperes"" /P ""emperes1nsid"" " & _
                "/TN ""BackupPST_Mensal"" /V /FO LIST > """ & Saida1 & """ 2>&1"
                                
        Set shellObj = CreateObject("WScript.Shell")
        shellObj.run comandoQueryComplementares, 0, True
        
        Me.Text1.Text = fso.OpenTextFile(Saida1, 1).ReadAll
        Screen.MousePointer = 0
    Else
        MsgBox "Nenhuma opção selecionada!", vbExclamation, "Sistemas"
        Screen.MousePointer = 0
    End If   ' <<< FECHA If principal

    Set fso = Nothing

End Sub


Private Sub Form_Load()
Me.Label1.BackColor = RGB(43, 87, 154)
End Sub
