VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8640
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   8640
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Left            =   9000
      Top             =   4560
   End
   Begin VB.TextBox TextBoxLOG 
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
      Height          =   2295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    On Error Resume Next

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If vEmpresa = "ELETRICOS" Then
        If fso.FileExists(vLOG) Then
            Dim texto As String
            texto = fso.OpenTextFile(vLOG, 1).ReadAll
    
            TextBoxLOG.Text = texto
            TextBoxLOG.SelStart = Len(TextBoxLOG.Text)
            TextBoxLOG.SelLength = 0

        Else
            Me.TextBoxLOG.Text = "Arquivo de saída não encontrado."
        End If
    
        Set fso = Nothing
    ElseIf vEmpresa = "TELECOM" Then
        If fso.FileExists(vLOGTEL) Then
            Dim texto2 As String
            texto2 = fso.OpenTextFile(vLOGTEL, 1).ReadAll
    
            TextBoxLOG.Text = texto2
            TextBoxLOG.SelStart = Len(TextBoxLOG.Text)
            TextBoxLOG.SelLength = 0

        Else
            Me.TextBoxLOG.Text = "Arquivo de saída não encontrado."
        End If
    
        Set fso = Nothing
    Else
        MsgBox "ERRO ao vizualizar o log!", vbCritical, "Suporte Sistemas"
    End If
    
End Sub

Private Sub Form_Load()
Timer1.Interval = 1000 ' Atualizar a cada segundo
Timer1.Enabled = True

End Sub
