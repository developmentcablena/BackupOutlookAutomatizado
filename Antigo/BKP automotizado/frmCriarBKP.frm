VERSION 5.00
Begin VB.Form frmCriarBKP 
   Caption         =   "Form1"
   ClientHeight    =   2685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   7080
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cbb_nome_maquina 
      Height          =   315
      ItemData        =   "frmCriarBKP.frx":0000
      Left            =   240
      List            =   "frmCriarBKP.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Nome da Maquina:"
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
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmCriarBKP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub List1_Click()

End Sub
