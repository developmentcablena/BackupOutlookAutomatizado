Attribute VB_Name = "Module1"
Global vLOG As String
Global vLOG2 As String
Global vLOGTEL As String
Public vEmpresa As String
Public CAMINHO_ARQ As String


'Public Const CAMINHO_ARQ As String = "\\196.200.80.28\TempBackupOutlook\Base_De_Dados_Computadores_" & vEmpresa & ".txt"

Public nomeMaquina As String, nomeTarefa As String, usuario As String, senha As String
Public script As String, args As String, horario As String, dataInicio As String
Public periodicidade As String, diasSemana As String, diaMes As String
Public logFile As String

Public comandoDelete As String
Public comandoCreate As String
Public comandoRun As String
Public comandoQueryBasic As String
Public comandoQueryComplementares As String
Public CaminhoSaidaLog As String
Global CaminhoSaidaBacupPST As String
Global arq As Object

Global vUsuario As String
Global vSenha As String
Global vNomeAD As String


Global nomeMaquinaMult As String



