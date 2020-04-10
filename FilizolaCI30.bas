Attribute VB_Name = "FilizolaCI30"
'Declara��o das fun��es (tem que ser no m�dulo)

Declare Function Filizola_ConfiguraBalanca Lib "PcScale.dll" Alias "ConfiguraBalanca" (ByVal balanca As Integer, ByVal Handle As Long) As Boolean
Declare Function Filizola_InicializaLeitura Lib "PcScale.dll" Alias "InicializaLeitura" (ByVal balanca As Integer) As Boolean
Declare Function Filizola_ObtemInformacao Lib "PcScale.dll" Alias "ObtemInformacao" (ByVal balanca As Integer, ByVal campo As Integer) As Double
Declare Function Filizola_FinalizaLeitura Lib "PcScale.dll" Alias "FinalizaLeitura" (ByVal balanca As Integer) As Boolean
Declare Function Filizola_EnviaPrecoCS Lib "PcScale.dll" Alias "EnviaPrecoCS" (ByVal balanca As Integer, ByVal preco As Double) As Boolean
Declare Sub Filizola_ExibeMsgErro Lib "PcScale.dll" Alias "ExibeMsgErro" (ByVal Handle As Long)

'As fun��es abaixo n�o s�o necess�rias. Elas est�o sendo usadas
'somente para exiber a configura��o da balan�a
Declare Sub Filizola_ObtemNomeBalanca Lib "PcScale.dll" Alias "ObtemNomeBalanca" (ByVal Modelo As Integer, ByVal Nome As String)
Declare Function Filizola_ObtemParametrosBalanca Lib "PcScale" Alias "ObtemParametrosBalanca" (ByVal balanca As Integer, ByRef Modelo As Integer, ByRef Porta As Integer, ByRef BaudRate As Long) As Boolean


