Attribute VB_Name = "EpsonNaoFiscal"
Option Explicit

Public Declare Function EpsonNaoFiscal_ConfiguraTaxaSerial Lib "InterfaceEpsonNF.dll" Alias "ConfiguraTaxaSerial" (ByVal dwTaxa As Long) As Long
Public Declare Function EpsonNaoFiscal_IniciaPorta Lib "InterfaceEpsonNF.dll" Alias "IniciaPorta" (ByVal pszPorta As String) As Long
Public Declare Function EpsonNaoFiscal_FechaPorta Lib "InterfaceEpsonNF.dll" Alias "FechaPorta" () As Long
Public Declare Function EpsonNaoFiscal_ImprimeTexto Lib "InterfaceEpsonNF.dll" Alias "ImprimeTexto" (ByVal pszTexto As String) As Long
Public Declare Function EpsonNaoFiscal_ImprimeTextoTag Lib "InterfaceEpsonNF.dll" Alias "ImprimeTextoTag" (ByVal pszTexto As String) As Long
Public Declare Function EpsonNaoFiscal_FormataTX Lib "InterfaceEpsonNF.dll" Alias "FormataTX" (ByVal pszTexto As String, ByVal dwTipoLetra As Long, ByVal dwItalico As Long, ByVal dwSublinhado As Long, ByVal dwExpandido As Long, ByVal dwEnfatizado As Long) As Long
Public Declare Function EpsonNaoFiscal_AcionaGuilhotina Lib "InterfaceEpsonNF.dll" Alias "AcionaGuilhotina" (ByVal dwTipoCorte As Long) As Long
Public Declare Function EpsonNaoFiscal_ComandoTX Lib "InterfaceEpsonNF.dll" Alias "ComandoTX" (ByVal pszComando As String, ByVal dwTamanho As Long) As Long
Public Declare Function EpsonNaoFiscal_Le_Status Lib "InterfaceEpsonNF.dll" Alias "Le_Status" () As Long
Public Declare Function EpsonNaoFiscal_Le_Status_Gaveta Lib "InterfaceEpsonNF.dll" Alias "Le_Status_Gaveta" () As Long
Public Declare Function EpsonNaoFiscal_ConfiguraCodigoBarras Lib "InterfaceEpsonNF.dll" Alias "ConfiguraCodigoBarras" (ByVal dwAltura As Long, ByVal dwLargura As Long, ByVal dwHRI As Long, ByVal dwFonte As Long, ByVal dwMargem As Long) As Long
Public Declare Function EpsonNaoFiscal_ImprimeCodigoBarrasCODABAR Lib "InterfaceEpsonNF.dll" Alias "ImprimeCodigoBarrasCODABAR" (ByVal pszCodigo As String) As Long
Public Declare Function EpsonNaoFiscal_ImprimeCodigoBarrasCODE128 Lib "InterfaceEpsonNF.dll" Alias "ImprimeCodigoBarrasCODE128" (ByVal pszCodigo As String) As Long
Public Declare Function EpsonNaoFiscal_ImprimeCodigoBarrasCODE39 Lib "InterfaceEpsonNF.dll" Alias "ImprimeCodigoBarrasCODE39" (ByVal pszCodigo As String) As Long
Public Declare Function EpsonNaoFiscal_ImprimeCodigoBarrasCODE93 Lib "InterfaceEpsonNF.dll" Alias "ImprimeCodigoBarrasCODE93" (ByVal pszCodigo As String) As Long
Public Declare Function EpsonNaoFiscal_ImprimeCodigoBarrasEAN13 Lib "InterfaceEpsonNF.dll" Alias "ImprimeCodigoBarrasEAN13" (ByVal pszCodigo As String) As Long
Public Declare Function EpsonNaoFiscal_ImprimeCodigoBarrasEAN8 Lib "InterfaceEpsonNF.dll" Alias "ImprimeCodigoBarrasEAN8" (ByVal pszCodigo As String) As Long
Public Declare Function EpsonNaoFiscal_ImprimeCodigoBarrasITF Lib "InterfaceEpsonNF.dll" Alias "ImprimeCodigoBarrasITF" (ByVal pszCodigo As String) As Long
Public Declare Function EpsonNaoFiscal_ImprimeCodigoBarrasUPCA Lib "InterfaceEpsonNF.dll" Alias "ImprimeCodigoBarrasUPCA" (ByVal pszCodigo As String) As Long
Public Declare Function EpsonNaoFiscal_ImprimeCodigoBarrasUPCE Lib "InterfaceEpsonNF.dll" Alias "ImprimeCodigoBarrasUPCE" (ByVal pszCodigo As String) As Long
Public Declare Function EpsonNaoFiscal_ImprimeCodigoBarrasPDF417 Lib "InterfaceEpsonNF.dll" Alias "ImprimeCodigoBarrasPDF417" (ByVal dwCorrecao As Long, ByVal dwAltura As Long, ByVal dwLargura As Long, ByVal dwColunas As Long, ByVal pszCodigo As String) As Long
Public Declare Function EpsonNaoFiscal_ImprimeCodigoQRCODE Lib "InterfaceEpsonNF.dll" Alias "ImprimeCodigoQRCODE" (ByVal dwRestauracao As Long, ByVal dwModulo As Long, ByVal dwTipo As Long, ByVal dwVersao As Long, ByVal dwModo As Long, ByVal pszCodigo As String) As Long
Public Declare Function EpsonNaoFiscal_GerarQRCodeArquivo Lib "InterfaceEpsonNF.dll" Alias "GerarQRCodeArquivo" (ByVal pszFileName As String, ByVal pszDados As String) As Long
Public Declare Function EpsonNaoFiscal_ImprimeBmpEspecial Lib "InterfaceEpsonNF.dll" Alias "ImprimeBmpEspecial" (ByVal pszFileName As String, ByVal dwX As Long, ByVal dwY As Long, ByVal dwAngulo As Long) As Long
Public Declare Function EpsonNaoFiscal_Habilita_Log Lib "InterfaceEpsonNF.dll" Alias "Habilita_Log" (ByVal dwEstado As Long, ByVal pszCaminho As String) As Long
Public Declare Function EpsonNaoFiscal_ImprimeCheque Lib "InterfaceEpsonNF.dll" Alias "ImprimeCheque" (ByVal szIndice As String, ByVal szValor As String, ByVal szData As String, ByVal szPara As String, ByVal szCidade As String, ByVal szAdicional As String) As Long
Public Declare Function EpsonNaoFiscal_LeMICR Lib "InterfaceEpsonNF.dll" Alias "LeMICR" (ByVal pszCodigo As String) As Long
Public Declare Function EpsonNaoFiscal_AcionaGaveta Lib "InterfaceEpsonNF.dll" Alias "AcionaGaveta" () As Long

