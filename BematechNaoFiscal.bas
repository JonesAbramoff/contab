Attribute VB_Name = "BematechNaoFiscal"
Option Explicit

'!!!!! INCLUIR CLAUSULA ALIAS A CADA NOVA FUNCAO UTILIZADA !!!!!!!!!!!!!!!!!!!!!!

'Declarando a MP2032.DLL e sua funções em Visual Basic
 
Public Declare Function BematechNaoFiscal_IniciaPorta Lib "MP2032.DLL" Alias "IniciaPorta" (ByVal Porta As String) As Integer
Public Declare Function BematechNaoFiscal_BematechTX Lib "MP2032.DLL" Alias "BematechTX" (ByVal comando As String) As Integer
Public Declare Function BematechNaoFiscal_ComandoTX Lib "MP2032.DLL" Alias "ComandoTX" (ByVal BufTrans As String, ByVal Flag As Integer) As Integer
Public Declare Function BematechNaoFiscal_CaracterGrafico Lib "MP2032.DLL" (ByVal BufTrans As String, ByVal TamBufTrans As Integer) As Integer
Public Declare Function BematechNaoFiscal_Le_Status Lib "MP2032.DLL" () As Integer
Public Declare Function BematechNaoFiscal_AutenticaDoc Lib "MP2032.DLL" (ByVal Texto As String, ByVal Tempo As Integer) As Integer
Public Declare Function BematechNaoFiscal_DocumentInserted Lib "MP2032.DLL" () As Integer
Public Declare Function BematechNaoFiscal_FechaPorta Lib "MP2032.DLL" Alias "FechaPorta" () As Integer
Public Declare Function BematechNaoFiscal_Le_Status_Gaveta Lib "MP2032.DLL" () As Integer
Public Declare Function BematechNaoFiscal_ConfiguraTamanhoExtrato Lib "MP2032.DLL" (ByVal NumeroLinhas As Integer) As Integer
Public Declare Function BematechNaoFiscal_HabilitaExtratoLongo Lib "MP2032.DLL" (ByVal Flag As Integer) As Integer
Public Declare Function BematechNaoFiscal_HabilitaEsperaImpressao Lib "MP2032.DLL" (ByVal Flag As Integer) As Integer
Public Declare Function BematechNaoFiscal_EsperaImpressao Lib "MP2032.DLL" () As Integer
Public Declare Function BematechNaoFiscal_ConfiguraModeloImpressora Lib "MP2032.DLL" Alias "ConfiguraModeloImpressora" (ByVal ModeloImpressora As Integer) As Integer
Public Declare Function BematechNaoFiscal_AcionaGuilhotina Lib "MP2032.DLL" Alias "AcionaGuilhotina" (ByVal Modo As Integer) As Integer
Public Declare Function BematechNaoFiscal_FormataTX Lib "MP2032.DLL" Alias "FormataTX" (ByVal BufTrans As String, ByVal TpoLtra As Integer, ByVal Italic As Integer, ByVal Sublin As Integer, ByVal Expand As Integer, ByVal Enfat As Integer) As Integer
Public Declare Function BematechNaoFiscal_HabilitaPresenterRetratil Lib "MP2032.DLL" (ByVal iFlag As Integer) As Integer
Public Declare Function BematechNaoFiscal_ProgramaPresenterRetratil Lib "MP2032.DLL" (ByVal iTempo As Integer) As Integer
Public Declare Function BematechNaoFiscal_VerificaPapelPresenter Lib "MP2032.DLL" () As Integer

' Função para Configuração dos Códigos de Barras

Public Declare Function BematechNaoFiscal_ConfiguraCodigoBarras Lib "MP2032.DLL" (ByVal Altura As Integer, ByVal Largura As Integer, ByVal PosicaoCaracteres As Integer, ByVal Fonte As Integer, ByVal Margem As Integer) As Integer

' Funções para impressão dos códigos de barras

Public Declare Function BematechNaoFiscal_ImprimeCodigoBarrasUPCA Lib "MP2032.DLL" (ByVal Codigo As String) As Integer
Public Declare Function BematechNaoFiscal_ImprimeCodigoBarrasUPCE Lib "MP2032.DLL" (ByVal Codigo As String) As Integer
Public Declare Function BematechNaoFiscal_ImprimeCodigoBarrasEAN13 Lib "MP2032.DLL" (ByVal Codigo As String) As Integer
Public Declare Function BematechNaoFiscal_ImprimeCodigoBarrasEAN8 Lib "MP2032.DLL" (ByVal Codigo As String) As Integer
Public Declare Function BematechNaoFiscal_ImprimeCodigoBarrasCODE39 Lib "MP2032.DLL" (ByVal Codigo As String) As Integer
Public Declare Function BematechNaoFiscal_ImprimeCodigoBarrasCODE93 Lib "MP2032.DLL" (ByVal Codigo As String) As Integer
Public Declare Function BematechNaoFiscal_ImprimeCodigoBarrasCODE128 Lib "MP2032.DLL" (ByVal Codigo As String) As Integer
Public Declare Function BematechNaoFiscal_ImprimeCodigoBarrasITF Lib "MP2032.DLL" (ByVal Codigo As String) As Integer
Public Declare Function BematechNaoFiscal_ImprimeCodigoBarrasCODABAR Lib "MP2032.DLL" (ByVal Codigo As String) As Integer
Public Declare Function BematechNaoFiscal_ImprimeCodigoBarrasISBN Lib "MP2032.DLL" (ByVal Codigo As String) As Integer
Public Declare Function BematechNaoFiscal_ImprimeCodigoBarrasMSI Lib "MP2032.DLL" (ByVal Codigo As String) As Integer
Public Declare Function BematechNaoFiscal_ImprimeCodigoBarrasPLESSEY Lib "MP2032.DLL" (ByVal Codigo As String) As Integer
Public Declare Function BematechNaoFiscal_ImprimeCodigoBarrasPDF417 Lib "MP2032.DLL" (ByVal NivelCorrecaoErros As Integer, ByVal Altura As Integer, ByVal Largura As Integer, ByVal Colunas As Integer, ByVal Codigo As String) As Integer
Public Declare Function BematechNaoFiscal_ImprimeCodigoQRCODE Lib "MP2032.DLL" Alias "ImprimeCodigoQRCODE" (ByVal errorCorrectionLevel As Integer, ByVal moduleSize As Integer, ByVal codeType As Integer, ByVal QRCodeVersion As Integer, ByVal encodingModes As Integer, ByVal codeQr As String) As Integer
 

' Funções para impressão de BitMap

Public Declare Function BematechNaoFiscal_ImprimeBitmap Lib "MP2032.DLL" Alias "ImprimeBitmap" (ByVal Name As String, ByVal mode As Integer) As Integer
Public Declare Function BematechNaoFiscal_ImprimeBmpEspecial Lib "MP2032.DLL" (ByVal Name As String, ByVal xScale As Integer, ByVal yScale As Integer, ByVal angle As Integer) As Integer
Public Declare Function BematechNaoFiscal_AjustaLarguraPapel Lib "MP2032.DLL" (ByVal width As Integer) As Integer
Public Declare Function BematechNaoFiscal_SelectDithering Lib "MP2032.DLL" (ByVal Tipo As Integer) As Integer
Public Declare Function BematechNaoFiscal_PrinterReset Lib "MP2032.DLL" () As Integer
Public Declare Function BematechNaoFiscal_LeituraStatusEstendido Lib "MP2032.DLL" (A() As Byte) As Integer
Public Declare Function BematechNaoFiscal_IoControl Lib "MP2032.DLL" (ByVal Flag As Integer, ByVal mode As Boolean) As Integer
Public Declare Function BematechNaoFiscal_DefineNVBitmap Lib "MP2032.DLL" (ByVal Count As Integer, filenames() As String) As Integer
Public Declare Function BematechNaoFiscal_PrintNVBitmap Lib "MP2032.DLL" Alias "PrintNVBitmap" (ByVal image As Integer, ByVal mode As Integer) As Integer
Public Declare Function BematechNaoFiscal_Define1NVBitmap Lib "MP2032.DLL" Alias "Define1NVBitmap" (ByVal fileName As String) As Integer
Public Declare Function BematechNaoFiscal_DefineDLBitmap Lib "MP2032.DLL" Alias "DefineDLBitmap" (ByVal fileName As String) As Integer
Public Declare Function BematechNaoFiscal_PrintDLBitmap Lib "MP2032.DLL" Alias "PrintDLBitmap" (ByVal mode As Integer) As Integer
Public Declare Function BematechNaoFiscal_SelecionaQualidade Lib "MP2032.DLL" Alias "SelecionaQualidadeImpressao" (ByVal mode As Integer) As Integer
' Função de Firmware

Public Declare Function BematechNaoFiscal_AtualizaFirmware Lib "MP2032.DLL" (ByVal fileName As String) As Integer
 

