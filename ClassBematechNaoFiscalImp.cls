VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassBematechNaoFiscalImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const BEMATECH_NAO_FISCAL_MAX_TEXTO = 30000

Private Declare Function eGerarQrCodeArquivo_DUAL_DarumaFramework Lib "DarumaFrameWork.dll" (ByVal sArq As String, ByVal sQRCode As String) As Integer

Private sBuffer As String
Public sPulaLinha As String

'parametros para qrcode
Public errorCorrectionLevel As Integer
Public moduleSize As Integer
Public codeType As Integer
Public QRCodeVersion As Integer
Public encodingModes As Integer

Public sDirTemp As String

Public Function ImprimeFormatado(ByVal sTexto As String, ByVal TpoLtra As Integer, ByVal Italic As Integer, ByVal Sublin As Integer, ByVal Expand As Integer, ByVal Enfat As Integer) As Integer

Dim iResult As Integer

    sTexto = sTexto & sPulaLinha
    
    iResult = DescarregaBuffer
    
    If iResult = 1 Then
        iResult = BematechNaoFiscal_FormataTX(sTexto, TpoLtra, Italic, Sublin, Expand, Enfat)
    End If

    ImprimeFormatado = iResult

End Function

Public Function ImprimeQRCode(ByVal sQRCode As String, ByVal sNomeArq As String) As Integer

Dim iResult As Integer, sArqAux As String, sComando As String

    iResult = DescarregaBuffer
    
    If iResult = 1 Then
'        iResult = BematechNaoFiscal_ImprimeCodigoQRCODE(errorCorrectionLevel, moduleSize, codeType, QRCodeVersion, encodingModes, sQRCode)
    
        sArqAux = sDirTemp & sNomeArq & "QRCode.bmp"
    
        If dir(sArqAux) = "" Then
        
            iResult = eGerarQrCodeArquivo_DUAL_DarumaFramework(sArqAux, sQRCode)
            If iResult = 1 Then
                sComando = Chr(27) + Chr(97) + Chr(1) 'centralizar
                iResult = BematechNaoFiscal_ComandoTX(sComando, Len(sComando))
                If iResult = 1 Then
                    iResult = BematechNaoFiscal_SelecionaQualidade(4)
                    If iResult = 1 Then
                        iResult = BematechNaoFiscal_DefineDLBitmap(sArqAux)
                        If iResult = 1 Then
                            iResult = BematechNaoFiscal_PrintDLBitmap(3)
                        End If
                    End If
            
                    Call BematechNaoFiscal_SelecionaQualidade(2)
                End If
                    
                iResult = BematechNaoFiscal_BematechTX(sPulaLinha)
                    
            End If
            
        End If
    
    End If

    ImprimeQRCode = iResult
    
End Function

Public Function ImprimeNormal(ByVal sTexto As String) As Integer

Dim iResult As Integer

    sTexto = sTexto & sPulaLinha
    
    iResult = 1
    
    If Len(sBuffer & sTexto) > BEMATECH_NAO_FISCAL_MAX_TEXTO Then
        iResult = DescarregaBuffer
    End If
    sBuffer = sBuffer & sTexto
    
    ImprimeNormal = iResult
    
End Function

Public Function TerminaImpressao() As Integer

Dim iResult As Integer

    iResult = DescarregaBuffer
    
    TerminaImpressao = iResult

End Function

Public Function DescarregaBuffer() As Integer

Dim iResult As Integer

    iResult = 1
    If Len(sBuffer) <> 0 Then
    
        iResult = BematechNaoFiscal_BematechTX(sBuffer)
        sBuffer = ""
    
    End If
    
    DescarregaBuffer = iResult
    
End Function

Private Sub Class_Initialize()
    sPulaLinha = Chr$(10)
    errorCorrectionLevel = 1
    moduleSize = 10
    codeType = 0
    QRCodeVersion = 10
    encodingModes = 1

    'buscar diretorio configurado
    sDirTemp = String(128, 0)
    Call GetPrivateProfileString("Forprint", "DirTsks", "c:\forpw40\", sDirTemp, 128, "ADM100.INI")
    sDirTemp = StringZ(sDirTemp)
    If right(sDirTemp, 1) <> "\" Then sDirTemp = sDirTemp & "\"

End Sub
