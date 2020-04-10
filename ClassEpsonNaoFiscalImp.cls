VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassEpsonNaoFiscalImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const EPSON_NAO_FISCAL_MAX_TEXTO = 30000

Private sBuffer As String
Public sPulaLinha As String

'parametros para qrcode
Public errorCorrectionLevel As Integer
Public moduleSize As Integer
Public codeType As Integer
Public QRCodeVersion As Integer
Public encodingModes As Integer

Public Function ImprimeFormatado(ByVal sTexto As String, ByVal TpoLtra As Integer, ByVal Italic As Integer, ByVal Sublin As Integer, ByVal Expand As Integer, ByVal Enfat As Integer) As Integer

Dim iResult As Integer

    sTexto = sTexto & sPulaLinha
    
    iResult = DescarregaBuffer
    
    If iResult = 1 Then
        iResult = EpsonNaoFiscal_FormataTX(sTexto, TpoLtra, Italic, Sublin, Expand, Enfat)
    End If

    ImprimeFormatado = iResult

End Function

Public Function ImprimeQRCode(ByVal sQRCode As String) As Integer

Dim iResult As Integer, sComando As String

    iResult = DescarregaBuffer
    
    If iResult = 1 Then
        
        'iResult = EpsonNaoFiscal_ImprimeCodigoQRCODE(errorCorrectionLevel, moduleSize, codeType, QRCodeVersion, encodingModes, sQRCode)
        sComando = "<ce><qrcode>" & sQRCode & "<lmodulo>5</lmodulo></qrcode></ce>" & sPulaLinha
        iResult = EpsonNaoFiscal_ImprimeTextoTag(sComando)
    
    End If

    ImprimeQRCode = iResult
    
End Function

Public Function ImprimeNormal(ByVal sTexto As String) As Integer

Dim iResult As Integer

    sTexto = sTexto & sPulaLinha
    
    iResult = 1
    
    If Len(sBuffer & sTexto) > EPSON_NAO_FISCAL_MAX_TEXTO Then
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
    
        iResult = EpsonNaoFiscal_ImprimeTextoTag(sBuffer)
        sBuffer = ""
    
    End If
    
    DescarregaBuffer = iResult
    
End Function

Private Sub Class_Initialize()
    sPulaLinha = Chr$(10)
'    errorCorrectionLevel = 3
'    moduleSize = 4
'    codeType = 4
'    QRCodeVersion = 1
'    encodingModes = 1
End Sub

