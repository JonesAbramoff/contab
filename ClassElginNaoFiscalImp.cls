VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassElginNaoFiscalImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type DOCINFO
          pDocName As String
          pOutputFile As String
          pDatatype As String
End Type

Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
          ByVal pDefault As Long) As Long
Private Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, _
         pDocInfo As DOCINFO) As Long
Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, _
         pcWritten As Long) As Long

Private Const ELGIN_NAO_FISCAL_MAX_TEXTO = 30000

Private sBuffer As String
Public sPulaLinha As String

'parametros para qrcode
Public errorCorrectionLevel As Integer
Public moduleSize As Integer
Public codeType As Integer
Public QRCodeVersion As Integer
Public encodingModes As Integer

Private gsNomeImpressora As String

Public Sub Define_Impressora(ByVal sNomeImpressora As String)
    gsNomeImpressora = sNomeImpressora
End Sub

Public Function ImprimeFormatado(ByVal sTexto As String, ByVal TpoLtra As Integer, ByVal Italic As Integer, ByVal Sublin As Integer, ByVal Expand As Integer, ByVal Enfat As Integer) As Integer

Dim iResult As Integer

    iResult = DescarregaBuffer
    
    If iResult = 1 Then
        
        sTexto = DesacentuaTexto(sTexto)
        
        If TpoLtra = 1 Then sTexto = Chr(27) & Chr$(15) & sTexto & Chr$(18) 'condensado
        
        sTexto = sTexto & sPulaLinha
        
        sBuffer = sBuffer & sTexto
    
    End If

    ImprimeFormatado = iResult

End Function

Public Function ImprimeQRCode(ByVal sQRCode As String) As Integer

Dim iResult As Integer, sComando As String

    iResult = DescarregaBuffer
    
    If iResult = 1 Then
        
        sBuffer = sBuffer & Chr(27) & Chr$(97) & Chr$(1) 'centrar horizontalmente
        sBuffer = sBuffer & Chr$(29) & Chr$(111) & Chr$(0) & Chr$(moduleSize) & Chr$(0) & Chr$(2) & Chr$(29) & Chr$(107) & Chr$(11) & Chr$(76) & Chr$(107) & " " & sQRCode & " " & Chr$(0)
        sBuffer = sBuffer & Chr(27) & Chr$(97) & Chr$(0) 'alinhar à esquerda esquerda
    
    End If

    ImprimeQRCode = iResult
    
End Function

Public Function ImprimeNormal(ByVal sTexto As String) As Integer

Dim iResult As Integer

    sTexto = DesacentuaTexto(sTexto)
        
    sTexto = sTexto & sPulaLinha
    
    iResult = 1
    
    If Len(sBuffer & sTexto) > ELGIN_NAO_FISCAL_MAX_TEXTO Then
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

    Dim lhPrinter As Long
    Dim lReturn As Long
    Dim lpcWritten As Long
    Dim lDoc As Long
    Dim MyDocInfo As DOCINFO
    
    iResult = 1
    If Len(sBuffer) <> 0 Then
        
        'descarregar buffer
        lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
        If lReturn = 0 Then
            'Me.List1.AddItem ("Impressora não reconhecida.")
            
        Else
        
            MyDocInfo.pDocName = "Corporator"
            MyDocInfo.pOutputFile = vbNullString
            MyDocInfo.pDatatype = "RAW"
            lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
            Call StartPagePrinter(lhPrinter)
            lReturn = WritePrinter(lhPrinter, ByVal sBuffer, Len(sBuffer), lpcWritten)
            lReturn = EndPagePrinter(lhPrinter)
            lReturn = EndDocPrinter(lhPrinter)
            lReturn = ClosePrinter(lhPrinter)
        
        End If
        
        sBuffer = ""
    
    End If
    
    DescarregaBuffer = iResult
    
End Function

Private Sub ObterImpressora()

Dim impr As Printer

    For Each impr In Printers
        If impr.DeviceName = "VOX(U) 1" Or impr.DeviceName = gsNomeImpressora Then
            Set Printer = impr
            Exit For
        End If
    Next

End Sub

Private Sub Class_Initialize()

    sPulaLinha = vbLf
'    errorCorrectionLevel = 3
    moduleSize = 3
'    codeType = 4
'    QRCodeVersion = 1
'    encodingModes = 1
    Call ObterImpressora
    
    sBuffer = Chr$(27) & Chr$(64) 'iniciar a impressora
    sBuffer = sBuffer & Chr$(27) & Chr$(116) & Chr(2) 'code page 850
    
End Sub


