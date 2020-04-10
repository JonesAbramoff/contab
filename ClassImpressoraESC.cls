VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassImpressoraESC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'https://reference.epson-biz.com/modules/ref_escpos/index.php?content_id=87

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

Private giBufferSize As Integer

Private sBuffer As String
Public sPulaLinha As String

'parametros para qrcode
Public errorCorrectionLevel As Integer
Public moduleSize As Integer
Public codeType As Integer
Public QRCodeVersion As Integer
Public encodingModes As Integer

Private giModeloImpressora As Integer

Private gsLigaCondensado As String
Private gsDesligaCondensado As String
Private gsCorteTotal As String
Private gsCorteParcial As String
Private gsCentralizarHorizontal As String
Private gsAlinharEsquerda As String
Private gsLigaEnfatizado As String
Private gsDesligaEnfatizado As String
Private gsAbrirGaveta As String
Private gsForcaESCBEMA As String

Public Sub AbrirGaveta()

    sBuffer = sBuffer & gsAbrirGaveta

End Sub

Public Sub AcionarGuilhotina(ByVal sCorte As String)

    Call DescarregaBuffer
    
    sBuffer = sBuffer & IIf(sCorte = "T", gsCorteTotal, gsCorteParcial)

End Sub

Public Sub Define_Impressora(ByVal sNomeImpressora As String, ByVal iModeloImpressora As Integer)

Dim impr As Printer
Dim sNomePadraoImpressora As String

    giModeloImpressora = iModeloImpressora
    
    giBufferSize = 4096
    errorCorrectionLevel = 3
    moduleSize = 3
    codeType = 4
    QRCodeVersion = 1
    encodingModes = 1
    
    For Each impr In Printers
        If UCase(impr.DeviceName) = UCase(sNomeImpressora) Then
            Set Printer = impr
            Exit For
        End If
    Next
    
    sBuffer = Chr$(27) & Chr$(64) 'iniciar a impressora
    
    Select Case iModeloImpressora
    
        Case IMPRESSORA_ELGIN_VOX
            gsLigaCondensado = Chr$(27) & Chr$(15)
            gsDesligaCondensado = Chr$(18)
            gsCentralizarHorizontal = Chr$(27) & Chr$(97) & Chr$(1)
            gsAlinharEsquerda = Chr$(27) & Chr$(97) & Chr$(0)
        
        Case IMPRESSORA_BEMATECH_ESCBEMA
            'gsLigaCondensado = Chr$(15)
            gsForcaESCBEMA = Chr$(29) & Chr$(249) & Chr$(32) & Chr$(48)
            gsLigaCondensado = gsForcaESCBEMA & Chr$(15)
            gsDesligaCondensado = Chr$(18)
            gsCorteParcial = Chr$(27) & "m"
            gsCorteTotal = Chr$(27) & "w"
            gsCentralizarHorizontal = Chr$(27) & Chr$(97) & Chr$(1)
            gsAlinharEsquerda = Chr$(27) & Chr$(97) & Chr$(0)
            gsAbrirGaveta = Chr$(27) & "v" & Chr$(140)
                
            'forçar temporariamente como ESC/BEMA
            sBuffer = sBuffer & gsForcaESCBEMA
        
        Case IMPRESSORA_EPSON_ESCPOS
            gsLigaCondensado = Chr$(27) & "M" & Chr$(1)
            gsDesligaCondensado = Chr$(27) & "M" & Chr$(0)
            gsCorteParcial = Chr$(29) & "V" + Chr$(1)
            gsCorteTotal = Chr$(29) & "V" + Chr$(0)
            gsCentralizarHorizontal = Chr$(27) & Chr$(97) & Chr$(1)
            gsAlinharEsquerda = Chr$(27) & Chr$(97) & Chr$(0)
            gsAbrirGaveta = Chr$(27) & Chr$(112) & Chr$(48) & Chr$(55) & Chr$(121)
        
        Case IMPRESSORA_DARUMA_ESCPOS
            gsLigaCondensado = Chr$(15)
            gsDesligaCondensado = Chr$(18)
            gsCorteParcial = Chr$(27) & "m"
            gsCorteTotal = Chr$(27) & "m"
            gsCentralizarHorizontal = Chr$(27) & Chr$(106) & Chr$(49)
            gsAlinharEsquerda = Chr$(27) & Chr$(106) & Chr$(48)
            gsLigaEnfatizado = Chr$(27) & Chr$(69)
            gsDesligaEnfatizado = Chr$(27) & Chr$(70)
            gsAbrirGaveta = Chr$(27) & "p"
        
    End Select
    
    sBuffer = sBuffer & Chr$(27) & Chr$(116) & Chr$(2) 'code page 850
    sBuffer = sBuffer & Chr$(27) & Chr$(72) '48 colunas

End Sub

Public Function ImprimeFormatado(ByVal sTexto As String, ByVal TpoLtra As Integer, ByVal Italic As Integer, ByVal Sublin As Integer, ByVal Expand As Integer, ByVal Enfat As Integer) As Integer

Dim iResult As Integer

    iResult = DescarregaBuffer
    
    If iResult = 1 Then
        
        sTexto = DesacentuaTexto(sTexto)
        
        sTexto = sTexto & sPulaLinha
    
        If TpoLtra = 1 Then sTexto = gsLigaCondensado & sTexto & gsDesligaCondensado             'condensado
        
        If Enfat = 1 Then sTexto = gsLigaEnfatizado & sTexto & gsDesligaEnfatizado             'negrito
        
        sTexto = sTexto & sPulaLinha
        
        sBuffer = sBuffer & sTexto
    
    End If

    ImprimeFormatado = iResult

End Function

Public Function ImprimeQRCode(ByVal sQRCode As String) As Integer

Dim iResult As Integer, sComando As String
Dim iTam1 As Integer, iTam2 As Integer, iLen As Integer

    iResult = DescarregaBuffer
    
    If iResult = 1 Then
        
        sBuffer = sBuffer & gsCentralizarHorizontal
        
        Select Case giModeloImpressora
        
            Case IMPRESSORA_ELGIN_VOX
                sBuffer = sBuffer & Chr$(29) & Chr$(111) & Chr$(0) & Chr$(moduleSize) & Chr$(0) & Chr$(2) & Chr$(29) & Chr$(107) & Chr$(11) & Chr$(76) & Chr$(107) & " " & sQRCode & " " & Chr$(0)
                
            Case IMPRESSORA_BEMATECH_ESCBEMA
            
                iLen = Len(sQRCode)
                If iLen > 255 Then
                
                    iTam1 = iLen Mod 256
                    iTam2 = iLen \ 256
                    
                Else
                    iTam1 = iLen
                    iTam2 = 0
                End If
            
                sBuffer = sBuffer & Chr$(29) & Chr$(107) & Chr$(81) & Chr$(moduleSize) & Chr$(8) & Chr$(8) & Chr$(1) & Chr$(iTam1) & Chr$(iTam2) & sQRCode
            
            Case IMPRESSORA_EPSON_ESCPOS
            
                iLen = Len(sQRCode) + 3
                If iLen > 255 Then
                
                    iTam1 = iLen Mod 256
                    iTam2 = iLen \ 256
                    
                Else
                    iTam1 = iLen
                    iTam2 = 0
                End If
                
                'https://reference.epson-biz.com/modules/ref_escpos/index.php?content_id=140 em diante
            
                sBuffer = sBuffer & Chr$(29) & Chr$(40) & Chr$(107) & Chr$(4) & Chr$(0) & Chr$(49) & Chr$(65) & Chr$(50) & Chr$(0)
                sBuffer = sBuffer & Chr$(29) & Chr$(40) & Chr$(107) & Chr$(3) & Chr$(0) & Chr$(49) & Chr$(67) & Chr$(4)
                sBuffer = sBuffer & Chr$(29) & Chr$(40) & Chr$(107) & Chr$(3) & Chr$(0) & Chr$(49) & Chr$(69) & Chr$(48)
                sBuffer = sBuffer & Chr$(29) & Chr$(40) & Chr$(107) & Chr$(iTam1) & Chr$(iTam2) & Chr$(49) & Chr$(80) & Chr$(48) & sQRCode
                sBuffer = sBuffer & Chr$(29) & Chr$(40) & Chr$(107) & Chr$(3) & Chr$(0) & Chr$(49) & Chr$(81) & Chr$(48)
            
            Case IMPRESSORA_DARUMA_ESCPOS
            
                iLen = Len(sQRCode) + 2
                If iLen > 255 Then
                
                    iTam1 = iLen Mod 256
                    iTam2 = iLen \ 256
                    
                Else
                    iTam1 = iLen
                    iTam2 = 0
                End If
            
                sBuffer = sBuffer & Chr$(27) & Chr$(129) & Chr$(iTam1) & Chr$(iTam2) & Chr$(moduleSize) & "M" & sQRCode
        
        End Select
        
        sBuffer = sBuffer & gsAlinharEsquerda
    
    End If

    ImprimeQRCode = iResult
    
End Function

Public Function ImprimeNormal(ByVal sTexto As String) As Integer

Dim iResult As Integer

    sTexto = DesacentuaTexto(sTexto)
        
    sTexto = sTexto & sPulaLinha
    
    iResult = 1
    
    If Len(sBuffer & sTexto) > giBufferSize Then
        iResult = DescarregaBuffer
    End If
    sBuffer = sBuffer & sTexto
    
    ImprimeNormal = iResult
    
End Function

Public Function TerminaImpressao() As Integer

Dim iResult As Integer

    iResult = DescarregaBuffer
    
    If giModeloImpressora = IMPRESSORA_BEMATECH_ESCBEMA Then
        
        'forçar ao que estava antes de ser setado temporariamente
        sBuffer = sBuffer & Chr$(29) & Chr$(249) & Chr$(31) & Chr$(49)
    
        Call DescarregaBuffer
    
    End If
    
    TerminaImpressao = iResult

End Function

Public Function DescarregaBuffer() As Integer

Dim iResult As Integer

On Error GoTo Erro_DescarregaBuffer

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
            
            MsgBox ("Impressora não reconhecida.")
        
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
    
    Exit Function
    
Erro_DescarregaBuffer:

    MsgBox (Error$)
    
    DescarregaBuffer = iResult
    
    Exit Function
    
End Function

Private Sub Class_Initialize()

    sPulaLinha = vbLf
    
End Sub

Public Function ImprimeCodBarras(ByVal sCodBarras As String, ByVal sTipoCodBarras As String) As Integer

Dim iResult As Integer, sComando As String
Dim iTam1 As Integer, iTam2 As Integer, iLen As Integer, sAux As String

    iResult = DescarregaBuffer
    
    If iResult = 1 Then
        
        sBuffer = sBuffer & gsCentralizarHorizontal
        
        Select Case giModeloImpressora
        
            Case IMPRESSORA_ELGIN_VOX
                
                
            Case IMPRESSORA_BEMATECH_ESCBEMA
            
            
            Case IMPRESSORA_EPSON_ESCPOS
                Select Case sTipoCodBarras
                
                    Case "CODE128"
                    
                        sBuffer = sBuffer & Chr$(29) & "w" & Chr$(1) 'Module width
                        sBuffer = sBuffer & Chr$(29) & "h" & Chr$(48) 'Bar lenght(Height)
                        
                        sAux = left(sCodBarras, 22)
                        'GS k m n d1...dn
                        sBuffer = sBuffer & Chr$(29) & "kI" & Chr$(Len(sAux) + 2) & Chr$(123) & Chr$(67) & sAux
                        sBuffer = sBuffer & sPulaLinha
                        If Len(sCodBarras) > 22 Then
                            sAux = Mid(sCodBarras, 23)
                            'GS k m n d1...dn
                            sBuffer = sBuffer & Chr$(29) & "kI" & Chr$(Len(sAux) + 2) & Chr$(123) & Chr$(67) & sAux
                        End If
                        'Call DescarregaBuffer
                
                End Select
            
            
            Case IMPRESSORA_DARUMA_ESCPOS
            
                Select Case sTipoCodBarras
                
                    Case "CODE128"
                        'ESC b n1 n2 n3 n4 s1...sn NULL
                        sBuffer = sBuffer & Chr$(27) & "b" & Chr$(5) & Chr$(0) & Chr$(0) & Chr$(0) & sCodBarras & Chr$(0)
                
                End Select
            
        End Select
        
        sBuffer = sBuffer & gsAlinharEsquerda
    
    End If

    ImprimeCodBarras = iResult
    
End Function




