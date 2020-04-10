VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelNFCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function eGerarQrCodeArquivo_DUAL_DarumaFramework Lib "DarumaFramework.dll" (ByVal sArq As String, ByVal sQRCode As String) As Integer

Private lNumIntNFUltAux As Long
Private sArqUltAux As String

Function ObterImagemQRCode(ByVal lNumIntNF As Long, sArqAux As String) As Long
'obtem texto associado à baixa do titulo ligado à comissao

Dim lErro As Long

On Error GoTo Erro_ObterImagemQRCode

    If lNumIntNF <> lNumIntNFUltAux Then
    
        'Obtem Dados da Comissao
        lErro = ObterImagemQRCode_Aux(lNumIntNF)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
      
   End If

    sArqAux = sArqUltAux

    ObterImagemQRCode = SUCESSO
    
    Exit Function
    
Erro_ObterImagemQRCode:
    
    ObterImagemQRCode = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
        
    Exit Function
    
End Function

Private Function ObterImagemQRCode_Aux(ByVal lNumIntNF As Long) As Long

Dim lErro As Long
Dim lComando As Long, sQRCode As String, sArqAux As String, iRetornoDual As Integer, schNFe As String, sBuffer As String

On Error GoTo Erro_ObterImagemQRCode_Aux

    lComando = Comando_AbrirExt(GL_lConexaoRel)
    If lComando = 0 Then gError 201566
                
    sQRCode = String(500, 0)
    schNFe = String(STRING_NFE_CHNFE, 0)
    lErro = Comando_Executar(lComando, "SELECT chNFe, QRCode FROM NFeFedProtNFe WHERE NumIntNF = ? AND QRCode <> ''", schNFe, sQRCode, lNumIntNF)
    If lErro <> AD_SQL_SUCESSO Then gError 201567
    
    lErro = Comando_BuscarProximo(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201568
    
    If lErro = AD_SQL_SUCESSO Then
    
        'buscar diretorio configurado
        sBuffer = String(128, 0)
        Call GetPrivateProfileString("Forprint", "DirTsks", "c:\forpw40\", sBuffer, 128, "ADM100.INI")
        sBuffer = StringZ(sBuffer)
        If right(sBuffer, 1) <> "\" Then sBuffer = sBuffer & "\"
        
        sArqAux = sBuffer & schNFe & "QRCode.bmp"
    
        If Dir(sArqAux) = "" Then
        
            iRetornoDual = eGerarQrCodeArquivo_DUAL_DarumaFramework(sArqAux, sQRCode)
            If iRetornoDual <> 1 Then gError 201569
    
        End If
    
    End If
    
    sArqUltAux = sArqAux
    lNumIntNFUltAux = lNumIntNF

    Call Comando_Fechar(lComando)
    
    ObterImagemQRCode_Aux = SUCESSO
    
    Exit Function
    
Erro_ObterImagemQRCode_Aux:
    
    ObterImagemQRCode_Aux = gErr
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select
        
    Call Comando_Fechar(lComando)
    
    Exit Function
    
End Function

