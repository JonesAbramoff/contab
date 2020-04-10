VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelPosCtb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private alComando(0 To 1) As Long

Private iUltTipoLcto As Integer
Private llUltNumIntDoc As Long
Private iUltDocFilEmp As Integer
Private sUltContaForn As String

Private Sub Class_Terminate()

Dim iIndice As Integer

    For iIndice = LBound(alComando) To UBound(alComando)
        If alComando(iIndice) <> 0 Then Comando_Fechar (alComando(iIndice))
    Next
    
End Sub

Public Function PosCPObterCtaForn(sContaForn As String, ByVal iTipoLcto As Integer, ByVal lNumIntDoc As Long, ByVal iDocFilEmp As Integer) As Long
'obter a conta contabil de fornecedor ligada ao titulo

Dim lErro As Long, sConta As String, dValor As Double, iIndice As Integer

On Error GoTo Erro_PosCPObterCtaForn

    If iUltTipoLcto = iTipoLcto And llUltNumIntDoc = lNumIntDoc And iUltDocFilEmp = iDocFilEmp Then
    
        sContaForn = sUltContaForn
        
    Else
    
        For iIndice = LBound(alComando) To UBound(alComando)
            If alComando(iIndice) = 0 Then
                alComando(iIndice) = Comando_AbrirExt(GL_lConexaoRel)
                If alComando(iIndice) = 0 Then gError 93672
            End If
        Next
        
        sContaForn = ""
    
        Select Case iTipoLcto
        
            Case 1, 2, 3, 4
            
                sConta = String(STRING_CONTA, 0)
                lErro = Comando_Executar(alComando(0), "SELECT Valor, Conta FROM Lancamentos L, NFiscal N, TransacaoCTB TCTB, TiposDocInfo T WHERE L.FilialEmpresa = ? AND N.NumIntDocCPR = ? AND N.ClasseDocCpr = 1 AND L.Transacao = TCTB.Codigo AND TCTB.OrigemLcto = 10 AND N.NumIntDoc = L.NumIntDoc AND T.SubTipoContabil = TCTB.SubTipo AND T.Codigo = N.TipoNFiscal ORDER BY L.Valor Desc", dValor, sConta, iDocFilEmp, lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 201504
                
                lErro = Comando_BuscarProximo(alComando(0))
                If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201505
                
                If lErro <> AD_SQL_SUCESSO Then
                
                    sConta = String(STRING_CONTA, 0)
                    lErro = Comando_Executar(alComando(1), "SELECT Valor, Conta FROM Lancamentos L, TransacaoCTB TCTB WHERE L.FilialEmpresa = ? AND L.Transacao = TCTB.Codigo AND TCTB.OrigemLcto = 15 AND L.NumIntDoc = ? ORDER BY L.Valor Desc", dValor, sConta, iDocFilEmp, lNumIntDoc)
                    If lErro <> AD_SQL_SUCESSO Then gError 20506
                    
                    lErro = Comando_BuscarProximo(alComando(1))
                    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 201507
                
                    If lErro = AD_SQL_SUCESSO Then sContaForn = sConta
                
                Else
                
                    sContaForn = sConta
                    
                End If
            
        End Select
        
        iUltTipoLcto = iTipoLcto
        llUltNumIntDoc = lNumIntDoc
        iUltDocFilEmp = iDocFilEmp
        sUltContaForn = sContaForn
    
    End If
    
    PosCPObterCtaForn = SUCESSO
    
    Exit Function
    
Erro_PosCPObterCtaForn:

    PosCPObterCtaForn = gErr

    Select Case gErr

        Case 201504 To 201507
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELPOSCP", gErr)
        
        Case 93672
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 201503)

    End Select
    
    Exit Function
    
End Function
