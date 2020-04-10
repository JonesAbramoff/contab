VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassRelOpCromaton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private lComando As Long

'chave dos ultimos dados lidos
Private sUltProd As String
Private iUltEmb As Integer
Private dUltQtdProd As Double
Private iUltFilialEmpresa As Integer
Private sUltCodigoOP As String

'ultimos resultados obtidos
Private gsDescEmb As String
Private gdQtdeEmb As Double
Private gdQtdeFrac As Double

Private Function OPObterInfoRel(ByVal sProduto As String, ByVal iEmbalagem As Integer, ByVal dQtdProd As Double, ByVal iFilialEmpresa As Integer, ByVal sCodigoOP As String) As Long

Dim lErro As Long, dCapacidade As Double, sSigla As String, sDescricao As String

On Error GoTo Erro_OPObterInfoRel

    If sProduto <> sUltProd Or iEmbalagem <> iUltEmb Or dQtdProd <> dUltQtdProd Or iFilialEmpresa <> iUltFilialEmpresa Or sCodigoOP <> sUltCodigoOP Then
    
        If lComando = 0 Then
            lComando = Comando_AbrirExt(GL_lConexaoRel)
            If lComando = 0 Then gError 184462
        End If
        sSigla = String(STRING_EMBALAGEM_SIGLA, 0)
        sDescricao = String(STRING_EMBALAGEM_DESCRICAO, 0)
        
        If iEmbalagem <> 0 Then
            lErro = Comando_Executar(lComando, "SELECT ProdutoEmbalagem.Capacidade, Embalagens.Sigla, Embalagens.Descricao FROM ProdutoEmbalagem, Embalagens WHERE ProdutoEmbalagem.Embalagem = Embalagens.Codigo AND Produto = ? AND Embalagem = ?", _
                dCapacidade, sSigla, sDescricao, sProduto, iEmbalagem)
        Else
            lErro = Comando_Executar(lComando, "SELECT ProdutoEmbalagem.Capacidade, Embalagens.Sigla, Embalagens.Descricao FROM ProdutoEmbalagem, Embalagens, OPFichaProcesso WHERE OPFichaProcesso.Embalagem = Embalagens.Codigo AND ProdutoEmbalagem.Embalagem = Embalagens.Codigo AND ProdutoEmbalagem.Produto = ? AND OPFichaProcesso.FilialEmpresa = ? AND OPFichaProcesso.CodigoOP = ?", _
                dCapacidade, sSigla, sDescricao, sProduto, iFilialEmpresa, sCodigoOP)
        End If
        If lErro <> AD_SQL_SUCESSO Then gError 184465
        
        lErro = Comando_BuscarProximo(lComando)
        If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 184466
        If lErro = AD_SQL_SUCESSO Then
            gsDescEmb = sDescricao
        Else

            Select Case iEmbalagem
            
                Case 1
                    dCapacidade = 20
                    gsDescEmb = "SACOS MÉDIOS COM IMPRESSÃO"
                    
                Case 2
                    dCapacidade = 25
                    gsDescEmb = "SACOS GRANDES COM IMPRESSÃO"
                
                Case Else
                    dCapacidade = 25
                    gsDescEmb = ""
                    
            End Select
        
        End If
        
        gdQtdeEmb = Fix(dQtdProd / dCapacidade)
        gdQtdeFrac = dQtdProd - (gdQtdeEmb * dCapacidade)
                
        sUltProd = sProduto
        iUltEmb = iEmbalagem
        dUltQtdProd = dQtdProd
        iUltFilialEmpresa = iFilialEmpresa
        sUltCodigoOP = sCodigoOP
                
    End If
    
    OPObterInfoRel = SUCESSO
    
    Exit Function
    
Erro_OPObterInfoRel:

    OPObterInfoRel = gErr

    Select Case gErr

        Case 184462
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
            
        Case 184465, 184466
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELOPCROMATON", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184461)

    End Select
    
    Exit Function

End Function

Public Function ObterQtdEmb_Cro(dQtdEmb As Double, ByVal sProduto As String, ByVal iEmbalagem As Integer, ByVal dQtdProd As Double, ByVal iFilialEmpresa As Integer, ByVal sCodigoOP As String) As Long

Dim lErro As Long

On Error GoTo Erro_ObterQtdEmb_Cro

    lErro = OPObterInfoRel(sProduto, iEmbalagem, dQtdProd, iFilialEmpresa, sCodigoOP)
    If lErro <> SUCESSO Then gError 184463

    dQtdEmb = gdQtdeEmb
    
    ObterQtdEmb_Cro = SUCESSO
    
    Exit Function
    
Erro_ObterQtdEmb_Cro:

    ObterQtdEmb_Cro = gErr

    Select Case gErr

        Case 184463
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184460)

    End Select
    
    Exit Function

End Function

Public Function ObterQtdFrac_Cro(dQtdFrac As Double, ByVal sProduto As String, ByVal iEmbalagem As Integer, ByVal dQtdProd As Double, ByVal iFilialEmpresa As Integer, ByVal sCodigoOP As String) As Long

Dim lErro As Long

On Error GoTo Erro_ObterQtdFrac_Cro

    lErro = OPObterInfoRel(sProduto, iEmbalagem, dQtdProd, iFilialEmpresa, sCodigoOP)
    If lErro <> SUCESSO Then gError 184463

    dQtdFrac = gdQtdeFrac
    
    ObterQtdFrac_Cro = SUCESSO
    
    Exit Function
    
Erro_ObterQtdFrac_Cro:

    ObterQtdFrac_Cro = gErr

    Select Case gErr

        Case 184463
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184460)

    End Select
    
    Exit Function

End Function

Public Function ObterDescEmb_Cro(sAux As String, ByVal sProduto As String, ByVal iEmbalagem As Integer, ByVal dQtdProd As Double, ByVal iFilialEmpresa As Integer, ByVal sCodigoOP As String) As Long

Dim lErro As Long

On Error GoTo Erro_ObterDescEmb_Cro

    lErro = OPObterInfoRel(sProduto, iEmbalagem, dQtdProd, iFilialEmpresa, sCodigoOP)
    If lErro <> SUCESSO Then gError 184463

    sAux = gsDescEmb
    
    ObterDescEmb_Cro = SUCESSO
    
    Exit Function
    
Erro_ObterDescEmb_Cro:

    ObterDescEmb_Cro = gErr

    Select Case gErr

        Case 184463
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 184460)

    End Select
    
    Exit Function

End Function

Private Sub Class_Terminate()

    If lComando <> 0 Then Call Comando_Fechar(lComando)

End Sub
