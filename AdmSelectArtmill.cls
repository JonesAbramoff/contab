VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Produto_Le_InfoUsu(ByVal objProduto As ClassProduto) As Long

Dim lErro As Long, lComando As Long, sCodAnterior As String, sDetalheCor As String, sDimEmbalagem As String

On Error GoTo Erro_Produto_Le_InfoUsu

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 124043
     
    sCodAnterior = String(ARTMILL_STRING_PRODUTO_CODANTERIOR, 0)
    sDetalheCor = String(ARTMILL_STRING_DETALHE_COR, 0)
    sDimEmbalagem = String(ARTMILL_STRING_DIM_EMBALAGEM, 0)
     
    lErro = Comando_Executar(lComando, "SELECT CodAnterior, DetalheCor, DimEmbalagem FROM Produtos WHERE Codigo = ?", _
        sCodAnterior, sDetalheCor, sDimEmbalagem, objProduto.sCodigo)
    If lErro <> AD_SQL_SUCESSO Then gError 124044

    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 124045
    If lErro <> AD_SQL_SUCESSO Then gError 124046
        
    Set objProduto.objInfoUsu = New ClassProdutoInfoUsu
    
    With objProduto.objInfoUsu
        .sCodAnterior = sCodAnterior
        .sDetalheCor = sDetalheCor
        .sDimEmbalagem = sDimEmbalagem
    End With
    
    Call Comando_Fechar(lComando)
    
    Produto_Le_InfoUsu = SUCESSO
     
    Exit Function
    
Erro_Produto_Le_InfoUsu:

    Produto_Le_InfoUsu = gErr
     
    Select Case gErr
          
        '??? completar
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 142610)
     
    End Select
     
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

Function Browser_Trata_Parametros_Customizado(ByVal sNomeTela As String, ByVal objBrowse1 As Object) As Long
'le os valores dos parametros de selecao relacionados a sNomeTela coloca os resultados na coleção

Dim lComando As Long
Dim lErro As Long
Dim tBrowseParamSelecao As typeBrowseParamSelecao
    
On Error GoTo Erro_Browser_Trata_Parametros_Customizado

    If sNomeTela = "EstoqueProdutoFilialLista" Then
        If giFilialEmpresa = 52 Then
            objBrowse1.colSelecao.Add 51
        Else
            objBrowse1.colSelecao.Add giFilialEmpresa
        End If
    Else

        'carrega os valores dos parametros de selecao que ainda faltam
        lErro = CF("BrowseParamSelecao_Le", sNomeTela, objBrowse1.colSelecao)
        If lErro <> SUCESSO Then gError 89977

    End If
    
    Browser_Trata_Parametros_Customizado = SUCESSO
    
    Exit Function
    
Erro_Browser_Trata_Parametros_Customizado:

    Browser_Trata_Parametros_Customizado = gErr

    Select Case gErr
    
        Case 89977
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142611)
        
    End Select
    
    Exit Function

End Function

Function Relatorio_Customiza_ExecutarAux(ByVal objRelOpcoes As AdmRelOpcoes, sNomeTsk As String) As Long

Dim lErro As Long
Dim sSerie As String


On Error GoTo Erro_Relatorio_Customiza_ExecutarAux

    'se forem relatorios de impressao de nota fiscal
    If UCase(left(sNomeTsk, 2)) = "NF" Then
        
        lErro = objRelOpcoes.ObterParametro("TSERIE", sSerie)
        If lErro <> SUCESSO Then gError 133003
        
        If sSerie = "1" Or sSerie = "2" Or sSerie = "3" Or sSerie = "4" Then
        
            lErro = objRelOpcoes.IncluirParametro("NFILIAL", CStr(FilialEmpresa_ConvFRFO1(giFilialEmpresa)))
            If lErro <> AD_BOOL_TRUE Then gError 126936
        
            lErro = objRelOpcoes.IncluirParametro("TSERIE", "1")
            If lErro <> AD_BOOL_TRUE Then gError 133005
        
        Else
        
            lErro = objRelOpcoes.IncluirParametro("NFILIAL", CStr(giFilialEmpresa))
            If lErro <> AD_BOOL_TRUE Then gError 133004
        
        End If
        
    ElseIf UCase(left(sNomeTsk, 5)) = "DANFE" Then

        lErro = objRelOpcoes.IncluirParametro("NFILIAL", CStr(FilialEmpresa_ConvFRFO1(giFilialEmpresa)))
        If lErro <> AD_BOOL_TRUE Then gError 126936
    
        lErro = objRelOpcoes.IncluirParametro("TSERIE", "1-e")
        If lErro <> AD_BOOL_TRUE Then gError 133005
        
    Else

        lErro = objRelOpcoes.IncluirParametro("NFILIAL", CStr(giFilialEmpresa))
        If lErro <> AD_BOOL_TRUE Then gError 126935

    End If

    Relatorio_Customiza_ExecutarAux = SUCESSO
    
    Exit Function
    
Erro_Relatorio_Customiza_ExecutarAux:

    Relatorio_Customiza_ExecutarAux = gErr
    
    Select Case gErr
    
        Case 126935, 126936, 133003, 133004, 133005
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142612)
        
    End Select
    
    Exit Function

End Function

Private Function FilialEmpresa_ConvFRFO1(ByVal iFilialEmpresa As Integer) As Integer

    If iFilialEmpresa <= DELTA_FILIALREAL_OFICIAL1 Then
        FilialEmpresa_ConvFRFO1 = iFilialEmpresa
    Else
        FilialEmpresa_ConvFRFO1 = iFilialEmpresa - DELTA_FILIALREAL_OFICIAL1
    End If
    
End Function

Function FilialEmpresa_Customiza(iFilialEmpresa As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_FilialEmpresa_Customiza

    If iFilialEmpresa = 52 Then iFilialEmpresa = 51

    FilialEmpresa_Customiza = SUCESSO
    
    Exit Function
    
Erro_FilialEmpresa_Customiza:

    FilialEmpresa_Customiza = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142613)
        
    End Select
    
    Exit Function

End Function

Function FilialEmpresaAlmox_Customiza(iFilialEmpresa As Integer) As Long

Dim lErro As Long

On Error GoTo Erro_FilialEmpresaAlmox_Customiza

    If iFilialEmpresa = 52 Then iFilialEmpresa = 51
    If iFilialEmpresa = 2 Then iFilialEmpresa = 1

    FilialEmpresaAlmox_Customiza = SUCESSO
    
    Exit Function
    
Erro_FilialEmpresaAlmox_Customiza:

    FilialEmpresaAlmox_Customiza = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142613)
        
    End Select
    
    Exit Function

End Function

