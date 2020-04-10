VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTFatPagTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get DiferencaValor(ByVal objCT As Object) As Object
     Set DiferencaValor = objCT.objUserControl.Controls("DiferencaValor")
End Property

Public Property Get DiferencaPerc(ByVal objCT As Object) As Object
     Set DiferencaPerc = objCT.objUserControl.Controls("DiferencaPerc")
End Property

Public Property Get FornGrid(ByVal objCT As Object) As Object
     Set FornGrid = objCT.objUserControl.Controls("FornGrid")
End Property

Public Function FATPAG_Inicializa_GridNF(ByVal objCT As Object, ByVal objGridInt As AdmGrid) As Long
'Faz as Inicializações no Grid de Notas Fiscais Fatura

    'Indica o Form do Grid
    Set objGridInt.objForm = objCT

    'Indica os nomes das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Filial ")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("FF")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Selecionada")
    
    'Indica os campos relacionados a cada coluna
    objGridInt.colCampo.Add (objCT.FilialNF.Name)
    objGridInt.colCampo.Add (FornGrid(objCT).Name)
    objGridInt.colCampo.Add (objCT.FilialFornecedor.Name)
    objGridInt.colCampo.Add (objCT.NumNotaFiscal.Name)
    objGridInt.colCampo.Add (objCT.DataEmissaoNF.Name)
    objGridInt.colCampo.Add (objCT.ValorNF.Name)
    objGridInt.colCampo.Add (objCT.Selecionada.Name)
    
    'Inicializa os valores das colunas
    objCT.iGrid_FilialNF_Col = 1
    objCT.gobjinfousu.iGrid_Fornecedor_Col = 2
    objCT.iGrid_FilialFornecedor_Col = 3
    objCT.iGrid_Numero_Col = 4
    objCT.iGrid_Emissao_Col = 5
    objCT.iGrid_ValorNF_Col = 6
    objCT.iGrid_Selecionada_Col = 7

    'Indica o Grid ao qual faz referencia
    objGridInt.objGrid = objCT.GridNF

    'Linhas visíveis do Grid
    objGridInt.iLinhasVisiveis = 11

    'Todas as linhas do Grid
    objGridInt.objGrid.Rows = 2001

    'Indica a largura da coluna 0
    objCT.GridNF.ColWidth(0) = 300

    'Indica a largura automática do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR

    'Chama a rotina que faz as demais inicializações
    Call Grid_Inicializa(objGridInt)

    FATPAG_Inicializa_GridNF = SUCESSO

    Exit Function

End Function

Public Function FATPAG_BotaoConsultarNFiscal_Click(ByVal objCT As Object, ByVal objNFsPag As ClassNFsPag) As Long

On Error GoTo Erro_FATPAG_BotaoConsultarNFiscal_Click

    objNFsPag.lFornecedor = LCodigo_Extrai(objCT.GridNF.TextMatrix(objCT.GridNF.Row, objCT.gobjinfousu.iGrid_Fornecedor_Col))

    FATPAG_BotaoConsultarNFiscal_Click = SUCESSO

    Exit Function

Erro_FATPAG_BotaoConsultarNFiscal_Click:

    FATPAG_BotaoConsultarNFiscal_Click = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202149)

    End Select

    Exit Function
    
End Function

Public Function FATPAG_Move_Tela_Memoria(ByVal objCT As Object, ByVal objNFsPag As ClassNFsPag, ByVal iLinha As Integer) As Long

On Error GoTo Erro_FATPAG_Move_Tela_Memoria

    objNFsPag.lFornecedor = LCodigo_Extrai(objCT.GridNF.TextMatrix(iLinha, objCT.gobjinfousu.iGrid_Fornecedor_Col))

    FATPAG_Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_FATPAG_Move_Tela_Memoria:

    FATPAG_Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202149)

    End Select

    Exit Function
    
End Function

Public Function FATPAG_Calcula_Dif(ByVal objCT) As Long

    DiferencaValor(objCT).Caption = Format(StrParaDbl(objCT.ValorTotal.Text) - StrParaDbl(objCT.ValorTotalNFSelecionadas.Caption), "STANDARD")
    
    If StrParaDbl(objCT.ValorTotalNFSelecionadas.Caption) <> 0 Then
        DiferencaPerc(objCT).Caption = Format(StrParaDbl(DiferencaValor(objCT).Caption) / StrParaDbl(objCT.ValorTotalNFSelecionadas.Caption), "PERCENT")
    Else
        DiferencaPerc(objCT).Caption = ""
    End If

    FATPAG_Calcula_Dif = SUCESSO
    
End Function

Public Function FATPAG_Limpa_Tela_FaturasPag(ByVal objCT) As Long

    DiferencaValor(objCT).Caption = ""
    DiferencaPerc(objCT).Caption = ""

    FATPAG_Limpa_Tela_FaturasPag = SUCESSO
    
End Function

Public Function FATPAG_Carrega_GridNF(ByVal objCT As Object, ByVal objNFsPag As ClassNFsPag, ByVal iLinha As Integer) As Long

Dim lErro As Long
Dim objForn As New ClassFornecedor

On Error GoTo Erro_FATPAG_Carrega_GridNF

    objForn.lCodigo = objNFsPag.lFornecedor

    lErro = CF("Fornecedor_Le", objForn)
    If lErro <> SUCESSO And lErro <> 12729 And lErro <> 12732 Then gError 202150

    objCT.GridNF.TextMatrix(iLinha, objCT.gobjinfousu.iGrid_Fornecedor_Col) = CStr(objNFsPag.lFornecedor) & SEPARADOR & objForn.sNomeReduzido

    FATPAG_Carrega_GridNF = SUCESSO
    
    Exit Function

Erro_FATPAG_Carrega_GridNF:

    FATPAG_Carrega_GridNF = gErr

    Select Case gErr
    
        Case 202150

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202151)

    End Select

    Exit Function
    
End Function

Public Function BotaoNFRegistrarDif_Click(ByVal objCT As Object) As Long

Dim lErro As Long
Dim vbMsgResult As VbMsgBoxResult
Dim lFornecedor As Long
Dim dValor As Double
Dim objForn As New ClassFornecedor
Dim objNFsPag As New ClassNFsPag

On Error GoTo Erro_BotaoNFRegistrarDif_Click

    dValor = StrParaDbl(DiferencaValor(objCT).Caption)
    If dValor < DELTA_VALORMONETARIO Then gError 202152
    
    objForn.sNomeReduzido = objCT.Fornecedor.Text
    
    lErro = CF("Fornecedor_Le_NomeReduzido", objForn)
    If lErro <> SUCESSO And lErro <> 6681 Then gError 202153

    'Não achou o Fornecedor --> erro
    If lErro <> SUCESSO Then gError 202154

    vbMsgResult = Rotina_Aviso(vbYesNo, "AVISO_CRIACAO_NFPAG", objCT.Fornecedor.Text, DiferencaValor(objCT).Caption)
    If vbMsgResult = vbNo Then gError 202155
    
    lErro = CF("TRVFatPag_Cadastra_NF", objForn.lCodigo, dValor, objNFsPag)
    If lErro <> SUCESSO Then gError 202156

    objCT.objGridNFiscais.iLinhasExistentes = objCT.objGridNFiscais.iLinhasExistentes + 1

    objCT.GridNF.TextMatrix(objCT.objGridNFiscais.iLinhasExistentes, objCT.iGrid_FilialNF_Col) = CStr(objNFsPag.iFilialEmpresa)
    objCT.GridNF.TextMatrix(objCT.objGridNFiscais.iLinhasExistentes, objCT.iGrid_Numero_Col) = CStr(objNFsPag.lNumNotaFiscal)
    objCT.GridNF.TextMatrix(objCT.objGridNFiscais.iLinhasExistentes, objCT.iGrid_Emissao_Col) = Format(objNFsPag.dtDataEmissao, "dd/mm/yyyy")
    objCT.GridNF.TextMatrix(objCT.objGridNFiscais.iLinhasExistentes, objCT.iGrid_FilialFornecedor_Col) = CStr(objNFsPag.iFilial)
    objCT.GridNF.TextMatrix(objCT.objGridNFiscais.iLinhasExistentes, objCT.iGrid_ValorNF_Col) = Format(objNFsPag.dValorTotal - objNFsPag.dValorIRRF, "Standard")
    objCT.GridNF.TextMatrix(objCT.objGridNFiscais.iLinhasExistentes, objCT.iGrid_Selecionada_Col) = "1"

    lErro = FATPAG_Carrega_GridNF(objCT, objNFsPag, objCT.objGridNFiscais.iLinhasExistentes)
    If lErro <> SUCESSO Then gError 202147

    objCT.NumNFSelecionadas.Caption = CStr(StrParaInt(objCT.NumNFSelecionadas.Caption) + 1)
    objCT.ValorTotalNFSelecionadas.Caption = Format(StrParaDbl(objCT.ValorTotalNFSelecionadas.Caption) + objNFsPag.dValorTotal - objNFsPag.dValorIRRF, "Standard")

    Call Grid_Refresh_Checkbox(objCT.objGridNFiscais)
    
    Call FATPAG_Calcula_Dif(objCT)

    'Call objCT.BotaoAtualizarNFs_Click

    BotaoNFRegistrarDif_Click = SUCESSO
    
    Exit Function

Erro_BotaoNFRegistrarDif_Click:

    BotaoNFRegistrarDif_Click = gErr

    Select Case gErr
    
        Case 202152
            Call Rotina_Erro(vbOKOnly, "ERRO_SEM_VALOR_DIFERENCA", gErr)
            
        Case 202153, 202155, 202156

        Case 202154
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_CADASTRADO1", gErr, objForn.sNomeReduzido)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202157)

    End Select

    Exit Function
    
End Function

Public Function FP_Calcula_Mnemonico_Cust(objMnemonicoValor As ClassMnemonicoValor, objTitPag As ClassTituloPagar, ColParcelaPagar As ColParcelaPagar, colNFPag As ColNFsPag) As Long

Dim lErro As Long
Dim dValorTarifaOver As Double
Dim objNFsPag As ClassNFsPag
Dim objTitExp As ClassTRVTitulosExp
Dim objParcelaPagar As ClassParcelaPagar
Dim dValorTotalParc As Double, dValorTotalOrig As Double

On Error GoTo Erro_FP_Calcula_Mnemonico_Cust

    Select Case UCase(objMnemonicoValor.sMnemonico)

        Case "VALOR_TARIFA"
        
            For Each objNFsPag In colNFPag
                
                lErro = CF("NFPag_Le_Numero", objNFsPag)
                If lErro <> SUCESSO And lErro <> 18338 Then gError ERRO_SEM_MENSAGEM
                
                Set objTitExp = New ClassTRVTitulosExp
                
                objTitExp.iTipoDocOrigem = TRV_TIPO_DOC_DESTINO_NFSPAG
                objTitExp.lNumIntDocOrigem = objNFsPag.lNumIntDoc
                objTitExp.iExcluido = DESMARCADO
            
                lErro = CF("TRVTitulosExp_Le", objTitExp)
                If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
                
                dValorTarifaOver = dValorTarifaOver + objTitExp.dValorTarifa
                
            Next
            
            objMnemonicoValor.colValor.Add dValorTarifaOver
            
        Case "VALOR_DIF"
            
            For Each objParcelaPagar In ColParcelaPagar
                dValorTotalParc = dValorTotalParc + objParcelaPagar.dValor
                dValorTotalOrig = dValorTotalOrig + objParcelaPagar.dValorOriginal
            Next
            
            objMnemonicoValor.colValor.Add dValorTotalParc - dValorTotalOrig
            
        Case Else
            gError ERRO_SEM_MENSAGEM

    End Select

    FP_Calcula_Mnemonico_Cust = SUCESSO

    Exit Function

Erro_FP_Calcula_Mnemonico_Cust:

    FP_Calcula_Mnemonico_Cust = gErr

    Select Case gErr
       
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 187164)

    End Select

    Exit Function

End Function
