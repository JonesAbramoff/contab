VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTClientesTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim gobjCT As Object

Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoRepresentante As AdmEvento
Attribute objEventoRepresentante.VB_VarHelpID = -1
Private WithEvents objEventoCorrentista As AdmEvento
Attribute objEventoCorrentista.VB_VarHelpID = -1
Private WithEvents objEventoEmpresaPai As AdmEvento
Attribute objEventoEmpresaPai.VB_VarHelpID = -1
Private WithEvents objEventoAcordo As AdmEvento
Attribute objEventoAcordo.VB_VarHelpID = -1
Private WithEvents objEventoCondicaoPagtoCC As AdmEvento
Attribute objEventoCondicaoPagtoCC.VB_VarHelpID = -1
Private WithEvents objEventoExcRep As AdmEvento
Attribute objEventoExcRep.VB_VarHelpID = -1
Private WithEvents objEventoExcCor As AdmEvento
Attribute objEventoExcCor.VB_VarHelpID = -1
Private WithEvents objEventoExcRepProduto As AdmEvento
Attribute objEventoExcRepProduto.VB_VarHelpID = -1
Private WithEvents objEventoExcCorProduto As AdmEvento
Attribute objEventoExcCorProduto.VB_VarHelpID = -1
Private WithEvents objEventoExcAgProduto As AdmEvento
Attribute objEventoExcAgProduto.VB_VarHelpID = -1

'Public Property Get Email2(ByVal objCT As Object, ByVal iIndice As Integer) As Object
'     Set Email2 = objCT.objUserControl.Controls("Email2")(iIndice)
'End Property

Public Property Get PercFatorDevCMCC(ByVal objCT As Object) As Object
     Set PercFatorDevCMCC = objCT.objUserControl.Controls("PercFatorDevCMCC")
End Property

Public Property Get GridComissao(ByVal objCT As Object) As Object
     Set GridComissao = objCT.objUserControl.Controls("GridComissao")
End Property

Public Property Get GridExcecoes(ByVal objCT As Object) As Object
     Set GridExcecoes = objCT.objUserControl.Controls("GridExcecoes")
End Property

Public Property Get Emissor(ByVal objCT As Object) As Object
     Set Emissor = objCT.objUserControl.Controls("Emissor")
End Property

Public Property Get PercComiss(ByVal objCT As Object) As Object
     Set PercComiss = objCT.objUserControl.Controls("PercComiss")
End Property

Public Property Get Produto(ByVal objCT As Object) As Object
     Set Produto = objCT.objUserControl.Controls("Produto")
End Property

Public Property Get DescricaoProduto(ByVal objCT As Object) As Object
     Set DescricaoProduto = objCT.objUserControl.Controls("DescricaoProduto")
End Property

Public Property Get PercComissProd(ByVal objCT As Object) As Object
     Set PercComissProd = objCT.objUserControl.Controls("PercComissProd")
End Property

Public Property Get Frame12(ByVal objCT As Object) As Object
     Set Frame12 = objCT.objUserControl.Controls("Frame12")
End Property

Public Property Get Representante(ByVal objCT As Object) As Object
     Set Representante = objCT.objUserControl.Controls("Representante")
End Property

Public Property Get Correntista(ByVal objCT As Object) As Object
     Set Correntista = objCT.objUserControl.Controls("Correntista")
End Property

Public Property Get PercComiRep(ByVal objCT As Object) As Object
     Set PercComiRep = objCT.objUserControl.Controls("PercComiRep")
End Property

Public Property Get PercComiCorr(ByVal objCT As Object) As Object
     Set PercComiCorr = objCT.objUserControl.Controls("PercComiCorr")
End Property

Public Property Get PercComiAg(ByVal objCT As Object) As Object
     Set PercComiAg = objCT.objUserControl.Controls("PercComiAg")
End Property

Public Property Get ConsiderarAporte(ByVal objCT As Object) As Object
     Set ConsiderarAporte = objCT.objUserControl.Controls("ConsiderarAporte")
End Property

Public Property Get EmpresaPai(ByVal objCT As Object) As Object
     Set EmpresaPai = objCT.objUserControl.Controls("EmpresaPai")
End Property

Public Property Get Codigo(ByVal objCT As Object) As Object
     Set Codigo = objCT.objUserControl.Controls("Codigo")
End Property

Public Property Get FilialEmpresa(ByVal objCT As Object) As Object
     Set FilialEmpresa = objCT.objUserControl.Controls("FilialEmpresa")
End Property

Public Property Get FilialEmpresaNF(ByVal objCT As Object) As Object
     Set FilialEmpresaNF = objCT.objUserControl.Controls("FilialEmpresaNF")
End Property

Public Property Get FilialEmpresaFAT(ByVal objCT As Object) As Object
     Set FilialEmpresaFAT = objCT.objUserControl.Controls("FilialEmpresaFAT")
End Property

Public Property Get Grupo(ByVal objCT As Object) As Object
     Set Grupo = objCT.objUserControl.Controls("Grupo")
End Property

Public Property Get GridExcRep(ByVal objCT As Object) As Object
     Set GridExcRep = objCT.objUserControl.Controls("GridExcRep")
End Property

Public Property Get ExcRep(ByVal objCT As Object) As Object
     Set ExcRep = objCT.objUserControl.Controls("ExcRep")
End Property

Public Property Get ExcRepProduto(ByVal objCT As Object) As Object
     Set ExcRepProduto = objCT.objUserControl.Controls("ExcRepProduto")
End Property

Public Property Get ExcRepPercComis(ByVal objCT As Object) As Object
     Set ExcRepPercComis = objCT.objUserControl.Controls("ExcRepPercComis")
End Property

Public Property Get GridExcCor(ByVal objCT As Object) As Object
     Set GridExcCor = objCT.objUserControl.Controls("GridExcCor")
End Property

Public Property Get ExcCor(ByVal objCT As Object) As Object
     Set ExcCor = objCT.objUserControl.Controls("ExcCor")
End Property

Public Property Get ExcCorProduto(ByVal objCT As Object) As Object
     Set ExcCorProduto = objCT.objUserControl.Controls("ExcCorProduto")
End Property

Public Property Get ExcCorPercComis(ByVal objCT As Object) As Object
     Set ExcCorPercComis = objCT.objUserControl.Controls("ExcCorPercComis")
End Property

Public Property Get GridExcAg(ByVal objCT As Object) As Object
     Set GridExcAg = objCT.objUserControl.Controls("GridExcAg")
End Property

Public Property Get ExcAgProduto(ByVal objCT As Object) As Object
     Set ExcAgProduto = objCT.objUserControl.Controls("ExcAgProduto")
End Property

Public Property Get ExcAgPercComis(ByVal objCT As Object) As Object
     Set ExcAgPercComis = objCT.objUserControl.Controls("ExcAgPercComis")
End Property

Public Property Get CondicaoPagtoCC(ByVal objCT As Object) As Object
     Set CondicaoPagtoCC = objCT.objUserControl.Controls("CondicaoPagtoCC")
End Property

Public Property Get GridCallCenter(ByVal objCT As Object) As Object
     Set GridCallCenter = objCT.objUserControl.Controls("GridCallCenter")
End Property

Public Property Get DataCallCenterDe(ByVal objCT As Object) As Object
     Set DataCallCenterDe = objCT.objUserControl.Controls("DataCallCenterDe")
End Property

Public Property Get DataCallCenterAte(ByVal objCT As Object) As Object
     Set DataCallCenterAte = objCT.objUserControl.Controls("DataCallCenterAte")
End Property

Public Property Get EmiPercCI(ByVal objCT As Object) As Object
     Set EmiPercCI = objCT.objUserControl.Controls("EmiPercCI")
End Property

Public Property Get EmiCargo(ByVal objCT As Object) As Object
     Set EmiCargo = objCT.objUserControl.Controls("EmiCargo")
End Property

Public Property Get EmiCartao(ByVal objCT As Object) As Object
     Set EmiCartao = objCT.objUserControl.Controls("EmiCartao")
End Property

Public Property Get EmiCPF(ByVal objCT As Object) As Object
     Set EmiCPF = objCT.objUserControl.Controls("EmiCPF")
End Property

Public Function CLI_Grava_Registro1(ByVal objCT As Object, objCliente As ClassCliente, colEnderecos As Collection) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objEndereco As ClassEndereco
Dim objEnderecoTRV As ClassEnderecoTRV
Dim objCliEmissoresTRV As ClassCliEmissoresTRV
Dim objClienteTRV As New ClassClienteTRV
Dim objCliente1 As New ClassCliente
Dim objTRVCliRep As ClassTRVClienteRepProd
Dim objTRVCliCor As ClassTRVClienteCorProd
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objExcComissaoCli As ClassTRVExcComissaoCli
Dim objDataCall As ClassTRVCliDataCallCenter

On Error GoTo Erro_CLI_Grava_Registro1

    'Verifica se no grid de comissao as informacoes estao preenchidas
    For iIndice = 1 To objCT.gobjInfoUsu.objGridComissao.iLinhasExistentes
        
        If Len(Trim(GridComissao(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Emissor_Col))) = 0 Then gError 195805
        If Len(Trim(GridComissao(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PercComiss_Col))) = 0 Then gError 195806
         
        Set objCliEmissoresTRV = objCT.gobjInfoUsu.gcolCliEmissoresTRV.Item(iIndice)
        
        objCliEmissoresTRV.lFornEmissor = LCodigo_Extrai(GridComissao(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Emissor_Col))
        objCliEmissoresTRV.dPercComissao = PercentParaDbl(GridComissao(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PercComiss_Col))
        objCliEmissoresTRV.lCargo = LCodigo_Extrai(GridComissao(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_EmiCargo_Col))
        objCliEmissoresTRV.sCPF = GridComissao(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_EmiCPF_Col)
        objCliEmissoresTRV.sNumCartao = GridComissao(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_EmiCartao_Col)
        objCliEmissoresTRV.dPercCI = PercentParaDbl(GridComissao(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_EmiPercCI_Col))
        
        
        objCliEmissoresTRV.sCPF = Replace(Replace(Replace(Replace(objCliEmissoresTRV.sCPF, ".", ""), "-", ""), "/", ""), "\", "")
    Next
    
    lErro = Move_CliEmissoresExcTRV(objCT, GridComissao(objCT).Row)
    If lErro <> SUCESSO Then gError 195809

    Set objClienteTRV.colCliEmissoresTRV = objCT.gobjInfoUsu.gcolCliEmissoresTRV

    objClienteTRV.lCodRepresentante = LCodigo_Extrai(Representante(objCT).Text)

    If Len(Trim(PercComiRep(objCT).Text)) > 0 Then
    
        If Len(Trim(Representante(objCT).Text)) = 0 Then gError 195863
        
        objClienteTRV.dPercComiRep = CDbl(PercComiRep(objCT).Text) / 100
        
    End If

    objClienteTRV.lCodCorrentista = LCodigo_Extrai(Correntista(objCT).Text)

    If Len(Trim(PercComiCorr(objCT).Text)) > 0 Then
    
        If Len(Trim(Correntista(objCT).Text)) = 0 Then gError 195864

        objClienteTRV.dPercComiCorr = CDbl(PercComiCorr(objCT).Text) / 100

    End If
    
    If Len(Trim(PercComiAg(objCT).Text)) > 0 Then objClienteTRV.dPercComiAg = CDbl(PercComiAg(objCT).Text) / 100

    objClienteTRV.iConsiderarAporte = ConsiderarAporte(objCT).Value

    Set objCliente.objInfoUsu = objClienteTRV
    
    objClienteTRV.lEmpresaPai = LCodigo_Extrai(EmpresaPai(objCT).Text)
    
    objClienteTRV.iFilialNF = Codigo_Extrai(FilialEmpresaNF(objCT).Text)
    objClienteTRV.iFilialEmpresa = Codigo_Extrai(FilialEmpresa(objCT).Text)
    objClienteTRV.iFilialFat = Codigo_Extrai(FilialEmpresaFAT(objCT).Text)
    objClienteTRV.sGrupo = Grupo(objCT).Text
    objClienteTRV.iCondPagtoCC = Codigo_Extrai(CondicaoPagtoCC(objCT).Text)
    
    'Verifica se no grid de comissao as informacoes estao preenchidas
    For iIndice = 1 To objCT.gobjInfoUsu.objGridExcRep.iLinhasExistentes
        
        Set objTRVCliRep = New ClassTRVClienteRepProd
               
        lErro = CF("Produto_Formata", GridExcRep(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ExcRepProduto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 195801
        
        objTRVCliRep.iSeq = iIndice
        objTRVCliRep.sProduto = sProdutoFormatado
        'objTRVCliRep.dPercComis = PercentParaDbl(GridExcRep(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ExcRepPercComiss_Col))
        objTRVCliRep.dPercComis = StrParaDbl(GridExcRep(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ExcRepPercComiss_Col)) / 100
        objTRVCliRep.lRepresentante = LCodigo_Extrai(GridExcRep(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ExcRep_Col))
        
        If objTRVCliRep.lRepresentante = 0 Or Len(Trim(objTRVCliRep.sProduto)) = 0 Then gError 213791
    
        objClienteTRV.colExcRep.Add objTRVCliRep
        
    Next
    
    'Verifica se no grid de comissao as informacoes estao preenchidas
    For iIndice = 1 To objCT.gobjInfoUsu.objGridExcCor.iLinhasExistentes
        
        Set objTRVCliCor = New ClassTRVClienteCorProd
        
        lErro = CF("Produto_Formata", GridExcCor(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ExcCorProduto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 195801
        
        objTRVCliCor.iSeq = iIndice
        objTRVCliCor.sProduto = sProdutoFormatado
        'objTRVCliCor.dPercComis = PercentParaDbl(GridExcCor(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ExcCorPercComiss_Col))
        objTRVCliCor.dPercComis = StrParaDbl(GridExcCor(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ExcCorPercComiss_Col)) / 100
        objTRVCliCor.lCorrentista = LCodigo_Extrai(GridExcCor(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ExcCor_Col))
    
        If objTRVCliCor.lCorrentista = 0 Or Len(Trim(objTRVCliCor.sProduto)) = 0 Then gError 213792
    
        objClienteTRV.colExcCor.Add objTRVCliCor
        
    Next
    
    'Verifica se no grid de comissao as informacoes estao preenchidas
    For iIndice = 1 To objCT.gobjInfoUsu.objGridExcAg.iLinhasExistentes
        
        Set objExcComissaoCli = New ClassTRVExcComissaoCli
        
        lErro = CF("Produto_Formata", GridExcAg(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ExcAgProduto_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 195801
        
        objExcComissaoCli.iSeq = iIndice
        objExcComissaoCli.sProduto = sProdutoFormatado
        'objExcComissaoCli.dPercComissao = PercentParaDbl(GridExcAg(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ExcAgPercComiss_Col))
        objExcComissaoCli.dPercComissao = StrParaDbl(GridExcAg(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ExcAgPercComiss_Col)) / 100
    
        If Len(Trim(objExcComissaoCli.sProduto)) = 0 Then gError 213793
    
        objClienteTRV.colExcComissaoCli.Add objExcComissaoCli
        
    Next
    
    For iIndice = 1 To objCT.gobjInfoUsu.objGridCallCenter.iLinhasExistentes
        Set objDataCall = New ClassTRVCliDataCallCenter
        objDataCall.dtDataDe = StrParaDate(GridCallCenter(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_DataCallCenterDe_Col))
        objDataCall.dtDataAte = StrParaDate(GridCallCenter(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_DataCallCenterAte_Col))
        
        If objDataCall.dtDataDe <> DATA_NULA And objDataCall.dtDataAte <> DATA_NULA Then
            If objDataCall.dtDataDe > objDataCall.dtDataAte Then gError 202708
        End If
        
        If objDataCall.dtDataDe <> DATA_NULA Or objDataCall.dtDataAte <> DATA_NULA Then
            objClienteTRV.colDataCallCenter.Add objDataCall
        End If
    Next
    
    If Len(Trim(PercFatorDevCMCC(objCT).Text)) > 0 Then

        objClienteTRV.dPercFatorDevCMCC = CDbl(PercFatorDevCMCC(objCT).Text) / 100

    End If

    CLI_Grava_Registro1 = SUCESSO
    
    Exit Function
    
Erro_CLI_Grava_Registro1:
    
    CLI_Grava_Registro1 = gErr
    
    Select Case gErr
    
        Case 195805
            Call Rotina_Erro(vbOKOnly, "ERRO_EMISSOR_NAO_PREENCHIDO_GRID_EMISSOR", gErr, iIndice)
        
        Case 195806
            Call Rotina_Erro(vbOKOnly, "ERRO_PERCCOMISS_NAO_PREENCHIDO_GRID_EMISSOR", gErr, iIndice)
        
        Case 195809, 195855, 195857, 197004
        
        Case 195856
            Call Rotina_Erro(vbOKOnly, "ERRO_REPRESENTANTE_NAO_CADASTRADO", gErr, objCliente1.sNomeReduzido)
        
        Case 195858
            Call Rotina_Erro(vbOKOnly, "ERRO_CORRENTISTA_NAO_CADASTRADO", gErr, objCliente1.sNomeReduzido)
        
        Case 195863
            Call Rotina_Erro(vbOKOnly, "ERRO_REPRESENTANTE_NAO_INFORMADO", gErr)
        
        Case 195864
            Call Rotina_Erro(vbOKOnly, "ERRO_CORRENTISTA_NAO_INFORMADO", gErr)
        
        Case 197005
            Call Rotina_Erro(vbOKOnly, "ERRO_EMPRESAPAI_NAO_CADASTRADO", gErr, objCliente1.sNomeReduzido)
        
        Case 197007
            Call Rotina_Erro(vbOKOnly, "ERRO_EMPRESAPAI_MESMO_FILHO", gErr)
            
        Case 202708
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_DE_MENOR_GRID", gErr, iIndice)
        
        Case 213791
            Call Rotina_Erro(vbOKOnly, "Informações incompletas na linha " & CStr(iIndice) & " do grid de exceções por representante", gErr, iIndice)
        
        Case 213792
            Call Rotina_Erro(vbOKOnly, "Informações incompletas na linha " & CStr(iIndice) & " do grid de exceções por correntista", gErr, iIndice)
        
        Case 213793
            Call Rotina_Erro(vbOKOnly, "Informações incompletas na linha " & CStr(iIndice) & " do grid de exceções por agência", gErr, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 196229)

    End Select

    Exit Function

End Function

Public Function CLI_Exibe_Dados_Cliente_Enderecos(ByVal objCT As Object, objCliente As ClassCliente, colEnderecos) As Long
    CLI_Exibe_Dados_Cliente_Enderecos = SUCESSO
End Function

Public Function CLI_Form_Load(ByVal objCT As Object) As Long

Dim lErro As Long
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As AdmCodigoNome

On Error GoTo Erro_CLI_Form_Load

    Set objCT.gobjInfoUsu.gcolCliEmissoresTRV = New Collection

    Set objCT.gobjInfoUsu.objGridComissao = New AdmGrid
    Set objCT.gobjInfoUsu.objGridExcecoes = New AdmGrid
    Set objCT.gobjInfoUsu.objGridExcRep = New AdmGrid
    Set objCT.gobjInfoUsu.objGridExcCor = New AdmGrid
    Set objCT.gobjInfoUsu.objGridExcAg = New AdmGrid
    Set objCT.gobjInfoUsu.objGridCallCenter = New AdmGrid

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto(objCT))
    If lErro <> SUCESSO Then gError 195842

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ExcRepProduto(objCT))
    If lErro <> SUCESSO Then gError 195842

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ExcCorProduto(objCT))
    If lErro <> SUCESSO Then gError 195842

    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ExcAgProduto(objCT))
    If lErro <> SUCESSO Then gError 195842

    lErro = Inicializa_Grid_Comissao(objCT, objCT.gobjInfoUsu.objGridComissao)
    If lErro <> SUCESSO Then gError 192973

    lErro = Inicializa_Grid_Excecoes(objCT, objCT.gobjInfoUsu.objGridExcecoes)
    If lErro <> SUCESSO Then gError 192974
    
    lErro = Inicializa_Grid_ExcRep(objCT, objCT.gobjInfoUsu.objGridExcRep)
    If lErro <> SUCESSO Then gError 192974
    
    lErro = Inicializa_Grid_ExcCor(objCT, objCT.gobjInfoUsu.objGridExcCor)
    If lErro <> SUCESSO Then gError 192974

    lErro = Inicializa_Grid_ExcAg(objCT, objCT.gobjInfoUsu.objGridExcAg)
    If lErro <> SUCESSO Then gError 192974

    lErro = Inicializa_Grid_CallCenter(objCT, objCT.gobjInfoUsu.objGridCallCenter)
    If lErro <> SUCESSO Then gError 192974

    'Lê cada codigo e descricao reduzida da tabela CondicoesPagto
    lErro = CF("CondicoesPagto_Le_Recebimento", colCodigoDescricao)
    If lErro <> SUCESSO Then gError 192974

    'Preenche a ComboBox CondicaoPagto com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        CondicaoPagtoCC(objCT).AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        CondicaoPagtoCC(objCT).ItemData(CondicaoPagtoCC(objCT).NewIndex) = objCodigoDescricao.iCodigo
    Next
    
    'Carrega a combo de filial
    lErro = Carrega_FilialEmpresa(objCT, FilialEmpresa(objCT))
    If lErro <> SUCESSO Then gError 192974

    lErro = Carrega_FilialEmpresa(objCT, FilialEmpresaNF(objCT))
    If lErro <> SUCESSO Then gError 192974

    lErro = Carrega_FilialEmpresa(objCT, FilialEmpresaFAT(objCT))
    If lErro <> SUCESSO Then gError 192974
    
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TRVEMICARGO, EmiCargo(objCT))
    If lErro <> SUCESSO Then gError 192974

    CLI_Form_Load = SUCESSO
     
    Exit Function
    
Erro_CLI_Form_Load:

    CLI_Form_Load = gErr
     
    Select Case gErr
    
        Case 192973 To 192975, 195842

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192976)
     
    End Select
     
    Exit Function

End Function

Private Function Inicializa_Grid_Comissao(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Emissor")
    objGridInt.colColuna.Add ("% C.D.")
    objGridInt.colColuna.Add ("% C.I.")
    objGridInt.colColuna.Add ("Cargo")
    objGridInt.colColuna.Add ("No Cartão")
    objGridInt.colColuna.Add ("CPF")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Emissor(objCT).Name)
    objGridInt.colCampo.Add (PercComiss(objCT).Name)
    objGridInt.colCampo.Add (EmiPercCI(objCT).Name)
    objGridInt.colCampo.Add (EmiCargo(objCT).Name)
    objGridInt.colCampo.Add (EmiCartao(objCT).Name)
    objGridInt.colCampo.Add (EmiCPF(objCT).Name)

    'Colunas do Grid
    objCT.gobjInfoUsu.iGrid_Emissor_Col = 1
    objCT.gobjInfoUsu.iGrid_PercComiss_Col = 2
    objCT.gobjInfoUsu.iGrid_EmiPercCI_Col = 3
    objCT.gobjInfoUsu.iGrid_EmiCargo_Col = 4
    objCT.gobjInfoUsu.iGrid_EmiCartao_Col = 5
    objCT.gobjInfoUsu.iGrid_EmiCPF_Col = 6

    'Grid do GridInterno
    objGridInt.objGrid = GridComissao(objCT)

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 1001

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridComissao(objCT).ColWidth(0) = 300

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Comissao = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_CallCenter(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("De")
    objGridInt.colColuna.Add ("Até")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (DataCallCenterDe(objCT).Name)
    objGridInt.colCampo.Add (DataCallCenterAte(objCT).Name)

    'Colunas do Grid
    objCT.gobjInfoUsu.iGrid_DataCallCenterDe_Col = 1
    objCT.gobjInfoUsu.iGrid_DataCallCenterAte_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridCallCenter(objCT)

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 11

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 2

    'Largura da primeira coluna
    GridCallCenter(objCT).ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_CallCenter = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_Excecoes(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("%Comissão")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto(objCT).Name)
    objGridInt.colCampo.Add (DescricaoProduto(objCT).Name)
    objGridInt.colCampo.Add (PercComissProd(objCT).Name)

    'Colunas do Grid
    objCT.gobjInfoUsu.iGrid_Produto_Col = 1
    objCT.gobjInfoUsu.iGrid_DescricaoProduto_Col = 2
    objCT.gobjInfoUsu.iGrid_PercComissProd_Col = 3

    'Grid do GridInterno
    objGridInt.objGrid = GridExcecoes(objCT)

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridExcecoes(objCT).ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Excecoes = SUCESSO

    Exit Function

End Function

Public Function CLI_Saida_Celula(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_CLI_Saida_Celula

    'Verifica qual o Grid em questão
    Select Case objGridInt.objGrid.Name

        'Se for o GridComissao
        Case GridComissao(objCT).Name

            lErro = Saida_Celula_GridComissao(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195770

        'Se for o GridExcecoes
        Case GridExcecoes(objCT).Name

            lErro = Saida_Celula_GridExcecoes(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195771

        'Se for o GridExcRep
        Case GridExcRep(objCT).Name

            lErro = Saida_Celula_GridExcRep(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195771

        'Se for o GridExcCor
        Case GridExcCor(objCT).Name

            lErro = Saida_Celula_GridExcCor(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195771

        'Se for o GridExcAg
        Case GridExcAg(objCT).Name

            lErro = Saida_Celula_GridExcAg(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195771

        'Se for o GridCallCenter
        Case GridCallCenter(objCT).Name

            lErro = Saida_Celula_GridCallCenter(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195771

    End Select
    
    CLI_Saida_Celula = SUCESSO

    Exit Function

Erro_CLI_Saida_Celula:

    CLI_Saida_Celula = gErr

    Select Case gErr

        Case 195770, 195771

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195772)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridComissao(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridComissao

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for o emissor
        Case objCT.gobjInfoUsu.iGrid_Emissor_Col
            lErro = Saida_Celula_Emissor(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195760

        'Se for o % de comissão direta
        Case objCT.gobjInfoUsu.iGrid_PercComiss_Col
            lErro = Saida_Celula_PercComiss(objCT, objGridInt, PercComiss(objCT))
            If lErro <> SUCESSO Then gError 195761

        'Se for o % de comissão indireta
        Case objCT.gobjInfoUsu.iGrid_EmiPercCI_Col
            lErro = Saida_Celula_PercComiss(objCT, objGridInt, EmiPercCI(objCT))
            If lErro <> SUCESSO Then gError 195761

        Case objCT.gobjInfoUsu.iGrid_EmiCPF_Col
            lErro = Saida_Celula_CPF(objGridInt, EmiCPF(objCT))
            If lErro <> SUCESSO Then gError 195761

        Case objCT.gobjInfoUsu.iGrid_EmiCargo_Col
            lErro = Saida_Celula_Padrao(objGridInt, EmiCargo(objCT))
            If lErro <> SUCESSO Then gError 195761

        Case objCT.gobjInfoUsu.iGrid_EmiCartao_Col
            lErro = Saida_Celula_Padrao(objGridInt, EmiCartao(objCT))
            If lErro <> SUCESSO Then gError 195761

    End Select

    Saida_Celula_GridComissao = SUCESSO

    Exit Function

Erro_Saida_Celula_GridComissao:

    Saida_Celula_GridComissao = gErr

    Select Case gErr

        Case 195760, 195761

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195762)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Emissor(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da célula Emissor que está deixando de ser a corrente

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iIndice As Integer
Dim objCliEmissoresTRV As ClassCliEmissoresTRV

On Error GoTo Erro_Saida_Celula_Emissor

    Set objGridInt.objControle = Emissor(objCT)

    If Len(Trim(Emissor(objCT).Text)) > 0 Then

        'Tenta ler o Fornecedor (NomeReduzido ou Código)
        lErro = TP_Fornecedor_Le2(Emissor(objCT), objFornecedor, 0)
        If lErro <> SUCESSO Then gError 195763

        'Verifica se já está em outra linha do Grid
        For iIndice = 1 To objGridInt.iLinhasExistentes
            If iIndice <> GridComissao(objCT).Row Then
                If GridComissao(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Emissor_Col) = Emissor(objCT).Text Then gError 195773
            End If
        Next
    
        If GridComissao(objCT).Row - GridComissao(objCT).FixedRows = objGridInt.iLinhasExistentes Then
    
            Frame12(objCT).Caption = "Exceções por produto (Emissor: " & left(Emissor(objCT).Text, 22) & ")"
    
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    
            Set objCliEmissoresTRV = New ClassCliEmissoresTRV
            objCT.gobjInfoUsu.gcolCliEmissoresTRV.Add objCliEmissoresTRV
    
        End If

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195764

    Saida_Celula_Emissor = SUCESSO

    Exit Function

Erro_Saida_Celula_Emissor:

    Saida_Celula_Emissor = gErr

    Select Case gErr

        Case 195763, 195764
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 195773
            Call Rotina_Erro(vbOKOnly, "ERRO_EMISSOR_JA_EXISTE_GRID", gErr, Emissor(objCT).Text, iIndice)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195765)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_PercComiss(ByVal objCT As Object, objGridInt As AdmGrid, objControle As Object) As Long
'Faz a crítica da célula Percentual Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double
Dim objCliEmissoresTRV As ClassCliEmissoresTRV

On Error GoTo Erro_Saida_Celula_PercComiss

    Set objGridInt.objControle = objControle

    If Len(objControle.Text) > 0 Then
        
        'Critica a porcentagem
        lErro = Porcentagem_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 195766

        dPercentDesc = CDbl(objControle.Text)
        
        objControle.Text = Format(dPercentDesc, "Fixed")
        
        If GridComissao(objCT).Row - GridComissao(objCT).FixedRows = objGridInt.iLinhasExistentes Then
    
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
            
            Set objCliEmissoresTRV = New ClassCliEmissoresTRV
            objCT.gobjInfoUsu.gcolCliEmissoresTRV.Add objCliEmissoresTRV
    
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195767

    Saida_Celula_PercComiss = SUCESSO

    Exit Function

Erro_Saida_Celula_PercComiss:

    Saida_Celula_PercComiss = gErr

    Select Case gErr

        Case 195766, 195767
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195768)

    End Select

    Exit Function

End Function

Public Function CLI_Traz_Cliente_Tela(ByVal objCT As Object, ByVal objCliente As ClassCliente) As Long

Dim lErro As Long
Dim objClienteTRV As ClassClienteTRV
Dim objCliEmissoresTRV As ClassCliEmissoresTRV
Dim iIndice As Integer
Dim objFornecedor As New ClassFornecedor
Dim objCliente1 As New ClassCliente
Dim iCodFilial As Integer
Dim objTRVCliRep As ClassTRVClienteRepProd
Dim objTRVCliCor As ClassTRVClienteCorProd
Dim sProdutoEnxuto As String
Dim objExcComissaoCli As ClassTRVExcComissaoCli
Dim objDataCall As ClassTRVCliDataCallCenter

On Error GoTo Erro_CLI_Traz_Cliente_Tela

    lErro = CF("Cliente_Le_Customizado", objCliente, True)
    If lErro <> SUCESSO Then gError 195784
    
    If Not (objCliente.objInfoUsu Is Nothing) Then
    
        Set objClienteTRV = objCliente.objInfoUsu
        
        Set objCT.gobjInfoUsu.gcolCliEmissoresTRV = objClienteTRV.colCliEmissoresTRV
        
        Call Grid_Limpa(objCT.gobjInfoUsu.objGridComissao)

        For Each objCliEmissoresTRV In objClienteTRV.colCliEmissoresTRV
        
            iIndice = iIndice + 1
        
            Emissor(objCT).Text = objCliEmissoresTRV.lFornEmissor
        
            'Tenta ler o Fornecedor (NomeReduzido ou Código)
            lErro = TP_Fornecedor_Le2(Emissor(objCT), objFornecedor, 0)
            If lErro <> SUCESSO Then gError 195785
        
            GridComissao(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Emissor_Col) = Emissor(objCT).Text
            GridComissao(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PercComiss_Col) = Format(objCliEmissoresTRV.dPercComissao, "Percent")
            GridComissao(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_EmiPercCI_Col) = Format(objCliEmissoresTRV.dPercCI, "Percent")
            GridComissao(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_EmiCartao_Col) = objCliEmissoresTRV.sNumCartao
        
            If Len(Trim(objCliEmissoresTRV.sCPF)) > 0 Then
                If Len(Trim(objCliEmissoresTRV.sCPF)) = STRING_CPF Then
                    GridComissao(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_EmiCPF_Col) = Format(objCliEmissoresTRV.sCPF, "000\.000\.000-00; ; ; ")
                Else
                    GridComissao(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_EmiCPF_Col) = Format(objCliEmissoresTRV.sCPF, "00\.000\.000\/0000-00; ; ; ")
                End If
            End If
            
            Call Combo_Seleciona_ItemData(EmiCargo(objCT), objCliEmissoresTRV.lCargo)
            GridComissao(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_EmiCargo_Col) = EmiCargo(objCT).Text
            
        Next

        'Atualiza o número de linhas existentes
        objCT.gobjInfoUsu.objGridComissao.iLinhasExistentes = iIndice
        
        objCT.gobjInfoUsu.giLinhaAntiga = 0
        
        If objClienteTRV.lCodRepresentante <> 0 Then
        
            Representante(objCT).Text = objClienteTRV.lCodRepresentante
        
            iCodFilial = FILIAL_MATRIZ

            'Tenta ler o Representante (NomeReduzido ou Código)
            lErro = TP_Cliente_Le2(Representante(objCT), objCliente1, iCodFilial)
            If lErro <> SUCESSO Then gError 195865
        
            PercComiRep(objCT).Text = CStr(100 * objClienteTRV.dPercComiRep)
            
        End If
        
        If objClienteTRV.lCodCorrentista <> 0 Then
        
            Correntista(objCT).Text = objClienteTRV.lCodCorrentista
        
            iCodFilial = FILIAL_MATRIZ
            
            Set objCliente1 = New ClassCliente

            'Tenta ler o Correntista (NomeReduzido ou Código)
            lErro = TP_Cliente_Le2(Correntista(objCT), objCliente1, iCodFilial)
            If lErro <> SUCESSO Then gError 195866
        
            PercComiCorr(objCT).Text = CStr(100 * objClienteTRV.dPercComiCorr)
            
        End If
        
        PercComiAg(objCT).Text = CStr(100 * objClienteTRV.dPercComiAg)
        
        ConsiderarAporte(objCT).Value = objClienteTRV.iConsiderarAporte
        
        If objClienteTRV.lEmpresaPai <> 0 Then
        
            EmpresaPai(objCT).Text = objClienteTRV.lEmpresaPai
        
            iCodFilial = FILIAL_MATRIZ

            'Tenta ler o Empresa Pai (NomeReduzido ou Código)
            lErro = TP_Cliente_Le2(EmpresaPai(objCT), objCliente1, iCodFilial)
            If lErro <> SUCESSO Then gError 197006
        
        End If
        
        If objClienteTRV.iFilialNF <> 0 Then
            FilialEmpresaNF(objCT).Text = objClienteTRV.iFilialNF
            Call FilialEmpresa_Validate(objCT, bSGECancelDummy, FilialEmpresaNF(objCT))
        Else
            FilialEmpresaNF(objCT).Text = ""
        End If
        
        If objClienteTRV.iFilialEmpresa <> 0 Then
            FilialEmpresa(objCT).Text = objClienteTRV.iFilialEmpresa
            Call FilialEmpresa_Validate(objCT, bSGECancelDummy, FilialEmpresa(objCT))
        Else
            FilialEmpresa(objCT).Text = ""
        End If
        
        If objClienteTRV.iFilialFat <> 0 Then
            FilialEmpresaFAT(objCT).Text = objClienteTRV.iFilialFat
            Call FilialEmpresa_Validate(objCT, bSGECancelDummy, FilialEmpresaFAT(objCT))
        Else
            FilialEmpresaFAT(objCT).Text = ""
        End If
        
        Grupo(objCT).PromptInclude = False
        Grupo(objCT).Text = objClienteTRV.sGrupo
        Grupo(objCT).PromptInclude = True
        
        If objClienteTRV.iCondPagtoCC = 0 Then
            CondicaoPagtoCC(objCT).Text = ""
        Else
            CondicaoPagtoCC(objCT).Text = CStr(objClienteTRV.iCondPagtoCC)
            Call CondicaoPagtoCC_Validate(objCT, bSGECancelDummy)
        End If
        
        iIndice = 0
        For Each objTRVCliRep In objClienteTRV.colExcRep
        
            iIndice = iIndice + 1
            
            Set objCliente1 = New ClassCliente
        
            ExcRep(objCT).Text = objTRVCliRep.lRepresentante

            'Tenta ler o Vendedor (NomeReduzido ou Código)
            lErro = TP_Cliente_Le2(ExcRep(objCT), objCliente1, iCodFilial)
            If lErro <> SUCESSO Then gError 197006
                    
            lErro = Mascara_RetornaProdutoEnxuto(objTRVCliRep.sProduto, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError 195802
    
            ExcRepProduto(objCT).PromptInclude = False
            ExcRepProduto(objCT).Text = sProdutoEnxuto
            ExcRepProduto(objCT).PromptInclude = True
        
            GridExcRep(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ExcRep_Col) = ExcRep(objCT).Text
            GridExcRep(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ExcRepPercComiss_Col) = Format(objTRVCliRep.dPercComis * 100, "0.00##") '"Percent")
            GridExcRep(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ExcRepProduto_Col) = ExcRepProduto(objCT)
        
        Next

        objCT.gobjInfoUsu.objGridExcRep.iLinhasExistentes = iIndice
        
        iIndice = 0
        For Each objTRVCliCor In objClienteTRV.colExcCor
        
            iIndice = iIndice + 1
            
            Set objCliente1 = New ClassCliente
        
            ExcCor(objCT).Text = objTRVCliCor.lCorrentista
            iCodFilial = FILIAL_MATRIZ

            'Tenta ler o Vendedor (NomeReduzido ou Código)
            lErro = TP_Cliente_Le2(ExcCor(objCT), objCliente1, iCodFilial)
            If lErro <> SUCESSO Then gError 197006
                    
            lErro = Mascara_RetornaProdutoEnxuto(objTRVCliCor.sProduto, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError 195802
    
            ExcCorProduto(objCT).PromptInclude = False
            ExcCorProduto(objCT).Text = sProdutoEnxuto
            ExcCorProduto(objCT).PromptInclude = True
        
            GridExcCor(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ExcCor_Col) = ExcCor(objCT).Text
            GridExcCor(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ExcCorPercComiss_Col) = Format(objTRVCliCor.dPercComis * 100, "0.00##") ' "Percent")
            GridExcCor(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ExcCorProduto_Col) = ExcCorProduto(objCT)
        
        Next
        
        objCT.gobjInfoUsu.objGridExcCor.iLinhasExistentes = iIndice

        iIndice = 0
        For Each objExcComissaoCli In objClienteTRV.colExcComissaoCli
        
            iIndice = iIndice + 1
                    
            lErro = Mascara_RetornaProdutoEnxuto(objExcComissaoCli.sProduto, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError 195802
    
            ExcAgProduto(objCT).PromptInclude = False
            ExcAgProduto(objCT).Text = sProdutoEnxuto
            ExcAgProduto(objCT).PromptInclude = True
        
            GridExcAg(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ExcAgPercComiss_Col) = Format(objExcComissaoCli.dPercComissao * 100, "0.00##") ' "Percent")
            GridExcAg(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_ExcAgProduto_Col) = ExcAgProduto(objCT)
        
        Next
        
        objCT.gobjInfoUsu.objGridExcAg.iLinhasExistentes = iIndice
        
        iIndice = 0
        Call Grid_Limpa(objCT.gobjInfoUsu.objGridCallCenter)
        For Each objDataCall In objClienteTRV.colDataCallCenter
        
            iIndice = iIndice + 1
        
            If objDataCall.dtDataDe <> DATA_NULA Then GridCallCenter(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_DataCallCenterDe_Col) = Format(objDataCall.dtDataDe, "dd/mm/yyyy")
            If objDataCall.dtDataAte <> DATA_NULA Then GridCallCenter(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_DataCallCenterAte_Col) = Format(objDataCall.dtDataAte, "dd/mm/yyyy")
        
        Next
        
        objCT.gobjInfoUsu.objGridCallCenter.iLinhasExistentes = iIndice
        
        PercFatorDevCMCC(objCT).Text = CStr(100 * objClienteTRV.dPercFatorDevCMCC)
        
    End If
    
    CLI_Traz_Cliente_Tela = SUCESSO
     
    Exit Function
    
Erro_CLI_Traz_Cliente_Tela:

    CLI_Traz_Cliente_Tela = gErr
     
    Select Case gErr
    
        Case 195784, 195785, 195865, 196866, 197006
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195786)
     
    End Select
     
    Exit Function

End Function

Public Sub GridComissao_Click(ByVal objCT As Object)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objCT.gobjInfoUsu.objGridComissao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridComissao, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridComissao_EnterCell(ByVal objCT As Object)
    Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridComissao, objCT.iAlterado)
End Sub

Public Sub GridComissao_GotFocus(ByVal objCT As Object)
    Call Grid_Recebe_Foco(objCT.gobjInfoUsu.objGridComissao)
End Sub

Public Sub GridComissao_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objCT.gobjInfoUsu.objGridComissao)
    
    If KeyCode = vbKeyDelete Then
        If objCT.gobjInfoUsu.gcolCliEmissoresTRV.Count >= GridComissao(objCT).Row Then
            objCT.gobjInfoUsu.gcolCliEmissoresTRV.Remove (GridComissao(objCT).Row)
        End If
    End If
    

End Sub

Public Sub GridComissao_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objCT.gobjInfoUsu.objGridComissao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridComissao, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridComissao_LeaveCell(ByVal objCT As Object)
    Call objCT.Saida_Celula(objCT.gobjInfoUsu.objGridComissao)
End Sub

Public Sub GridComissao_Validate(ByVal objCT As Object, Cancel As Boolean)
    Call Grid_Libera_Foco(objCT.gobjInfoUsu.objGridComissao)
End Sub

Public Sub GridComissao_RowColChange(ByVal objCT As Object)
    Call Grid_RowColChange(objCT.gobjInfoUsu.objGridComissao)
End Sub

Public Sub GridComissao_Scroll(ByVal objCT As Object)
    Call Grid_Scroll(objCT.gobjInfoUsu.objGridComissao)
End Sub

Public Sub GridExcecoes_Click(ByVal objCT As Object)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objCT.gobjInfoUsu.objGridExcecoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridExcecoes, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridExcecoes_EnterCell(ByVal objCT As Object)
    Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridExcecoes, objCT.iAlterado)
End Sub

Public Sub GridExcecoes_GotFocus(ByVal objCT As Object)
    Call Grid_Recebe_Foco(objCT.gobjInfoUsu.objGridExcecoes)
End Sub

Public Sub GridExcecoes_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objCT.gobjInfoUsu.objGridExcecoes)
End Sub

Public Sub GridExcecoes_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objCT.gobjInfoUsu.objGridExcecoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridExcecoes, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridExcecoes_LeaveCell(ByVal objCT As Object)
    Call objCT.Saida_Celula(objCT.gobjInfoUsu.objGridExcecoes)
End Sub

Public Sub GridExcecoes_Validate(ByVal objCT As Object, Cancel As Boolean)
    Call Grid_Libera_Foco(objCT.gobjInfoUsu.objGridExcecoes)
End Sub

Public Sub GridExcecoes_RowColChange(ByVal objCT As Object)
    Call Grid_RowColChange(objCT.gobjInfoUsu.objGridExcecoes)
End Sub

Public Sub GridExcecoes_Scroll(ByVal objCT As Object)
    Call Grid_Scroll(objCT.gobjInfoUsu.objGridExcecoes)
End Sub

Public Sub Emissor_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Emissor_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridComissao)
End Sub

Public Sub Emissor_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridComissao)
End Sub

Public Sub Emissor_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor

On Error GoTo Erro_Emissor_Validate

    If Len(Trim(Emissor(objCT).Text)) > 0 Then
   
        'Tenta ler o Fornecedor (NomeReduzido ou Código)
        lErro = TP_Fornecedor_Le2(Emissor(objCT), objFornecedor, 0)
        If lErro <> SUCESSO Then gError 195782

    End If

    Set objCT.gobjInfoUsu.objGridComissao.objControle = Emissor(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridComissao)
    If lErro <> SUCESSO Then Cancel = True
    
    Exit Sub
    
Erro_Emissor_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195782
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195783)
     
    End Select
    
End Sub

Public Sub PercComiss_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercComiss_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridComissao)
End Sub

Public Sub PercComiss_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridComissao)
End Sub

Public Sub PercComiss_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComiss_Validate

    Set objCT.gobjInfoUsu.objGridComissao.objControle = PercComiss(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridComissao)
    If lErro <> SUCESSO Then gError 195837
    
    Exit Sub
    
Erro_PercComiss_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195837
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195834)
     
    End Select
    
End Sub

Public Sub Produto_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Produto_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridExcecoes)
End Sub

Public Sub Produto_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridExcecoes)
End Sub

Public Sub Produto_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Produto_Validate

    Set objCT.gobjInfoUsu.objGridExcecoes.objControle = Produto(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridExcecoes)
    If lErro <> SUCESSO Then gError 195835
    
    Exit Sub
    
Erro_Produto_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195835
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195836)
     
    End Select
    
End Sub

Public Sub PercComissProd_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercComissProd_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridExcecoes)
End Sub

Public Sub PercComissProd_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridExcecoes)
End Sub

Public Sub PercComissProd_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComissProd_Validate

    Set objCT.gobjInfoUsu.objGridExcecoes.objControle = PercComissProd(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridExcecoes)
    If lErro <> SUCESSO Then gError 195838
    
    Exit Sub
    
Erro_PercComissProd_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195838
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195839)
     
    End Select
    
End Sub

Public Function CLI_Rotina_Grid_Enable(ByVal objCT As Object, iLinha As Integer, objControl As Object, iLocalChamada As Integer) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objCliEmissoresTRV As ClassCliEmissoresTRV
Dim objCliEmissoresExcTRV As ClassCliEmissoresExcTRV
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_CLI_Rotina_Grid_Enable

    Select Case objControl.Name

        Case Emissor(objCT).Name, PercComiss(objCT).Name
        
            If objCT.gobjInfoUsu.giLinhaAntiga <> iLinha Then
            
                If objCT.gobjInfoUsu.giLinhaAntiga > 0 And objCT.gobjInfoUsu.giLinhaAntiga <= objCT.gobjInfoUsu.objGridComissao.iLinhasExistentes Then
                
                    lErro = Move_CliEmissoresExcTRV(objCT, objCT.gobjInfoUsu.giLinhaAntiga)
                    If lErro <> SUCESSO Then gError 195809
                
               End If
               
               Call Grid_Limpa(objCT.gobjInfoUsu.objGridExcecoes)
               
                objCT.gobjInfoUsu.giLinhaAntiga = iLinha
               
                Frame12(objCT).Caption = "Exceções por produto (Emissor: )"
               
                If iLinha <= objCT.gobjInfoUsu.gcolCliEmissoresTRV.Count Then
                
                    Frame12(objCT).Caption = "Exceções por produto (Emissor: " & left(GridComissao(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Emissor_Col), 22) & ")"
                
                    Set objCliEmissoresTRV = objCT.gobjInfoUsu.gcolCliEmissoresTRV.Item(iLinha)
                    
                    Call Grid_Limpa(objCT.gobjInfoUsu.objGridExcecoes)
                    
                    For Each objCliEmissoresExcTRV In objCliEmissoresTRV.colCliEmissoresExcTRV
                    
                        iIndice = iIndice + 1
                    
                        lErro = Mascara_RetornaProdutoEnxuto(objCliEmissoresExcTRV.sProduto, sProdutoEnxuto)
                        If lErro <> SUCESSO Then gError 195802
    
                        Produto(objCT).PromptInclude = False
                        Produto(objCT).Text = sProdutoEnxuto
                        Produto(objCT).PromptInclude = True
                
                        objProduto.sCodigo = objCliEmissoresExcTRV.sProduto
                        
                        'Lê o Produto
                        lErro = CF("Produto_Le", objProduto)
                        If lErro <> SUCESSO And lErro <> 28030 Then gError 195803
                        
                        'Preenche o Grid
                        GridExcecoes(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Produto_Col) = Produto(objCT).Text
                        GridExcecoes(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_DescricaoProduto_Col) = objProduto.sDescricao
                        GridExcecoes(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PercComissProd_Col) = Format(objCliEmissoresExcTRV.dPercComissao, "Percent")
                    
                    Next
                    
                    objCT.gobjInfoUsu.objGridExcecoes.iLinhasExistentes = objCliEmissoresTRV.colCliEmissoresExcTRV.Count
                    
                End If
                    
            End If
            
        Case ExcRepProduto(objCT).Name
            
            If Len(Trim(GridExcRep(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_ExcRepProduto_Col))) > 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case ExcRep(objCT).Name, ExcRepPercComis(objCT).Name
            If Len(Trim(GridExcRep(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_ExcRepProduto_Col))) > 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        Case ExcCorProduto(objCT).Name
            
            If Len(Trim(GridExcCor(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_ExcCorProduto_Col))) > 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case ExcCor(objCT).Name, ExcCorPercComis(objCT).Name
            If Len(Trim(GridExcCor(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_ExcCorProduto_Col))) > 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If

        Case ExcAgProduto(objCT).Name
            
            If Len(Trim(GridExcAg(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_ExcAgProduto_Col))) > 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
            
        Case ExcAgPercComis(objCT).Name
            If Len(Trim(GridExcAg(objCT).TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_ExcAgProduto_Col))) > 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
    End Select
    
    CLI_Rotina_Grid_Enable = SUCESSO

    Exit Function

Erro_CLI_Rotina_Grid_Enable:

    CLI_Rotina_Grid_Enable = gErr

    Select Case gErr
    
        Case 195801, 195803
        
        Case 195802
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195804)

    End Select

    Exit Function

End Function

Private Function Move_CliEmissoresExcTRV(ByVal objCT As Object, ByVal iLinhaAntiga As Integer) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objCliEmissoresTRV As ClassCliEmissoresTRV
Dim objCliEmissoresExcTRV As ClassCliEmissoresExcTRV
Dim iIndice As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Move_CliEmissoresExcTRV

    If objCT.gobjInfoUsu.giLinhaAntiga > 0 And objCT.gobjInfoUsu.giLinhaAntiga <= objCT.gobjInfoUsu.gcolCliEmissoresTRV.Count Then

        Set objCliEmissoresTRV = objCT.gobjInfoUsu.gcolCliEmissoresTRV.Item(objCT.gobjInfoUsu.giLinhaAntiga)
        
        Set objCliEmissoresTRV.colCliEmissoresExcTRV = New Collection
                                
        For iIndice = 1 To objCT.gobjInfoUsu.objGridExcecoes.iLinhasExistentes
        
            lErro = CF("Produto_Formata", GridExcecoes(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Produto_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 195801
            
            If iProdutoPreenchido = PRODUTO_VAZIO Then gError 195807
            
            Set objCliEmissoresExcTRV = New ClassCliEmissoresExcTRV
            
            objCliEmissoresExcTRV.sProduto = sProdutoFormatado
            
            If Len(Trim(GridExcecoes(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PercComissProd_Col))) = 0 Then gError 195808
            
            objCliEmissoresExcTRV.sProduto = sProdutoFormatado
            objCliEmissoresExcTRV.dPercComissao = PercentParaDbl(GridExcecoes(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_PercComissProd_Col))
        
            objCliEmissoresTRV.colCliEmissoresExcTRV.Add objCliEmissoresExcTRV
            
        Next
                        
    End If
                        
    Move_CliEmissoresExcTRV = SUCESSO

    Exit Function

Erro_Move_CliEmissoresExcTRV:

    Move_CliEmissoresExcTRV = gErr

    Select Case gErr
    
        Case 195801
        
        Case 195807
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_GRID_EXCECOES", gErr, iIndice)
        
        Case 195808
            Call Rotina_Erro(vbOKOnly, "ERRO_PERCCOMISSPROD_NAO_PREENCHIDO_GRID_EXCECOES", gErr, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195810)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridExcecoes(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridExcecoes

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for a de Produto
        Case objCT.gobjInfoUsu.iGrid_Produto_Col
            lErro = Saida_Celula_Produto(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195788

        'Se for a de Unidade de Medida
        Case objCT.gobjInfoUsu.iGrid_PercComissProd_Col
            lErro = Saida_Celula_PercComissProd(objCT, objGridInt)
            If lErro <> SUCESSO Then gError 195789

    End Select

    Saida_Celula_GridExcecoes = SUCESSO

    Exit Function

Erro_Saida_Celula_GridExcecoes:

    Saida_Celula_GridExcecoes = gErr

    Select Case gErr

        Case 195788, 195789

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195790)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridCallCenter(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridCallCenter

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for a de datade
        Case objCT.gobjInfoUsu.iGrid_DataCallCenterDe_Col
            lErro = Saida_Celula_Data(objGridInt, DataCallCenterDe(objCT))
            If lErro <> SUCESSO Then gError 195788

        'Se for a de dataate
        Case objCT.gobjInfoUsu.iGrid_DataCallCenterAte_Col
            lErro = Saida_Celula_Data(objGridInt, DataCallCenterAte(objCT))
            If lErro <> SUCESSO Then gError 195789

    End Select

    Saida_Celula_GridCallCenter = SUCESSO

    Exit Function

Erro_Saida_Celula_GridCallCenter:

    Saida_Celula_GridCallCenter = gErr

    Select Case gErr

        Case 195788, 195789

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195790)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto que está deixando de ser a corrente

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iIndice As Integer
Dim iProdutoPreenchido As Integer
Dim sProdutoEnxuto As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto(objCT)

    If Len(Trim(Produto(objCT).ClipText)) > 0 Then

        'Critica o Produto
        lErro = CF("Produto_Critica_Filial2", Produto(objCT).Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 195791
        
        'Se o produto é gerencial ==> erro
        If lErro = 86295 Then gError 195792
               
        'Se o produto não foi encontrado ==> Pergunta se deseja criar
        If lErro = 51381 Then gError 195793

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 195794
    
        Produto(objCT).PromptInclude = False
        Produto(objCT).Text = sProdutoEnxuto
        Produto(objCT).PromptInclude = True
    
        'Verifica se já está em outra linha do Grid
        For iIndice = 1 To objGridInt.iLinhasExistentes
            If iIndice <> GridExcecoes(objCT).Row Then
                If GridExcecoes(objCT).TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Produto_Col) = Produto(objCT).Text Then gError 195795
            End If
        Next
    
        If GridExcecoes(objCT).Row - GridExcecoes(objCT).FixedRows = objGridInt.iLinhasExistentes Then
    
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    
        End If
    
        GridExcecoes(objCT).TextMatrix(GridExcecoes(objCT).Row, objCT.gobjInfoUsu.iGrid_DescricaoProduto_Col) = objProduto.sDescricao

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195796

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 195791, 195796
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 195792
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 195793
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto(objCT).Text)
            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 195794
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case 195795
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_PREENCHIDO_LINHA_GRID", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195797)

    End Select

    Exit Function

End Function

Function Saida_Celula_PercComissProd(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double

On Error GoTo Erro_Saida_Celula_PercComissProd

    Set objGridInt.objControle = PercComissProd(objCT)

    If Len(PercComissProd(objCT).Text) > 0 Then
        
        'Critica a porcentagem
        lErro = Porcentagem_Critica(PercComissProd(objCT).Text)
        If lErro <> SUCESSO Then gError 195798

        dPercentDesc = CDbl(PercComissProd(objCT).Text)
        
        PercComissProd(objCT).Text = Format(dPercentDesc, "Fixed")
        
        If GridExcecoes(objCT).Row - GridExcecoes(objCT).FixedRows = objGridInt.iLinhasExistentes Then
    
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195799

    Saida_Celula_PercComissProd = SUCESSO

    Exit Function

Erro_Saida_Celula_PercComissProd:

    Saida_Celula_PercComissProd = gErr

    Select Case gErr

        Case 195798, 195799
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195800)

    End Select

    Exit Function

End Function

Public Sub BotaoEmissores_Click(ByVal objCT As Object)

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

On Error GoTo Erro_BotaoEmissores_Click

    If objCT.objUserControl.ActiveControl Is Emissor(objCT) Then
    
        objFornecedor.lCodigo = LCodigo_Extrai(Emissor(objCT).Text)
    
        'recolhe o Nome Reduzido da tela
        objFornecedor.sNomeReduzido = Emissor(objCT).Text
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridComissao(objCT).Row = 0 Then gError 195825

        objFornecedor.lCodigo = LCodigo_Extrai(GridComissao(objCT).TextMatrix(GridComissao(objCT).Row, objCT.gobjInfoUsu.iGrid_Emissor_Col))
        
    End If

    Set gobjCT = objCT

    Set objEventoFornecedor = New AdmEvento

    'Chama a Tela de browse Fornecedores
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

    Exit Sub

Erro_BotaoEmissores_Click:

    Select Case gErr
    
        Case 195825
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195826)
    
    End Select
    
    Exit Sub

End Sub

Public Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As New ClassFornecedor
Dim bCancel As Boolean
Dim lErro As Long
Dim iIndice As Integer
Dim objCliEmissoresTRV As ClassCliEmissoresTRV


On Error GoTo Erro_objEventoFornecedor_evSelecao

    Set objFornecedor = obj1

    'Coloca o Fornecedor na tela
    Emissor(gobjCT).Text = objFornecedor.lCodigo
    
    If Len(Trim(Emissor(gobjCT).Text)) > 0 Then
   
        'Tenta ler o Fornecedor (NomeReduzido ou Código)
        lErro = TP_Fornecedor_Le2(Emissor(gobjCT), objFornecedor, 0)
        If lErro <> SUCESSO Then gError 195823

        If Not (gobjCT.objUserControl.ActiveControl Is Emissor(gobjCT)) Then

            'Verifica se já está em outra linha do Grid
            For iIndice = 1 To gobjCT.gobjInfoUsu.objGridComissao.iLinhasExistentes
                If iIndice <> GridComissao(gobjCT).Row Then
                    If GridComissao(gobjCT).TextMatrix(iIndice, gobjCT.gobjInfoUsu.iGrid_Emissor_Col) = Emissor(gobjCT).Text Then gError 195840
                End If
            Next


            GridComissao(gobjCT).TextMatrix(GridComissao(gobjCT).Row, gobjCT.gobjInfoUsu.iGrid_Emissor_Col) = Emissor(gobjCT).Text
            
            If GridComissao(gobjCT).Row - GridComissao(gobjCT).FixedRows = gobjCT.gobjInfoUsu.objGridComissao.iLinhasExistentes Then
        
                Frame12(gobjCT).Caption = "Exceções por produto (Emissor: " & left(Emissor(gobjCT).Text, 22) & ")"

                gobjCT.gobjInfoUsu.objGridComissao.iLinhasExistentes = gobjCT.gobjInfoUsu.objGridComissao.iLinhasExistentes + 1
        
                Set objCliEmissoresTRV = New ClassCliEmissoresTRV
                gobjCT.gobjInfoUsu.gcolCliEmissoresTRV.Add objCliEmissoresTRV
        
            End If
            
            
        End If

    End If

    gobjCT.objUserControl.Show
    
    Exit Sub
    
Erro_objEventoFornecedor_evSelecao:

    Select Case gErr
    
        Case 195823
    
        Case 195840
            Call Rotina_Erro(vbOKOnly, "ERRO_EMISSOR_JA_EXISTE_GRID", gErr, Emissor(gobjCT).Text, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195824)
    
    End Select

End Sub

Public Sub BotaoProdutos_Click(ByVal objCT As Object)

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoProdutos_Click

    If objCT.objUserControl.ActiveControl Is Produto(objCT) Then
    
        sProduto1 = Produto(objCT).Text
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridExcecoes(objCT).Row = 0 Then gError 195827

        sProduto1 = GridExcecoes(objCT).TextMatrix(GridExcecoes(objCT).Row, objCT.gobjInfoUsu.iGrid_Produto_Col)
        
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 195828
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto
    
    Set gobjCT = objCT

    Set objEventoProduto = New AdmEvento

    'Chama a Tela ProdutoVendaLista
    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoProduto)

    Exit Sub
    
Erro_BotaoProdutos_Click:

    Select Case gErr
    
        Case 195827
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 195828
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195829)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProdutoEnxuto As String
Dim iIndice As Integer

On Error GoTo Erro_objEventoProduto_evSelecao

    'verifica se tem alguma linha do Grid selecionada
    If GridExcecoes(gobjCT).Row > 0 Then

        Set objProduto = obj1

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 195831

        Produto(gobjCT).PromptInclude = False
        Produto(gobjCT).Text = sProdutoEnxuto
        Produto(gobjCT).PromptInclude = True
        
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 195832
        
        If Not (gobjCT.objUserControl.ActiveControl Is Produto(gobjCT)) Then
    
    
            'Verifica se já está em outra linha do Grid
            For iIndice = 1 To gobjCT.gobjInfoUsu.objGridExcecoes.iLinhasExistentes
                If iIndice <> GridExcecoes(gobjCT).Row Then
                    If GridExcecoes(gobjCT).TextMatrix(iIndice, gobjCT.gobjInfoUsu.iGrid_Produto_Col) = Produto(gobjCT).Text Then gError 195839
                End If
            Next

    
            'Preenche o Grid
            GridExcecoes(gobjCT).TextMatrix(GridExcecoes(gobjCT).Row, gobjCT.gobjInfoUsu.iGrid_Produto_Col) = Produto(gobjCT).Text
            GridExcecoes(gobjCT).TextMatrix(GridExcecoes(gobjCT).Row, gobjCT.gobjInfoUsu.iGrid_DescricaoProduto_Col) = objProduto.sDescricao
    
            If GridExcecoes(gobjCT).Row - GridExcecoes(gobjCT).FixedRows = gobjCT.gobjInfoUsu.objGridExcecoes.iLinhasExistentes Then
        
                gobjCT.gobjInfoUsu.objGridExcecoes.iLinhasExistentes = gobjCT.gobjInfoUsu.objGridExcecoes.iLinhasExistentes + 1
        
            End If
    
    
        End If
        

    End If

    gobjCT.objUserControl.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    GridExcecoes(gobjCT).TextMatrix(GridExcecoes(gobjCT).Row, gobjCT.gobjInfoUsu.iGrid_Produto_Col) = ""

    Select Case gErr

        Case 195830, 195832
        
        Case 195831
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case 195839
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_PREENCHIDO_LINHA_GRID", gErr, iIndice)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195833)

    End Select

    Exit Sub

End Sub

Public Sub UserControl_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If objCT.objUserControl.ActiveControl Is Emissor(objCT) Then
            Call BotaoEmissores_Click(objCT)
        ElseIf objCT.objUserControl.ActiveControl Is Produto(objCT) Then
            Call BotaoProdutos_Click(objCT)
        ElseIf objCT.objUserControl.ActiveControl Is Representante(objCT) Then
            Call LabelRep_Click(objCT)
        ElseIf objCT.objUserControl.ActiveControl Is Correntista(objCT) Then
            Call LabelCorr_Click(objCT)
        ElseIf objCT.objUserControl.ActiveControl Is EmpresaPai(objCT) Then
            Call LabelEmpresaPai_Click(objCT)
        ElseIf objCT.objUserControl.ActiveControl Is CondicaoPagtoCC(objCT) Then
            Call CondicaoPagtoCCLabel_Click(objCT)
        ElseIf objCT.objUserControl.ActiveControl Is ExcRep(objCT) Then
            Call BotaoExcRep_Click(objCT)
        ElseIf objCT.objUserControl.ActiveControl Is ExcCor(objCT) Then
            Call BotaoExcCor_Click(objCT)
        ElseIf objCT.objUserControl.ActiveControl Is ExcRepProduto(objCT) Then
            Call BotaoExcRepProduto_Click(objCT)
        ElseIf objCT.objUserControl.ActiveControl Is ExcCorProduto(objCT) Then
            Call BotaoExcCorProduto_Click(objCT)
        ElseIf objCT.objUserControl.ActiveControl Is ExcAgProduto(objCT) Then
            Call BotaoExcAgProduto_Click(objCT)
        End If
    
    End If

End Sub

Public Function CLI_Form_UnLoad(ByVal objCT As Object) As Long

    Set objEventoFornecedor = Nothing
    Set objEventoProduto = Nothing
    Set objEventoRepresentante = Nothing
    Set objEventoCorrentista = Nothing
    Set objEventoEmpresaPai = Nothing
    Set objEventoAcordo = Nothing
    Set objEventoCondicaoPagtoCC = Nothing
    Set objEventoExcRep = Nothing
    Set objEventoExcCor = Nothing
    Set objEventoExcRepProduto = Nothing
    Set objEventoExcCorProduto = Nothing
    Set objEventoExcAgProduto = Nothing
    Set gobjCT = Nothing

    CLI_Form_UnLoad = SUCESSO
    
End Function

Public Function CLI_Limpa_Tela_Cliente(ByVal objCT As Object) As Long
    
    Call Grid_Limpa(objCT.gobjInfoUsu.objGridComissao)
    Call Grid_Limpa(objCT.gobjInfoUsu.objGridExcecoes)
    Call Grid_Limpa(objCT.gobjInfoUsu.objGridExcRep)
    Call Grid_Limpa(objCT.gobjInfoUsu.objGridExcCor)
    Call Grid_Limpa(objCT.gobjInfoUsu.objGridExcAg)
    Call Grid_Limpa(objCT.gobjInfoUsu.objGridCallCenter)
    
    FilialEmpresa(objCT).ListIndex = -1
    FilialEmpresa(objCT).Text = ""
    
    FilialEmpresaNF(objCT).ListIndex = -1
    FilialEmpresaNF(objCT).Text = ""
    
    FilialEmpresaFAT(objCT).ListIndex = -1
    FilialEmpresaFAT(objCT).Text = ""
    
    CondicaoPagtoCC(objCT).ListIndex = -1
    CondicaoPagtoCC(objCT).Text = ""
    
    CLI_Limpa_Tela_Cliente = SUCESSO
    
End Function

Public Sub Representante_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Representante_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim dPercComissao As Double
Dim iCodFilial As Integer

On Error GoTo Erro_Representante_Validate

    If Len(Trim(Representante(objCT).Text)) > 0 Then

        iCodFilial = FILIAL_MATRIZ

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(Representante(objCT), objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 195843
        
        If objCliente.iAtivo = DESMARCADO Then gError 195844

    End If
    
    Exit Sub

Erro_Representante_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 195843
        
        Case 195844
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INATIVO", gErr, objCliente.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195845)
    
    End Select

End Sub

Public Sub LabelRep_Click(ByVal objCT As Object)

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
Dim sNomeBrowse As String

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(Representante(objCT).Text)) > 0 Then objCliente.sNomeReduzido = Representante(objCT).Text

    sNomeBrowse = "ClientesLista"

    Call CF("Cliente_Obtem_NomeBrowse", sNomeBrowse)

    Set gobjCT = objCT

    Set objEventoRepresentante = New AdmEvento

    'Chama Tela ClienteLista
    Call Chama_Tela(sNomeBrowse, colSelecao, objCliente, objEventoRepresentante)

End Sub

Public Sub objEventoRepresentante_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    Representante(gobjCT).Text = objCliente.sNomeReduzido

    gobjCT.objUserControl.Show

    Exit Sub

End Sub

Public Sub PerComiRep_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercComiRep_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComiRep_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(PercComiRep(objCT).Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(PercComiRep(objCT).Text)
    If lErro <> SUCESSO Then Error 195849

    'Formata
    PercComiRep(objCT).Text = Format(PercComiRep(objCT).Text, "Fixed")

    Exit Sub

Erro_PercComiRep_Validate:

    Cancel = True


    Select Case Err

        Case 195849
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 195850)
    
    End Select

    Exit Sub

End Sub

Public Sub PercComiRep_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Correntista_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Correntista_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim dPercComissao As Double
Dim iCodFilial As Integer

On Error GoTo Erro_Correntista_Validate

    If Len(Trim(Correntista(objCT).Text)) > 0 Then

        iCodFilial = FILIAL_MATRIZ

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(Correntista(objCT), objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 195846
        
        If objCliente.iAtivo = DESMARCADO Then gError 195847

    End If
    
    Exit Sub

Erro_Correntista_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 195846
        
        Case 195847
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INATIVO", gErr, objCliente.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195848)
    
    End Select

End Sub

Public Sub LabelCorr_Click(ByVal objCT As Object)

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
Dim sNomeBrowse As String

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(Correntista(objCT).Text)) > 0 Then objCliente.sNomeReduzido = Correntista(objCT).Text

    sNomeBrowse = "ClientesLista"

    Call CF("Cliente_Obtem_NomeBrowse", sNomeBrowse)

    Set gobjCT = objCT

    Set objEventoCorrentista = New AdmEvento

    'Chama Tela ClienteLista
    Call Chama_Tela(sNomeBrowse, colSelecao, objCliente, objEventoCorrentista)

End Sub

Public Sub objEventoCorrentista_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    Correntista(gobjCT).Text = objCliente.sNomeReduzido

    gobjCT.objUserControl.Show

    Exit Sub

End Sub

Public Sub PercComiCorr_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercComiCorr_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComiCorr_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(PercComiCorr(objCT).Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(PercComiCorr(objCT).Text)
    If lErro <> SUCESSO Then Error 195851

    'Formata
    PercComiCorr(objCT).Text = Format(PercComiCorr(objCT).Text, "Fixed")

    Exit Sub

Erro_PercComiCorr_Validate:

    Cancel = True

    Select Case Err

        Case 195851
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 195852)
    
    End Select

    Exit Sub

End Sub

Public Sub PercComiAg_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercComiAg_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercComiAg_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(PercComiAg(objCT).Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(PercComiAg(objCT).Text)
    If lErro <> SUCESSO Then Error 195853

    'Formata
    PercComiAg(objCT).Text = Format(PercComiAg(objCT).Text, "Fixed")

    Exit Sub

Erro_PercComiAg_Validate:

    Cancel = True

    Select Case Err

        Case 195853
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 195854)
    
    End Select

    Exit Sub

End Sub

Public Sub ConsiderarAporte_Click(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub EmpresaPai_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub EmpresaPai_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente
Dim iCodFilial As Integer

On Error GoTo Erro_EmpresaPai_Validate

    If Len(Trim(EmpresaPai(objCT).Text)) > 0 Then

        iCodFilial = FILIAL_MATRIZ

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_Cliente_Le2(EmpresaPai(objCT), objCliente, iCodFilial)
        If lErro <> SUCESSO Then gError 197000
        
        If objCliente.iAtivo = DESMARCADO Then gError 197001

    End If
    
    Exit Sub

Erro_EmpresaPai_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case 197000
        
        Case 197001
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INATIVO", gErr, objCliente.lCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197002)
    
    End Select

End Sub

Public Sub LabelEmpresaPai_Click(ByVal objCT As Object)

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
Dim sNomeBrowse As String

    'Preenche NomeReduzido com o cliente da tela
    If Len(Trim(EmpresaPai(objCT).Text)) > 0 Then objCliente.sNomeReduzido = EmpresaPai(objCT).Text

    sNomeBrowse = "ClientesLista"

    Call CF("Cliente_Obtem_NomeBrowse", sNomeBrowse)

    Set gobjCT = objCT

    Set objEventoEmpresaPai = New AdmEvento

    'Chama Tela ClienteLista
    Call Chama_Tela(sNomeBrowse, colSelecao, objCliente, objEventoEmpresaPai)

End Sub

Public Sub objEventoEmpresaPai_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    EmpresaPai(gobjCT).Text = objCliente.sNomeReduzido

    gobjCT.objUserControl.Show

    Exit Sub

End Sub

Public Sub BotaoAcordos_Click(ByVal objCT As Object)

Dim lErro As Long
Dim objTRVAcordo As New ClassTRVAcordos
Dim sSelecao As String
Dim colSelecao As New Collection
Dim objCliente1 As New ClassCliente

On Error GoTo Erro_BotaoAcordos_Click

    If Len(Trim(EmpresaPai(objCT).Text)) > 0 Then
    
'        objCliente1.sNomeReduzido = EmpresaPai(objCT).Text
'
'        'Lê o Cliente
'        lErro = CF("Cliente_Le_NomeReduzido", objCliente1)
'        If lErro <> SUCESSO And lErro <> 12348 Then gError 195857
        
        objTRVAcordo.lCliente = LCodigo_Extrai(EmpresaPai(objCT).Text)

        colSelecao.Add objTRVAcordo.lCliente

        sSelecao = "Cliente = ?"
        
    End If
        
    Set objEventoAcordo = New AdmEvento
        
    Call Chama_Tela("TRVAcordosLista", colSelecao, objTRVAcordo, objEventoAcordo, sSelecao)

    Exit Sub

Erro_BotaoAcordos_Click:

    Select Case gErr

        Case 195857

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196488)

    End Select

    Exit Sub
    
End Sub

Public Sub FilialEmpresa_Click(ByVal objCT As Object, ByVal objCombo As Object)

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_FilialEmpresa_Click

    'Verifica se a filial foi selecionada
    If objCombo.ListIndex = -1 Then Exit Sub

    'preeenche Codigo da Filial em objFilialEmpresa
    objFilialEmpresa.iCodFilial = Codigo_Extrai(objCombo.Text)

    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
    If lErro <> SUCESSO And lErro <> 27378 Then gError 196486

    'Se nao encontrou => erro
    If lErro = 27378 Then gError 196487

    Exit Sub

Erro_FilialEmpresa_Click:

    Select Case gErr

        Case 196486

        Case 196487
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196488)

    End Select

    Exit Sub

End Sub

Public Sub FilialEmpresa_Validate(ByVal objCT As Object, Cancel As Boolean, ByVal objCombo As Object)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_FilialEmpresa_Validate

    'Verifica se FilialEmpresa foi preenchida
    If Len(Trim(objCombo.Text)) > 0 Then

        If objCombo.ListIndex <> -1 Then Exit Sub

        lErro = Combo_Seleciona(objCombo, iCodigo)
        If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 196489

        'Se nao encontra o ítem com o código informado
        If lErro = 6730 Then

            'preeenche objFilialEmpresa
            objFilialEmpresa.iCodFilial = iCodigo

            'Le a FilialEmpresa
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO And lErro <> 27378 Then gError 196490

            'Se nao encontrou => erro
            If lErro = 27378 Then gError 196491

            If lErro = SUCESSO Then

                'Coloca na tela o codigo e o nome da FilialEmpresa
                objCombo.Text = objFilialEmpresa.lCodEmpresa & SEPARADOR & objFilialEmpresa.sNome

            End If

        End If

        'Se nao encontrou e nao era codigo
        If lErro = 6731 Then gError 196492

    End If

    Exit Sub

Erro_FilialEmpresa_Validate:

    Select Case gErr

        Case 196491
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA", gErr, iCodigo)

        Case 196489, 196490

        Case 196492
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_EMPRESA_NAO_CADASTRADA2", gErr, objFilialEmpresa.iCodFilial)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196493)

    End Select

    Exit Sub

End Sub

Private Function Carrega_FilialEmpresa(ByVal objCT As Object, ByVal objCombo As Object) As Long

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Carrega_FilialEmpresa

    objCombo.Clear

    objCombo.Text = ""

    'Lê o Código e o NOme de Toda FilialEmpresa do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 196505

    'Carrega a combo de Filial Empresa
    For Each objCodigoNome In colCodigoNome
        If objCodigoNome.iCodigo < Abs(giFilialAuxiliar) Then
            objCombo.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            objCombo.ItemData(objCombo.NewIndex) = objCodigoNome.iCodigo
        End If
    Next

    Carrega_FilialEmpresa = SUCESSO

    Exit Function

Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = gErr

    Select Case gErr

        Case 196505

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 196506)

    End Select

    Exit Function

End Function

Public Sub Grupo_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Private Function Inicializa_Grid_ExcRep(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Represent.")
    objGridInt.colColuna.Add ("%Comissão")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ExcRepProduto(objCT).Name)
    objGridInt.colCampo.Add (ExcRep(objCT).Name)
    objGridInt.colCampo.Add (ExcRepPercComis(objCT).Name)

    'Colunas do Grid
    objCT.gobjInfoUsu.iGrid_ExcRepProduto_Col = 1
    objCT.gobjInfoUsu.iGrid_ExcRep_Col = 2
    objCT.gobjInfoUsu.iGrid_ExcRepPercComiss_Col = 3

    'Grid do GridInterno
    objGridInt.objGrid = GridExcRep(objCT)

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridExcRep(objCT).ColWidth(0) = 200
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ExcRep = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_ExcCor(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Corrent.")
    objGridInt.colColuna.Add ("%Comissão")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ExcCorProduto(objCT).Name)
    objGridInt.colCampo.Add (ExcCor(objCT).Name)
    objGridInt.colCampo.Add (ExcCorPercComis(objCT).Name)

    'Colunas do Grid
    objCT.gobjInfoUsu.iGrid_ExcCorProduto_Col = 1
    objCT.gobjInfoUsu.iGrid_ExcCor_Col = 2
    objCT.gobjInfoUsu.iGrid_ExcCorPercComiss_Col = 3

    'Grid do GridInterno
    objGridInt.objGrid = GridExcCor(objCT)

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridExcCor(objCT).ColWidth(0) = 200
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ExcCor = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_ExcAg(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("%Comissão")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ExcAgProduto(objCT).Name)
    objGridInt.colCampo.Add (ExcAgPercComis(objCT).Name)

    'Colunas do Grid
    objCT.gobjInfoUsu.iGrid_ExcAgProduto_Col = 1
    objCT.gobjInfoUsu.iGrid_ExcAgPercComiss_Col = 2

    'Grid do GridInterno
    objGridInt.objGrid = GridExcAg(objCT)

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 201

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridExcAg(objCT).ColWidth(0) = 400
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ExcAg = SUCESSO

    Exit Function

End Function

Public Sub ExcRep_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ExcRep_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridExcRep)
End Sub

Public Sub ExcRep_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridExcRep)
End Sub

Public Sub ExcRep_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ExcRep_Validate

    Set objCT.gobjInfoUsu.objGridExcRep.objControle = ExcRep(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridExcRep)
    If lErro <> SUCESSO Then gError 195835
    
    Exit Sub
    
Erro_ExcRep_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195835
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195836)
     
    End Select
    
End Sub

Public Sub ExcRepProduto_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ExcRepProduto_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridExcRep)
End Sub

Public Sub ExcRepProduto_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridExcRep)
End Sub

Public Sub ExcRepProduto_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ExcRepProduto_Validate

    Set objCT.gobjInfoUsu.objGridExcRep.objControle = ExcRepProduto(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridExcRep)
    If lErro <> SUCESSO Then gError 195835
    
    Exit Sub
    
Erro_ExcRepProduto_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195835
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195836)
     
    End Select
    
End Sub

Public Sub ExcRepPercComis_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ExcRepPercComis_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridExcRep)
End Sub

Public Sub ExcRepPercComis_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridExcRep)
End Sub

Public Sub ExcRepPercComis_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ExcRepPercComis_Validate

    Set objCT.gobjInfoUsu.objGridExcRep.objControle = ExcRepPercComis(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridExcRep)
    If lErro <> SUCESSO Then gError 195835
    
    Exit Sub
    
Erro_ExcRepPercComis_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195835
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195836)
     
    End Select
    
End Sub

Public Sub ExcCor_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ExcCor_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridExcCor)
End Sub

Public Sub ExcCor_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridExcCor)
End Sub

Public Sub ExcCor_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ExcCor_Validate

    Set objCT.gobjInfoUsu.objGridExcCor.objControle = ExcCor(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridExcCor)
    If lErro <> SUCESSO Then gError 195835
    
    Exit Sub
    
Erro_ExcCor_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195835
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195836)
     
    End Select
    
End Sub

Public Sub ExcCorProduto_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ExcCorProduto_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridExcCor)
End Sub

Public Sub ExcCorProduto_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridExcCor)
End Sub

Public Sub ExcCorProduto_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ExcCorProduto_Validate

    Set objCT.gobjInfoUsu.objGridExcCor.objControle = ExcCorProduto(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridExcCor)
    If lErro <> SUCESSO Then gError 195835
    
    Exit Sub
    
Erro_ExcCorProduto_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195835
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195836)
     
    End Select
    
End Sub

Public Sub ExcCorPercComis_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ExcCorPercComis_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridExcCor)
End Sub

Public Sub ExcCorPercComis_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridExcCor)
End Sub

Public Sub ExcCorPercComis_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ExcCorPercComis_Validate

    Set objCT.gobjInfoUsu.objGridExcCor.objControle = ExcCorPercComis(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridExcCor)
    If lErro <> SUCESSO Then gError 195835
    
    Exit Sub
    
Erro_ExcCorPercComis_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195835
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195836)
     
    End Select
    
End Sub

Public Sub GridExcRep_Click(ByVal objCT As Object)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objCT.gobjInfoUsu.objGridExcRep, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridExcRep, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridExcRep_EnterCell(ByVal objCT As Object)
    Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridExcRep, objCT.iAlterado)
End Sub

Public Sub GridExcRep_GotFocus(ByVal objCT As Object)
    Call Grid_Recebe_Foco(objCT.gobjInfoUsu.objGridExcRep)
End Sub

Public Sub GridExcRep_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objCT.gobjInfoUsu.objGridExcRep)
End Sub

Public Sub GridExcRep_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objCT.gobjInfoUsu.objGridExcRep, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridExcRep, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridExcRep_LeaveCell(ByVal objCT As Object)
    Call objCT.Saida_Celula(objCT.gobjInfoUsu.objGridExcRep)
End Sub

Public Sub GridExcRep_Validate(ByVal objCT As Object, Cancel As Boolean)
    Call Grid_Libera_Foco(objCT.gobjInfoUsu.objGridExcRep)
End Sub

Public Sub GridExcRep_RowColChange(ByVal objCT As Object)
    Call Grid_RowColChange(objCT.gobjInfoUsu.objGridExcRep)
End Sub

Public Sub GridExcRep_Scroll(ByVal objCT As Object)
    Call Grid_Scroll(objCT.gobjInfoUsu.objGridExcRep)
End Sub

Public Sub GridExcCor_Click(ByVal objCT As Object)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objCT.gobjInfoUsu.objGridExcCor, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridExcCor, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridExcCor_EnterCell(ByVal objCT As Object)
    Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridExcCor, objCT.iAlterado)
End Sub

Public Sub GridExcCor_GotFocus(ByVal objCT As Object)
    Call Grid_Recebe_Foco(objCT.gobjInfoUsu.objGridExcCor)
End Sub

Public Sub GridExcCor_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objCT.gobjInfoUsu.objGridExcCor)
End Sub

Public Sub GridExcCor_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objCT.gobjInfoUsu.objGridExcCor, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridExcCor, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridExcCor_LeaveCell(ByVal objCT As Object)
    Call objCT.Saida_Celula(objCT.gobjInfoUsu.objGridExcCor)
End Sub

Public Sub GridExcCor_Validate(ByVal objCT As Object, Cancel As Boolean)
    Call Grid_Libera_Foco(objCT.gobjInfoUsu.objGridExcCor)
End Sub

Public Sub GridExcCor_RowColChange(ByVal objCT As Object)
    Call Grid_RowColChange(objCT.gobjInfoUsu.objGridExcCor)
End Sub

Public Sub GridExcCor_Scroll(ByVal objCT As Object)
    Call Grid_Scroll(objCT.gobjInfoUsu.objGridExcCor)
End Sub

Public Sub CondicaoPagtoCC_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CondicaoPagtoCC_Click(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CondicaoPagtoCC_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCondicaoPagto As New ClassCondicaoPagto
Dim iCodigo As Integer

On Error GoTo Erro_CondicaoPagtoCC_Validate

    'Verifica se foi preenchida a ComboBox CondicaoPagto
    If Len(Trim(CondicaoPagtoCC(objCT).Text)) = 0 Then Exit Sub

    'Verifica se está preenchida com o ítem selecionado na ComboBox CondicaoPagto
    If CondicaoPagtoCC(objCT).Text = CondicaoPagtoCC(objCT).List(CondicaoPagtoCC(objCT).ListIndex) Then Exit Sub

    'Verifica se existe o ítem na List da Combo. Se existir seleciona.
    lErro = Combo_Seleciona(CondicaoPagtoCC(objCT), iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 19223

    'Nao existe o ítem com o CÓDIGO na List da ComboBox
    If lErro = 6730 Then

        objCondicaoPagto.iCodigo = iCodigo

        'Tenta ler CondicaoPagto com esse código no BD
        lErro = CF("CondicaoPagto_Le", objCondicaoPagto)
        If lErro <> SUCESSO And lErro <> 19205 Then gError 19224
        
        'Não encontrou CondicaoPagto no BD
        If lErro <> SUCESSO Then Error 19225

        'Encontrou CondicaoPagto no BD e não é de Recebimento
        If objCondicaoPagto.iEmRecebimento = 0 Then gError 6927

        'Coloca no Text da Combo
        CondicaoPagtoCC(objCT).Text = CStr(objCondicaoPagto.iCodigo) & SEPARADOR & objCondicaoPagto.sDescReduzida

    End If

    'Não existe o ítem com a STRING na List da ComboBox
    If lErro = 6731 Then gError 19226

    Exit Sub

Erro_CondicaoPagtoCC_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 6927
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_RECEBIMENTO", gErr, iCodigo)

        Case 19223, 19224

        Case 19225  'Não encontrou CondicaoPagto no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_CONDICAO_PAGAMENTO")
            If vbMsgRes = vbYes Then
                'Chama a tela de CondicaoPagto
                Call Chama_Tela("CondicoesPagto", objCondicaoPagto)
            End If

        Case 19226
            Call Rotina_Erro(vbOKOnly, "ERRO_CONDICAO_PAGTO_NAO_ENCONTRADA", gErr, CondicaoPagtoCC(objCT).Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155556)

    End Select

    Exit Sub

End Sub

Public Sub CondicaoPagtoCCLabel_Click(ByVal objCT As Object)

Dim objCondicaoPagto As New ClassCondicaoPagto
Dim colSelecao As New Collection
    
    'Se a codicao estiver preenchida então passa para o objCondicaoPagamento
    If Len(Trim(CondicaoPagtoCC(objCT).Text)) > 0 Then objCondicaoPagto.iCodigo = Codigo_Extrai(CondicaoPagtoCC(objCT).Text)
   
    Set gobjCT = objCT

    Set objEventoCondicaoPagtoCC = New AdmEvento
    
    'Chama tela que lista as Condições de Pagamento
    Call Chama_Tela("CondicaoPagtoCRLista", colSelecao, objCondicaoPagto, objEventoCondicaoPagtoCC)

End Sub

Private Sub objEventoCondicaoPagtoCC_evSelecao(obj1 As Object)

Dim objCondicaoPagto As ClassCondicaoPagto
Dim bCancel As Boolean

    Set objCondicaoPagto = obj1

    'Preenche Text da Combo CondicaoPagto
    CondicaoPagtoCC(gobjCT).Text = CStr(objCondicaoPagto.iCodigo)
    
    'Chama rotina de Validate
    Call CondicaoPagtoCC_Validate(gobjCT, bCancel)

    gobjCT.objUserControl.Show

    Exit Sub

End Sub

Private Function Saida_Celula_ExcProduto(ByVal objCT As Object, objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Produto que está deixando de ser a corrente

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iIndice As Integer
Dim iProdutoPreenchido As Integer
Dim sProdutoEnxuto As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_ExcProduto

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.ClipText)) > 0 Then
    
        If objControle.Name = ExcAgProduto(objCT).Name Then
            If LCodigo_Extrai(EmpresaPai(objCT).Text) <> StrParaLong(objCT.Codigo.Text) Then
                Call Rotina_Aviso(vbOKOnly, "AVISO_ALTERACAO_SO_EMPPAI")
                gError 195791
            End If
        End If

        'Critica o Produto
        lErro = CF("Produto_Critica_Filial2", objControle.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 195791
        
        'Se o produto é gerencial ==> erro
        If lErro = 86295 Then gError 195792
               
        'Se o produto não foi encontrado ==> Pergunta se deseja criar
        If lErro = 51381 Then gError 195793

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 195794
    
        objControle.PromptInclude = False
        objControle.Text = sProdutoEnxuto
        objControle.PromptInclude = True
        
        'Verifica se já está em outra linha do Grid
        For iIndice = 1 To objGridInt.iLinhasExistentes
            If iIndice <> objGridInt.objGrid.Row Then
                If objGridInt.objGrid.TextMatrix(iIndice, objGridInt.objGrid.Col) = objControle.Text Then gError 195795
            End If
        Next
    
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195796

    Saida_Celula_ExcProduto = SUCESSO

    Exit Function

Erro_Saida_Celula_ExcProduto:

    Saida_Celula_ExcProduto = gErr

    Select Case gErr

        Case 195791, 195796
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 195792
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 195793
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto(objCT).Text)
            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 195794
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case 195795
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_PREENCHIDO_LINHA_GRID", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195797)

    End Select

    Exit Function

End Function

Function Saida_Celula_ExcPercComis(ByVal objCT As Object, objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Percentual Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double
Dim dPercentDescAnt As Double

On Error GoTo Erro_Saida_Celula_ExcPercComis

    Set objGridInt.objControle = objControle
    
    'MsgBox "1 - objControle.Text = " & objControle.Text

    dPercentDesc = StrParaDbl(objControle.Text)

    'MsgBox "2 - dPercentDesc = " & CStr(dPercentDesc)

    If objControle.Name = ExcAgPercComis(objCT).Name Then
    
        'MsgBox "2.1 "
    
        'dPercentDescAnt = PercentParaDbl(GridExcAg(objCT).TextMatrix(GridExcAg(objCT).Row, objCT.gobjInfoUsu.iGrid_ExcAgPercComiss_Col)) * 100
        dPercentDescAnt = StrParaDbl(GridExcAg(objCT).TextMatrix(GridExcAg(objCT).Row, objCT.gobjInfoUsu.iGrid_ExcAgPercComiss_Col))
    
        'MsgBox "2.2 "
    
        If Abs(dPercentDescAnt - dPercentDesc) > QTDE_ESTOQUE_DELTA Then
            If LCodigo_Extrai(EmpresaPai(objCT).Text) <> StrParaLong(objCT.Codigo.Text) Then
                Call Rotina_Aviso(vbOKOnly, "AVISO_ALTERACAO_SO_EMPPAI")
                objControle.Text = Format(dPercentDescAnt, "0.00##")
                gError 195798
            End If
        End If
        
        'MsgBox "2.3 "
        
    End If
        
    If Len(objControle.Text) > 0 Then
    
        'Critica a porcentagem
        'lErro = Porcentagem_Critica(objControle.Text)
        lErro = Porcentagem_Critica(CStr(dPercentDesc))
        If lErro <> SUCESSO Then gError 195798
       
        'MsgBox "3 "
       
        'MsgBox "4 - Format(dPercentDesc, 'General Number') = " & Format(dPercentDesc, "General Number")
       
        objControle.Text = Format(dPercentDesc, "0.00##")
     
    End If
    
    'MsgBox "5 "
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 195799

    Saida_Celula_ExcPercComis = SUCESSO

    Exit Function

Erro_Saida_Celula_ExcPercComis:

    Saida_Celula_ExcPercComis = gErr

    Select Case gErr

        Case 195798, 195799
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195800)

    End Select

    Exit Function

End Function

Function Saida_Celula_ExcCliente(ByVal objCT As Object, objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_Saida_Celula_ExcCliente

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.ClipText)) > 0 Then
    
        lErro = TP_Cliente_Le2(objControle, objCliente, 0)
        If lErro <> SUCESSO Then gError 192145

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 192146

    Saida_Celula_ExcCliente = SUCESSO

    Exit Function

Erro_Saida_Celula_ExcCliente:

    Saida_Celula_ExcCliente = gErr

    Select Case gErr

        Case 192145 To 192146
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192147)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridExcRep(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridExcRep

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for a de Cliente
        Case objCT.gobjInfoUsu.iGrid_ExcRep_Col
            lErro = Saida_Celula_ExcCliente(objCT, objGridInt, ExcRep(objCT))
            If lErro <> SUCESSO Then gError 195760
            
        'Se for a de Cliente
        Case objCT.gobjInfoUsu.iGrid_ExcRepProduto_Col
            lErro = Saida_Celula_ExcProduto(objCT, objGridInt, ExcRepProduto(objCT))
            If lErro <> SUCESSO Then gError 195760

        'Se for a de % comissão
        Case objCT.gobjInfoUsu.iGrid_ExcRepPercComiss_Col
            lErro = Saida_Celula_ExcPercComis(objCT, objGridInt, ExcRepPercComis(objCT))
            If lErro <> SUCESSO Then gError 195761

    End Select

    Saida_Celula_GridExcRep = SUCESSO

    Exit Function

Erro_Saida_Celula_GridExcRep:

    Saida_Celula_GridExcRep = gErr

    Select Case gErr

        Case 195760, 195761

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195762)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridExcCor(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridExcCor

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for a de Cliente
        Case objCT.gobjInfoUsu.iGrid_ExcCor_Col
            lErro = Saida_Celula_ExcCliente(objCT, objGridInt, ExcCor(objCT))
            If lErro <> SUCESSO Then gError 195760
            
        'Se for a de Cliente
        Case objCT.gobjInfoUsu.iGrid_ExcCorProduto_Col
            lErro = Saida_Celula_ExcProduto(objCT, objGridInt, ExcCorProduto(objCT))
            If lErro <> SUCESSO Then gError 195760

        'Se for a de % comissão
        Case objCT.gobjInfoUsu.iGrid_ExcCorPercComiss_Col
            lErro = Saida_Celula_ExcPercComis(objCT, objGridInt, ExcCorPercComis(objCT))
            If lErro <> SUCESSO Then gError 195761

    End Select

    Saida_Celula_GridExcCor = SUCESSO

    Exit Function

Erro_Saida_Celula_GridExcCor:

    Saida_Celula_GridExcCor = gErr

    Select Case gErr

        Case 195760, 195761

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195762)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridExcAg(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridExcAg

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col
            
        'Se for a de Cliente
        Case objCT.gobjInfoUsu.iGrid_ExcAgProduto_Col
            lErro = Saida_Celula_ExcProduto(objCT, objGridInt, ExcAgProduto(objCT))
            If lErro <> SUCESSO Then gError 195760

        'Se for a de % comissão
        Case objCT.gobjInfoUsu.iGrid_ExcAgPercComiss_Col
            lErro = Saida_Celula_ExcPercComis(objCT, objGridInt, ExcAgPercComis(objCT))
            If lErro <> SUCESSO Then gError 195761

    End Select

    Saida_Celula_GridExcAg = SUCESSO

    Exit Function

Erro_Saida_Celula_GridExcAg:

    Saida_Celula_GridExcAg = gErr

    Select Case gErr

        Case 195760, 195761

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195762)

    End Select

    Exit Function

End Function

Public Sub BotaoExcRep_Click(ByVal objCT As Object)

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
Dim sNomeBrowse As String


On Error GoTo Erro_BotaoExcRep_Click

    If objCT.objUserControl.ActiveControl Is Produto(objCT) Then
        objCliente.lCodigo = LCodigo_Extrai(ExcRep(objCT).Text)
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridExcRep(objCT).Row = 0 Then gError 195827
        objCliente.lCodigo = LCodigo_Extrai(GridExcRep(objCT).TextMatrix(GridExcRep(objCT).Row, objCT.gobjInfoUsu.iGrid_ExcRep_Col))
    End If
    
    sNomeBrowse = "ClientesLista"

    Call CF("Cliente_Obtem_NomeBrowse", sNomeBrowse)

    Set gobjCT = objCT

    Set objEventoExcRep = New AdmEvento

    'Chama Tela ClienteLista
    Call Chama_Tela(sNomeBrowse, colSelecao, objCliente, objEventoExcRep)
    
    Exit Sub
    
Erro_BotaoExcRep_Click:

    Select Case gErr
    
        Case 195827
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195829)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub BotaoExcCor_Click(ByVal objCT As Object)

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
Dim sNomeBrowse As String

On Error GoTo Erro_BotaoExcCor_Click

    If objCT.objUserControl.ActiveControl Is Produto(objCT) Then
        objCliente.lCodigo = LCodigo_Extrai(ExcCor(objCT).Text)
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridExcCor(objCT).Row = 0 Then gError 195827
        objCliente.lCodigo = LCodigo_Extrai(GridExcCor(objCT).TextMatrix(GridExcCor(objCT).Row, objCT.gobjInfoUsu.iGrid_ExcCor_Col))
    End If
    
    sNomeBrowse = "ClientesLista"

    Call CF("Cliente_Obtem_NomeBrowse", sNomeBrowse)

    Set gobjCT = objCT

    Set objEventoExcCor = New AdmEvento

    'Chama Tela ClienteLista
    Call Chama_Tela(sNomeBrowse, colSelecao, objCliente, objEventoExcCor)
    
    Exit Sub
    
Erro_BotaoExcCor_Click:

    Select Case gErr
    
        Case 195827
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195829)
    
    End Select
    
    Exit Sub
    
End Sub

Public Sub BotaoExcRepProduto_Click(ByVal objCT As Object)

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoExcRepProduto_Click

    If objCT.objUserControl.ActiveControl Is ExcRepProduto(objCT) Then
        sProduto1 = ExcRepProduto(objCT).Text
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridExcRep(objCT).Row = 0 Then gError 195827
        sProduto1 = GridExcRep(objCT).TextMatrix(GridExcRep(objCT).Row, objCT.gobjInfoUsu.iGrid_ExcRepProduto_Col)
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 195828
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto
    
    Set gobjCT = objCT

    Set objEventoExcRepProduto = New AdmEvento

    'Chama a Tela ProdutoVendaLista
    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoExcRepProduto)

    Exit Sub
    
Erro_BotaoExcRepProduto_Click:

    Select Case gErr
    
        Case 195827
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 195828
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195829)
    
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoExcCorProduto_Click(ByVal objCT As Object)

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoExcCorProduto_Click

    If objCT.objUserControl.ActiveControl Is ExcCorProduto(objCT) Then
        sProduto1 = ExcCorProduto(objCT).Text
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridExcCor(objCT).Row = 0 Then gError 195827
        sProduto1 = GridExcCor(objCT).TextMatrix(GridExcCor(objCT).Row, objCT.gobjInfoUsu.iGrid_ExcCorProduto_Col)
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 195828
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto
    
    Set gobjCT = objCT

    Set objEventoExcCorProduto = New AdmEvento

    'Chama a Tela ProdutoVendaLista
    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoExcCorProduto)

    Exit Sub
    
Erro_BotaoExcCorProduto_Click:

    Select Case gErr
    
        Case 195827
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 195828
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195829)
    
    End Select
    
    Exit Sub

End Sub

Public Sub BotaoExcAgProduto_Click(ByVal objCT As Object)

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoExcAgProduto_Click

    If objCT.objUserControl.ActiveControl Is ExcAgProduto(objCT) Then
        sProduto1 = ExcAgProduto(objCT).Text
    Else
        'Verifica se tem alguma linha selecionada no Grid
        If GridExcAg(objCT).Row = 0 Then gError 195827
        sProduto1 = GridExcAg(objCT).TextMatrix(GridExcAg(objCT).Row, objCT.gobjInfoUsu.iGrid_ExcAgProduto_Col)
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 195828
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto
    
    Set gobjCT = objCT

    Set objEventoExcAgProduto = New AdmEvento

    'Chama a Tela ProdutoVendaLista
    Call Chama_Tela("ProdutoLista1", colSelecao, objProduto, objEventoExcAgProduto)

    Exit Sub
    
Erro_BotaoExcAgProduto_Click:

    Select Case gErr
    
        Case 195827
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 195828
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195829)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoExcRepProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProdutoEnxuto As String
Dim iIndice As Integer

On Error GoTo Erro_objEventoExcRepProduto_evSelecao

    'verifica se tem alguma linha do Grid selecionada
    If GridExcRep(gobjCT).Row > 0 Then

        Set objProduto = obj1

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 195831

        ExcRepProduto(gobjCT).PromptInclude = False
        ExcRepProduto(gobjCT).Text = sProdutoEnxuto
        ExcRepProduto(gobjCT).PromptInclude = True
        
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 195832
        
        If Not (gobjCT.objUserControl.ActiveControl Is ExcRepProduto(gobjCT)) Then

            'Verifica se já está em outra linha do Grid
            For iIndice = 1 To gobjCT.gobjInfoUsu.objGridExcRep.iLinhasExistentes
                If iIndice <> GridExcRep(gobjCT).Row Then
                    If GridExcRep(gobjCT).TextMatrix(iIndice, gobjCT.gobjInfoUsu.iGrid_ExcRepProduto_Col) = ExcRepProduto(gobjCT).Text Then gError 195839
                End If
            Next
            
            GridExcRep(gobjCT).TextMatrix(GridExcRep(gobjCT).Row, gobjCT.gobjInfoUsu.iGrid_ExcRepProduto_Col) = ExcRepProduto(gobjCT).Text
            
            If GridExcRep(gobjCT).Row - GridExcRep(gobjCT).FixedRows = gobjCT.gobjInfoUsu.objGridExcRep.iLinhasExistentes Then
                gobjCT.gobjInfoUsu.objGridExcRep.iLinhasExistentes = gobjCT.gobjInfoUsu.objGridExcRep.iLinhasExistentes + 1
            End If
    
        End If

    End If

    gobjCT.objUserControl.Show

    Exit Sub

Erro_objEventoExcRepProduto_evSelecao:

    GridExcRep(gobjCT).TextMatrix(GridExcRep(gobjCT).Row, gobjCT.gobjInfoUsu.iGrid_ExcRepProduto_Col) = ""

    Select Case gErr

        Case 195830, 195832
        
        Case 195831
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case 195839
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_PREENCHIDO_LINHA_GRID", gErr, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195833)

    End Select

    Exit Sub

End Sub

Private Sub objEventoExcCorProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProdutoEnxuto As String
Dim iIndice As Integer

On Error GoTo Erro_objEventoExcCorProduto_evSelecao

    'verifica se tem alguma linha do Grid selecionada
    If GridExcCor(gobjCT).Row > 0 Then

        Set objProduto = obj1

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 195831

        ExcCorProduto(gobjCT).PromptInclude = False
        ExcCorProduto(gobjCT).Text = sProdutoEnxuto
        ExcCorProduto(gobjCT).PromptInclude = True
        
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 195832
        
        If Not (gobjCT.objUserControl.ActiveControl Is ExcCorProduto(gobjCT)) Then
    
            'Verifica se já está em outra linha do Grid
            For iIndice = 1 To gobjCT.gobjInfoUsu.objGridExcCor.iLinhasExistentes
                If iIndice <> GridExcCor(gobjCT).Row Then
                    If GridExcCor(gobjCT).TextMatrix(iIndice, gobjCT.gobjInfoUsu.iGrid_ExcCorProduto_Col) = ExcCorProduto(gobjCT).Text Then gError 195839
                End If
            Next
            
            GridExcCor(gobjCT).TextMatrix(GridExcCor(gobjCT).Row, gobjCT.gobjInfoUsu.iGrid_ExcCorProduto_Col) = ExcCorProduto(gobjCT).Text
            
            If GridExcCor(gobjCT).Row - GridExcCor(gobjCT).FixedRows = gobjCT.gobjInfoUsu.objGridExcCor.iLinhasExistentes Then
                gobjCT.gobjInfoUsu.objGridExcCor.iLinhasExistentes = gobjCT.gobjInfoUsu.objGridExcCor.iLinhasExistentes + 1
            End If
    
        End If

    End If

    gobjCT.objUserControl.Show

    Exit Sub

Erro_objEventoExcCorProduto_evSelecao:

    GridExcCor(gobjCT).TextMatrix(GridExcCor(gobjCT).Row, gobjCT.gobjInfoUsu.iGrid_ExcCorProduto_Col) = ""

    Select Case gErr

        Case 195830, 195832
        
        Case 195831
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case 195839
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_PREENCHIDO_LINHA_GRID", gErr, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195833)

    End Select

    Exit Sub

End Sub

Private Sub objEventoExcAgProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProdutoEnxuto As String
Dim iIndice As Integer

On Error GoTo Erro_objEventoExcAgProduto_evSelecao

    'verifica se tem alguma linha do Grid selecionada
    If GridExcAg(gobjCT).Row > 0 Then

        Set objProduto = obj1

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 195831

        ExcAgProduto(gobjCT).PromptInclude = False
        ExcAgProduto(gobjCT).Text = sProdutoEnxuto
        ExcAgProduto(gobjCT).PromptInclude = True
        
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 195832
        
        If Not (gobjCT.objUserControl.ActiveControl Is ExcAgProduto(gobjCT)) Then
    
            'Verifica se já está em outra linha do Grid
            For iIndice = 1 To gobjCT.gobjInfoUsu.objGridExcAg.iLinhasExistentes
                If iIndice <> GridExcAg(gobjCT).Row Then
                    If GridExcAg(gobjCT).TextMatrix(iIndice, gobjCT.gobjInfoUsu.iGrid_ExcAgProduto_Col) = ExcAgProduto(gobjCT).Text Then gError 195839
                End If
            Next
            
            GridExcAg(gobjCT).TextMatrix(GridExcAg(gobjCT).Row, gobjCT.gobjInfoUsu.iGrid_ExcAgProduto_Col) = ExcAgProduto(gobjCT).Text
            
            If GridExcAg(gobjCT).Row - GridExcAg(gobjCT).FixedRows = gobjCT.gobjInfoUsu.objGridExcAg.iLinhasExistentes Then
                gobjCT.gobjInfoUsu.objGridExcAg.iLinhasExistentes = gobjCT.gobjInfoUsu.objGridExcAg.iLinhasExistentes + 1
            End If
    
        End If

    End If

    gobjCT.objUserControl.Show

    Exit Sub

Erro_objEventoExcAgProduto_evSelecao:

    GridExcAg(gobjCT).TextMatrix(GridExcAg(gobjCT).Row, gobjCT.gobjInfoUsu.iGrid_ExcAgProduto_Col) = ""

    Select Case gErr

        Case 195830, 195832
        
        Case 195831
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case 195839
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_PREENCHIDO_LINHA_GRID", gErr, iIndice)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195833)

    End Select

    Exit Sub

End Sub

Private Sub objEventoExcRep_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    If gobjCT.ActiveControl Is ExcRep(gobjCT) Then
        ExcRep(gobjCT).Text = objCliente.lCodigo & SEPARADOR & objCliente.sNomeReduzido
    Else
        GridExcRep(gobjCT).TextMatrix(GridExcRep(gobjCT).Row, gobjCT.gobjInfoUsu.iGrid_ExcRep_Col) = objCliente.lCodigo & SEPARADOR & objCliente.sNomeReduzido
    End If
    
    If GridExcRep(gobjCT).Row - GridExcRep(gobjCT).FixedRows = gobjCT.gobjInfoUsu.objGridExcRep.iLinhasExistentes Then
        gobjCT.gobjInfoUsu.objGridExcRep.iLinhasExistentes = gobjCT.gobjInfoUsu.objGridExcRep.iLinhasExistentes + 1
    End If

    gobjCT.objUserControl.Show

    Exit Sub

End Sub

Private Sub objEventoExcCor_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente
Dim bCancel As Boolean

    Set objCliente = obj1

    'Preenche campo Cliente
    If gobjCT.ActiveControl Is ExcCor(gobjCT) Then
        ExcCor(gobjCT).Text = objCliente.lCodigo & SEPARADOR & objCliente.sNomeReduzido
    Else
        GridExcCor(gobjCT).TextMatrix(GridExcCor(gobjCT).Row, gobjCT.gobjInfoUsu.iGrid_ExcCor_Col) = objCliente.lCodigo & SEPARADOR & objCliente.sNomeReduzido
    End If
    
    If GridExcCor(gobjCT).Row - GridExcCor(gobjCT).FixedRows = gobjCT.gobjInfoUsu.objGridExcCor.iLinhasExistentes Then
        gobjCT.gobjInfoUsu.objGridExcCor.iLinhasExistentes = gobjCT.gobjInfoUsu.objGridExcCor.iLinhasExistentes + 1
    End If

    gobjCT.objUserControl.Show

    Exit Sub

End Sub

Public Sub ExcAgProduto_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ExcAgProduto_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridExcAg)
End Sub

Public Sub ExcAgProduto_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridExcAg)
End Sub

Public Sub ExcAgProduto_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ExcAgProduto_Validate

    Set objCT.gobjInfoUsu.objGridExcAg.objControle = ExcAgProduto(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridExcAg)
    If lErro <> SUCESSO Then gError 195835
    
    Exit Sub
    
Erro_ExcAgProduto_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195835
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195836)
     
    End Select
    
End Sub

Public Sub ExcAgPercComis_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ExcAgPercComis_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridExcAg)
End Sub

Public Sub ExcAgPercComis_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridExcAg)
End Sub

Public Sub ExcAgPercComis_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ExcAgPercComis_Validate

    Set objCT.gobjInfoUsu.objGridExcAg.objControle = ExcAgPercComis(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridExcAg)
    If lErro <> SUCESSO Then gError 195835
    
    Exit Sub
    
Erro_ExcAgPercComis_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195835
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195836)
     
    End Select
    
End Sub


Public Sub GridExcAg_Click(ByVal objCT As Object)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objCT.gobjInfoUsu.objGridExcAg, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridExcAg, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridExcAg_EnterCell(ByVal objCT As Object)
    Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridExcAg, objCT.iAlterado)
End Sub

Public Sub GridExcAg_GotFocus(ByVal objCT As Object)
    Call Grid_Recebe_Foco(objCT.gobjInfoUsu.objGridExcAg)
End Sub

Public Sub GridExcAg_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)

Dim bNaoExclui As Boolean

    bNaoExclui = False

    If KeyCode = vbKeyDelete Then
    
        If LCodigo_Extrai(EmpresaPai(objCT).Text) <> StrParaLong(objCT.Codigo.Text) Then
        
            Call Rotina_Aviso(vbOKOnly, "AVISO_ALTERACAO_SO_EMPPAI")
            bNaoExclui = True
        End If
    
    End If
    
    If Not bNaoExclui Then Call Grid_Trata_Tecla1(KeyCode, objCT.gobjInfoUsu.objGridExcAg)
End Sub

Public Sub GridExcAg_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objCT.gobjInfoUsu.objGridExcAg, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridExcAg, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridExcAg_LeaveCell(ByVal objCT As Object)
    Call objCT.Saida_Celula(objCT.gobjInfoUsu.objGridExcAg)
End Sub

Public Sub GridExcAg_Validate(ByVal objCT As Object, Cancel As Boolean)
    Call Grid_Libera_Foco(objCT.gobjInfoUsu.objGridExcAg)
End Sub

Public Sub GridExcAg_RowColChange(ByVal objCT As Object)
    Call Grid_RowColChange(objCT.gobjInfoUsu.objGridExcAg)
End Sub

Public Sub GridExcAg_Scroll(ByVal objCT As Object)
    Call Grid_Scroll(objCT.gobjInfoUsu.objGridExcAg)
End Sub

Public Sub GridCallCenter_Click(ByVal objCT As Object)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objCT.gobjInfoUsu.objGridCallCenter, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridCallCenter, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridCallCenter_EnterCell(ByVal objCT As Object)
    Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridCallCenter, objCT.iAlterado)
End Sub

Public Sub GridCallCenter_GotFocus(ByVal objCT As Object)
    Call Grid_Recebe_Foco(objCT.gobjInfoUsu.objGridCallCenter)
End Sub

Public Sub GridCallCenter_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objCT.gobjInfoUsu.objGridCallCenter)
End Sub

Public Sub GridCallCenter_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objCT.gobjInfoUsu.objGridCallCenter, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjInfoUsu.objGridCallCenter, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridCallCenter_LeaveCell(ByVal objCT As Object)
    Call objCT.Saida_Celula(objCT.gobjInfoUsu.objGridCallCenter)
End Sub

Public Sub GridCallCenter_Validate(ByVal objCT As Object, Cancel As Boolean)
    Call Grid_Libera_Foco(objCT.gobjInfoUsu.objGridCallCenter)
End Sub

Public Sub GridCallCenter_RowColChange(ByVal objCT As Object)
    Call Grid_RowColChange(objCT.gobjInfoUsu.objGridCallCenter)
End Sub

Public Sub GridCallCenter_Scroll(ByVal objCT As Object)
    Call Grid_Scroll(objCT.gobjInfoUsu.objGridCallCenter)
End Sub

Public Sub DataCallCenterDe_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataCallCenterDe_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridCallCenter)
End Sub

Public Sub DataCallCenterDe_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridCallCenter)
End Sub

Public Sub DataCallCenterDe_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataCallCenterDe_Validate

    Set objCT.gobjInfoUsu.objGridCallCenter.objControle = DataCallCenterDe(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridCallCenter)
    If lErro <> SUCESSO Then gError 195837
    
    Exit Sub
    
Erro_DataCallCenterDe_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195837
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195834)
     
    End Select
    
End Sub

Public Sub DataCallCenterAte_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataCallCenterAte_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridCallCenter)
End Sub

Public Sub DataCallCenterAte_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridCallCenter)
End Sub

Public Sub DataCallCenterAte_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataCallCenterAte_Validate

    Set objCT.gobjInfoUsu.objGridCallCenter.objControle = DataCallCenterAte(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridCallCenter)
    If lErro <> SUCESSO Then gError 195837
    
    Exit Sub
    
Erro_DataCallCenterAte_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195837
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195834)
     
    End Select
    
End Sub

Function Saida_Celula_Data(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Data

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.ClipText)) > 0 Then
    
        'Critica a Data informada
        lErro = Data_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 190671

        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 190673

    Saida_Celula_Data = SUCESSO

    Exit Function

Erro_Saida_Celula_Data:

    Saida_Celula_Data = gErr

    Select Case gErr

        Case 190671 To 190673
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190674)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Sub EmiPercCI_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub EmiPercCI_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridComissao)
End Sub

Public Sub EmiPercCI_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridComissao)
End Sub

Public Sub EmiPercCI_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmiPercCI_Validate

    Set objCT.gobjInfoUsu.objGridComissao.objControle = EmiPercCI(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridComissao)
    If lErro <> SUCESSO Then gError 195837
    
    Exit Sub
    
Erro_EmiPercCI_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195837
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195834)
     
    End Select
    
End Sub

Public Sub EmiCargo_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub EmiCargo_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridComissao)
End Sub

Public Sub EmiCargo_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridComissao)
End Sub

Public Sub EmiCargo_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmiCargo_Validate

    Set objCT.gobjInfoUsu.objGridComissao.objControle = EmiCargo(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridComissao)
    If lErro <> SUCESSO Then gError 195837
    
    Exit Sub
    
Erro_EmiCargo_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195837
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195834)
     
    End Select
    
End Sub

Public Sub EmiCPF_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub EmiCPF_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridComissao)
End Sub

Public Sub EmiCPF_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridComissao)
End Sub

Public Sub EmiCPF_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmiCPF_Validate

    Set objCT.gobjInfoUsu.objGridComissao.objControle = EmiCPF(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridComissao)
    If lErro <> SUCESSO Then gError 195837
    
    Exit Sub
    
Erro_EmiCPF_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195837
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195834)
     
    End Select
    
End Sub

Public Sub EmiCartao_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub EmiCartao_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjInfoUsu.objGridComissao)
End Sub

Public Sub EmiCartao_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjInfoUsu.objGridComissao)
End Sub

Public Sub EmiCartao_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmiCartao_Validate

    Set objCT.gobjInfoUsu.objGridComissao.objControle = EmiCartao(objCT)
    
    lErro = Grid_Campo_Libera_Foco(objCT.gobjInfoUsu.objGridComissao)
    If lErro <> SUCESSO Then gError 195837
    
    Exit Sub
    
Erro_EmiCartao_Validate:

    Cancel = True

    Select Case gErr
    
        Case 195837
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195834)
     
    End Select
    
End Sub

Function Saida_Celula_CPF(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long
Dim sCgc As String

On Error GoTo Erro_Saida_Celula_CPF

    Set objGridInt.objControle = objControle
    
    sCgc = Replace(Replace(Replace(Replace(objControle.Text, ".", ""), "-", ""), "/", ""), "\", "")
    
    If Len(Trim(sCgc)) > 0 Then
    
        Select Case Len(Trim(sCgc))
    
            Case STRING_CPF 'CPF
                
                'Critica Cpf
                lErro = Cpf_Critica(sCgc)
                If lErro <> SUCESSO Then gError 190672
                
                'Formata e coloca na Tela
                objControle.Text = Format(sCgc, "000\.000\.000-00; ; ; ")

    
            Case STRING_CGC 'CGC
                
                'Critica CGC
                lErro = Cgc_Critica(sCgc)
                If lErro <> SUCESSO Then gError 190672
                
                'Formata e Coloca na Tela
                objControle.Text = Format(sCgc, "00\.000\.000\/0000-00; ; ; ")
    
            Case Else
                gError 190671
    
        End Select
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 190673

    Saida_Celula_CPF = SUCESSO

    Exit Function

Erro_Saida_Celula_CPF:

    Saida_Celula_CPF = gErr

    Select Case gErr

        Case 190671
            Call Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", Err)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 190672
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190674)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Sub PercFatorDevCMCC_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_PercFatorDevCMCC_Validate

    'Verifica se foi preenchido a Comissao de Venda
    If Len(Trim(PercFatorDevCMCC(objCT).Text)) = 0 Then Exit Sub

    'Critica se é porcentagem
    lErro = Porcentagem_Critica(PercFatorDevCMCC(objCT).Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Formata
    PercFatorDevCMCC(objCT).Text = Format(PercFatorDevCMCC(objCT).Text, "Fixed")

    Exit Sub

Erro_PercFatorDevCMCC_Validate:

    Cancel = True

    Select Case Err

        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 195850)
    
    End Select

    Exit Sub

End Sub

Public Sub PercFatorDevCMCC_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub
