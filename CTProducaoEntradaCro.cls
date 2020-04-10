VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTProdEntCro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Property Get Responsavel(ByVal objCT As Object) As Object
     Set Responsavel = objCT.objUserControl.Controls("Responsavel")
End Property

Public Property Get Localizacao(ByVal objCT As Object) As Object
     Set Localizacao = objCT.objUserControl.Controls("Localizacao")
End Property

Public Sub Responsavel_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Responsavel_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Responsavel_Validate
 
    Exit Sub

Erro_Responsavel_Validate:

    Cancel = True

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Sub

End Sub

Public Function ProdEnt_Move_Tela_Memoria(ByVal objCT As Object, ByVal objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim objMovEstoqueCro As New ClassMovEstoqueCro

On Error GoTo Erro_ProdEnt_Move_Tela_Memoria

    objMovEstoqueCro.sResponsavel = Responsavel(objCT).Text
    
    Set objMovEstoque.objInfoUsu = objMovEstoqueCro

    ProdEnt_Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_ProdEnt_Move_Tela_Memoria:

    ProdEnt_Move_Tela_Memoria = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function
    
End Function

Public Function ProdEnt_Preenche_Tela(ByVal objCT As Object, ByVal objMovEstoque As ClassMovEstoque) As Long

Dim lErro As Long
Dim objMovEstoqueCro As ClassMovEstoqueCro

On Error GoTo Erro_ProdEnt_Preenche_Tela

    If Not (objMovEstoque.objInfoUsu Is Nothing) Then
    
        Set objMovEstoqueCro = objMovEstoque.objInfoUsu
    
        Responsavel(objCT).Text = objMovEstoqueCro.sResponsavel
    
    End If

    ProdEnt_Preenche_Tela = SUCESSO

    Exit Function

Erro_ProdEnt_Preenche_Tela:

    ProdEnt_Preenche_Tela = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function
    
End Function

Public Function ProdEnt_Inicializa_GridMovimentos(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = objCT

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Ordem Produçao")
    objGridInt.colColuna.Add ("Lote/ Série Inicial") 'Alterado por Wagner 14/03/2006
    objGridInt.colColuna.Add ("Almoxarifado")
    objGridInt.colColuna.Add ("Benef.")
    objGridInt.colColuna.Add ("Horas Máquina")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("Conta Contábil de Estoque")
    objGridInt.colColuna.Add ("Conta Contabil de Produção")
    objGridInt.colColuna.Add ("Estorno")
    objGridInt.colColuna.Add ("Localização")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (objCT.Produto.Name)
    objGridInt.colCampo.Add (objCT.DescricaoItem.Name)
    objGridInt.colCampo.Add (objCT.UnidadeMed.Name)
    objGridInt.colCampo.Add (objCT.Quantidade.Name)
    objGridInt.colCampo.Add (objCT.OPCodigo.Name)
    objGridInt.colCampo.Add (objCT.Lote.Name)
    objGridInt.colCampo.Add (objCT.Almoxarifado.Name)
    objGridInt.colCampo.Add (objCT.Benef.Name)
    objGridInt.colCampo.Add (objCT.HorasMaq.Name)
    objGridInt.colCampo.Add (objCT.DataProducao.Name)
    objGridInt.colCampo.Add (objCT.Ccl.Name)
    objGridInt.colCampo.Add (objCT.ContaContabilEst.Name)
    objGridInt.colCampo.Add (objCT.ContaContabilProducao.Name)
    objGridInt.colCampo.Add (objCT.Estorno.Name)
    objGridInt.colCampo.Add (Localizacao(objCT).Name)
    
    'Colunas do Grid
    objCT.iGrid_Sequencial_Col = 0
    objCT.iGrid_Produto_Col = 1
    objCT.iGrid_Descricao_Col = 2
    objCT.iGrid_UnidadeMed_Col = 3
    objCT.iGrid_Quantidade_Col = 4
    objCT.iGrid_OP_Col = 5
    objCT.iGrid_Lote_Col = 6
    objCT.iGrid_Almoxarifado_Col = 7
    objCT.iGrid_Benef_Col = 8
    objCT.iGrid_HorasMaquina_Col = 9
    objCT.iGrid_DataProducao_Col = 10
    objCT.iGrid_Ccl_Col = 11
    objCT.iGrid_ContaContabilEst_Col = 12
    objCT.iGrid_ContaContabilProducao_Col = 13
    objCT.iGrid_Estorno_Col = 14
    objCT.gobjInfoUsu.iGrid_Localizacao_Col = 15
    
    'Grid do GridInterno
    objGridInt.objGrid = objCT.GridMovimentos

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAX_ITENS_MOV_ESTOQUE + 1

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    objCT.GridMovimentos.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    ProdEnt_Inicializa_GridMovimentos = SUCESSO

    Exit Function

End Function

Public Function ProdEnt_Preenche_GridMovimentos(ByVal objCT As Object, ByVal objItemMovEstoque As ClassItemMovEstoque, ByVal iIndice As Integer) As Long

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objRastroLoteCro As ClassRastreamentoLoteCro

On Error GoTo Erro_ProdEnt_Preenche_GridMovimentos

    objRastroLote.sProduto = objItemMovEstoque.sProduto
    objRastroLote.iFilialOP = giFilialEmpresa
    objRastroLote.sCodigo = objItemMovEstoque.sOPCodigo
       
    lErro = CF("RastreamentoLote_Le", objRastroLote)
    If lErro <> SUCESSO And lErro <> 75710 Then gError 182760

    lErro = CF("RastreamentoLote_Le_Customizado", objRastroLote)
    If lErro <> SUCESSO Then gError 185909

    Set objRastroLoteCro = objRastroLote.objInfoUsu
    
    objCT.GridMovimentos.TextMatrix(iIndice, objCT.gobjInfoUsu.iGrid_Localizacao_Col) = objRastroLoteCro.sLocalizacao

    ProdEnt_Preenche_GridMovimentos = SUCESSO

    Exit Function

Erro_ProdEnt_Preenche_GridMovimentos:

    ProdEnt_Preenche_GridMovimentos = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error)

    End Select

    Exit Function
    
End Function

Public Function ProdEnt_Rotina_Grid_Enable(ByVal objCT As Object, ByVal iLinha As Integer, ByVal objControl As Object, ByVal iLocalChamada As Integer, ByVal iProdutoPreenchido As Integer, ByVal sProdutoFormatado As String, ByVal lNumIntDoc As Long) As Long

Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_ProdEnt_Rotina_Grid_Enable

    If objControl.Name = "Localizacao" Then

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
            objProduto.sCodigo = sProdutoFormatado
    
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 78342
        
            If objProduto.iRastro = PRODUTO_RASTRO_OP And lNumIntDoc = 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
        Else
            objControl.Enabled = False
        End If
        
    End If
    
    ProdEnt_Rotina_Grid_Enable = SUCESSO
        
    Exit Function

Erro_ProdEnt_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr)

    End Select

    Exit Function

End Function

Public Function ProdEnt_Move_RastroEstoque_Memoria(ByVal objCT As Object, ByVal iLinha As Integer, ByVal objRastreamentoMovto As ClassRastreamentoMovto) As Long

Dim objRastroLoteCro As New ClassRastreamentoLoteCro

     objRastroLoteCro.sLocalizacao = objCT.GridMovimentos.TextMatrix(iLinha, objCT.gobjInfoUsu.iGrid_Localizacao_Col)
     
     Set objRastreamentoMovto.objInfoUsu = objRastroLoteCro
     
     ProdEnt_Move_RastroEstoque_Memoria = SUCESSO
     
End Function

Public Sub Localizacao_Change(ByVal objCT As Object)

    objCT.iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Localizacao_GotFocus(ByVal objCT As Object)

    Call Grid_Campo_Recebe_Foco(objCT.objGrid)

End Sub

Public Sub Localizacao_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.objGrid)

End Sub

Public Sub Localizacao_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.objGrid.objControle = Localizacao(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function ProdEnt_Saida_Celula_Grid(ByVal objCT As Object, objGridInt As AdmGrid) As Long

    Select Case objCT.GridMovimentos.Col
    
        Case objCT.gobjInfoUsu.iGrid_Localizacao_Col
        
            Set objGridInt.objControle = Localizacao(objCT)

    End Select

End Function
