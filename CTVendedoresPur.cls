VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTVendedoresPur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim gobjCT As Object

Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

Public Property Get DataInicio(ByVal objCT As Object) As Object
     Set DataInicio = objCT.objUserControl.Controls("DataInicio")
End Property

Public Property Get Obs(ByVal objCT As Object) As Object
     Set Obs = objCT.objUserControl.Controls("Obs")
End Property

Public Property Get ExcProd(ByVal objCT As Object) As Object
     Set ExcProd = objCT.objUserControl.Controls("ExcProd")
End Property

Public Property Get ExcProdDesc(ByVal objCT As Object) As Object
     Set ExcProdDesc = objCT.objUserControl.Controls("ExcProdDesc")
End Property

Public Property Get ExcMeta(ByVal objCT As Object) As Object
     Set ExcMeta = objCT.objUserControl.Controls("ExcMeta")
End Property

Public Property Get ExcPercComiss(ByVal objCT As Object) As Object
     Set ExcPercComiss = objCT.objUserControl.Controls("ExcPercComiss")
End Property

Public Property Get ExcMetaP(ByVal objCT As Object) As Object
     Set ExcMetaP = objCT.objUserControl.Controls("ExcMetaP")
End Property

Public Property Get ExcPercComissP(ByVal objCT As Object) As Object
     Set ExcPercComissP = objCT.objUserControl.Controls("ExcPercComissP")
End Property

Public Property Get GridExc(ByVal objCT As Object) As Object
     Set GridExc = objCT.objUserControl.Controls("GridExc")
End Property

Public Function VND_Form_Load(ByVal objCT As Object) As Long

Dim lErro As Long

On Error GoTo Erro_VND_Form_Load

    Set objCT.gobjinfousu.objGridExc = New AdmGrid
    
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ExcProd(objCT))
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Inicializa_Grid_Excecoes(objCT, objCT.gobjinfousu.objGridExc)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    VND_Form_Load = SUCESSO
     
    Exit Function
    
Erro_VND_Form_Load:

    VND_Form_Load = gErr
     
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195888)
     
    End Select
     
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
    objGridInt.colColuna.Add ("%Padrão")
    objGridInt.colColuna.Add ("%Comissão")
    objGridInt.colColuna.Add ("Meta Padrão")
    objGridInt.colColuna.Add ("Meta Mínima")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (ExcProd(objCT).Name)
    objGridInt.colCampo.Add (ExcProdDesc(objCT).Name)
    objGridInt.colCampo.Add (ExcPercComissP(objCT).Name)
    objGridInt.colCampo.Add (ExcPercComiss(objCT).Name)
    objGridInt.colCampo.Add (ExcMetaP(objCT).Name)
    objGridInt.colCampo.Add (ExcMeta(objCT).Name)

    'Colunas do Grid
    objCT.gobjinfousu.iGrid_ExcProd_Col = 1
    objCT.gobjinfousu.iGrid_ExcProdDesc_Col = 2
    objCT.gobjinfousu.iGrid_ExcPercComissP_Col = 3
    objCT.gobjinfousu.iGrid_ExcPercComiss_Col = 4
    objCT.gobjinfousu.iGrid_ExcMetaP_Col = 5
    objCT.gobjinfousu.iGrid_ExcMeta_Col = 6

    'Grid do GridInterno
    objGridInt.objGrid = GridExc(objCT)

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 351

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridExc(objCT).ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Excecoes = SUCESSO

    Exit Function

End Function

Public Function VND_Grava_Registro(ByVal objCT As Object, objVendedor As ClassVendedor) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objVendedorPur As New ClassVendedorPur
Dim objVendPurExc As ClassVendedorPurExc
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_VND_Grava_Registro
    
    objVendedorPur.dtDataInicio = StrParaDate(DataInicio(objCT).Text)
    objVendedorPur.sObservacao = Obs(objCT).Text
    
    'Verifica se no grid de comissao as informacoes estao preenchidas
    For iIndice = 1 To objCT.gobjinfousu.objGridExc.iLinhasExistentes
        
        If Len(Trim(GridExc(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ExcProd_Col))) <> 0 Then
        
            Set objVendPurExc = New ClassVendedorPurExc
                
            lErro = CF("Produto_Formata", GridExc(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ExcProd_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            objVendPurExc.sProduto = sProdutoFormatado
            
            If Len(Trim(GridExc(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ExcMeta_Col))) > 0 Then
                objVendPurExc.dMetaComissao = StrParaDbl(GridExc(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ExcMeta_Col))
            Else
                objVendPurExc.dMetaComissao = -1 'Não preecnhido
            End If
            
            If Len(Trim(GridExc(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ExcPercComiss_Col))) > 0 Then
                objVendPurExc.dPercComissao = PercentParaDbl(GridExc(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ExcPercComiss_Col))
            Else
                objVendPurExc.dPercComissao = -1 'Não preecnhido
            End If
            
            objVendedorPur.colExcecoes.Add objVendPurExc
            
        End If
         
    Next
    
    Set objVendedor.objInfoUsu = objVendedorPur

    VND_Grava_Registro = SUCESSO
    
    Exit Function
    
Erro_VND_Grava_Registro:
    
    VND_Grava_Registro = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195871)

    End Select

    Exit Function

End Function

Public Function VND_Traz_Vendedor_Tela(ByVal objCT As Object, ByVal objVendedor As ClassVendedor) As Long

Dim lErro As Long
Dim iCodigo As Integer
Dim bCancel As Boolean
Dim objVendedorPur As New ClassVendedorPur
Dim objVendPurExc As ClassVendedorPurExc
Dim objProd As ClassProduto
Dim iIndice As Integer, sProdutoEnxuto As String

On Error GoTo Erro_VND_Traz_Vendedor_Tela

    lErro = CF("Vendedor_Le_Customizado", objVendedor)
    If lErro <> SUCESSO Then gError 195877
    
    If Not (objVendedor.objInfoUsu Is Nothing) Then
    
        Set objVendedorPur = objVendedor.objInfoUsu
        
        Call DateParaMasked(DataInicio(objCT), objVendedorPur.dtDataInicio)
        
        Obs(objCT).Text = objVendedorPur.sObservacao
       
        iIndice = 0
        
        Call Grid_Limpa(objCT.gobjinfousu.objGridExc)
        
        For Each objVendPurExc In objVendedorPur.colExcecoes
        
            iIndice = iIndice + 1
        
            If objVendPurExc.dMetaComissao <> -1 Then GridExc(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ExcMeta_Col) = Formata_Estoque(objVendPurExc.dMetaComissao)
            If objVendPurExc.dPercComissao <> -1 Then GridExc(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ExcPercComiss_Col) = Format(objVendPurExc.dPercComissao, "Percent")

            lErro = Mascara_RetornaProdutoEnxuto(objVendPurExc.sProduto, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError 195802

            ExcProd(objCT).promptinclude = False
            ExcProd(objCT).Text = sProdutoEnxuto
            ExcProd(objCT).promptinclude = True
            
            Set objProd = New ClassProduto
    
            objProd.sCodigo = objVendPurExc.sProduto
            
            'Lê o Produto
            lErro = CF("Produto_Le", objProd)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            
            'Preenche o Grid
            GridExc(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ExcProd_Col) = ExcProd(objCT).Text
            GridExc(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ExcProdDesc_Col) = objProd.sDescricao

            GridExc(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ExcMetaP_Col) = Formata_Estoque(objProd.dMetaComissao)
            GridExc(objCT).TextMatrix(iIndice, objCT.gobjinfousu.iGrid_ExcPercComissP_Col) = Format(objProd.dPercComissao, "Percent")

        Next
        
        objCT.gobjinfousu.objGridExc.iLinhasExistentes = iIndice
                
    End If
       
    VND_Traz_Vendedor_Tela = SUCESSO
     
    Exit Function
    
Erro_VND_Traz_Vendedor_Tela:

    VND_Traz_Vendedor_Tela = gErr
     
    Select Case gErr
    
        Case 195877, 195878, 195976, 195977, 195978
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195879)
     
    End Select
     
    Exit Function

End Function

Public Function VND_Limpa_Tela_Vendedor(ByVal objCT As Object) As Long
        
    Call Grid_Limpa(objCT.gobjinfousu.objGridExc)

    VND_Limpa_Tela_Vendedor = SUCESSO
    
End Function

Public Sub Obs_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Obs_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Obs_Validate


    Exit Sub

Erro_Obs_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195886)
    
    End Select

End Sub

Public Sub DataInicio_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataInicio_GotFocus(ByVal objCT As Object)
    Call MaskEdBox_TrataGotFocus(DataInicio(objCT), objCT.iAlterado)
End Sub

Public Sub DataInicio_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicio_Validate

    If Len(Trim(DataInicio(objCT).Text)) > 0 Then

        lErro = Data_Critica(DataInicio(objCT).Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    Exit Sub

Erro_DataInicio_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195886)
    
    End Select

End Sub

Public Sub UserControl_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If objCT.objUserControl.ActiveControl Is ExcProd(objCT) Then
            Call BotaoProdutos_Click(objCT)
        End If
    
    End If

End Sub

Public Function VND_Saida_Celula(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_VND_Saida_Celula

    'Verifica qual o Grid em questão
    Select Case objGridInt.objGrid.Name

        'Se for o GridComissao
        
        Case GridExc(objCT).Name

            lErro = Saida_Celula_GridExc(objCT, objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End Select
    
    VND_Saida_Celula = SUCESSO

    Exit Function

Erro_VND_Saida_Celula:

    VND_Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195904)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridExc(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridExc

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for a de Produto
        Case objCT.gobjinfousu.iGrid_ExcProd_Col
            lErro = Saida_Celula_ExcProd(objCT, objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Case objCT.gobjinfousu.iGrid_ExcMeta_Col
            lErro = Saida_Celula_ExcMeta(objCT, objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Se for a de Unidade de Medida
        Case objCT.gobjinfousu.iGrid_ExcPercComiss_Col
            lErro = Saida_Celula_ExcPercComiss(objCT, objGridInt)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End Select

    Saida_Celula_GridExc = SUCESSO

    Exit Function

Erro_Saida_Celula_GridExc:

    Saida_Celula_GridExc = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195909)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ExcProd(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto que está deixando de ser a corrente

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iIndice As Integer
Dim iProdutoPreenchido As Integer
Dim sProdutoEnxuto As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_ExcProd

    Set objGridInt.objControle = ExcProd(objCT)

    If Len(Trim(ExcProd(objCT).ClipText)) > 0 Then

        'Critica o Produto
        lErro = CF("Produto_Critica_Filial2", ExcProd(objCT).Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError ERRO_SEM_MENSAGEM
        
        'Se o produto é gerencial ==> erro
        If lErro = 86295 Then gError 195792
               
        'Se o produto não foi encontrado ==> Pergunta se deseja criar
        If lErro = 51381 Then gError 195793

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 195794
    
        ExcProd(objCT).promptinclude = False
        ExcProd(objCT).Text = sProdutoEnxuto
        ExcProd(objCT).promptinclude = True
        
        'Verifica se já está em outra linha do Grid
        For iIndice = 1 To objGridInt.iLinhasExistentes
            If iIndice <> objGridInt.objGrid.Row Then
                If objGridInt.objGrid.TextMatrix(iIndice, objGridInt.objGrid.Col) = ExcProd(objCT).Text Then gError 195795
            End If
        Next
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
        
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, objCT.gobjinfousu.iGrid_ExcProdDesc_Col) = objProduto.sDescricao
    
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, objCT.gobjinfousu.iGrid_ExcMetaP_Col) = Formata_Estoque(objProduto.dMetaComissao)
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, objCT.gobjinfousu.iGrid_ExcPercComissP_Col) = Format(objProduto.dPercComissao, "Percent")
    
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_ExcProd = SUCESSO

    Exit Function

Erro_Saida_Celula_ExcProd:

    Saida_Celula_ExcProd = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 195792
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 195793
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", ExcProd(objCT).Text)
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

Function Saida_Celula_ExcPercComiss(ByVal objCT As Object, objGridInt As AdmGrid) As Long
'Faz a crítica da célula Percentual Desconto que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double

On Error GoTo Erro_Saida_Celula_ExcPercComiss

    Set objGridInt.objControle = ExcPercComiss(objCT)

    If Len(ExcPercComiss(objCT).Text) > 0 Then
        
        'Critica a porcentagem
        lErro = Porcentagem_Critica(ExcPercComiss(objCT).Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        dPercentDesc = CDbl(ExcPercComiss(objCT).Text)
        
        ExcPercComiss(objCT).Text = Format(dPercentDesc, "Fixed")
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_ExcPercComiss = SUCESSO

    Exit Function

Erro_Saida_Celula_ExcPercComiss:

    Saida_Celula_ExcPercComiss = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195931)

    End Select

    Exit Function

End Function

Function Saida_Celula_ExcMeta(ByVal objCT As Object, objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_ExcMeta

    Set objGridInt.objControle = ExcMeta(objCT)

    If Len(ExcMeta(objCT).Text) > 0 Then
        
        'Critica a porcentagem
        lErro = Valor_NaoNegativo_Critica(ExcMeta(objCT).Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_ExcMeta = SUCESSO

    Exit Function

Erro_Saida_Celula_ExcMeta:

    Saida_Celula_ExcMeta = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195931)

    End Select

    Exit Function

End Function

Public Sub GridExc_Click(ByVal objCT As Object)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objCT.gobjinfousu.objGridExc, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjinfousu.objGridExc, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridExc_EnterCell(ByVal objCT As Object)
    Call Grid_Entrada_Celula(objCT.gobjinfousu.objGridExc, objCT.iAlterado)
End Sub

Public Sub GridExc_GotFocus(ByVal objCT As Object)
    Call Grid_Recebe_Foco(objCT.gobjinfousu.objGridExc)
End Sub

Public Sub GridExc_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objCT.gobjinfousu.objGridExc)
End Sub

Public Sub GridExc_KeyPress(ByVal objCT As Object, KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objCT.gobjinfousu.objGridExc, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objCT.gobjinfousu.objGridExc, objCT.iAlterado)
    End If
    
End Sub

Public Sub GridExc_LeaveCell(ByVal objCT As Object)
    Call objCT.Saida_Celula(objCT.gobjinfousu.objGridExc)
End Sub

Public Sub GridExc_Validate(ByVal objCT As Object, Cancel As Boolean)
    Call Grid_Libera_Foco(objCT.gobjinfousu.objGridExc)
End Sub

Public Sub GridExc_RowColChange(ByVal objCT As Object)
    Call Grid_RowColChange(objCT.gobjinfousu.objGridExc)
End Sub

Public Sub GridExc_Scroll(ByVal objCT As Object)
    Call Grid_Scroll(objCT.gobjinfousu.objGridExc)
End Sub

Public Sub BotaoProdutos_Click(ByVal objCT As Object)

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoProdutos_Click

    If objCT.objUserControl.ActiveControl Is ExcProd(objCT) Then
    
        sProduto1 = ExcProd(objCT).Text
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridExc(objCT).Row = 0 Then gError 195827

        sProduto1 = GridExc(objCT).TextMatrix(GridExc(objCT).Row, objCT.gobjinfousu.iGrid_ExcProd_Col)
        
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 195828
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto
    
    Set gobjCT = objCT

    Set objEventoProduto = New AdmEvento

    'Chama a Tela ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)

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
    If GridExc(gobjCT).Row > 0 Then

        Set objProduto = obj1

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 195831

        ExcProd(gobjCT).promptinclude = False
        ExcProd(gobjCT).Text = sProdutoEnxuto
        ExcProd(gobjCT).promptinclude = True
        
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 195832
        
        If Not (gobjCT.objUserControl.ActiveControl Is ExcProd(gobjCT)) Then
    
            'Verifica se já está em outra linha do Grid
            For iIndice = 1 To gobjCT.gobjinfousu.objGridExcecoes.iLinhasExistentes
                If iIndice <> GridExc(gobjCT).Row Then
                    If GridExc(gobjCT).TextMatrix(iIndice, gobjCT.gobjinfousu.iGrid_ExcProd_Col) = ExcProd(gobjCT).Text Then gError 195839
                End If
            Next
    
            'Preenche o Grid
            GridExc(gobjCT).TextMatrix(GridExc(gobjCT).Row, gobjCT.gobjinfousu.iGrid_ExcProd_Col) = ExcProd(gobjCT).Text
            GridExc(gobjCT).TextMatrix(GridExc(gobjCT).Row, gobjCT.gobjinfousu.iGrid_ExcProdDesc_Col) = objProduto.sDescricao
    
            GridExc(gobjCT).TextMatrix(GridExc(gobjCT).Row, gobjCT.gobjinfousu.iGrid_ExcMetaP_Col) = Formata_Estoque(objProduto.dMetaComissao)
            GridExc(gobjCT).TextMatrix(GridExc(gobjCT).Row, gobjCT.gobjinfousu.iGrid_ExcPercComissP_Col) = Format(objProduto.dPercComissao, "Percent")
    
            If GridExc(gobjCT).Row - GridExc(gobjCT).FixedRows = gobjCT.gobjinfousu.objGridExc.iLinhasExistentes Then
                gobjCT.gobjinfousu.objGridExc.iLinhasExistentes = gobjCT.gobjinfousu.objGridExc.iLinhasExistentes + 1
            End If
    
        End If

    End If

    gobjCT.objUserControl.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    GridExc(gobjCT).TextMatrix(GridExc(gobjCT).Row, gobjCT.gobjinfousu.iGrid_ExcProd_Col) = ""

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

Public Sub ExcProd_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ExcProd_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjinfousu.objGridExc)
End Sub

Public Sub ExcProd_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjinfousu.objGridExc)
End Sub

Public Sub ExcProd_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjinfousu.objGridExc.objControle = ExcProd(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjinfousu.objGridExc)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ExcProdDesc_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ExcProdDesc_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjinfousu.objGridExc)
End Sub

Public Sub ExcProdDesc_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjinfousu.objGridExc)
End Sub

Public Sub ExcProdDesc_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjinfousu.objGridExc.objControle = ExcProdDesc(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjinfousu.objGridExc)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ExcMeta_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ExcMeta_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjinfousu.objGridExc)
End Sub

Public Sub ExcMeta_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjinfousu.objGridExc)
End Sub

Public Sub ExcMeta_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjinfousu.objGridExc.objControle = ExcMeta(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjinfousu.objGridExc)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub ExcPercComiss_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ExcPercComiss_GotFocus(ByVal objCT As Object)
    Call Grid_Campo_Recebe_Foco(objCT.gobjinfousu.objGridExc)
End Sub

Public Sub ExcPercComiss_KeyPress(ByVal objCT As Object, KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objCT.gobjinfousu.objGridExc)
End Sub

Public Sub ExcPercComiss_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long

    Set objCT.gobjinfousu.objGridExc.objControle = ExcPercComiss(objCT)
    lErro = Grid_Campo_Libera_Foco(objCT.gobjinfousu.objGridExc)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub UpDown_DownClick(ByVal objCT As Object)

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown_DownClick

    'Diminui a data em um dia
    lErro = Data_Up_Down_Click(DataInicio(objCT), DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDown_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157190)

    End Select

    Exit Sub

End Sub

Public Sub UpDown_UpClick(ByVal objCT As Object)

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDown_UpClick

    'Aumenta a data em um dia
    lErro = Data_Up_Down_Click(DataInicio(objCT), AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDown_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157191)

    End Select

    Exit Sub

End Sub
