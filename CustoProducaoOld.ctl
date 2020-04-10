VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl CustoProducaoOld 
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8685
   ScaleHeight     =   4260
   ScaleWidth      =   8685
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6540
      ScaleHeight     =   495
      ScaleWidth      =   1650
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   165
      Width           =   1710
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "CustoProducaoOld.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   615
         Picture         =   "CustoProducaoOld.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "CustoProducaoOld.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.TextBox Descricao 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   1620
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1170
      Width           =   2400
   End
   Begin VB.CommandButton RepeteCustos 
      Caption         =   "Repete Custos do Mês Anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2745
      TabIndex        =   6
      Top             =   3705
      Width           =   2985
   End
   Begin VB.TextBox UnidMedida 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   7200
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1170
      Width           =   915
   End
   Begin MSMask.MaskEdBox CustoAnterior 
      Height          =   225
      Left            =   4065
      TabIndex        =   2
      Top             =   1170
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Produto 
      Height          =   225
      Left            =   480
      TabIndex        =   0
      Top             =   1170
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox CustoProducao 
      Height          =   225
      Left            =   5700
      TabIndex        =   3
      Top             =   1170
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid GridProdutos 
      Height          =   2730
      Left            =   165
      TabIndex        =   5
      Top             =   870
      Width           =   8010
      _ExtentX        =   14129
      _ExtentY        =   4815
      _Version        =   393216
      Rows            =   11
      Cols            =   5
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
      AllowUserResizing=   1
   End
   Begin VB.Label Ano 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   840
      TabIndex        =   14
      Top             =   390
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Ano:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   375
      TabIndex        =   13
      Top             =   420
      Width           =   375
   End
   Begin VB.Label Mes 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2580
      TabIndex        =   12
      Top             =   390
      Width           =   1185
   End
   Begin VB.Label Label3 
      Caption         =   "Mês:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2100
      TabIndex        =   11
      Top             =   420
      Width           =   375
   End
End
Attribute VB_Name = "CustoProducaoOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim gcolProdutosProduzidos As Collection
Dim giApuradoMesAnterior As Integer

'campos do grid
Dim iGrid_Sequencial_Col As Integer
Dim iGrid_Produto_Col As Integer
Dim iGrid_Descricao_Col As Integer
Dim iGrid_UnidMedida_Col As Integer
Dim iGrid_CustoAnterior_Col As Integer
Dim iGrid_CustoProducao_Col As Integer

Dim objGrid As AdmGrid

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 48725
    
    iAlterado = 0
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err
    
        Case 48725
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158682)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro  As Long
Dim iAno As Integer
Dim iMes As Integer
Dim colProdutoCustoAtual As New colProdutoCusto
Dim objProdutoCustoAtual As New ClassProdutoCusto
Dim iIndice As Integer
Dim iIndice1 As Integer
Dim sProduto As String
Dim iPreenchido As Integer

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Copia o ano e o mes da tela
    iAno = CStr(Ano.Caption)
    
    lErro = MesNumero(Mes.Caption, iMes)
    If lErro <> SUCESSO Then Error 48726
    
    'Pega o codigo e a descricao da colecao produtoCustoAnterior
    For Each objProdutoCustoAtual In gcolProdutosProduzidos
        colProdutoCustoAtual.Add objProdutoCustoAtual.sCodProduto, objProdutoCustoAtual.sDescProduto, objProdutoCustoAtual.dCusto, objProdutoCustoAtual.sCodProduto
    Next
    
    'Para cada Produto atualiza o custo de Producao
    For iIndice = 1 To objGrid.iLinhasExistentes
    
        objProdutoCustoAtual.sCodProduto = GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col)
        
        lErro = CF("Produto_Formata",objProdutoCustoAtual.sCodProduto, sProduto, iPreenchido)
        If lErro <> SUCESSO Then Error 17796
        
        objProdutoCustoAtual.sCodProduto = sProduto
        
        'Verifica se o Produto custo foi preenchido
        If Len(Trim(GridProdutos.TextMatrix(iIndice, iGrid_CustoProducao_Col))) = 0 Then
            'se nao foi zera o custo
            objProdutoCustoAtual.dCusto = 0
        Else
            'se foi informado entao preenche com o que está no grid
            objProdutoCustoAtual.dCusto = StrParaDbl(GridProdutos.TextMatrix(iIndice, iGrid_CustoProducao_Col))
        End If
            
        'preenche na colecao com a chave sendo o codigo
        colProdutoCustoAtual(objProdutoCustoAtual.sCodProduto).dCusto = objProdutoCustoAtual.dCusto
        
    Next
    
    lErro = CF("ProdutosProduzidosCustos_Grava",iAno, iMes, colProdutoCustoAtual)
    If lErro <> SUCESSO Then Error 48727
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:
        
    Gravar_Registro = Err
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 48726, 48727
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158683)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 48728

    'Limpa a Tela
    lErro = CustoProducao_Limpa()
    If lErro <> SUCESSO Then Error 48729
    
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 48728, 48729

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158684)

    End Select

    Exit Sub

End Sub

Function CustoProducao_Limpa() As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_CustoProducao_Limpa

    'limpa somente a coluna de custo de Producao
    For iIndice = 1 To objGrid.iLinhasExistentes
    
        GridProdutos.TextMatrix(iIndice, iGrid_CustoProducao_Col) = ""
        
    Next
    
    Exit Function
    
Erro_CustoProducao_Limpa:

    CustoProducao_Limpa = Err
    
    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158685)

    End Select

    Exit Function
    
End Function

Public Sub Form_Load()

Dim lErro As Long
Dim objEstoqueMes As New ClassEstoqueMes
Dim objEstoqueMesAnterior As New ClassEstoqueMes
Dim sMes As String
Dim objProdutosProduzidos As ClassProdutoCusto
Dim objProdutosCustoAnterior As New ClassProdutoCusto

On Error GoTo Erro_Form_Load
    
    Set gcolProdutosProduzidos = New Collection
    
    'Formata os custos
    CustoAnterior.Format = FORMATO_CUSTO
    CustoProducao.Format = FORMATO_CUSTO

    'Inicializa Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd",Produto)
    If lErro <> SUCESSO Then Error 48730
        
    'Le o primeiro mes que ainda nao foi apurado
    lErro = CF("EstoqueMesNaoApurado_Le",objEstoqueMes)
    If lErro <> SUCESSO And lErro <> 25221 Then Error 48731

    'se não encontrou ------> Erro
    If lErro = 25221 Then Error 48732

    'Pega o Nome do mes em questão
    lErro = MesNome(objEstoqueMes.iMes, sMes)
    If lErro <> SUCESSO Then Error 48733

    'Preenche o Mes e o Ano na tela
    Mes.Caption = sMes
    Ano.Caption = CStr(objEstoqueMes.iAno)

    ' Preenche o objeto para pegar o Mes anterior
    If objEstoqueMes.iMes > 1 Then
        objEstoqueMesAnterior.iMes = objEstoqueMes.iMes - 1
        objEstoqueMesAnterior.iAno = objEstoqueMes.iAno
    Else
        If objEstoqueMes.iMes = 1 Then
            objEstoqueMesAnterior.iMes = 12
            objEstoqueMesAnterior.iAno = objEstoqueMes.iAno - 1
        End If
        
    End If

    objEstoqueMesAnterior.iFilialEmpresa = objEstoqueMes.iFilialEmpresa
    
    'le o mes anterior
    lErro = CF("EstoqueMes_Le",objEstoqueMesAnterior)
    If lErro <> SUCESSO And lErro <> 36513 Then Error 48734

    If lErro = 36513 Then
        giApuradoMesAnterior = 0
    Else
        giApuradoMesAnterior = 1
    End If
    
    'le os produtos Produzidos e o seus custos do mes em questao
    lErro = CF("ProdProduzidos_Custos_Mes_Le",objEstoqueMes.iMes, objEstoqueMes.iAno, gcolProdutosProduzidos)
    If lErro <> SUCESSO Then gError 48735
    
    'se existe mes anterior
    If giApuradoMesAnterior = 1 Then
        
        'le o custo dos produtos do mes anterior
        lErro = CF("ProdProduzidos_Custos_MesAnt_Le",objEstoqueMesAnterior.iMes, objEstoqueMesAnterior.iAno, gcolProdutosProduzidos)
        If lErro <> SUCESSO Then Error 48736
        
    End If
    
    'Inicialização do GridProduto
    Set objGrid = New AdmGrid
    
    lErro = Inicializa_GridProdutos(objGrid, gcolProdutosProduzidos.Count)
    If lErro <> SUCESSO Then Error 48737
    
    lErro = Preenche_GridProdutos(gcolProdutosProduzidos)
    If lErro <> SUCESSO Then Error 48738
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
        
        Case 48730, 48731
        
        Case 48732
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CUSTO_PRODUCAO_APURADO", Err)
        
        Case 48733, 48734, 48735, 48736, 48737, 48738
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158686)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)
       
    Set objGrid = Nothing
    Set gcolProdutosProduzidos = Nothing
    
    Unload Me
    
End Sub

Private Sub GridProdutos_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGrid, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGrid, iAlterado)
        End If
    
End Sub

Private Sub GridProdutos_EnterCell()
    
    Call Grid_Entrada_Celula(objGrid, iAlterado)
    
End Sub

Private Sub GridProdutos_GotFocus()
    
    Call Grid_Recebe_Foco(objGrid)

End Sub

Private Sub GridProdutos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer


    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridProdutos_LeaveCell()

    Call Saida_Celula(objGrid)

End Sub

Private Sub GridProdutos_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGrid)
    
End Sub

Private Sub RepeteCustos_Click()

Dim lErro  As Long
Dim objProdutoCustoAnterior As New ClassProdutoCusto
Dim iIndice As Integer
Dim iMesAnterior As Integer
Dim iMesAtual As Integer
Dim iAnoAnterior As Integer
Dim iAnoAtual As Integer
Dim sMesAnterior As String

On Error GoTo Erro_RepeteCustos_Click
    
    'Pega o Numero do mes em questão
    lErro = MesNumero(Mes.Caption, iMesAtual)
    If lErro <> SUCESSO Then Error 48739
    
    iAnoAtual = CInt(Ano.Caption)
    
    'Para pegar o Mes anterior e o Ano
    If CInt(iMesAtual) > 1 Then
        iMesAnterior = iMesAtual - 1
        iAnoAnterior = iAnoAtual
    Else
        If iMesAtual = 1 Then
            iMesAnterior = 12
            iAnoAnterior = iAnoAtual - 1
        End If
    End If
    
    'Pega o Numero do mes em questão
    lErro = MesNome(iMesAnterior, sMesAnterior)
    If lErro <> SUCESSO Then Error 48740
    
    If giApuradoMesAnterior = 1 Then
        
        For Each objProdutoCustoAnterior In gcolProdutosProduzidos
    
            iIndice = iIndice + 1
            If objProdutoCustoAnterior.dCustoMesAnterior <> 0 Then
            
                GridProdutos.TextMatrix(iIndice, iGrid_CustoProducao_Col) = Format(objProdutoCustoAnterior.dCustoMesAnterior, FORMATO_CUSTO)
            End If
            
        Next
    Else
        If giApuradoMesAnterior = 0 Then Error 48741
        
    End If
    
    Exit Sub

Erro_RepeteCustos_Click:

    Select Case Err
        
        Case 48739, 48740
        
        Case 48741
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CP_MES_ANTERIOR_NAO_APURADO", Err, sMesAnterior, iAnoAnterior)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158687)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_GridProdutos(objGridInt As AdmGrid, iLinhas As Integer) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("U.M.")
    objGridInt.colColuna.Add ("Custo Mês Anterior")
    objGridInt.colColuna.Add ("Custo Produção")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (Descricao.Name)
    objGridInt.colCampo.Add (UnidMedida.Name)
    objGridInt.colCampo.Add (CustoAnterior.Name)
    objGridInt.colCampo.Add (CustoProducao.Name)

    'Colunas do Grid
    iGrid_Sequencial_Col = 0
    iGrid_Produto_Col = 1
    iGrid_Descricao_Col = 2
    iGrid_UnidMedida_Col = 3
    iGrid_CustoAnterior_Col = 4
    iGrid_CustoProducao_Col = 5
    
    'Grid do GridInterno
    objGridInt.objGrid = GridProdutos
    
    'Todas as linhas do grid
    
    If iLinhas > 10 Then
        objGridInt.objGrid.Rows = iLinhas + 1
    Else
        objGridInt.objGrid.Rows = 11
    End If
    
    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridProdutos.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_GridProdutos = SUCESSO

    Exit Function

End Function

Private Function Preenche_GridProdutos(colProdutoCusto As Collection) As Long

Dim iIndice As Integer
Dim sProdutoMascarado As String
Dim objProdutoCusto As ClassProdutoCusto
Dim lErro As Long

On Error GoTo Erro_Preenche_GridProdutos
        
    'Preenche GridProdutos ainda nao preenchidos
    For Each objProdutoCusto In colProdutoCusto
                
        If objProdutoCusto.dCusto = 0 Then
            
            iIndice = iIndice + 1
    
            sProdutoMascarado = String(STRING_PRODUTO, 0)
    
            lErro = Mascara_MascararProduto(objProdutoCusto.sCodProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then Error 48742
    
            Produto.PromptInclude = False
            Produto.Text = sProdutoMascarado
            Produto.PromptInclude = True
    
            GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado
            GridProdutos.TextMatrix(iIndice, iGrid_Descricao_Col) = objProdutoCusto.sDescProduto
            GridProdutos.TextMatrix(iIndice, iGrid_UnidMedida_Col) = objProdutoCusto.sSiglaUMEstoque
            
            
            If giApuradoMesAnterior = 1 Then
            
                If gcolProdutosProduzidos.Item(objProdutoCusto.sCodProduto).dCustoMesAnterior <> 0 Then GridProdutos.TextMatrix(iIndice, iGrid_CustoAnterior_Col) = Format(gcolProdutosProduzidos.Item(objProdutoCusto.sCodProduto).dCustoMesAnterior, FORMATO_CUSTO)
                
            End If
            
            If objProdutoCusto.dCusto <> 0 Then GridProdutos.TextMatrix(iIndice, iGrid_CustoProducao_Col) = Format(objProdutoCusto.dCusto, FORMATO_CUSTO)
        
        End If
        
    Next
    
    'Preenche GridProdutos com produtos com o custo producao preenchido
    For Each objProdutoCusto In colProdutoCusto
        
        If objProdutoCusto.dCusto <> 0 Then
            
            iIndice = iIndice + 1
    
            sProdutoMascarado = String(STRING_PRODUTO, 0)
    
            lErro = Mascara_MascararProduto(objProdutoCusto.sCodProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then Error 48743
    
            Produto.PromptInclude = False
            Produto.Text = sProdutoMascarado
            Produto.PromptInclude = True
    
            GridProdutos.TextMatrix(iIndice, iGrid_Produto_Col) = sProdutoMascarado
            GridProdutos.TextMatrix(iIndice, iGrid_Descricao_Col) = objProdutoCusto.sDescProduto
            GridProdutos.TextMatrix(iIndice, iGrid_UnidMedida_Col) = objProdutoCusto.sSiglaUMEstoque
            
            If giApuradoMesAnterior = 1 Then
            
                If gcolProdutosProduzidos.Item(objProdutoCusto.sCodProduto).dCusto <> 0 Then GridProdutos.TextMatrix(iIndice, iGrid_CustoAnterior_Col) = Format(gcolProdutosProduzidos.Item(objProdutoCusto.sCodProduto).dCustoMesAnterior, FORMATO_CUSTO)
            
            End If
            
            If objProdutoCusto.dCusto <> 0 Then GridProdutos.TextMatrix(iIndice, iGrid_CustoProducao_Col) = Format(objProdutoCusto.dCusto, FORMATO_CUSTO)
        
        End If
        
        Next
    
    
    objGrid.iLinhasExistentes = colProdutoCusto.Count

    Preenche_GridProdutos = SUCESSO

    Exit Function

Erro_Preenche_GridProdutos:

    Preenche_GridProdutos = Err

    Select Case Err

        Case 48742, 48743
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_MASCARARPRODUTO", Err, objProdutoCusto.sCodProduto)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 158688)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
        
        Select Case GridProdutos.Col
    
            Case iGrid_CustoProducao_Col
    
                lErro = Saida_Celula_CustoProducao(objGridInt)
                If lErro <> SUCESSO Then Error 48744
                    
            End Select
        
            lErro = Grid_Finaliza_Saida_Celula(objGridInt)
            If lErro <> SUCESSO Then Error 48745
       
    End If
    
    Saida_Celula = SUCESSO
    
    Exit Function

Erro_Saida_Celula:
    
    Saida_Celula = Err
    
    Select Case Err

        Case 48744, 48745
        
        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158689)

    End Select

    Exit Function

End Function

Function Saida_Celula_CustoProducao(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_CustoProducao

    Set objGridInt.objControle = CustoProducao

    If Len(Trim(CustoProducao.Text)) > 0 Then

        lErro = Valor_Positivo_Critica(CustoProducao)
        If lErro <> SUCESSO Then Error 48746
        
        CustoProducao.Text = Format(CustoProducao.Text, "Standard")

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 48747

    Saida_Celula_CustoProducao = SUCESSO

    Exit Function

Erro_Saida_Celula_CustoProducao:

    Saida_Celula_CustoProducao = Err

    Select Case Err
        
        Case 48746, 48747
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, 158690)

    End Select

    Exit Function

End Function

Private Sub CustoProducao_Change()

    iAlterado = 0
    
End Sub

Private Sub CustoProducao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub
Private Sub CustoProducao_KeyPress(KeyAscii As Integer)
    
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub CustoProducao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = CustoProducao
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CUSTO_PRODUCAO
    Set Form_Load_Ocx = Me
    Caption = "Custo de Produção"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CustoProducao"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object
    Set Controls = UserControl.Controls
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
    Height = UserControl.Height
End Property

Public Property Get Width() As Long
    Width = UserControl.Width
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ActiveControl
Public Property Get ActiveControl() As Object
    Set ActiveControl = UserControl.ActiveControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Private Sub Unload(objme As Object)
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property

'**** fim do trecho a ser copiado *****

Public Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO
    
End Function



Private Sub Ano_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Ano, Source, X, Y)
End Sub

Private Sub Ano_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Ano, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Mes_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Mes, Source, X, Y)
End Sub

Private Sub Mes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Mes, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub
