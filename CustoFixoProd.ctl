VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl CustoFixoProdOcx 
   ClientHeight    =   5130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   KeyPreview      =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   6000
   Begin VB.CommandButton BotaoListar 
      Caption         =   "Custos Fixos Cadastrados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3375
      TabIndex        =   23
      Top             =   4665
      Width           =   2430
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3600
      ScaleHeight     =   495
      ScaleWidth      =   2190
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   150
      Width           =   2250
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   105
         Picture         =   "CustoFixoProd.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Gravar"
         Top             =   105
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   630
         Picture         =   "CustoFixoProd.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Excluir"
         Top             =   105
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1155
         Picture         =   "CustoFixoProd.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Limpar"
         Top             =   105
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1680
         Picture         =   "CustoFixoProd.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fechar"
         Top             =   105
         Width           =   420
      End
   End
   Begin VB.Frame FrameCusto1 
      Caption         =   "Custos"
      Height          =   1470
      Left            =   135
      TabIndex        =   15
      Top             =   3090
      Width           =   5685
      Begin MSMask.MaskEdBox CustoAplicado 
         Height          =   300
         Left            =   2100
         TabIndex        =   21
         Top             =   930
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   20
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   " "
      End
      Begin VB.Label LabelCustoAplicado 
         AutoSize        =   -1  'True
         Caption         =   "Novo Custo Fixo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   585
         TabIndex        =   22
         Top             =   990
         Width           =   1470
      End
      Begin VB.Label CustoRateio 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2115
         TabIndex        =   19
         Top             =   390
         Width           =   1125
      End
      Begin VB.Label LabelCustoRateio 
         AutoSize        =   -1  'True
         Caption         =   "Calculado pelo Rateio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   135
         TabIndex        =   18
         Top             =   420
         Width           =   1950
      End
      Begin VB.Label CustoAnterior 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4290
         TabIndex        =   17
         Top             =   390
         Width           =   1125
      End
      Begin VB.Label LabelCustoAnterior 
         AutoSize        =   -1  'True
         Caption         =   "Anterior:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3435
         TabIndex        =   16
         Top             =   420
         Width           =   735
      End
   End
   Begin VB.Frame FrameProduto 
      Caption         =   "Produto"
      Height          =   1335
      Left            =   135
      TabIndex        =   9
      Top             =   1620
      Width           =   5715
      Begin MSMask.MaskEdBox Produto 
         Height          =   300
         Left            =   1230
         TabIndex        =   2
         Top             =   360
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   " "
      End
      Begin VB.Label LabelDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   14
         Top             =   885
         Width           =   930
      End
      Begin VB.Label LabelProduto 
         AutoSize        =   -1  'True
         Caption         =   "Produto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   315
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   13
         Top             =   390
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "U.M.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3780
         TabIndex        =   12
         Top             =   420
         Width           =   480
      End
      Begin VB.Label LabelUMEstoque 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4410
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Descricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1230
         TabIndex        =   10
         Top             =   855
         Width           =   4275
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data de Referência"
      Height          =   750
      Left            =   135
      TabIndex        =   0
      Top             =   750
      Width           =   5685
      Begin MSMask.MaskEdBox Data 
         Height          =   300
         Left            =   3060
         TabIndex        =   1
         Top             =   270
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   4110
         TabIndex        =   7
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label LabelData 
         AutoSize        =   -1  'True
         Caption         =   "Rotina de rateio executada em:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   300
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   8
         Top             =   345
         Width           =   2685
      End
   End
End
Attribute VB_Name = "CustoFixoProdOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_Caption As String
Event Unload()

Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoCusto As AdmEvento
Attribute objEventoCusto.VB_VarHelpID = -1

'variaveis de controle de alteração
Dim iAlterado As Integer
Dim iProdutoAlterado As Integer

Private gdtDataPadrao As Date

Public Function Trata_Parametros(Optional objCustoFixoProd As ClassCustoFixoProd) As Long
'Espera receber algum parametro  em objCustoFixoProd

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum parametro
    If Not (objCustoFixoProd Is Nothing) Then
        
        'faz uma leitura no Bd a apartir p/ carregar o obj
        lErro = CF("CustoFixoProd_Le", objCustoFixoProd)
        If lErro <> SUCESSO And lErro <> 116391 Then gError 116412
        
        'traz as informações do BD para a tela
        lErro = Traz_CustoFixo_Tela(objCustoFixoProd)
        If lErro <> SUCESSO Then gError 116387
    
    End If

    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:
    
    Trata_Parametros = gErr
    
    Select Case gErr
        
        Case 116387, 116412
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158637)
        
    End Select
        
    Exit Function
        
End Function

Private Sub Form_Load()
'carrega as configurações iniciais da tela

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Inicializa Máscara de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 116358
    
    'inicializa o evento de browser
    Set objEventoProduto = New AdmEvento
    Set objEventoCusto = New AdmEvento
        
    'preenche o campo data c/ a data padrao
    lErro = Carregar_Data_Padrao()
    If lErro <> SUCESSO Then gError 106648
    
    Data.PromptInclude = False
    Data.Text = Format(gdtDataPadrao, "dd/mm/yy")
    Data.PromptInclude = True
        
    'zera as variaveis de alteração
    iAlterado = 0
    iProdutoAlterado = 0
        
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:
    
    lErro_Chama_Tela = gErr
    
    Select Case gErr
                    
        Case 116358, 106648
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158638)

    End Select
    
    Exit Sub
    
End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objCustoFixoProd As New ClassCustoFixoProd

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "CustoFixoProd"

    'Lê os dados da Tela
    lErro = Move_Tela_Memoria(objCustoFixoProd)
    If lErro <> SUCESSO Then gError 116359
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Produto", objCustoFixoProd.sProduto, STRING_PRODUTO, "Produto"
    colCampoValor.Add "DataReferencia", objCustoFixoProd.dtDataReferencia, 0, "DataReferencia"
    colCampoValor.Add "Custo", objCustoFixoProd.dCusto, 0, "Custo"
    colCampoValor.Add "CustoCalculado", objCustoFixoProd.dCustoCalculado, 0, "CustoCalculado"
    colCampoValor.Add "Automatico", objCustoFixoProd.iAutomatico, 0, "Automatico"
    
    'adiciona FilialEmpresa
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub
    
Erro_Tela_Extrai:
    
    Select Case gErr

        Case 116359

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158639)
            
    End Select

    Exit Sub

End Sub

Private Function Move_Tela_Memoria(ByVal objCustoFixoProd As ClassCustoFixoProd) As Long
'Move os dados da tela p/ a memoria

Dim lErro As Long
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria
         
    'Retira a mascara do produto
    lErro = CF("Produto_Formata", Produto.Text, sProduto, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 116360

    'carrega o obj c/ os dados da tela
    objCustoFixoProd.iFilialEmpresa = giFilialEmpresa
    objCustoFixoProd.dtDataReferencia = StrParaDate(Data.Text)
    objCustoFixoProd.sProduto = sProduto
    objCustoFixoProd.dCusto = StrParaDbl(CustoAplicado.Text)
    
    Move_Tela_Memoria = SUCESSO
        
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 116360

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158640)
    
    End Select
    
    Exit Function
    
End Function

Public Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objCustoFixoProd As New ClassCustoFixoProd

On Error GoTo Erro_Tela_Preenche

    'preenche o obj c/ os valores correspondentes
    objCustoFixoProd.sProduto = colCampoValor.Item("Produto").vValor
    objCustoFixoProd.dtDataReferencia = colCampoValor.Item("DataReferencia").vValor
    objCustoFixoProd.dCusto = colCampoValor.Item("Custo").vValor
    objCustoFixoProd.dCustoCalculado = colCampoValor.Item("CustoCalculado").vValor
    objCustoFixoProd.iAutomatico = colCampoValor.Item("Automatico").vValor
    objCustoFixoProd.iFilialEmpresa = giFilialEmpresa

    'Traz os dados para tela
    lErro = Traz_CustoFixo_Tela(objCustoFixoProd)
    If lErro <> SUCESSO Then gError 116361

    Exit Function

Erro_Tela_Preenche:

    Select Case gErr

        Case 116361

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158641)

    End Select

    Exit Function

End Function

Private Function Traz_CustoFixo_Tela(ByVal objCustoFixoProd As ClassCustoFixoProd) As Long
'Exibe os dados na tela

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim dCusto As Double

On Error GoTo Erro_Traz_CustoFixo_Tela
              
    'Guarda o código do produto em objproduto
    objProduto.sCodigo = objCustoFixoProd.sProduto
    
    'Critica o formato do codigo
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 116362
    
    'Se não encontrou o produto => erro
    If lErro = 28030 Then gError 116363
    
    'preenche o produto e as labels c/ os dados obtidos
    Produto.PromptInclude = False
    Produto.Text = objProduto.sCodigo
    Produto.PromptInclude = True
    
    'preenche a descrição (label)
    Descricao.Caption = objProduto.sDescricao
    
    'preenche o UN. de medida (label)
    LabelUMEstoque.Caption = objProduto.sSiglaUMEstoque
                       
    'preenche a data c/ a data de referencia
    Data.PromptInclude = False
    Data.Text = Format(objCustoFixoProd.dtDataReferencia, "dd/mm/yy")
    Data.PromptInclude = True
    
    'preenche o custo aplicado
    CustoAplicado.Text = Format(objCustoFixoProd.dCusto, "Standard")
    
    'preenche o custo de Rateio
    CustoRateio.Caption = Format(objCustoFixoProd.dCustoCalculado, "Standard")
    
    'verifica se tem algum custo anterior
    lErro = CF("CustoFixoProd_Le_Anterior", objCustoFixoProd, dCusto)
    If lErro <> SUCESSO And lErro <> 116395 Then gError 116364
    
    'preenche o custo anterior
    CustoAnterior.Caption = Format(dCusto, "Standard")
    
    'zera as variaveis de controle de alteração
    iAlterado = 0
    iProdutoAlterado = 0

    Traz_CustoFixo_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_CustoFixo_Tela:

    Traz_CustoFixo_Tela = gErr

    Select Case gErr
        
        Case 116362, 116364
        
        Case 116363
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158642)
    
    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()
'Sub que inicializa a exclusão de registros

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objCustoFixoProd As New ClassCustoFixoProd
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_BotaoExcluir_Click
    
    'verifica o preenchimento da data
    If Len(Trim(Data.ClipText)) = 0 Then gError 116365
    
    'Verifica preenchimento do codigo do produto
    If Len(Trim(Produto.ClipText)) = 0 Then gError 116366

    'Retira a mascara do produto
    lErro = CF("Produto_Formata", Produto.Text, sProduto, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 116368

    'preenche o obj c/ os dados a serem passados como parametro
    objCustoFixoProd.sProduto = sProduto
    objCustoFixoProd.iFilialEmpresa = giFilialEmpresa
    objCustoFixoProd.dtDataReferencia = StrParaDate(Data.Text)
    
    'LE a tabela CustoFixoProd
    lErro = CF("CustoFixoProd_Le", objCustoFixoProd)
    If lErro <> SUCESSO And lErro <> 116391 Then gError 116369

    'Se não achou --> Erro
    If lErro = 116391 Then gError 116370
    
    'verifica se o CustoFixo é automatico, se for, erro
    If objCustoFixoProd.iAutomatico = CUSTOFIXOPROD_AUTOMATICO Then gError 116386
   
    'pergunta se relamente deseja excluir
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_CUSTOFIXOPROD", Produto.Text, Data.Text)

    'se sim
    If vbMsgRes = vbYes Then
        
        'tranforma o ponteiro em ampulheta
        GL_objMDIForm.MousePointer = vbHourglass
        
        'exclui o registro
        lErro = CF("CustoFixoProd_Exclui", objCustoFixoProd)
        If lErro <> SUCESSO Then gError 116371
                                                        
        'limpa a tela
        Call Limpa_Tela_CustoFixoProd
            
        'volta o ponteiro ao padrão
        GL_objMDIForm.MousePointer = vbDefault
    
    End If
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 116368, 116369, 116371

        Case 116366
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
            Produto.SetFocus
        
        Case 116365
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
            Data.SetFocus
        
        Case 116370
            Call Rotina_Erro(vbOKOnly, "ERRO_CUSTOFIXOPROD_NAO_EXISTENTE", gErr, Produto.Text, Data.Text)
        
        Case 116386
            Call Rotina_Erro(vbOKOnly, "ERRO_CUSTOFIXOPROD_AUTOMATICO", gErr, Produto.Text, Data.Text)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158643)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'inicializa a etapa de gravacao

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama a função p/ a gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 116371

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 116371

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158644)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long
'função que critica grava os registros

Dim lErro As Long
Dim objCustoFixoProd As New ClassCustoFixoProd

On Error GoTo Erro_Gravar_Registro

    'transforma o ponteiro em ampulheta
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se a data está preenchida
    If Len(Trim(Data.ClipText)) = 0 Then gError 116372
    
    'Verifica se o produto está preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then gError 116373
    
    'verifica se o custo aplicado está preenchido
    If StrParaDbl(CustoAplicado.Text) <= 0 Then gError 116374
    
    'Chama Move_Tela_Memoria para passar os dados da tela para o obj
    lErro = Move_Tela_Memoria(objCustoFixoProd)
    If lErro <> SUCESSO Then gError 116375

    'Chama a função de gravacao
    lErro = CF("CustoFixoProd_Grava", objCustoFixoProd)
    If lErro <> SUCESSO Then gError 116376

    'Limpa a Tela
    Call Limpa_Tela_CustoFixoProd
    
    'volta o ponteiro p/ o padrão
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 116375, 116376

        Case 116372
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
            Data.SetFocus

        Case 116373
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
            Produto.SetFocus

        Case 116374
            Call Rotina_Erro(vbOKOnly, "ERRO_CUSTOAPLICADO_NAO_PREENCHIDO", gErr)
            CustoAplicado.SetFocus
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158645)

    End Select
    
    Exit Function

End Function

Private Sub BotaoLimpar_Click()
'sub para limpar a tela

Dim lErro As Long

On Error GoTo Erro_Botao_Limpar

    'pergunta se deseja salvar
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 116377
    
    'limpa a tela
    Call Limpa_Tela_CustoFixoProd
    
    Exit Sub
        
Erro_Botao_Limpar:

    Select Case gErr

        Case 116377
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158646)

    End Select
    
    Exit Sub

End Sub

Private Sub Limpa_Tela_CustoFixoProd()
'sub que limpa a tela inteira

On Error GoTo Erro_Limpa_Tela_CustoEmbMP
    
    'limpa as text box
    Call Limpa_Tela(Me)

    'limpa o restante(labels e MaskEds)
    Data.PromptInclude = False
    Data.Text = Format(gdtDataPadrao, "dd/mm/yy")
    Data.PromptInclude = True
    Descricao.Caption = ""
    LabelUMEstoque.Caption = ""
    CustoAnterior.Caption = ""
    CustoRateio.Caption = ""
    
    'zera as variaveis de alteracao
    iAlterado = 0
    iProdutoAlterado = 0
    
    Data.SetFocus

    Exit Sub
    
Erro_Limpa_Tela_CustoEmbMP:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158647)
            
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()
'fecha a tela

Dim lErro As Long
        
On Error GoTo Erro_Botao_Fechar
        
    'pergunta se deseja salvar
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 116378
    
    Unload Me
    
    Exit Sub
    
Erro_Botao_Fechar:

    Select Case gErr
    
        Case 116378
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158648)
        
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoListar_Click()
    Call LabelData_Click
End Sub

Private Sub LabelData_Click()
'sub chamadora do browser

Dim objCustoFixoProd As New ClassCustoFixoProd
Dim colSelecao As New Collection
Dim sSelecao As String

    'verifica se a data está preenchida
    If Len(Trim(Data.ClipText)) > 0 Then
             
        'preenche o obj
         objCustoFixoProd.dtDataReferencia = StrParaDate(Data.Text)
             
        'adiciona a Data na selecao
        sSelecao = "DataReferencia = ?"
        
        'Adiciona o Filtro na collection
        colSelecao.Add (StrParaDate(Data.Text))

    Else
        'preenche o obj c/ a data corrente p/ trazer todos os registros
        objCustoFixoProd.dtDataReferencia = gdtDataAtual
    
    End If
    
    'chama o browser
    Call Chama_Tela("CustoFixoProdLista", colSelecao, objCustoFixoProd, objEventoCusto, sSelecao)
     
End Sub

Private Sub objEventoCusto_evSelecao(obj1 As Object)
'preenche a tela c/ os dados selecionados no browser

Dim objCustoFixoProd As ClassCustoFixoProd
Dim lErro As Long

On Error GoTo Erro_objEventoDVVCliente_evSelecao

    Set objCustoFixoProd = obj1

    'traz os dados p/ a tela
    lErro = Traz_CustoFixo_Tela(objCustoFixoProd)
    If lErro <> SUCESSO Then gError 119520

    Me.Show

    Exit Sub

Erro_objEventoDVVCliente_evSelecao:

    Select Case gErr
    
        Case 119520
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158649)
            
    End Select

    Exit Sub

End Sub

Private Sub Produto_Validate(Cancel As Boolean)
'verifica se o produto é válido

Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Produto_Validate

    'se o produto nao estiver preenchido, sai da rotina
    If Len(Trim(Produto.ClipText)) = 0 Then
    
        Call Limpa_Produto
        Exit Sub
    
    End If

    'se o produto não foi alterado => sai da função
    If iProdutoAlterado <> REGISTRO_ALTERADO Then Exit Sub
    
    'Critica o formato do codigo
    lErro = CF("Produto_Critica_Filial", Produto.Text, objProduto, iProdutoPreenchido)
    If lErro <> SUCESSO And lErro <> 51381 Then gError 116379
            
    'lErro = 51381 => inexistente
    If lErro = 51381 Then gError 116380
        
    'exibe os dados do produto na tela
    Produto.PromptInclude = False
    Produto.Text = objProduto.sCodigo
    Produto.PromptInclude = True
    
    'exibe a descrição
    Descricao.Caption = objProduto.sDescricao
    
    'exibe a uni. de medida
    LabelUMEstoque.Caption = objProduto.sSiglaUMEstoque
    
    'zera a variavel de alteração do produto
    iProdutoAlterado = 0
        
    Exit Sub
    
Erro_Produto_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 116379
            'limpa o frame do produto
            Call Limpa_Produto
            
        Case 116380
           'Não encontrou Produto no BD e pergunta se deseja criar um novo
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)
            
            'se sim
            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)
            'senão
            Else
                'limpa o frame do produto
                Call Limpa_Produto
            End If
         
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158650)
            
    End Select

    Exit Sub

End Sub

Private Sub Limpa_Produto()
'rotina que limpa apenas o frame do produto

    'limpa o frame do produto
    Produto.PromptInclude = False
    Produto.Text = ""
    Produto.PromptInclude = True
    
    Descricao.Caption = ""
    LabelUMEstoque.Caption = ""

End Sub

Private Sub UpDownData_DownClick()
'Dimunui a data

Dim lErro As Long

On Error GoTo Erro_UpDownData_DownClick

    'Diminui a data em 1 dia
    lErro = Data_Up_Down_Click(Data, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 116381

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 116381
            Data.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158651)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()
'aumenta a data

Dim lErro As Long

On Error GoTo Erro_UpDownData_UpClick

    'Aumenta a data em 1 dia
    lErro = Data_Up_Down_Click(Data, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 116382

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 116382
            Data.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158652)

    End Select

    Exit Sub

End Sub

Private Sub CustoAplicado_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Data_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
    iProdutoAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Data_GotFocus()
    Call MaskEdBox_TrataGotFocus(Data)
End Sub

Private Sub Data_Validate(Cancel As Boolean)
'verifica se o campo Data está correto

Dim lErro As Long
Dim objCustoFixoProd As New ClassCustoFixoProd

On Error GoTo Erro_Data_Validate

    'Verifica se o campo Data foi preenchida
    If Len(Trim(Data.ClipText)) > 0 Then
        
        'Critica a Data
        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError 116383

    End If

    Exit Sub

Erro_Data_Validate:
    
    Cancel = True

    Select Case gErr

        Case 116383

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158653)

    End Select

    Exit Sub
    
End Sub

Private Sub CustoAplicado_Validate(Cancel As Boolean)
'verifica se o valor de Custo é valido

Dim lErro As Long

On Error GoTo Erro_CustoAplicado_Validate

    'verifica se o custo aplicado foi preenchido
    If Len(Trim(CustoAplicado.Text)) <> 0 Then

        'não pode ser valor negativo
        lErro = Valor_NaoNegativo_Critica(CustoAplicado.Text)
        If lErro <> SUCESSO Then gError 116384

    End If

    Exit Sub

Erro_CustoAplicado_Validate:

    Cancel = True
    
    Select Case gErr

        Case 116384

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158654)
    
    End Select

    Exit Sub

End Sub

Private Sub LabelProduto_Click()
'sub chamadora do browser Produto

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_LabelProduto_Click

    'Verifica se o produto foi preenchido
    If Len(Trim(Produto.ClipText)) <> 0 Then

        'formata o produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 116385

        'Preenche o código de objProduto
        objProduto.sCodigo = sProdutoFormatado

    End If

    'chama a tela de produtos
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_LabelProduto_Click:

    Select Case gErr

        Case 116385

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158655)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)
'evento de inclusão de um item selecionado no browser Produto

Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
    
    'Preenche campo Produto
    Produto.PromptInclude = False
    Produto.Text = CStr(objProduto.sCodigo)
    Produto.PromptInclude = True
    Call Produto_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 158656)

    End Select
    
    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Produto Then
            Call LabelProduto_Click
        End If
    End If

End Sub

'**** inicio do trecho a ser copiado *****
Public Sub Form_Activate()
    Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    gi_ST_SetaIgnoraClick = 1
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    
    Set objEventoProduto = Nothing
    Set objEventoCusto = Nothing
    
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Custo Fixo de Produtos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CustoFixoProd"
    
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

Public Sub Unload(objme As Object)
   ' Parent.UnloadDoFilho
    
   RaiseEvent Unload
    
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property
'***** fim do trecho a ser copiado ******

Private Sub LabelProduto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Produto, Source, X, Y)
End Sub

Private Sub LabelProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Produto, Button, Shift, X, Y)
End Sub

Private Sub LabelDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelDescricao, Source, X, Y)
End Sub

Private Sub LabelDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelDescricao, Button, Shift, X, Y)
End Sub

Private Sub LabelUMEstoque_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelUMEstoque, Source, X, Y)
End Sub

Private Sub LabelUMEstoque_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelUMEstoque, Button, Shift, X, Y)
End Sub

Private Sub LabelData_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelData, Source, X, Y)
End Sub

Private Sub LabelData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelData, Button, Shift, X, Y)
End Sub

Private Sub LabelCustoAnterior_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCustoAnterior, Source, X, Y)
End Sub

Private Sub LabelCustoAnterior_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCustoAnterior, Button, Shift, X, Y)
End Sub

Private Sub LabelCustoRateio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCustoRateio, Source, X, Y)
End Sub

Private Sub LabelCustoRateio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCustoRateio, Button, Shift, X, Y)
End Sub

Private Sub LabelCustoAplicado_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCustoAplicado, Source, X, Y)
End Sub

Private Sub LabelCustoAplicado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCustoAplicado, Button, Shift, X, Y)
End Sub

Private Function Carregar_Data_Padrao() As Long

Dim lErro As Long, dtDataUltCalc As Date

On Error GoTo Erro_Carregar_Data_Padrao

    gdtDataPadrao = gdtDataAtual
    
    lErro = CF("CustoFixo_ObterDataUltCalc", giFilialEmpresa, dtDataUltCalc)
    If lErro <> SUCESSO Then gError 106649
    
    If dtDataUltCalc <> DATA_NULA Then gdtDataPadrao = dtDataUltCalc
    
    Carregar_Data_Padrao = SUCESSO
     
    Exit Function
    
Erro_Carregar_Data_Padrao:

    Carregar_Data_Padrao = gErr
     
    Select Case gErr
          
        Case 106649
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158657)
     
    End Select
     
    Exit Function

End Function
