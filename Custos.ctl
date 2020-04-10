VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl Custos 
   ClientHeight    =   5715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   KeyPreview      =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   6795
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4965
      ScaleHeight     =   495
      ScaleWidth      =   1650
      TabIndex        =   22
      Top             =   180
      Width           =   1710
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "Custos.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   615
         Picture         =   "Custos.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1140
         Picture         =   "Custos.ctx":068C
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Custos"
      Height          =   3420
      Left            =   180
      TabIndex        =   8
      Top             =   2160
      Width           =   4410
      Begin VB.ComboBox Mes 
         Height          =   315
         ItemData        =   "Custos.ctx":080A
         Left            =   2415
         List            =   "Custos.ctx":080C
         OLEDropMode     =   1  'Manual
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   330
         Width           =   1620
      End
      Begin VB.ComboBox Ano 
         Height          =   315
         ItemData        =   "Custos.ctx":080E
         Left            =   855
         List            =   "Custos.ctx":0810
         OLEDropMode     =   1  'Manual
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   315
         Width           =   825
      End
      Begin MSMask.MaskEdBox CustoRProducao 
         Height          =   315
         Left            =   2355
         TabIndex        =   11
         Top             =   975
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0.0000"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CustoStd 
         Height          =   315
         Left            =   2355
         TabIndex        =   12
         Top             =   1443
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0.0000"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CustoReposicaoMes 
         Height          =   315
         Left            =   2355
         TabIndex        =   26
         Top             =   2850
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   "#,##0.0000"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Custo de Reposição:"
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
         Left            =   495
         TabIndex        =   27
         Top             =   2895
         Width           =   1785
      End
      Begin VB.Label CustoMRProducao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2355
         TabIndex        =   20
         Top             =   2379
         Width           =   1650
      End
      Begin VB.Label LabelCustoMProd 
         AutoSize        =   -1  'True
         Caption         =   "Custo Médio Produção:"
         Enabled         =   0   'False
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
         Left            =   300
         TabIndex        =   19
         Top             =   2400
         Width           =   1995
      End
      Begin VB.Label LabelCustoStd 
         AutoSize        =   -1  'True
         Caption         =   "Custo Standard:"
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
         Left            =   930
         TabIndex        =   18
         Top             =   1485
         Width           =   1380
      End
      Begin VB.Label LabelCustoReal 
         AutoSize        =   -1  'True
         Caption         =   "Custo Real de Produção:"
         Enabled         =   0   'False
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
         Left            =   150
         TabIndex        =   17
         Top             =   1035
         Width           =   2145
      End
      Begin VB.Label LabelCustoMedio 
         AutoSize        =   -1  'True
         Caption         =   "Custo Médio:"
         Enabled         =   0   'False
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
         Left            =   1170
         TabIndex        =   16
         Top             =   1940
         Width           =   1125
      End
      Begin VB.Label CustoMedio 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2355
         TabIndex        =   15
         Top             =   1911
         Width           =   1650
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1920
         TabIndex        =   14
         Top             =   375
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   390
         TabIndex        =   13
         Top             =   360
         Width           =   405
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Produto"
      Height          =   1410
      Left            =   180
      TabIndex        =   2
      Top             =   210
      Width           =   4410
      Begin MSMask.MaskEdBox Produto 
         Height          =   315
         Left            =   1245
         TabIndex        =   0
         Top             =   315
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
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
         Left            =   2880
         TabIndex        =   7
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label1 
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
         TabIndex        =   6
         Top             =   780
         Width           =   930
      End
      Begin VB.Label Descricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   480
         Left            =   1230
         TabIndex        =   5
         Top             =   765
         Width           =   3030
      End
      Begin VB.Label ProdutoLabel 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
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
         Left            =   480
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   4
         Top             =   360
         Width           =   660
      End
      Begin VB.Label LblUMEstoque 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3435
         TabIndex        =   3
         Top             =   300
         Width           =   825
      End
   End
   Begin MSMask.MaskEdBox CustoReposicao 
      Height          =   315
      Left            =   2595
      TabIndex        =   1
      Top             =   1785
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   556
      _Version        =   393216
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
      Format          =   "#,##0.0000"
      PromptChar      =   " "
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      Caption         =   "Custo de Reposição Atual:"
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
      Left            =   270
      TabIndex        =   21
      Top             =   1830
      Width           =   2280
   End
End
Attribute VB_Name = "Custos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iProdutoAlterado As Integer
Dim gobjMovEstoque As ClassMovEstoque
Dim giAlmoxarifado As Integer

'Mnemônicos
Private Const PRODUTO_CODIGO As String = "Produto_Codigo"
Private Const PRODUTO_DESCRICAO As String = "Produto_Descricao"
Private Const ALMOXARIFADO_NOMEREDUZIDO As String = "Almoxarifado_NomeRed"
Private Const VALORAJUSTE_ALMOXARIFADO As String = "ValorAjuste_Almox"
Private Const CONTACONTABILEST1 As String = "ContaContabilEst"
Private Const CTAAJUSTECUSTOSTANDARD As String = "CtaAjusteCStandard"

Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long
'Calcula os valores dos mnemonicos usados na contabilidade

Dim lErro As Long
Dim dValorAjusteAlmoxarifado As Double
Dim sContaMascarada As String
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objEstoqueProduto As New ClassEstoqueProduto
Dim objItemMovEstoque As ClassItemMovEstoque
Dim objMnemonico As New ClassMnemonicoCTBValor

On Error GoTo Erro_Calcula_Mnemonico
    
    'Verifica qual mnemonico deve ser calculado
    Select Case objMnemonicoValor.sMnemonico
    
        
        'Se for o código do produto
        Case PRODUTO_CODIGO
        
            'Guarda na coleção o código do produto que está sendo exibido na tela
            objMnemonicoValor.colValor.Add Produto.Text
        
        'Se for a descrição do produto
        Case PRODUTO_DESCRICAO
        
            'Guarda na coleção a descrição do produto que está sendo exibido na tela
            objMnemonicoValor.colValor.Add Descricao.Caption
        
        'Se for o nome do almoxarifado
        Case ALMOXARIFADO_NOMEREDUZIDO
        
            'guarda em objAlmoxarifado o código do Almoxarifado para o qual estão sendo contabilizados os movimentos
            objAlmoxarifado.iCodigo = giAlmoxarifado
            
            'Lê o Almoxarifado no BD para obter o nome reduzido
            lErro = CF("Almoxarifado_Le", objAlmoxarifado)
            If lErro <> SUCESSO And lErro <> 25056 Then gError 79898
            
            'Se não encontrou => erro
            If lErro = 25056 Then gError 79899
                        
            'Guarda na coleção o nome reduzido do almoxarifado
            objMnemonicoValor.colValor.Add objAlmoxarifado.sNomeReduzido
        
        'Se for o valor do ajuste a ser contabilizado
        Case VALORAJUSTE_ALMOXARIFADO
        
            'Faz um loop para cada movimento de estoque
            For Each objItemMovEstoque In gobjMovEstoque.colItens
                
                'Se o Almoxarifado que está tendo seus movimentos contabilizados é o mesmo almoxarifado do movimento de estoque
                If giAlmoxarifado = objItemMovEstoque.iAlmoxarifado Then
                    'Acumula o valor do movimento de estoque
                    dValorAjusteAlmoxarifado = dValorAjusteAlmoxarifado + objItemMovEstoque.dCusto
                End If
            
            Next
            
            'Guarda na coleção o valor total a ser contabilizado para o almoxarifado
            objMnemonicoValor.colValor.Add dValorAjusteAlmoxarifado
            
        'Se for a conta contábil de estoque
        Case CONTACONTABILEST1
            
            'Seta objItemMovEstoque como o primeiro movimento de estoque contido em colItens
            Set objItemMovEstoque = gobjMovEstoque.colItens.Item(1)
            
            'Guarda no objEstoqueProduto os códigos do Produto e do Almoxarifado
            objEstoqueProduto.sProduto = objItemMovEstoque.sProduto
            objEstoqueProduto.iAlmoxarifado = giAlmoxarifado
            
            'Lê as informações do par Produto x Almoxarifado
            lErro = CF("EstoqueProduto_Le", objEstoqueProduto)
            If lErro <> SUCESSO And lErro <> 21306 Then gError 79896
            
            'Se não encontrou => erro
            If lErro = 21306 Then gError 79897
            
            If Len(Trim(objEstoqueProduto.sContaContabil)) <> 0 Then
            
                'Mascara a conta contábil
                lErro = Mascara_MascararConta(objEstoqueProduto.sContaContabil, sContaMascarada)
                If lErro <> SUCESSO Then gError 79904
            
            Else
                sContaMascarada = ""
            End If
            
            'Guarda na coleção a conta contábil de estoque do produto
            objMnemonicoValor.colValor.Add sContaMascarada
        
        'Se for a conta contábil de ajuste de custo standard
        Case CTAAJUSTECUSTOSTANDARD
        
            'Guarda em objMnemonico o nome do Mnemonico que deve ser procurado
            objMnemonico.sMnemonico = CTAAJUSTECUSTOSTANDARD
            
            'Lê em MnemonicoCTBValor o mnem6onico que está sendo procurado
            lErro = CF("MnemonicoCTBValor_Le", objMnemonico)
            If lErro <> SUCESSO And lErro <> 39690 Then gError 79900
            
            'Se não encontrou o mnemônico => erro
            If lErro = 39690 Then gError 79901
            
            'Guarda na coleção a conta contábil de ajuste do custo standard
            objMnemonicoValor.colValor.Add objMnemonico.sValor
        
        'Se não for nenhum dos mnemônicos acima
        Case Else
        
            'Erro
            gError 79902
        
        End Select
    
    Calcula_Mnemonico = SUCESSO
    
    Exit Function
    
Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = gErr
    
    Select Case gErr
        
        Case 79896, 79898, 79900
        
        Case 79902
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO
        
        Case 79901, 79899, 79897
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158711)
            
    End Select
    
    Exit Function
    
End Function
Public Function Trata_Parametros(Optional objProduto As ClassProduto) As Long
'se for passado um produto traze-lo p/a tela

Dim lErro As Long
Dim sProdutoEnxuto As String

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum Produto
    If Not (objProduto Is Nothing) Then

        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 41698

        If lErro = 28030 Then Error 41700
        
        lErro = Traz_Produto_Tela(objProduto)
        If lErro <> SUCESSO Then Error 41699

    Else
        Call Limpa_Tela_Custos
    End If

    iAlterado = 0
    iProdutoAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 41698, 41699
        
        Case 41700
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", Err, objProduto.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158712)

    End Select

    iAlterado = 0
    iProdutoAlterado = 0

    Exit Function

End Function

Public Function Gravar_Registro()

Dim lErro As Long
Dim objSldMesEst As New ClassSldMesEst
Dim iMes As Integer
Dim dCustoReposicao As Double, objAux As Object

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'critica o preenchimento
    If Len(Trim(Produto.ClipText)) = 0 Then Error 41701
    If Ano.ListIndex = -1 Then Error 41702
    If Mes.ListIndex = -1 Then Error 41703
    
    'move dados da tela para a memoria
    lErro = Move_Tela_Memoria(iMes, dCustoReposicao, objSldMesEst)
    If lErro <> SUCESSO Then Error 41706
    
    lErro = Trata_Alteracao(objSldMesEst, objSldMesEst.iFilialEmpresa, objSldMesEst.iAno, objSldMesEst.sProduto)
    If lErro <> SUCESSO Then Error 32326
    
    Set objAux = Me
    lErro = CF("Custos_Grava", iMes, dCustoReposicao, objSldMesEst, objAux)
    If lErro <> SUCESSO Then Error 41707

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 32326

        Case 41701
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)
            Produto.SetFocus
            
        Case 41702
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ANO_NAO_PREENCHIDO", Err)
            Ano.SetFocus
                    
        Case 41703
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MES_NAO_PREENCHIDO", Err)
            Mes.SetFocus
            
        Case 41706, 41707
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158713)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_Custos()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Custos

    Call Limpa_Tela(Me)

    LblUMEstoque.Caption = ""
    Descricao.Caption = ""
    CustoMedio.Caption = ""
    CustoMRProducao.Caption = ""

    'Mostra o ultimo Ano
    Ano.ListIndex = Ano.ListCount - 1

    iAlterado = 0
    iProdutoAlterado = 0

    Exit Sub

Erro_Limpa_Tela_Custos:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158714)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 41709

    Call Limpa_Tela_Custos

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 41709

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158715)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then Error 41710

    Call Limpa_Tela_Custos
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then Error 41781

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case 41710, 41781

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158716)

    End Select

    Exit Sub

End Sub

Private Sub CustoReposicao_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoReposicao_Validate(Cancel As Boolean)

Dim lErro As Long
    
On Error GoTo Erro_CustoReposicao_Validate

    If Len(Trim(CustoReposicao.ClipText)) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(CustoReposicao.Text)
        If lErro <> SUCESSO Then Error 52860
        
    End If
    
    Exit Sub
    
Erro_CustoReposicao_Validate:

    Cancel = True


    Select Case Err
        
        Case 52860
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158717)
    
    End Select
    
    Exit Sub

End Sub

Private Sub CustoRProducao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoRProducao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CustoRProducao_Validate

    If Len(Trim(CustoRProducao.ClipText)) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(CustoRProducao.Text)
        If lErro <> SUCESSO Then Error 41711
        
    End If

    Exit Sub
    
Erro_CustoRProducao_Validate:

    Cancel = True


    Select Case Err
    
        Case 41711
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158718)
            
    End Select
    
    Exit Sub

End Sub

Private Sub CustoStd_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub CustoStd_Validate(Cancel As Boolean)

Dim lErro As Long
    
On Error GoTo Erro_CustoStd_Validate

    If Len(Trim(CustoStd.ClipText)) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(CustoStd.Text)
        If lErro <> SUCESSO Then Error 41712
        
    End If
    
    Exit Sub
    
Erro_CustoStd_Validate:

    Cancel = True


    Select Case Err
        
        Case 41712
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158719)
    
    End Select
    
    Exit Sub

End Sub

Public Sub Form_Activate()

    'Carrega índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoProduto = New AdmEvento

    'inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then Error 41713

''    lErro = CF("Carga_Arvore_Produto_Inventariado",TvwProduto.Nodes)
''    If lErro <> SUCESSO Then Error 41714

    lErro = Inicializa_Tela()
    If lErro <> SUCESSO Then Error 41715
    
    CustoRProducao.Format = FORMATO_CUSTO
    
    CustoReposicao.Format = FORMATO_CUSTO
    
    CustoStd.Format = FORMATO_CUSTO
    
    iProdutoAlterado = 0
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 41713, 41715

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158720)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Function Inicializa_Tela() As Long

Dim lErro As Long
Dim colAnos As New Collection
Dim iAnoAtual As Integer, iIndice As Integer

On Error GoTo Erro_Inicializa_Tela

    'Le todos os anos da filial corrente
    lErro = CF("EstoqueMes_Le_Anos", colAnos)
    If lErro <> SUCESSO And lErro <> 41745 Then Error 41716

    If lErro = 41745 Then Error 41717
    
    Ano.Clear
    
    'Carrega todos os anos
    For iIndice = 1 To colAnos.Count
        Ano.AddItem CStr(colAnos.Item(iIndice))
        Ano.ItemData(Ano.NewIndex) = colAnos.Item(iIndice)
    Next

    'Mostra o ultimo Ano
    Ano.ListIndex = Ano.ListCount - 1

    Inicializa_Tela = SUCESSO

    Exit Function

Erro_Inicializa_Tela:

    Inicializa_Tela = Err

    Select Case Err

        Case 41716
        
        Case 41717
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_ANOS_INEXISTENTES", Err, giFilialEmpresa)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158721)

    End Select

    Exit Function

End Function

Private Sub Ano_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Ano_Click()

Dim lErro As Long
Dim iAno As Integer, iMes As Integer, iIndice As Integer
Dim colMeses As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objSldMesEst As New ClassSldMesEst
Dim sMes As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_Ano_Click

    If Ano.ListIndex = -1 Then Exit Sub

    Call Limpa_Custos

    iAno = Ano.ItemData(Ano.ListIndex)

    'lista os Meses daquele Ano presentes na tabela EstoqueMes
    lErro = CF("EstoqueMes_Le_Meses", iAno, colMeses)
    If lErro <> SUCESSO And lErro <> 41750 Then Error 41718

    If lErro = 41750 Then Error 41719

    Mes.Clear
    
    For iIndice = 1 To colMeses.Count
        
        lErro = MesNome(colMeses.Item(iIndice), sMes)
        If lErro <> SUCESSO Then Error 41648
        
        Mes.AddItem sMes
        Mes.ItemData(Mes.NewIndex) = colMeses.Item(iIndice)
        
    Next

    If Mes.ListIndex = -1 Then Exit Sub

    iMes = Mes.ItemData(Mes.ListIndex)

    'Se o produto estiver preenchido, trazer os dados de custo de SldMesEst
    If Len(Trim(Produto.ClipText)) = 0 Then Exit Sub

    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 41720
            
    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 52125
        
        If lErro = 28030 Then Error 52126
        
        If iAlterado = 0 Then
        
            lErro = Preenche_Custos(iMes, iAno, objProduto)
            If lErro <> SUCESSO Then Error 52127
            
            iAlterado = 0
            
        Else
        
            lErro = Preenche_Custos(iMes, iAno, objProduto)
            If lErro <> SUCESSO Then Error 55191
        
        End If
        
    End If

    Exit Sub

Erro_Ano_Click:

    Select Case Err

        Case 41718, 41720, 41648, 52125, 52127, 55191
        
        Case 41719
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ESTOQUEMES_MESES_INEXISTENTES", Err, giFilialEmpresa, iAno)
        
        Case 52126
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)
      
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158722)

    End Select

    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoProduto = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
    If lErro <> SUCESSO Then Error 41721

    Exit Sub

Erro_Form_Unload:

    Select Case Err

        Case 41721

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158723)

    End Select

    Exit Sub

End Sub

Private Sub Mes_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Function Desabilita_Custos(iApropriacao As Integer) As Long
'desabilita o custo de acordo com sua apropriacao

Dim lErro As Long

On Error GoTo Erro_Desabilita_Custos

    Select Case iApropriacao
        
        'caso custo medio - -> CustoStd e CustoRealProducao desabilitados
        Case APROPR_CUSTO_MEDIO
            LabelCustoMedio.Enabled = True
'            CustoStd.Enabled = False
'            LabelCustoStd.Enabled = False
            CustoRProducao.Enabled = False
            LabelCustoReal.Enabled = False
            LabelCustoMProd.Enabled = False
            
            CustoRProducao.PromptInclude = False
            CustoRProducao.Text = ""
            CustoRProducao.PromptInclude = True
'            CustoStd.PromptInclude = False
'            CustoStd.Text = ""
'            CustoStd.PromptInclude = True
            CustoMRProducao.Caption = ""
            
        'caso custo STANDARD - -> CustoStd abilitado e CustoRealProducao desabilitado
        Case APROPR_CUSTO_STANDARD
'            CustoStd.Enabled = True
'            LabelCustoStd.Enabled = True
            CustoRProducao.Enabled = False
            LabelCustoReal.Enabled = False
            LabelCustoMedio.Enabled = False
            LabelCustoMProd.Enabled = False
        
            CustoRProducao.PromptInclude = False
            CustoRProducao.Text = ""
            CustoRProducao.PromptInclude = True
            CustoMedio.Caption = ""
            CustoMRProducao.Caption = ""
        
        'caso custo Real - -> CustoStd desabilitado e CustoRealProducao abilitado
        Case APROPR_CUSTO_REAL
            CustoRProducao.Enabled = True
            LabelCustoReal.Enabled = True
            LabelCustoMProd.Enabled = True
'            CustoStd.Enabled = False
'            LabelCustoStd.Enabled = False
            LabelCustoMedio.Enabled = False
            
'            CustoStd.PromptInclude = False
'            CustoStd.Text = ""
'            CustoStd.PromptInclude = True
            CustoMedio.Caption = ""
            
        'Case Apropriacao diferente - - - > Erro
        Case Else
            Error 52128
        
        End Select
   
    Exit Function
    
    Desabilita_Custos = SUCESSO
    
    Exit Function

Erro_Desabilita_Custos:

    Desabilita_Custos = Err
    
    Select Case Err
        
        Case 52128
            lErro = Rotina_Erro(vbOKOnly, "ERRO_APROPRIACAO_CUSTO_INEXISTENTE", Err, Produto.Text)
            Produto.SetFocus
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158724)

    End Select

    Exit Function

End Function

Private Function Preenche_Custos(iMes As Integer, iAno As Integer, objProduto As ClassProduto) As Long
'Le do BD os custos utilizados na tela e os preenche

Dim lErro As Long
Dim objSldMesEst As New ClassSldMesEst
Dim dCustoMRProducao As Double

On Error GoTo Erro_Preenche_Custos

    Call Limpa_Custos
    
    objSldMesEst.iFilialEmpresa = giFilialEmpresa
    objSldMesEst.iAno = iAno
    objSldMesEst.sProduto = objProduto.sCodigo

    'Lê os custos referentes
    lErro = CF("SldMesEst_Le_Custos", iMes, objSldMesEst)
    If lErro <> SUCESSO And lErro <> 41755 Then Error 41722

    If lErro = 41755 Then Error 41723
    
    CustoStd.Text = CStr(objSldMesEst.dCustoStandard(iMes))
    
    Select Case objProduto.iApropriacaoCusto
    
        Case APROPR_CUSTO_MEDIO
            CustoMedio.Caption = Formata_Custo(CStr(objSldMesEst.dCustoMedio(iMes)))
            
'        Case APROPR_CUSTO_STANDARD
'            CustoStd.Text = CStr(objSldMesEst.dCustoStandard(iMes))
            
        Case APROPR_CUSTO_REAL
            CustoRProducao.Text = CStr(objSldMesEst.dCustoProducao(iMes))
        
            'calcula o ultimo custo medio Apurado
            lErro = CF("CustoMedioProducaoApurado_Le_Mes", objProduto.sCodigo, iMes, iAno, dCustoMRProducao)
            If lErro <> SUCESSO Then Error 52134
            
            CustoMRProducao.Caption = Formata_Custo(CStr(dCustoMRProducao))
            
    End Select
       
    CustoReposicaoMes.Text = CStr(objSldMesEst.dCustoReposicao(iMes))
    
    Preenche_Custos = SUCESSO
    
    Exit Function

Erro_Preenche_Custos:

    Preenche_Custos = Err
    
    Select Case Err

        Case 41722, 52134
        
        Case 41723
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CUSTOS_INEXISTENTES", Err, giFilialEmpresa, iAno, objProduto.sCodigo)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158725)

    End Select

    Exit Function

End Function

Private Sub Mes_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iMes As Integer, iAno As Integer
Dim objSldMesEst As New ClassSldMesEst
Dim objProduto As New ClassProduto

On Error GoTo Erro_Mes_Click

    'Se não há nenhum mes selecionado
    If Mes.ListIndex = -1 Then Exit Sub

    Call Limpa_Custos

    iMes = Mes.ItemData(Mes.ListIndex)
    iAno = Ano.ItemData(Ano.ListIndex)

    'Se o produto estiver preenchido, trazer os dados de custo de SldMesEst
    If Len(Trim(Produto.ClipText)) = 0 Then Exit Sub

    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then Error 41724

    If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
        
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then Error 52129
        
        If lErro = 28030 Then Error 52130
        
        If iAlterado = 0 Then
        
            lErro = Preenche_Custos(iMes, iAno, objProduto)
            If lErro <> SUCESSO Then Error 52131
            
            iAlterado = 0
            
        Else
        
            lErro = Preenche_Custos(iMes, iAno, objProduto)
            If lErro <> SUCESSO Then Error 55190
        
        End If
        
    End If

    Exit Sub

Erro_Mes_Click:

    Select Case Err

        Case 41724, 52129
        
        Case 52130
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", Err, objProduto.sCodigo)
        
        Case 52131, 55190
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158726)

    End Select

    Exit Sub

End Sub

Private Sub Produto_Change()

    iAlterado = REGISTRO_ALTERADO
    iProdutoAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim iAno As Integer, iMes As Integer

On Error GoTo Erro_Produto_Validate

    If iProdutoAlterado = REGISTRO_ALTERADO Then

        'Verifica preenchimento de Produto
        If Len(Trim(Produto.ClipText)) > 0 Then

            sProduto = Produto.Text

            'Critica o formato do Produto e se existe no BD
            lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
            If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then Error 41725

            'O Produto é Gerencial
            If lErro = 25043 Then Error 41726

            'O Produto não está cadastrado
            If lErro = 25041 Then Error 41727

            If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then Error 41729

            Descricao.Caption = objProduto.sDescricao
            LblUMEstoque.Caption = objProduto.sSiglaUMEstoque
            
            lErro = Desabilita_Custos(objProduto.iApropriacaoCusto)
            If lErro <> SUCESSO Then Error 52133
            
            Call Mes_Click
            
        Else
        
            Descricao.Caption = ""
            LblUMEstoque.Caption = ""
            
        End If

    End If

    Exit Sub

Erro_Produto_Validate:

    Cancel = True


    Select Case Err
        
        Case 41725

        Case 41726
            
        Case 41727
            'Não encontrou Produto no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)

            Else
                Descricao.Caption = ""
                LblUMEstoque.Caption = ""
            End If
            
        Case 41729
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", Err)
        
        Case 52133
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158727)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoLabel_Click()

Dim colSelecao As New Collection
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long
Dim sSelecao  As String

On Error GoTo Erro_ProdutoLabel_Click

    sSelecao = "ControleEstoque <> ?"
    colSelecao.Add PRODUTO_CONTROLE_SEM_ESTOQUE
    
    'Verifica se Produto está preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 55188

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objProduto.sCodigo = sProdutoFormatado
        Else
            objProduto.sCodigo = ""
        End If

    End If

    'Chama a tela de browse
    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProduto, sSelecao)

    Exit Sub
    
Erro_ProdutoLabel_Click:

    Select Case Err

        Case 55188
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158728)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProduto As String

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    Call Limpa_Custos
    
    lErro = Traz_Produto_Tela(objProduto)
    If lErro <> SUCESSO Then Error 41731

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then Error 41780

    Me.Show

    iAlterado = 0

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case Err

        Case 41731, 41780
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158729)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Custos()
    
    If iAlterado = 0 Then
        CustoRProducao.PromptInclude = False
        CustoRProducao.Text = ""
        CustoRProducao.PromptInclude = True
        CustoStd.PromptInclude = False
        CustoStd.Text = ""
        CustoStd.PromptInclude = True
        CustoMedio.Caption = ""
        CustoMRProducao.Caption = ""
        iAlterado = 0
    Else
        CustoRProducao.PromptInclude = False
        CustoRProducao.Text = ""
        CustoRProducao.PromptInclude = True
        CustoStd.PromptInclude = False
        CustoStd.Text = ""
        CustoStd.PromptInclude = True
        CustoMedio.Caption = ""
        CustoMRProducao.Caption = ""
    End If

End Sub

Private Function Traz_Produto_Tela(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim objTipoProduto As New ClassTipoDeProduto
Dim iAno As Integer
Dim iMes As Integer

On Error GoTo Erro_Traz_Produto_Tela

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
    If lErro <> SUCESSO Then Error 41736

    'Coloca o Codigo na tela
    Produto.PromptInclude = False
    Produto.Text = sProdutoEnxuto
    Produto.PromptInclude = True

    'Coloca os demais dados do Produto na tela
    Descricao.Caption = objProduto.sDescricao
    LblUMEstoque.Caption = objProduto.sSiglaUMEstoque
        
    'Preenche o Custo Reposição
    CustoReposicao.Text = CStr(objProduto.dCustoReposicao)
        
    lErro = Desabilita_Custos(objProduto.iApropriacaoCusto)
    If lErro <> SUCESSO Then Error 52162
    
    If Mes.ListIndex <> -1 And Ano.ListIndex <> -1 Then

        iAno = Ano.ItemData(Ano.ListIndex)
        iMes = Mes.ItemData(Mes.ListIndex)

        lErro = Preenche_Custos(iMes, iAno, objProduto)
        If lErro <> SUCESSO Then Error 52132

    End If
    
    Traz_Produto_Tela = SUCESSO

    Exit Function

Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = Err

    Select Case Err
        
        Case 41736
            lErro = Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", Err, objProduto.sCodigo)

        Case 52132, 52162

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158730)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(iMes As Integer, dCustoReposicao As Double, objSldMesEst As ClassSldMesEst) As Long
'move os dados da tela para as variaveis

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    objSldMesEst.iFilialEmpresa = giFilialEmpresa

    'Verifica se Produto está preenchido
    If Len(Trim(Produto.ClipText)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then Error 41737

        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            objSldMesEst.sProduto = sProdutoFormatado
        Else
            objSldMesEst.sProduto = ""
        End If

    End If

    If Ano.ListIndex <> -1 Then objSldMesEst.iAno = Ano.ItemData(Ano.ListIndex)

    If Mes.ListIndex <> -1 Then iMes = Mes.ItemData(Mes.ListIndex)

    If iMes > 0 Then

        If Len(Trim(CustoRProducao.ClipText)) > 0 Then
            objSldMesEst.dCustoProducao(iMes) = CDbl(CustoRProducao.FormattedText)
        Else
            objSldMesEst.dCustoProducao(iMes) = 0
        End If
        
        If Len(Trim(CustoStd.ClipText)) > 0 Then
            objSldMesEst.dCustoStandard(iMes) = CDbl(CustoStd.FormattedText)
        Else
            objSldMesEst.dCustoStandard(iMes) = 0
        End If
        
        If Len(Trim(CustoReposicaoMes.ClipText)) > 0 Then
            objSldMesEst.dCustoReposicao(iMes) = CDbl(CustoReposicaoMes.FormattedText)
        Else
            objSldMesEst.dCustoReposicao(iMes) = 0
        End If
        
    End If
    
    'Move custoReposicao para Memoria
    dCustoReposicao = StrParaDbl(CustoReposicao.FormattedText)
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err

        Case 41737
            Produto.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158731)

    End Select

    Exit Function

End Function

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long
'esta funcao é chamada a cada atualizacao de baixaparcpag e é responsavel por gerar a contabilizacao correspondente

Dim lErro As Long, lDoc As Long
Dim objItemMovEstoque1 As ClassItemMovEstoque

On Error GoTo Erro_GeraContabilizacao

    Set gobjMovEstoque = vParams(0)
    
    giAlmoxarifado = 0
    
    For Each objItemMovEstoque1 In gobjMovEstoque.colItens
    
        If giAlmoxarifado <> objItemMovEstoque1.iAlmoxarifado Then
    
            giAlmoxarifado = objItemMovEstoque1.iAlmoxarifado
            
            lErro = objContabAutomatica.Obter_Doc(lDoc, gobjMovEstoque.iFilialEmpresa)
            If lErro <> SUCESSO Then gError 83580
            
            'grava a contabilizacao na filial pagadora
            lErro = objContabAutomatica.Gravar_Registro(Me, "Custos", objItemMovEstoque1.lNumIntDoc, 0, 0, 0, lDoc, gobjMovEstoque.iFilialEmpresa)
            If lErro <> SUCESSO Then gError 83581
            
        End If
        
    Next

    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = gErr
     
    Select Case gErr
          
        Case 83580, 83581
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158732)
     
    End Select
     
    Exit Function

End Function


'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objSldMesEst As New ClassSldMesEst
Dim iMes As Integer
Dim sDescricao As String
Dim sUMEstoque As String
Dim dCustoReposicao As Double
Dim iApropriacao As Integer

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "CustoMensal"

    'Lê os dados da Tela EstoqueProduto
    lErro = Move_Tela_Memoria(iMes, dCustoReposicao, objSldMesEst)
    If lErro <> SUCESSO Then Error 41738

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "FilialEmpresa", objSldMesEst.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Produto", objSldMesEst.sProduto, STRING_PRODUTO, "Produto"
    colCampoValor.Add "Ano", objSldMesEst.iAno, 0, "Ano"
    colCampoValor.Add "Mes", iMes, 0, "Mes"
    colCampoValor.Add "Descricao", sDescricao, STRING_PRODUTO_DESCRICAO, "Descricao"
    colCampoValor.Add "SiglaUMEstoque", sUMEstoque, STRING_UM_SIGLA, "SiglaUMEstoque"
    colCampoValor.Add "Apropriacao", iApropriacao, 0, "Apropriacao"

    'Filtro FilialEmpresa
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 41738

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158733)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objSldMesEst As New ClassSldMesEst
Dim iMes As Integer
Dim sDescricao As String
Dim sUMEstoque As String
Dim iApropriacao As Integer

On Error GoTo Erro_Tela_Preenche

    objSldMesEst.sProduto = colCampoValor.Item("Produto").vValor
    objSldMesEst.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
    objSldMesEst.iAno = colCampoValor.Item("Ano").vValor
    iMes = colCampoValor.Item("Mes").vValor
    sDescricao = colCampoValor.Item("Descricao").vValor
    sUMEstoque = colCampoValor.Item("SiglaUMEstoque").vValor
    iApropriacao = colCampoValor.Item("Apropriacao").vValor
    
    'Traz os dados para tela
    lErro = Traz_CustoMensal_Tela(iMes, sDescricao, sUMEstoque, objSldMesEst, iApropriacao)
    If lErro <> SUCESSO Then Error 41739
    
    iAlterado = 0
    iProdutoAlterado = 0

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 41739

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158734)

    End Select

    Exit Sub

End Sub

Private Function Traz_CustoMensal_Tela(ByVal iMes As Integer, sDescricao As String, sUMEstoque As String, objSldMesEst As ClassSldMesEst, ByVal iApropriacao As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dCustoMRProducao As Double

On Error GoTo Erro_Traz_CustoMensal_Tela

    'le os custos referentes
    lErro = CF("SldMesEst_Le_Custos", iMes, objSldMesEst)
    If lErro <> SUCESSO And lErro <> 41755 Then Error 41740

    If lErro = 41755 Then Error 41741

    Produto.PromptInclude = False
    Produto.Text = objSldMesEst.sProduto
    Produto.PromptInclude = True

    lErro = Desabilita_Custos(iApropriacao)
    If lErro <> SUCESSO Then Error 52162

    For iIndice = 0 To Ano.ListCount - 1
        If Ano.ItemData(iIndice) = objSldMesEst.iAno Then
            Ano.ListIndex = iIndice
            Exit For
        End If
    Next
    
    Mes.ListIndex = -1
    For iIndice = 0 To Mes.ListCount - 1
        If Mes.ItemData(iIndice) = iMes Then
            Mes.ListIndex = iIndice
            Exit For
        End If
    Next

    Descricao.Caption = sDescricao
    LblUMEstoque.Caption = sUMEstoque

    Traz_CustoMensal_Tela = SUCESSO

    Exit Function

Erro_Traz_CustoMensal_Tela:

    Traz_CustoMensal_Tela = Err

    Select Case Err

        Case 41740, 55187
        
        Case 41741
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CUSTOS_INEXISTENTES", Err, objSldMesEst.iFilialEmpresa, objSldMesEst.iAno, objSldMesEst.sProduto)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158735)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CUSTOS
    Set Form_Load_Ocx = Me
    Caption = "Custos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Custos"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Produto Then
            Call ProdutoLabel_Click
        End If
    End If

End Sub


Private Sub CustoMRProducao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CustoMRProducao, Source, X, Y)
End Sub

Private Sub CustoMRProducao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CustoMRProducao, Button, Shift, X, Y)
End Sub

Private Sub LabelCustoMProd_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCustoMProd, Source, X, Y)
End Sub

Private Sub LabelCustoMProd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCustoMProd, Button, Shift, X, Y)
End Sub

Private Sub LabelCustoStd_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCustoStd, Source, X, Y)
End Sub

Private Sub LabelCustoStd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCustoStd, Button, Shift, X, Y)
End Sub

Private Sub LabelCustoReal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCustoReal, Source, X, Y)
End Sub

Private Sub LabelCustoReal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCustoReal, Button, Shift, X, Y)
End Sub

Private Sub LabelCustoMedio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCustoMedio, Source, X, Y)
End Sub

Private Sub LabelCustoMedio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCustoMedio, Button, Shift, X, Y)
End Sub

Private Sub CustoMedio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CustoMedio, Source, X, Y)
End Sub

Private Sub CustoMedio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CustoMedio, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Descricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Descricao, Source, X, Y)
End Sub

Private Sub Descricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Descricao, Button, Shift, X, Y)
End Sub

Private Sub ProdutoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoLabel, Source, X, Y)
End Sub

Private Sub ProdutoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoLabel, Button, Shift, X, Y)
End Sub

Private Sub LblUMEstoque_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LblUMEstoque, Source, X, Y)
End Sub

Private Sub LblUMEstoque_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LblUMEstoque, Button, Shift, X, Y)
End Sub

Private Sub Label35_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label35, Source, X, Y)
End Sub

Private Sub Label35_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label35, Button, Shift, X, Y)
End Sub

Private Sub CustoReposicaoMes_Change()
    
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CustoReposicaoMes_Validate(Cancel As Boolean)

Dim lErro As Long
    
On Error GoTo Erro_CustoReposicaoMes_Validate

    If Len(Trim(CustoReposicaoMes.ClipText)) > 0 Then
    
        lErro = Valor_NaoNegativo_Critica(CustoReposicaoMes.Text)
        If lErro <> SUCESSO Then Error 52860
        
    End If
    
    Exit Sub
    
Erro_CustoReposicaoMes_Validate:

    Cancel = True


    Select Case Err
        
        Case 52860
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158736)
    
    End Select
    
    Exit Sub

End Sub


