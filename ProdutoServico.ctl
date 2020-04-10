VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl ProdutoServico 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9480
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9480
   Begin VB.Frame Frame8 
      Caption         =   "Serviço"
      Height          =   585
      Left            =   75
      TabIndex        =   19
      Top             =   60
      Width           =   7515
      Begin MSMask.MaskEdBox Produto 
         Height          =   315
         Left            =   885
         TabIndex        =   0
         Top             =   195
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
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
         Left            =   135
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   21
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Descricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2880
         TabIndex        =   20
         Top             =   195
         Width           =   4530
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tributação"
      Height          =   5325
      Left            =   90
      TabIndex        =   16
      Top             =   660
      Width           =   9270
      Begin VB.Frame Frame3 
         Caption         =   "Localização padrão para incidência do imposto "
         Height          =   465
         Left            =   90
         TabIndex        =   34
         Top             =   2220
         Width           =   9090
         Begin VB.OptionButton optEmpresaImp 
            Caption         =   "Na cidade da prestadora do serviço"
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
            Left            =   765
            TabIndex        =   6
            Top             =   210
            Value           =   -1  'True
            Width           =   3540
         End
         Begin VB.OptionButton optClienteImp 
            Caption         =   "Na cidade do cliente"
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
            Left            =   4920
            TabIndex        =   7
            Top             =   210
            Width           =   3540
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Localização padrão para prestação do serviço"
         Height          =   465
         Left            =   90
         TabIndex        =   31
         Top             =   1755
         Width           =   9090
         Begin VB.OptionButton optCliente 
            Caption         =   "Na cidade do cliente"
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
            Left            =   4920
            TabIndex        =   5
            Top             =   210
            Width           =   3540
         End
         Begin VB.OptionButton optEmpresa 
            Caption         =   "Na cidade da prestadora do serviço"
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
            Left            =   765
            TabIndex        =   4
            Top             =   210
            Value           =   -1  'True
            Width           =   3540
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Lista de Código de Serviços por Cidade"
         Height          =   2625
         Left            =   90
         TabIndex        =   22
         Top             =   2670
         Width           =   9090
         Begin VB.CommandButton BotaoCodServ 
            Caption         =   "Código do Serviço"
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
            Left            =   1485
            TabIndex        =   10
            ToolTipText     =   "Abre o Browse de Código de Serviços"
            Top             =   2205
            Width           =   2160
         End
         Begin VB.CommandButton BotaoLimparGrid 
            Caption         =   "Limpar Grid"
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
            Left            =   7635
            TabIndex        =   11
            ToolTipText     =   "Limpa o Grid"
            Top             =   2205
            Width           =   1380
         End
         Begin VB.CommandButton BotaoCidades 
            Caption         =   "Cidades"
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
            Left            =   60
            TabIndex        =   9
            ToolTipText     =   "Abre o Browse de Cidades"
            Top             =   2205
            Width           =   1380
         End
         Begin VB.TextBox CodServDesc 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   5505
            TabIndex        =   29
            Top             =   1680
            Width           =   3270
         End
         Begin MSMask.MaskEdBox CodServ 
            Height          =   315
            Left            =   2970
            TabIndex        =   27
            Top             =   1695
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CidadeIBGECod 
            Height          =   315
            Left            =   1965
            TabIndex        =   26
            Top             =   1695
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            Enabled         =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Cidade 
            Height          =   315
            Left            =   495
            TabIndex        =   25
            Top             =   1665
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Aliquota 
            Height          =   315
            Left            =   4455
            TabIndex        =   28
            Top             =   1695
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            MaxLength       =   20
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridSRV 
            Height          =   1410
            Left            =   45
            TabIndex        =   8
            Top             =   225
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   2487
            _Version        =   393216
            Rows            =   8
            Cols            =   6
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            ScrollTrack     =   -1  'True
            FocusRect       =   2
         End
         Begin VB.Label DescServico 
            BorderStyle     =   1  'Fixed Single
            Height          =   540
            Left            =   60
            TabIndex        =   30
            Top             =   1665
            Width           =   8970
         End
      End
      Begin MSMask.MaskEdBox CNAE 
         Height          =   315
         Left            =   870
         TabIndex        =   1
         Top             =   180
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         Format          =   "0000-0/00"
         Mask            =   "#######"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ISSQN 
         Height          =   315
         Left            =   870
         TabIndex        =   2
         Top             =   705
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Format          =   "0000"
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox NBS 
         Height          =   315
         Left            =   870
         TabIndex        =   3
         Top             =   1245
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         Format          =   "0\.0000\.00\.00"
         Mask            =   "#########"
         PromptChar      =   " "
      End
      Begin VB.Label LabelNBS 
         AutoSize        =   -1  'True
         Caption         =   "NBS:"
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
         Left            =   330
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   33
         Top             =   1290
         Width           =   450
      End
      Begin VB.Label DescNBS 
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Left            =   2190
         TabIndex        =   32
         Top             =   1245
         Width           =   6945
      End
      Begin VB.Label DescISSQN 
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Left            =   1530
         TabIndex        =   24
         Top             =   705
         Width           =   7605
      End
      Begin VB.Label LabelISSQN 
         AutoSize        =   -1  'True
         Caption         =   "ISSQN:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   750
         Width           =   645
      End
      Begin VB.Label LabelCNAE 
         AutoSize        =   -1  'True
         Caption         =   "CNAE:"
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
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   18
         Top             =   255
         Width           =   570
      End
      Begin VB.Label DescCNAE 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   510
         Left            =   2190
         TabIndex        =   17
         Top             =   180
         Width           =   6945
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7665
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   90
      Width           =   1695
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1125
         Picture         =   "ProdutoServico.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   615
         Picture         =   "ProdutoServico.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "ProdutoServico.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
End
Attribute VB_Name = "ProdutoServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'todas as variáveis devem ser declaradas

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Dim iProdutoAlterado As Integer

'Grid de Tipos de Mão-de-Obra
Dim objGridSRV As AdmGrid
Dim iGrid_Cidade_Col As Integer
Dim iGrid_CidadeIBGECod_Col As Integer
Dim iGrid_CodServ_Col As Integer
Dim iGrid_CodServDesc_Col As Integer
Dim iGrid_Aliquota_Col As Integer

Dim WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Dim WithEvents objEventoCNAE As AdmEvento
Attribute objEventoCNAE.VB_VarHelpID = -1
Dim WithEvents objEventoISSQN As AdmEvento
Attribute objEventoISSQN.VB_VarHelpID = -1
Dim WithEvents objEventoCidade As AdmEvento
Attribute objEventoCidade.VB_VarHelpID = -1
Dim WithEvents objEventoCodServ As AdmEvento
Attribute objEventoCodServ.VB_VarHelpID = -1
Dim WithEvents objEventoNBS As AdmEvento
Attribute objEventoNBS.VB_VarHelpID = -1

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa a Tela
    Call Limpa_Tela_Produto

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM 'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208134)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

   'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa a Tela
    Call Limpa_Tela_Produto

    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM 'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208135)

    End Select

    Exit Sub

End Sub

Public Function Trata_Parametros(Optional objProduto As ClassProduto) As Long

Dim lErro As Long, sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum Produto
    If Not (objProduto Is Nothing) Then
       
        'Traz os dados para a Tela
        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        'Coloca na tela o Produto selecionado
        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True
        
        sProduto = Produto.Text

        'Critica o formato do Produto e se existe no BD
        lErro = CF("Produto_Critica2", sProduto, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError ERRO_SEM_MENSAGEM

        'Não encontrou o Produto
        If lErro = 25041 Then gError 208136
    
        lErro = Traz_Produto_Tela(objProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 208136
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, sProduto)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208137)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoProduto = New AdmEvento
    Set objEventoCNAE = New AdmEvento
    Set objEventoCidade = New AdmEvento
    Set objEventoCodServ = New AdmEvento
    Set objEventoISSQN = New AdmEvento
    Set objEventoNBS = New AdmEvento

    'Inicializa as máscaras de Produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Set objGridSRV = New AdmGrid
    
    lErro = Inicializa_GridSRV(objGridSRV)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    iAlterado = 0
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208138)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub BotaoLimparGrid_Click()
    Call Grid_Limpa(objGridSRV)
    DescServico.Caption = ""
End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim vbMsgRes As VbMsgBoxResult
Dim sCodProduto As String
Dim iProdutoPreenchido As Integer
Dim iCodigo As Integer
Dim objProdutoFilial As New ClassProdutoFilial
Dim iIndice As Integer
Dim objTipoDeProduto As New ClassTipoDeProduto

On Error GoTo Erro_Produto_Validate

    'Se Produto não foi alterado, sai
    If iProdutoAlterado <> REGISTRO_ALTERADO Then Exit Sub

    'Limpa descrição
    Descricao.Caption = ""

    'Verifica preenchimento de Produto
    If Len(Trim(Produto.ClipText)) > 0 Then

        sCodProduto = Produto.Text
        
        'Critica o formato do Produto e se existe no BD
        lErro = CF("Produto_Critica2", sCodProduto, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError ERRO_SEM_MENSAGEM

        'Não encontrou o Produto
        If lErro = 25041 Then gError 208139

        lErro = Traz_Produto_Tela(objProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    iProdutoAlterado = 0

    Exit Sub

Erro_Produto_Validate:

    Cancel = True
    
    Select Case gErr

        Case 208139
            'Não encontrou Produto no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)
            
            End If
            
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208140)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long
'Verifica se dados de Estoque necessários foram preenchidos

Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Gravar_Registro
    
    'Verifica se o Código do Produto foi preenchido
    If Len(Trim(Produto.ClipText)) = 0 Then gError 208141
    
    'Lê os dados da Tela relacionados ao Estoque
    lErro = Move_Tela_Memoria(objProduto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = Trata_Alteracao(objProduto, objProduto.sCodigo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Grava os dados do Controle de Estoque do Produto no BD
    lErro = CF("ProdutoCNAE_Grava", objProduto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr

        Case 208141
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208142)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(ByVal objProduto As ClassProduto) As Long

'Move os dados da tela para memória
Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim iLinha As Integer
Dim objCodTrib As ClassCodTribMun
Dim objCidade As ClassCidades

On Error GoTo Erro_Move_Tela_Memoria

    'Passa para o formato do BD
    lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    objProduto.sCodigo = sProduto
    
    objProduto.sNBS = NBS.ClipText
    objProduto.sISSQN = ISSQN.ClipText
    objProduto.objProdutoCNAE.sCNAE = CNAE.ClipText
    objProduto.objProdutoCNAE.sProduto = sProduto
    
    If optCliente.Value Then
        objProduto.objProdutoCNAE.iLocServCliente = MARCADO
    Else
        objProduto.objProdutoCNAE.iLocServCliente = DESMARCADO
    End If
    
    If optClienteImp.Value Then
        objProduto.objProdutoCNAE.iLocIncidImpCliente = MARCADO
    Else
        objProduto.objProdutoCNAE.iLocIncidImpCliente = DESMARCADO
    End If
    
    Set objProduto.objProdutoCNAE.colCidades = New Collection
    For iLinha = 1 To objGridSRV.iLinhasExistentes
    
        Set objCodTrib = New ClassCodTribMun
        
        Set objCidade = New ClassCidades
        objCidade.sDescricao = GridSRV.TextMatrix(iLinha, iGrid_Cidade_Col)
        
        If Len(Trim(objCidade.sDescricao)) = 0 Then gError 208143
        
        lErro = CF("Cidade_Le_Nome", objCidade)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objCodTrib.lCidade = objCidade.iCodigo
        objCodTrib.sProduto = sProduto
        objCodTrib.sCodTribMun = GridSRV.TextMatrix(iLinha, iGrid_CodServ_Col)
        'objCodTrib.dAliquota = StrParaDbl(Val(GridSRV.TextMatrix(iLinha, iGrid_Aliquota_Col)) / 100)
        objCodTrib.dAliquota = PercentParaDbl(GridSRV.TextMatrix(iLinha, iGrid_Aliquota_Col))
        
        If Len(Trim(objCodTrib.sCodTribMun)) = 0 Then gError 208144

        objProduto.objProdutoCNAE.colCidades.Add objCodTrib

    Next
 
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 208143
            Call Rotina_Erro(vbOKOnly, "EERO_CIDADE_NAO_PREENCHIDA_GRID", gErr, iLinha)
        
        Case 208144
            Call Rotina_Erro(vbOKOnly, "EERO_CODTRIBMUN_NAO_PREENCHIDA_GRID", gErr, iLinha)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208145)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

'Extrai os campos da tela que correspondem aos campos no BD
Dim lErro As Long
Dim objProduto As New ClassProduto

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Produtos"

    'Lê os dados da Tela
    lErro = Move_Tela_Memoria(objProduto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objProduto.sCodigo, STRING_PRODUTO, "Codigo"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "Natureza", OP_IGUAL, NATUREZA_PROD_SERVICO

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208146)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim bCancel As Boolean

On Error GoTo Erro_Tela_Preenche

    'Passa o produto da coleção de campos-valores para objprodutofilial.sproduto
    objProduto.sCodigo = colCampoValor.Item("Codigo").vValor

    'Se o produto existir
    If objProduto.sCodigo <> "" Then

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Coloca na tela o Produto selecionado
        Produto.PromptInclude = False
        Produto.Text = sProduto
        Produto.PromptInclude = True

        'Lê Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
    
        If lErro = 28030 Then gError 208147
  
        lErro = Traz_Produto_Tela(objProduto)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr
        
        Case 208147
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208148)

    End Select

    Exit Sub

End Sub

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

    Set objEventoProduto = Nothing
    Set objEventoCidade = Nothing
    Set objEventoCNAE = Nothing
    Set objEventoCodServ = Nothing
    Set objEventoISSQN = Nothing
    Set objEventoNBS = Nothing
    
    Set objGridSRV = Nothing

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)

End Sub

Private Sub Produto_Change()

    iProdutoAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ProdutoLabel_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_ProdutoLabel_Click

    'Verifica se o produto foi preenchido
    If Len(Trim(Produto.ClipText)) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoLista_Consulta", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_ProdutoLabel_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208149)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim bCancel As Boolean

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 208150

    lErro = CF("Traz_Produto_MaskEd", objProduto.sCodigo, Produto, Descricao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Traz_Produto_Tela(objProduto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 208150
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208151)

    End Select

    Exit Sub

End Sub

Private Function Limpa_Tela_Produto() As Long

    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridSRV)
    
    Descricao.Caption = ""
    
    DescCNAE.Caption = ""
    DescNBS.Caption = ""
    DescISSQN.Caption = ""
    DescServico.Caption = ""
    
    optEmpresa.Value = True
    optEmpresaImp.Value = True
   
    iAlterado = 0

End Function

Private Function Traz_Produto_Tela(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objCodTrib As ClassCodTribMun
Dim objCodServ As ClassCodServMun
Dim objCidade As ClassCidades
Dim iLinha As Integer
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Traz_Produto_Tela

    If objProduto.iNatureza <> NATUREZA_PROD_SERVICO Then gError 208152

    Call Grid_Limpa(objGridSRV)
    DescServico.Caption = ""
    
    'Preenche ProdutoDescricao com Descrição do Produto
    Descricao.Caption = objProduto.sDescricao

    'Lê Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM

    ISSQN.PromptInclude = False
    ISSQN.Text = objProduto.sISSQN
    ISSQN.PromptInclude = True
    Call ISSQN_Validate(bSGECancelDummy)

    NBS.PromptInclude = False
    NBS.Text = objProduto.sNBS
    NBS.PromptInclude = True
    Call NBS_Validate(bSGECancelDummy)

    lErro = CF("ProdutoCNAE_Le", objProduto)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    If lErro = SUCESSO Then
    
        CNAE.PromptInclude = False
        CNAE.Text = objProduto.objProdutoCNAE.sCNAE
        CNAE.PromptInclude = True
        Call CNAE_Validate(bSGECancelDummy)
       
        If objProduto.objProdutoCNAE.iLocServCliente = MARCADO Then
            optCliente.Value = True
        Else
            optEmpresa.Value = True
        End If
        
        If objProduto.objProdutoCNAE.iLocIncidImpCliente = MARCADO Then
            optClienteImp.Value = True
        Else
            optEmpresaImp.Value = True
        End If
        
        iLinha = 0
        For Each objCodTrib In objProduto.objProdutoCNAE.colCidades
        
            iLinha = iLinha + 1
            
            Set objCidade = New ClassCidades
            Set objCodServ = New ClassCodServMun
            
            objCidade.iCodigo = objCodTrib.lCidade
            
            lErro = CF("Cidade_Le", objCidade)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
            objCodServ.sCodIBGE = objCidade.sCodIBGE
            objCodServ.sCodServ = objCodTrib.sCodTribMun
            
            lErro = CF("CodServMun_Le", objCodServ)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
            
            GridSRV.TextMatrix(iLinha, iGrid_Cidade_Col) = objCidade.sDescricao
            GridSRV.TextMatrix(iLinha, iGrid_CidadeIBGECod_Col) = objCidade.sCodIBGE
            GridSRV.TextMatrix(iLinha, iGrid_CodServ_Col) = objCodTrib.sCodTribMun
            GridSRV.TextMatrix(iLinha, iGrid_CodServDesc_Col) = objCodServ.sDescricao1 & objCodServ.sDescricao2
            
            If objCodTrib.dAliquota > 0 Then
                GridSRV.TextMatrix(iLinha, iGrid_Aliquota_Col) = Format(objCodTrib.dAliquota, "PERCENT")
            Else
                GridSRV.TextMatrix(iLinha, iGrid_Aliquota_Col) = ""
            End If
    
        Next
        
        objGridSRV.iLinhasExistentes = objProduto.objProdutoCNAE.colCidades.Count
        
        iAlterado = 0
        
    Else
            optEmpresa.Value = True
            optEmpresaImp.Value = True
            
            objFilialEmpresa.iCodFilial = giFilialEmpresa
            
            'Le o nome da filial
            lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
            If lErro <> SUCESSO And lErro <> 27378 Then gError ERRO_SEM_MENSAGEM
        
            CNAE.PromptInclude = False
            CNAE.Text = objFilialEmpresa.sCNAE
            CNAE.PromptInclude = True
            Call CNAE_Validate(bSGECancelDummy)
    
    End If
            
    Traz_Produto_Tela = SUCESSO
            
    Exit Function

Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = gErr

    Select Case gErr
    
        Case 208152
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_SERVICO", gErr, Produto.Text)
        
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208153)
            
    End Select
    
    Exit Function
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ESTOQUE_DADOS_PRINCIPAIS
    Set Form_Load_Ocx = Me
    Caption = "Classificação Nacional de Atividades Econômicas Por Produto"
    Call Form_Load
    
End Function

Public Function Name() As String
    Name = "ProdutoServico"
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
        
    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Produto Then
            Call ProdutoLabel_Click
        ElseIf Me.ActiveControl Is CNAE Then
            Call LabelCNAE_Click
        ElseIf Me.ActiveControl Is ISSQN Then
            Call LabelISSQN_Click
        ElseIf Me.ActiveControl Is CodServ Then
            Call BotaoCodServ_Click
        ElseIf Me.ActiveControl Is Cidade Then
            Call BotaoCidades_Click
        ElseIf Me.ActiveControl Is NBS Then
            Call LabelNBS_Click
        End If
    End If

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

Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Public Property Get Height() As Long
   Height = UserControl.Height
End Property

Public Property Get Width() As Long
   Width = UserControl.Width
End Property

Private Function Inicializa_GridSRV(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Cidade")
    objGrid.colColuna.Add ("IBGE")
    objGrid.colColuna.Add ("Serviço")
    objGrid.colColuna.Add ("Descrição")
    objGrid.colColuna.Add ("Alíquota")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Cidade.Name)
    objGrid.colCampo.Add (CidadeIBGECod.Name)
    objGrid.colCampo.Add (CodServ.Name)
    objGrid.colCampo.Add (CodServDesc.Name)
    objGrid.colCampo.Add (Aliquota.Name)

    'Colunas do Grid
    iGrid_Cidade_Col = 1
    iGrid_CidadeIBGECod_Col = 2
    iGrid_CodServ_Col = 3
    iGrid_CodServDesc_Col = 4
    iGrid_Aliquota_Col = 5

    objGrid.objGrid = GridSRV

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridSRV.ColWidth(0) = 300

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridSRV = SUCESSO

End Function

Private Sub GridSRV_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridSRV, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridSRV, iAlterado)
    End If

End Sub

Private Sub GridSRV_GotFocus()
    
    Call Grid_Recebe_Foco(objGridSRV)

End Sub

Private Sub GridSRV_EnterCell()

    Call Grid_Entrada_Celula(objGridSRV, iAlterado)

End Sub

Private Sub GridSRV_LeaveCell()
    
    Call Saida_Celula(objGridSRV)

End Sub

Private Sub GridSRV_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer


    Call Grid_Trata_Tecla(KeyAscii, objGridSRV, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridSRV, iAlterado)
    End If

End Sub

Private Sub GridSRV_RowColChange()

    Call Grid_RowColChange(objGridSRV)
    
    If GridSRV.Row <> 0 Then DescServico.Caption = GridSRV.TextMatrix(GridSRV.Row, iGrid_CodServDesc_Col)

End Sub

Private Sub GridSRV_Scroll()

    Call Grid_Scroll(objGridSRV)

End Sub

Private Sub GridSRV_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridSRV)
        
End Sub

Private Sub GridSRV_LostFocus()

    Call Grid_Libera_Foco(objGridSRV)

End Sub

Private Sub Cidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cidade_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSRV)
End Sub

Private Sub Cidade_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSRV)
End Sub

Private Sub Cidade_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSRV.objControle = Cidade
    lErro = Grid_Campo_Libera_Foco(objGridSRV)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub CidadeIBGECod_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CidadeIBGECod_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSRV)
End Sub

Private Sub CidadeIBGECod_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSRV)
End Sub

Private Sub CidadeIBGECod_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSRV.objControle = CidadeIBGECod
    lErro = Grid_Campo_Libera_Foco(objGridSRV)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub CodServ_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodServ_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSRV)
End Sub

Private Sub CodServ_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSRV)
End Sub

Private Sub CodServ_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSRV.objControle = CodServ
    lErro = Grid_Campo_Libera_Foco(objGridSRV)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub CodServDesc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodServDesc_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSRV)
End Sub

Private Sub CodServDesc_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSRV)
End Sub

Private Sub CodServDesc_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSRV.objControle = CodServDesc
    lErro = Grid_Campo_Libera_Foco(objGridSRV)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub Aliquota_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Aliquota_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridSRV)
End Sub

Private Sub Aliquota_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridSRV)
End Sub

Private Sub Aliquota_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridSRV.objControle = Aliquota
    lErro = Grid_Campo_Libera_Foco(objGridSRV)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Public Sub LabelISSQN_Click()

Dim objISSQN As New ClassISSQN
Dim colSelecao As Collection

    objISSQN.sCodigo = ISSQN.ClipText

    Call Chama_Tela("ISSQNLista", colSelecao, objISSQN, objEventoISSQN, , "Codigo")

End Sub

Private Sub objEventoISSQN_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objISSQN As ClassISSQN
    
On Error GoTo Erro_objEventoISSQN_evSelecao
    
    Set objISSQN = obj1

    ISSQN.PromptInclude = False
    ISSQN.Text = objISSQN.sCodigo
    ISSQN.PromptInclude = True
    Call ISSQN_Validate(bSGECancelDummy)

    Me.Show
   
    Exit Sub

Erro_objEventoISSQN_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208154)

    End Select

    Exit Sub

End Sub

Public Sub ISSQN_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub ISSQN_GotFocus()
    Call MaskEdBox_TrataGotFocus(ISSQN)
End Sub

Public Sub ISSQN_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objISSQN As New ClassISSQN
    
On Error GoTo Erro_ISSQN_Validate

    If Len(Trim(ISSQN.ClipText)) > 0 Then
    
        objISSQN.sCodigo = ISSQN.ClipText
        
        lErro = CF("ISSQN_Le", objISSQN)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        If lErro <> SUCESSO Then gError 208155
        
        DescISSQN.Caption = objISSQN.sDescricao
        
    Else
    
        DescISSQN.Caption = ""
        
    End If
    
    Exit Sub
    
Erro_ISSQN_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 208155
            Call Rotina_Erro(vbOKOnly, "ERRO_ISSQN_NAO_CADASTRADO", gErr, objISSQN.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208156)
    
    End Select
    
    Exit Sub

End Sub

Public Sub CNAE_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CNAE_GotFocus()
    Call MaskEdBox_TrataGotFocus(CNAE)
End Sub

Public Sub CNAE_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sDescricao As String

On Error GoTo Erro_CNAE_Validate

    If Len(Trim(CNAE.Text)) <> 0 Then
    
        lErro = CF("CNAE_Le", CNAE.Text, sDescricao)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
        If lErro = ERRO_LEITURA_SEM_DADOS Then gError 208157
        
    End If
    
    DescCNAE.Caption = sDescricao
    
    Exit Sub
    
Erro_CNAE_Validate:
    
    Cancel = True
    
    Select Case gErr
    
        Case 208157
            Call Rotina_Erro(vbOKOnly, "ERRO_CNAE_NAO_CADASTRADO", gErr, CNAE.Text)
        
        Case ERRO_SEM_MENSAGEM  'tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208158)
            
    End Select
    
    Exit Sub
    
End Sub

Public Sub LabelCNAE_Click()

Dim objCNAE As New ClassCNAE
Dim colSelecao As Collection

    objCNAE.sCodigo = CNAE.ClipText

    Call Chama_Tela("CNAELista", colSelecao, objCNAE, objEventoCNAE, , "Codigo")

End Sub

Private Sub objEventoCNAE_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCNAE As ClassCNAE
    
On Error GoTo Erro_objEventoCNAE_evSelecao
    
    Set objCNAE = obj1

    CNAE.PromptInclude = False
    CNAE.Text = objCNAE.sCodigo
    CNAE.PromptInclude = True
    Call CNAE_Validate(bSGECancelDummy)

    Me.Show
   
    Exit Sub

Erro_objEventoCNAE_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208159)

    End Select

    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        'Verifica se é o GridItens
        If objGridInt.objGrid.Name = GridSRV.Name Then

            Select Case GridSRV.Col
    
                Case iGrid_Cidade_Col
    
                    lErro = Saida_Celula_Cidade(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                Case iGrid_CodServ_Col
    
                    lErro = Saida_Celula_CodServ(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Case iGrid_Aliquota_Col
    
                    lErro = Saida_Celula_Aliquota(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        
            End Select

        End If
        
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 208160

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 208160
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208161)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iChamada As Integer)

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa o controle da coluna em questão
    Select Case objControl.Name

        Case Cidade.Name
        
            If Len(Trim(GridSRV.TextMatrix(iLinha, iGrid_Cidade_Col))) > 0 Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If

        Case Aliquota.Name, CodServ.Name
        
            If Len(Trim(GridSRV.TextMatrix(iLinha, iGrid_Cidade_Col))) > 0 Then
                objControl.Enabled = True
            Else
                objControl.Enabled = False
            End If
            
        Case Else
            objControl.Enabled = False

        End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208162)

    End Select

    Exit Sub

End Sub

Private Function Saida_Celula_Aliquota(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Aliquota

    Set objGridInt.objControle = Aliquota
                    
    'Se o campo foi preenchido
    If Len(Aliquota.Text) > 0 Then

        'Critica o valor
        lErro = Porcentagem_Critica(Aliquota.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_Aliquota = SUCESSO

    Exit Function

Erro_Saida_Celula_Aliquota:

    Saida_Celula_Aliquota = gErr

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208163)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Cidade(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Cidade

    Set objGridInt.objControle = Cidade
                    
    'Se o campo foi preenchido
    If Len(Cidade.Text) > 0 Then
        
        lErro = Trata_Cidade(Cidade.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_Cidade = SUCESSO

    Exit Function

Erro_Saida_Celula_Cidade:

    Saida_Celula_Cidade = gErr

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208164)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Function Trata_Cidade(ByVal sCidade As String) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objCidade As New ClassCidades
Dim objCodServ As New ClassCodServMun

On Error GoTo Erro_Trata_Cidade

    For iLinha = 1 To objGridSRV.iLinhasExistentes
    
        If iLinha <> GridSRV.Row Then
        
            If UCase(GridSRV.TextMatrix(iLinha, iGrid_Cidade_Col)) = UCase(sCidade) Then gError 208174
        
        End If
    
    Next

    objCidade.sDescricao = sCidade
        
    lErro = CF("Cidade_Le_Nome", objCidade)
    If lErro <> SUCESSO And lErro <> ERRO_OBJETO_NAO_CADASTRADO Then gError ERRO_SEM_MENSAGEM
    
    If lErro = ERRO_OBJETO_NAO_CADASTRADO Then gError 208165
    
    objCodServ.sCodIBGE = objCidade.sCodIBGE
    objCodServ.sISSQNBase = ISSQN.ClipText
    
    lErro = CF("CodServMun_Le_Padrao", objCodServ)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    'GridSRV.TextMatrix(GridSRV.Row, iGrid_Cidade_Col) = objCidade.sDescricao
    GridSRV.TextMatrix(GridSRV.Row, iGrid_CidadeIBGECod_Col) = objCidade.sCodIBGE
    GridSRV.TextMatrix(GridSRV.Row, iGrid_CodServ_Col) = objCodServ.sCodServ
    GridSRV.TextMatrix(GridSRV.Row, iGrid_CodServDesc_Col) = objCodServ.sDescricao1 & objCodServ.sDescricao2
    
    If objCodServ.dAliquota > 0 Then
        GridSRV.TextMatrix(GridSRV.Row, iGrid_Aliquota_Col) = Format(objCodServ.dAliquota, "PERCENT")
    Else
        GridSRV.TextMatrix(GridSRV.Row, iGrid_Aliquota_Col) = ""
    End If

    'verifica se precisa preencher o grid com uma nova linha
    If objGridSRV.objGrid.Row - objGridSRV.objGrid.FixedRows = objGridSRV.iLinhasExistentes Then
        objGridSRV.iLinhasExistentes = objGridSRV.iLinhasExistentes + 1
    End If

    Trata_Cidade = SUCESSO

    Exit Function

Erro_Trata_Cidade:

    Trata_Cidade = gErr

    Select Case gErr
    
        Case 208165
            Call Rotina_Erro(vbOKOnly, "ERRO_CIDADE_NAO_CADASTRADA2", gErr, sCidade)
            
        Case 208174
            Call Rotina_Erro(vbOKOnly, "ERRO_CIDADE_REPETIDA", gErr, iLinha)
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208166)

    End Select

    Exit Function
    
End Function

Private Function Saida_Celula_CodServ(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objCodServ As New ClassCodServMun

On Error GoTo Erro_Saida_Celula_CodServ

    Set objGridInt.objControle = CodServ
                    
    'Se o campo foi preenchido
    If Len(CodServ.Text) > 0 Then
               
        objCodServ.sCodIBGE = GridSRV.TextMatrix(GridSRV.Row, iGrid_CidadeIBGECod_Col)
        objCodServ.sCodServ = CodServ.ClipText
        
        lErro = CF("CodServMun_Le", objCodServ)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        
        If lErro <> ERRO_LEITURA_SEM_DADOS Then
            GridSRV.TextMatrix(GridSRV.Row, iGrid_CodServDesc_Col) = objCodServ.sDescricao1 & objCodServ.sDescricao2
            
            If objCodServ.dAliquota > 0 Then
                GridSRV.TextMatrix(GridSRV.Row, iGrid_Aliquota_Col) = Format(objCodServ.dAliquota, "PERCENT")
'            Else
'                GridSRV.TextMatrix(GridSRV.Row, iGrid_Aliquota_Col) = ""
            End If
        Else
            GridSRV.TextMatrix(GridSRV.Row, iGrid_CodServDesc_Col) = ""
        End If
    
    Else
        GridSRV.TextMatrix(GridSRV.Row, iGrid_CodServDesc_Col) = ""
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Saida_Celula_CodServ = SUCESSO

    Exit Function

Erro_Saida_Celula_CodServ:

    Saida_Celula_CodServ = gErr

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208167)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Private Sub BotaoCidades_Click()

Dim objCidade As New ClassCidades
Dim colSelecao As New Collection
Dim sCidade As String

On Error GoTo Erro_BotaoCidades_Click
    
    'Verifica se tem alguma linha selecionada no Grid
    If GridSRV.Row = 0 Then gError 208168
    
    If Me.ActiveControl Is Cidade Then
        sCidade = Cidade.Text
    Else
        sCidade = GridSRV.TextMatrix(GridSRV.Row, iGrid_Cidade_Col)
    End If
    
    objCidade.sDescricao = sCidade

    'Chama a Tela de browse
    Call Chama_Tela("CidadeLista", colSelecao, objCidade, objEventoCidade)
    
    Exit Sub

Erro_BotaoCidades_Click:

    Select Case gErr

        Case 208168
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208169)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoCodServ_Click()

Dim objCodServMun As New ClassCodServMun
Dim colSelecao As New Collection
Dim sCodServ As String, sFiltro As String

On Error GoTo Erro_BotaoCodServ_Click
    
    'Verifica se tem alguma linha selecionada no Grid
    If GridSRV.Row = 0 Then gError 208170
    
    If Me.ActiveControl Is CodServ Then
        sCodServ = CodServ.Text
    Else
        sCodServ = GridSRV.TextMatrix(GridSRV.Row, iGrid_CodServ_Col)
    End If
       
    objCodServMun.sCodServ = sCodServ
    
    If Len(Trim(ISSQN.ClipText)) > 0 Then
        colSelecao.Add ISSQN.Text
        colSelecao.Add ""
        sFiltro = "(ISSQNBase = ? OR ISSQNBase = ?)"
    End If
    
    colSelecao.Add GridSRV.TextMatrix(GridSRV.Row, iGrid_CidadeIBGECod_Col)

    'Chama a Tela de browse
    Call Chama_Tela("CodServMunLista", colSelecao, objCodServMun, objEventoCodServ, sFiltro)
    
    Exit Sub

Erro_BotaoCodServ_Click:

    Select Case gErr

        Case 208170
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208171)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoCidade_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCidade As ClassCidades

On Error GoTo Erro_objEventoCidade_evSelecao

    Set objCidade = obj1
    
    Cidade.Text = objCidade.sDescricao
    
    If Not (Me.ActiveControl Is Cidade) Then
        GridSRV.TextMatrix(GridSRV.Row, iGrid_Cidade_Col) = objCidade.sDescricao
    End If
        
    lErro = Trata_Cidade(objCidade.sDescricao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iAlterado = REGISTRO_ALTERADO
    
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCidade_evSelecao:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208172)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCodServ_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCodServ As ClassCodServMun

On Error GoTo Erro_objEventoCodServ_evSelecao

    Set objCodServ = obj1
        
    CodServ.PromptInclude = False
    CodServ.Text = objCodServ.sCodServ
    CodServ.PromptInclude = True
        
    GridSRV.TextMatrix(GridSRV.Row, iGrid_CodServ_Col) = objCodServ.sCodServ
    GridSRV.TextMatrix(GridSRV.Row, iGrid_CodServDesc_Col) = objCodServ.sDescricao1 & objCodServ.sDescricao2
        
    If objCodServ.dAliquota > 0 Then
        GridSRV.TextMatrix(GridSRV.Row, iGrid_Aliquota_Col) = Format(objCodServ.dAliquota, "PERCENT")
'    Else
'        GridSRV.TextMatrix(GridSRV.Row, iGrid_Aliquota_Col) = ""
    End If
            
    iAlterado = REGISTRO_ALTERADO
    
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodServ_evSelecao:

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208173)

    End Select

    Exit Sub

End Sub

Public Sub LabelNBS_Click()

Dim objNBS As New ClassNBS
Dim colSelecao As Collection

    objNBS.sCodigo = NBS.ClipText

    Call Chama_Tela("NBSLista", colSelecao, objNBS, objEventoNBS, , "Codigo")

End Sub

Private Sub objEventoNBS_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objNBS As ClassNBS
    
On Error GoTo Erro_objEventoNBS_evSelecao
    
    Set objNBS = obj1

    NBS.PromptInclude = False
    NBS.Text = objNBS.sCodigo
    NBS.PromptInclude = True
    Call NBS_Validate(bSGECancelDummy)

    Me.Show
   
    Exit Sub

Erro_objEventoNBS_evSelecao:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208154)

    End Select

    Exit Sub

End Sub

Public Sub NBS_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NBS_GotFocus()
    Call MaskEdBox_TrataGotFocus(NBS)
End Sub

Public Sub NBS_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objNBS As New ClassNBS
    
On Error GoTo Erro_NBS_Validate

    If Len(Trim(NBS.ClipText)) > 0 Then
    
        objNBS.sCodigo = NBS.ClipText
        
        lErro = CF("NBS_Le", objNBS)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
        If lErro <> SUCESSO Then gError 208155
        
        DescNBS.Caption = objNBS.sDescricao
        
    Else
    
        DescNBS.Caption = ""
        
    End If
    
    Exit Sub
    
Erro_NBS_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case 208155
            Call Rotina_Erro(vbOKOnly, "ERRO_NBS_NAO_CADASTRADO", gErr, objNBS.sCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208156)
    
    End Select
    
    Exit Sub

End Sub
