VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl OCManualArtlux 
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   KeyPreview      =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   6645
   Begin VB.Frame Frame3 
      Caption         =   "Conclusão de Etapas"
      Height          =   600
      Left            =   75
      TabIndex        =   19
      Top             =   3345
      Width           =   6495
      Begin VB.CheckBox CForro 
         Caption         =   "Forro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4020
         TabIndex        =   21
         Top             =   255
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.CheckBox CCorte 
         Caption         =   "Corte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1260
         TabIndex        =   20
         Top             =   240
         Width           =   1590
      End
      Begin VB.CheckBox CMontagem 
         Caption         =   "Montagem"
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
         Height          =   300
         Left            =   4020
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4380
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   60
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "OCManualArtlux.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   615
         Picture         =   "OCManualArtlux.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1590
         Picture         =   "OCManualArtlux.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "OCManualArtlux.ctx":0462
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordem de Corte"
      Height          =   1005
      Left            =   75
      TabIndex        =   9
      Top             =   2190
      Width           =   6495
      Begin MSMask.MaskEdBox Quantidade 
         Height          =   315
         Left            =   1275
         TabIndex        =   10
         Top             =   225
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   15
         PromptChar      =   " "
      End
      Begin VB.Label Status 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1260
         TabIndex        =   24
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
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
         Left            =   585
         TabIndex        =   23
         Top             =   660
         Width           =   615
      End
      Begin VB.Label Data 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4830
         TabIndex        =   13
         Top             =   210
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data:"
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
         Left            =   4290
         TabIndex        =   12
         Top             =   270
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade:"
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
         Left            =   165
         TabIndex        =   11
         Top             =   270
         Width           =   1050
      End
   End
   Begin VB.CommandButton BotaoAutomatico 
      Caption         =   "Ordens de Corte Automáticas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   4005
      TabIndex        =   8
      Top             =   4050
      Width           =   1830
   End
   Begin VB.CommandButton BotaoManual 
      Caption         =   "Ordens de Corte Manuais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   780
      TabIndex        =   7
      Top             =   4065
      Width           =   1830
   End
   Begin VB.Frame Frame2 
      Caption         =   "Produto"
      Height          =   1410
      Left            =   105
      TabIndex        =   1
      Top             =   690
      Width           =   6465
      Begin MSMask.MaskEdBox Produto 
         Height          =   315
         Left            =   1230
         TabIndex        =   0
         Top             =   315
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label TipoCouro 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4770
         TabIndex        =   26
         Top             =   285
         Width           =   1605
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "T.C.:"
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
         Left            =   4275
         TabIndex        =   25
         Top             =   345
         Width           =   435
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
         Left            =   2985
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   780
         Width           =   930
      End
      Begin VB.Label Descricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   600
         Left            =   1230
         TabIndex        =   4
         Top             =   765
         Width           =   5160
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
         TabIndex        =   3
         Top             =   360
         Width           =   660
      End
      Begin VB.Label LblUMEstoque 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3540
         TabIndex        =   2
         Top             =   300
         Width           =   705
      End
   End
End
Attribute VB_Name = "OCManualArtlux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Dim gobjOC As New ClassOCArtlux

Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoOCM As AdmEvento
Attribute objEventoOCM.VB_VarHelpID = -1

Public Function Trata_Parametros(Optional objOC As ClassOCArtlux) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Verifica se foi passado algum Produto
    If Not (objOC Is Nothing) Then
        
        lErro = Traz_OC_Tela(objOC)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206738)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Function Gravar_Registro()

Dim lErro As Long
Dim objOC As New ClassOCArtlux
Dim iMes As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'critica o preenchimento
    If Len(Trim(Produto.ClipText)) = 0 Then gError 206739
    If StrParaDbl(Quantidade.Text) = 0 Then gError 206740
    If gobjOC.lNumIntDoc <> 0 Then gError 206741  'Não pode alterar
'    If gobjOC.iManual = DESMARCADO And gobjOC.lNumIntDoc <> 0 Then gError 206741 'É alteração e não é manual, não pode alterar
'    If gobjOC.colItens.Count <> 0 Then gError 206742 'Já foi iniciado a produção, não pode alterar
        
    'move dados da tela para a memoria
    lErro = Move_Tela_Memoria(objOC)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If gobjOC.sProduto <> objOC.sProduto And gobjOC.lNumIntDoc <> 0 Then gError 206743 'Não se pode trocar o produto de uma ordem de corte
    
    lErro = CF("OrdensDeCorteArtlux_Grava", objOC)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 206739
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
            Produto.SetFocus
        
        Case 206740
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDO1", gErr)
            Produto.SetFocus
        
        Case 206741
            Call Rotina_Erro(vbOKOnly, "ERRO_OC_JA_GRAVADA", gErr)
            'Call Rotina_Erro(vbOKOnly, "ERRO_OC_NAO_MANUAL", gErr)
            
''        Case 206742
''            Call Rotina_Erro(vbOKOnly, "ERRO_OC_EM_PRODUCAO", gErr)
        
        Case 206743
            Call Rotina_Erro(vbOKOnly, "ERRO_OC_TROCA_PRODUTO", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206744)

    End Select

    Exit Function

End Function

Private Sub Limpa_Tela_OC()

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_OC

    Call Limpa_Tela(Me)

    LblUMEstoque.Caption = ""
    Descricao.Caption = ""
    Data.Caption = Format(Date, "dd/mm/yyyy")
    Status.Caption = "Nova"
    
    TipoCouro.Caption = ""
    
    CCorte.Value = vbUnchecked
    CForro.Value = vbUnchecked
    CMontagem.Value = vbUnchecked
    
    Set gobjOC = New ClassOCArtlux

    iAlterado = 0

    Exit Sub

Erro_Limpa_Tela_OC:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206745)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    If gobjOC.lNumIntDoc = 0 Then gError 206773
    'If gobjOC.iManual = DESMARCADO Then gError 206742

    lErro = CF("OrdensDeCorteArtlux_Exclui", gobjOC)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Limpa_Tela_OC

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr
    
        Case 206742
            Call Rotina_Erro(vbOKOnly, "ERRO_OC_NAO_MANUAL", gErr)
    
        Case 206773
            Call Rotina_Erro(vbOKOnly, "ERRO_OC_NAO_GRAVADA", gErr)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206774)

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Limpa_Tela_OC

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206746)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Limpa_Tela_OC
    
    'Fecha comando de setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206747)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()
'
'    'Carrega índices da tela
'    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()
'
'    gi_ST_SetaIgnoraClick = 1

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoProduto = New AdmEvento
    Set objEventoOCM = New AdmEvento

    'inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Limpa_Tela_OC
    
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206748)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoProduto = Nothing
    Set objEventoOCM = Nothing
    
    Set gobjOC = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206749)

    End Select

    Exit Sub

End Sub

Private Sub Quantidade_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Quantidade_GotFocus()
    Call MaskEdBox_TrataGotFocus(Quantidade, iAlterado)
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Quantidade_Validate

    'Veifica se Quantidade está preenchida
    If Len(Trim(Quantidade.Text)) = 0 Then Exit Sub

    'Critica a Quantidade
    lErro = Valor_Positivo_Critica(Quantidade.Text)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_Quantidade_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206750)

    End Select

    Exit Sub

End Sub

Private Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim vbMsgRes As VbMsgBoxResult
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iProdutoPreenchido As Integer
Dim sTipoCouro As String

On Error GoTo Erro_Produto_Validate

    'Verifica preenchimento de Produto
    If Len(Trim(Produto.ClipText)) > 0 Then

        sProduto = Produto.Text

        'Critica o formato do Produto e se existe no BD
        lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 25041 And lErro <> 25043 Then gError ERRO_SEM_MENSAGEM

        'O Produto é Gerencial
        If lErro = 25043 Then gError 206752

        'O Produto não está cadastrado
        If lErro = 25041 Then gError 206753

        If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 206754

        lErro = CF("Produto_TipoCouro_Le", objProduto.sCodigo, sTipoCouro)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        TipoCouro.Caption = sTipoCouro

        Descricao.Caption = objProduto.sDescricao
        LblUMEstoque.Caption = objProduto.sSiglaUMEstoque
        
    Else
    
        Descricao.Caption = ""
        LblUMEstoque.Caption = ""
        TipoCouro.Caption = ""
        
    End If

    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM

        Case 206752
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, sProduto)
        
        Case 206753
            'Não encontrou Produto no BD
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)

            Else
                Descricao.Caption = ""
                LblUMEstoque.Caption = ""
            End If
            
        Case 206754
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206751)

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
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

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

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206755)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProduto As String

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1
    
    lErro = Traz_Produto_Tela(objProduto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    iAlterado = 0

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206756)

    End Select

    Exit Sub

End Sub

Private Function Traz_Produto_Tela(ByVal objProduto As ClassProduto) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim sTipoCouro As String

On Error GoTo Erro_Traz_Produto_Tela

    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then Error ERRO_SEM_MENSAGEM

    lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
    If lErro <> SUCESSO Then gError 206757

    'Coloca o Codigo na tela
    Produto.PromptInclude = False
    Produto.Text = sProdutoEnxuto
    Produto.PromptInclude = True

    'Coloca os demais dados do Produto na tela
    Descricao.Caption = objProduto.sDescricao
    LblUMEstoque.Caption = objProduto.sSiglaUMEstoque
    
    lErro = CF("Produto_TipoCouro_Le", objProduto.sCodigo, sTipoCouro)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    TipoCouro.Caption = sTipoCouro
          
    Traz_Produto_Tela = SUCESSO

    Exit Function

Erro_Traz_Produto_Tela:

    Traz_Produto_Tela = gErr

    Select Case gErr
        
        Case 206757
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206758)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(ByVal objOC As ClassOCArtlux) As Long
'move os dados da tela para as variaveis

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    'Critica o formato do Produto
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    objOC.sProduto = sProdutoFormatado
    objOC.iManual = MARCADO
    objOC.sUsuManual = gsUsuario
    objOC.dtDataManual = Date
    objOC.iFilialEmpresa = giFilialEmpresa
    objOC.dQuantidade = StrParaDbl(Quantidade.Text)
    
    If CCorte.Value = Checked Then
        objOC.sUsuCorte = gsUsuario
        objOC.dHoraIniCorte = CDbl(Time)
        objOC.dtDataIniCorte = Date
        objOC.dtDataFimCorte = Date
        objOC.dHoraFimCorte = CDbl(Time)
    Else
        objOC.dtDataIniCorte = DATA_NULA
        objOC.dtDataFimCorte = DATA_NULA
    End If
    
'    If CForro.Value = Checked Then
'        objOC.sUsuForro = gsUsuario
'        objOC.dHoraIniForro = CDbl(Time)
'        objOC.dtDataIniForro = Date
'        objOC.dtDataFimForro = Date
'        objOC.dHoraFimForro = CDbl(Time)
'    Else
'        objOC.dtDataIniForro = DATA_NULA
'        objOC.dtDataFimForro = DATA_NULA
'    End If

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206759)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD
'
End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'
End Sub

Private Function Traz_OC_Tela(ByVal objOC As ClassOCArtlux) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim sProduto As String
Dim objProduto As New ClassProduto

On Error GoTo Erro_Traz_OC_Tela

    Call Limpa_Tela_OC

    objProduto.sCodigo = objOC.sProduto
    
    lErro = Traz_Produto_Tela(objProduto)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Quantidade.Text = Formata_Estoque(objOC.dQuantidade)

    Data.Caption = Format(objOC.dtDataManual, "dd/mm/yyyy")
    Status.Caption = IIf(objOC.iManual = MARCADO, "Manual", "Automática") & SEPARADOR & "Gravada"
    
    If objOC.dtDataFimCorte <> DATA_NULA Then CCorte.Value = vbChecked
    If objOC.dtDataFimForro <> DATA_NULA Then CForro.Value = vbChecked
    If Abs(objOC.dQuantidade - objOC.dQuantidadeProd) < DELTA_VALORMONETARIO Then CMontagem.Value = vbChecked
    
    Set gobjOC = objOC
    
    Traz_OC_Tela = SUCESSO

    Exit Function

Erro_Traz_OC_Tela:

    Traz_OC_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206760)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CUSTOS
    Set Form_Load_Ocx = Me
    Caption = "Ordem de Corte - Manual"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "OCManualArtlux"
    
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

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
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

Private Sub BotaoManual_Click()

Dim colSelecao As New Collection
Dim objOC As New ClassOCArtlux
Dim lErro As Long
Dim sSelecao  As String

On Error GoTo Erro_BotaoManual_Click

    sSelecao = "Manual = ?"
    colSelecao.Add MARCADO

    'Chama a tela de browse
    Call Chama_Tela("OCArtluxLista", colSelecao, objOC, objEventoOCM, sSelecao)

    Exit Sub
    
Erro_BotaoManual_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206761)

    End Select

    Exit Sub

End Sub

Private Sub BotaoAutomatico_Click()

Dim colSelecao As New Collection
Dim objOC As New ClassOCArtlux
Dim lErro As Long
Dim sSelecao  As String

On Error GoTo Erro_BotaoAutomatico_Click

    sSelecao = "Manual <> ?"
    colSelecao.Add MARCADO

    'Chama a tela de browse
    Call Chama_Tela("OCArtluxLista", colSelecao, objOC, objEventoOCM, sSelecao)

    Exit Sub
    
Erro_BotaoAutomatico_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206762)

    End Select

    Exit Sub

End Sub

Private Sub objEventoOCM_evSelecao(obj1 As Object)

Dim lErro As Long

On Error GoTo Erro_objEventoOCM_evSelecao
   
    lErro = Traz_OC_Tela(obj1)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    iAlterado = 0

    Exit Sub

Erro_objEventoOCM_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206763)

    End Select

    Exit Sub

End Sub
