VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl OrdemProducaoBloqueio 
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8220
   KeyPreview      =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   8220
   Begin VB.TextBox Justificativa 
      Height          =   285
      Left            =   1365
      TabIndex        =   3
      Top             =   3990
      Width           =   6660
   End
   Begin VB.CommandButton BotaoDesmarcarTodos 
      Caption         =   "Desmarcar Todos"
      Height          =   675
      Index           =   0
      Left            =   2115
      Picture         =   "OrdemProducaoBloqueio.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4620
      Width           =   1830
   End
   Begin VB.CommandButton BotaoMarcarTodos 
      Caption         =   "Marcar Todas"
      Height          =   675
      Left            =   240
      Picture         =   "OrdemProducaoBloqueio.ctx":11E2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4620
      Width           =   1830
   End
   Begin VB.ComboBox Situacao 
      Height          =   315
      Left            =   5055
      TabIndex        =   6
      Top             =   4650
      Width           =   1935
   End
   Begin VB.CommandButton botaoAlterar 
      Caption         =   "Alterar Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7155
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4620
      Width           =   855
   End
   Begin VB.CommandButton BotaoTrazerOP 
      Height          =   315
      Left            =   3180
      Picture         =   "OrdemProducaoBloqueio.ctx":21FC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Trazer Dados"
      Top             =   315
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6345
      ScaleHeight     =   495
      ScaleWidth      =   1590
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   150
      Width           =   1650
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1050
         Picture         =   "OrdemProducaoBloqueio.ctx":26EE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   570
         Picture         =   "OrdemProducaoBloqueio.ctx":286C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "OrdemProducaoBloqueio.ctx":2D9E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox CodOP 
      Height          =   315
      Left            =   2085
      TabIndex        =   0
      Top             =   315
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   9
      PromptChar      =   " "
   End
   Begin MSComctlLib.TreeView OP 
      Height          =   3060
      Left            =   255
      TabIndex        =   2
      Top             =   750
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   5398
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5205
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OrdemProducaoBloqueio.ctx":2EF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OrdemProducaoBloqueio.ctx":32B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OrdemProducaoBloqueio.ctx":3670
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OrdemProducaoBloqueio.ctx":3A2C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label JustificativaLabel 
      Caption         =   "Justificativa:"
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
      Left            =   240
      TabIndex        =   14
      Top             =   4005
      Width           =   1275
   End
   Begin VB.Label SituacaoLabel 
      Caption         =   "Situação:"
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
      Left            =   4185
      TabIndex        =   13
      Top             =   4680
      Width           =   795
   End
   Begin VB.Label CodigoOPLabel 
      AutoSize        =   -1  'True
      Caption         =   "Ordem de Produção:"
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
      TabIndex        =   12
      Top             =   375
      Width           =   1755
   End
End
Attribute VB_Name = "OrdemProducaoBloqueio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim colComponentes As New Collection 'cada elemento é objItemOP e guarda informacoes correspondentes aos dados de cada nó da treeview EstruturaProduto

Dim iAlterado As Integer
Dim iCodigoAlterado As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1

Private Sub BotaoAlterar_Click()

Dim iIndice As Integer
Dim lErro As Long
Dim objNode As Node

On Error GoTo Erro_BotaoAlterar_Click

    If Len(Trim(Situacao.Text)) = 0 Then Exit Sub

    For iIndice = 1 To OP.Nodes.Count
    
        If OP.Nodes.Item(iIndice).Checked Then
            OP.Nodes.Item(iIndice).Image = Converte_Situacao(Situacao.Text) + 1

            colComponentes.Item(OP.Nodes.Item(iIndice).Tag).iSituacao = Converte_Situacao(Situacao.Text)
            
            iAlterado = REGISTRO_ALTERADO
        End If
    
    Next
    
    Exit Sub

Erro_BotaoAlterar_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163923)

    End Select

    Exit Sub

End Sub

Private Sub BotaoDesmarcarTodos_Click(Index As Integer)

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoDesmarcarTodos_Click

    For iIndice = 1 To OP.Nodes.Count
            
        OP.Nodes.Item(iIndice).Checked = False
    
    Next
    
    Exit Sub

Erro_BotaoDesmarcarTodos_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163924)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()
'implementa gravação de uma nova ou atualizacao de uma OP

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    'Rotina de gravação da OP
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 129227

    'limpa a tela
    lErro = Limpa_Tela_OrdemDeProducao
    If lErro <> SUCESSO Then gError 129228
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 129227, 129228

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163925)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'testa se houva alguma alteração
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 129224

    'limpa a tela
    lErro = Limpa_Tela_OrdemDeProducao
    If lErro <> SUCESSO Then gError 129225
    
    iAlterado = 0

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 129224, 129225

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163926)

    End Select

    Exit Sub

End Sub

Function Limpa_Arvore_OP() As Long
'Limpa a Arvore do OP

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Arvore_OP

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    If lErro <> SUCESSO Then gError 129223

    OP.Nodes.Clear
    Set colComponentes = Nothing

    Limpa_Arvore_OP = SUCESSO

    Exit Function

Erro_Limpa_Arvore_OP:

    Limpa_Arvore_OP = gErr
    
    Select Case gErr

        Case 129223

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163927)

    End Select

    Exit Function

End Function

Private Sub BotaoMarcarTodos_Click()

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_BotaoMarcarTodos_Click
    
    For iIndice = 1 To OP.Nodes.Count
            
        OP.Nodes.Item(iIndice).Checked = True
    
    Next
    
    Exit Sub

Erro_BotaoMarcarTodos_Click:

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163928)

    End Select

    Exit Sub

End Sub

Private Sub BotaoTrazerOP_Click()

Dim lErro As Long
Dim objOrdemDeProducao As New ClassOrdemDeProducao

On Error GoTo Erro_BotaoTrazerOP_Click

    If Len(Trim(CodOP.Text)) > 0 Then

        objOrdemDeProducao.sCodigo = CodOP.Text
        objOrdemDeProducao.iFilialEmpresa = giFilialEmpresa

        'tenta ler a OP desejada
        lErro = CF("OrdemDeProducao_Le_ComItens", objOrdemDeProducao)
        If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 129229

        If lErro = SUCESSO And objOrdemDeProducao.iTipo = OP_TIPO_OC Then gError 129230

        'ordem de producao baixada
        If lErro = 55316 Then gError 129231

        'se não existir
        If lErro <> SUCESSO Then gError 129232

        'traz a OP para a tela
        lErro = Traz_Tela_OrdemDeProducao(objOrdemDeProducao)
        If lErro <> SUCESSO And lErro <> 21966 Then gError 129233
        
        iAlterado = 0

        Call ComandoSeta_Fechar(Me.Name)

    End If
    
    Exit Sub

Erro_BotaoTrazerOP_Click:

    Select Case gErr
    
        Case 129229 To 129233

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163929)

    End Select

    Exit Sub
    
End Sub

Private Sub CodigoOPLabel_Click()

Dim objOrdemDeProducao As New ClassOrdemDeProducao
Dim colSelecao As New Collection
Dim sSelecao As String

    'preenche o objOrdemDeProducao com o código da tela , se estiver preenchido
    If Len(Trim(CodOP.Text)) <> 0 Then objOrdemDeProducao.sCodigo = CodOP.Text
    
    sSelecao = "Tipo = 0"
    
    'lista as OP's
    Call Chama_Tela("OrdemProducaoLista", colSelecao, objOrdemDeProducao, objEventoCodigo, sSelecao)

End Sub

Private Function Converte_Situacao(sSituacao As String) As Long

    Select Case sSituacao
    
        Case STRING_NORMAL
            Converte_Situacao = ITEMOP_SITUACAO_NORMAL
            
        Case STRING_DESABILITADA
            Converte_Situacao = ITEMOP_SITUACAO_DESAB
            
        Case STRING_SACRAMENTADA
            Converte_Situacao = ITEMOP_SITUACAO_SACR
            
        Case STRING_BAIXADA
            Converte_Situacao = ITEMOP_SITUACAO_BAIXADA
    
    End Select


End Function

Public Sub CargaCombo_Situacao(objSituacao As Object)
'Carga dos itens da combo Situação

    objSituacao.AddItem STRING_NORMAL
    objSituacao.ItemData(objSituacao.NewIndex) = ITEMOP_SITUACAO_NORMAL
    objSituacao.AddItem STRING_DESABILITADA
    objSituacao.ItemData(objSituacao.NewIndex) = ITEMOP_SITUACAO_DESAB
    objSituacao.AddItem STRING_SACRAMENTADA
    objSituacao.ItemData(objSituacao.NewIndex) = ITEMOP_SITUACAO_SACR
    objSituacao.AddItem STRING_BAIXADA
    objSituacao.ItemData(objSituacao.NewIndex) = ITEMOP_SITUACAO_BAIXADA

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
     
    'Carrega Ítens da Combo
    Call CargaCombo_Situacao(Situacao)
        
    iAlterado = 0
    iCodigoAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163930)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_UnLoad(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoCodigo = Nothing

   'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
   
    If lErro <> SUCESSO Then gError 129234

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case 129234

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163931)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objOrdemDeProducao As New ClassOrdemDeProducao
Dim objItemOP As ClassItemOP
Dim objOP As New ClassOrdemDeProducao

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(CodOP.Text)) = 0 Then gError 129235
    
    objOP.sCodigo = CodOP.Text
    objOP.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("OrdemDeProducao_Grava_Arvore", colComponentes, objOP)
    If lErro <> SUCESSO Then gError 129236
    
    iAlterado = 0
    iCodigoAlterado = 0

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr
    
        Case 129235 To 129236
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163932)

    End Select

    Exit Function

End Function

Function Move_Arvore_Memoria(objOrdemDeProducao As ClassOrdemDeProducao) As Long
'move itens do Grid para objOrdemDeProducao

Dim lErro As Long

On Error GoTo Erro_Move_Arvore_Memoria

    'NÃO FAZ NADA
    '=====> OS objItemOPs ESTÃO NA MEMÓRIA ===> colComponentes

    Move_Arvore_Memoria = SUCESSO

    Exit Function

Erro_Move_Arvore_Memoria:

    Move_Arvore_Memoria = gErr

    Select Case gErr
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163933)

    End Select

    Exit Function

End Function

Function Limpa_Tela_OrdemDeProducao(Optional iFechaSetas As Integer = FECHAR_SETAS) As Long
'Limpa a Tela

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Limpa_Tela_OrdemDeProducao

    If iFechaSetas = FECHAR_SETAS Then
        'Fecha o comando das setas se estiver aberto
        lErro = ComandoSeta_Fechar(Me.Name)
        If lErro <> SUCESSO Then gError 129218
    End If
    
    Call Limpa_Arvore_OP
    
    Call Limpa_Tela(Me)
    
    iAlterado = 0
    iCodigoAlterado = 0

    Limpa_Tela_OrdemDeProducao = SUCESSO

    Exit Function

Erro_Limpa_Tela_OrdemDeProducao:

    Select Case gErr
    
        Case 129218

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163934)

    End Select

    Exit Function

End Function

Function Trata_Parametros(Optional objOrdemDeProducao As ClassOrdemDeProducao) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Trata_Parametros

    If Not (objOrdemDeProducao Is Nothing) Then

        'traz OP para a tela
        lErro = Traz_Tela_OrdemDeProducao(objOrdemDeProducao)
        If lErro <> SUCESSO And lErro <> 21966 Then gError 129219

        If lErro = 21966 Then

            'Se não existe exibe apenas o código
            CodOP.Text = objOrdemDeProducao.sCodigo

        End If

        Call ComandoSeta_Fechar(Me.Name)
                
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 129219

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163935)

    End Select

    iAlterado = 0

    Exit Function

End Function


'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objOrdemDeProducao As New ClassOrdemDeProducao

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "OrdemProducaoOP"

    objOrdemDeProducao.sCodigo = CodOP.Text
    
    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objOrdemDeProducao.sCodigo, STRING_ORDEM_DE_PRODUCAO, "Codigo"
    
    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163936)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objOrdemDeProducao As ClassOrdemDeProducao) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria

    objOrdemDeProducao.sCodigo = CodOP.Text
    objOrdemDeProducao.iFilialEmpresa = giFilialEmpresa
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163937)

    End Select

    Exit Function

End Function

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objOrdemDeProducao As New ClassOrdemDeProducao

On Error GoTo Erro_Tela_Preenche

    objOrdemDeProducao.sCodigo = colCampoValor.Item("Codigo").vValor
    objOrdemDeProducao.iFilialEmpresa = giFilialEmpresa
    
    'Traz dados da Ordem de Produção para a Tela
    lErro = Traz_Tela_OrdemDeProducao(objOrdemDeProducao)
    If lErro <> SUCESSO And lErro <> 21966 Then gError 129220

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 129220

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163938)

    End Select

    Exit Sub

End Sub

Function Traz_Tela_OrdemDeProducao(objOrdemDeProducao As ClassOrdemDeProducao) As Long
'preenche a tela com os dados da OP

Dim lErro As Long
Dim colOP As New Collection

On Error GoTo Erro_Traz_Tela_OrdemDeProducao

    lErro = Limpa_Tela_OrdemDeProducao(NAO_FECHAR_SETAS)
    If lErro <> SUCESSO Then gError 129238

    CodOP.Text = objOrdemDeProducao.sCodigo
    
    iAlterado = 0
    iCodigoAlterado = 0
    
    'Lê as Ordens de Produção Filhas
    lErro = CF("OrdemProducao_Le_Filhos", objOrdemDeProducao, colOP)
    If lErro <> SUCESSO Then gError 129239
    
    '##############################################
    '++++++++++++++++> TESTE
    'Call GeraColOPTeste(objOrdemDeProducao, colOP)
    '###################################################
    
    'Povoa a árvore
    lErro = CarregaArvore(objOrdemDeProducao, colOP, 0, 0, 0)
    If lErro <> SUCESSO Then gError 129240

    Traz_Tela_OrdemDeProducao = SUCESSO

    Exit Function

Erro_Traz_Tela_OrdemDeProducao:

    Traz_Tela_OrdemDeProducao = gErr

    Select Case gErr
    
        Case 129238 To 129240

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163939)

    End Select

    Exit Function

End Function

Private Sub CodOP_Change()

    iCodigoAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub Justificativa_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Justificativa_Validate(Cancel As Boolean)

    If OP.SelectedItem Is Nothing Then Exit Sub

    colComponentes.Item(OP.SelectedItem.Tag).sJustificativaBloqueio = Justificativa.Text

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOrdemDeProducao As ClassOrdemDeProducao

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objOrdemDeProducao = obj1

    iCodigoAlterado = 1

    'traz OP para a tela
    lErro = Traz_Tela_OrdemDeProducao(objOrdemDeProducao)
    If lErro <> SUCESSO Then gError 129221

    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Call BotaoTrazerOP_Click
    
    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 129221

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 163940)
    
    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ORDEM_PRODUCAO
    Set Form_Load_Ocx = Me
    Caption = "Bloqueio de Ordem de Produção"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "OrdemProducaoBloqueio"
    
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

Private Sub OP_NodeCheck(ByVal Node As MSComctlLib.Node)
'Marca e desmarca descendentes (Recursivo)

Dim iIndice As Integer

    For iIndice = 1 To OP.Nodes.Count
            
        If Not (OP.Nodes.Item(iIndice).Parent Is Nothing) Then
            If OP.Nodes.Item(iIndice).Parent = Node.Text Then
                OP.Nodes.Item(iIndice).Checked = Node.Checked
                
                Call OP_NodeCheck(OP.Nodes.Item(iIndice))
            End If
        End If
    
    Next

End Sub

Private Sub OP_NodeClick(ByVal Node As MSComctlLib.Node)

    Justificativa.Text = colComponentes.Item(Node.Tag).sJustificativaBloqueio

End Sub

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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is CodOP Then
            Call CodigoOPLabel_Click
        End If
    End If

End Sub

Private Sub CodigoOPLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodigoOPLabel, Source, X, Y)
End Sub

Private Sub CodigoOPLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(CodigoOPLabel, Button, Shift, X, Y)
End Sub

Private Sub CodOP_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(CodOP, Source, X, Y)
End Sub

Private Sub Justificativa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Justificativa, Source, X, Y)
End Sub

Private Sub Justificativa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Justificativa, Button, Shift, X, Y)
End Sub

Private Sub JustificativaLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(JustificativaLabel, Source, X, Y)
End Sub

Private Sub JustificativaLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(JustificativaLabel, Button, Shift, X, Y)
End Sub

Private Sub Situacao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Situacao, Source, X, Y)
End Sub

Private Sub SituacaoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(SituacaoLabel, Source, X, Y)
End Sub

Private Sub SituacaoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(SituacaoLabel, Button, Shift, X, Y)
End Sub

Private Sub CodOp_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOrdemDeProducao As New ClassOrdemDeProducao

On Error GoTo Erro_Codigo_Validate

    'Se houve alteração nos dados da tela
    If (iCodigoAlterado = REGISTRO_ALTERADO) Then

        If Len(Trim(CodOP.Text)) > 0 Then

            objOrdemDeProducao.sCodigo = CodOP.Text
            objOrdemDeProducao.iFilialEmpresa = giFilialEmpresa

            'tenta ler a OP desejada
            lErro = CF("OrdemDeProducao_Le_ComItens", objOrdemDeProducao)
            If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 129253

            If lErro = SUCESSO And objOrdemDeProducao.iTipo = OP_TIPO_OC Then gError 129254

            'ordem de producao baixada
            If lErro = 55316 Then gError 129255

            'se não existir
            If lErro <> SUCESSO Then gError 129256

            Call ComandoSeta_Fechar(Me.Name)

        End If

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True
    
    Select Case gErr

        Case 129253
    
        Case 129254
            Call Rotina_Erro(vbOKOnly, "ERRO_ORDEMDECORTE", gErr, CodOP.Text)

        Case 129255
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ORDEMDEPRODUCAO_BAIXADA", gErr, CodOP.Text)

        Case 129256
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 163941)

    End Select

    Exit Sub

End Sub

Function CarregaArvore(ByVal objOrdemProducao As ClassOrdemDeProducao, colOP As Collection, ByVal iNivel As Integer, iSeq As Integer, ByVal lNumIntDocPai As Long) As Long
'Monta Árvore (Função Recursiva)

Dim objNode As Node
Dim lErro As Long
Dim objOP As ClassOrdemDeProducao
Dim objItemOP As ClassItemOP
Dim objItemOPFilha As ClassItemOP
Dim iSeqFilho As Integer
   
On Error GoTo Erro_CarregaArvore

    'Para cada Item nó Pai Insere
    For Each objItemOP In objOrdemProducao.colItens
            
        If iNivel = 0 Then
        
            Set objNode = OP.Nodes.Add(, tvwFirst, "X" & CStr(objItemOP.lNumIntDoc), objItemOP.sCodigo & SEPARADOR & objItemOP.sDescricao, objItemOP.iSituacao + 1)

            OP.Nodes.Item(objNode.Index).Expanded = True
            colComponentes.Add objItemOP, "X" & objItemOP.lNumIntDoc
            objNode.Tag = "X" & objItemOP.lNumIntDoc

            iSeqFilho = objNode.Index

        Else
            
            If lNumIntDocPai = objItemOP.lNumIntDocPai Then

                Set objNode = OP.Nodes.Add(iSeq, tvwChild, "X" & objItemOP.lNumIntDoc, objOrdemProducao.sCodigo & SEPARADOR & objItemOP.sDescricao, objItemOP.iSituacao + 1)
                colComponentes.Add objItemOP, "X" & objItemOP.lNumIntDoc
                objNode.Tag = "X" & objItemOP.lNumIntDoc

                iSeqFilho = objNode.Index
            
            End If

        End If
        
        
        For Each objOP In colOP
    
            If objOP.sOPGeradora = objOrdemProducao.sCodigo And objOP.iFilialEmpresa = objOrdemProducao.iFilialEmpresa Then
    
                lErro = CarregaArvore(objOP, colOP, iNivel + 1, iSeqFilho, objItemOP.lNumIntDoc)
                If lErro <> SUCESSO Then gError 129257
                
            End If
            
        Next
    
    Next

    CarregaArvore = SUCESSO
    
    Exit Function

Erro_CarregaArvore:

    CarregaArvore = gErr

    Select Case gErr
    
        Case 129257

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 163942)

    End Select


    Exit Function
    
End Function


'####################################################################
'LIXO
'APENAS PARA TESTES
'SCRIPT PARA POVOAR A COLEÇÃO ColOP

'Sub GeraColOPTeste(objOrdemDeProducao As ClassOrdemDeProducao, colOP As Collection)
'
'Dim objItemOP As New ClassItemOP
'Dim objItemOP1 As New ClassItemOP
'Dim objItemOP2 As New ClassItemOP
'Dim objItemOP3 As New ClassItemOP
'Dim objOP1 As New ClassOrdemDeProducao
'Dim objOP2 As New ClassOrdemDeProducao
'Dim objOP3 As New ClassOrdemDeProducao
'Dim sProdutoPai As String
'Dim iIndice As Integer
'
'    'FILHA 01
'    'OP
'    objOP1.sOPGeradora = objOrdemDeProducao.sCodigo
'    objOP1.sCodigo = "OPFILHA01"
'    objOP1.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
'    objOP1.iStatusOP = objOrdemDeProducao.iStatusOP
'    'ITEM OP
'    objItemOP1.sDescricao = "FILHO 1"
'    objItemOP1.lNumIntDoc = 1001
'    objItemOP1.iSituacao = 1
'    objItemOP1.sCodigo = "X01"
'    For Each objItemOP In objOrdemDeProducao.colItens
'        objItemOP1.lNumIntDocPai = objItemOP.lNumIntDoc
'        Exit For
'    Next
'
'    objOP1.colItens.Add objItemOP1
'
'    'FILHA 02
'    'OP
'    objOP2.sOPGeradora = objOrdemDeProducao.sCodigo
'    objOP2.sCodigo = "OPFILHA02"
'    objOP2.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
'    objOP2.iStatusOP = objOrdemDeProducao.iStatusOP
'    'ITEM OP
'    objItemOP2.sDescricao = "FILHO 2"
'    objItemOP2.lNumIntDoc = 1002
'    objItemOP2.iSituacao = 2
'    objItemOP2.sCodigo = "X02"
'    For Each objItemOP In objOrdemDeProducao.colItens
'        objItemOP2.lNumIntDocPai = objItemOP.lNumIntDoc
'        Exit For
'    Next
'
'    objOP2.colItens.Add objItemOP2
'
'    'NETA 01
'    'OP
'    objOP3.sOPGeradora = objOP1.sCodigo
'    objOP3.sCodigo = "OPNETA01"
'    objOP3.iFilialEmpresa = objOrdemDeProducao.iFilialEmpresa
'    objOP3.iStatusOP = objOrdemDeProducao.iStatusOP
'    'ITEM OP
'    objItemOP3.sDescricao = "NETO 1"
'    objItemOP3.lNumIntDoc = 2001
'    objItemOP3.iSituacao = 3
'    objItemOP3.sCodigo = "X03"
'    objItemOP3.lNumIntDocPai = objItemOP1.lNumIntDoc
'
'    objOP3.colItens.Add objItemOP3
'
'    colOP.Add objOP1
'    colOP.Add objOP2
'    colOP.Add objOP3
'
'End Sub
