VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl CustoTabelaPrecoOcx 
   ClientHeight    =   6015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6675
   ScaleHeight     =   6015
   ScaleWidth      =   6675
   Begin VB.Frame FrameTabela 
      Caption         =   "Tabela"
      Height          =   855
      Left            =   120
      TabIndex        =   26
      Top             =   720
      Width           =   6375
      Begin VB.ComboBox TabelaPreco 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   375
         Width           =   735
      End
      Begin VB.Label DescricaoTabela 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1800
         TabIndex        =   29
         Top             =   375
         Width           =   2355
      End
      Begin VB.Label LabelTabela 
         AutoSize        =   -1  'True
         Caption         =   "Tabela:"
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
         Left            =   240
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   28
         Top             =   435
         Width           =   660
      End
   End
   Begin VB.Frame FramePrecos 
      Caption         =   "Preços"
      Height          =   1575
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   6375
      Begin MSMask.MaskEdBox DataReferencia 
         Height          =   315
         Left            =   4440
         TabIndex        =   15
         Top             =   990
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label LabelFilial 
         AutoSize        =   -1  'True
         Caption         =   "Filial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   480
         TabIndex        =   22
         Top             =   420
         Width           =   465
      End
      Begin VB.Label LabelvalorCalculado 
         AutoSize        =   -1  'True
         Caption         =   "Calculado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   1050
         Width           =   915
      End
      Begin VB.Label ValorFilial 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1200
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label ValorCalculado 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1200
         TabIndex        =   19
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label LabelDataReferencia 
         AutoSize        =   -1  'True
         Caption         =   "Data do Cálculo:"
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
         Index           =   1
         Left            =   2880
         TabIndex        =   18
         Top             =   1050
         Width           =   1440
      End
      Begin VB.Label LabelEmpresa 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3480
         TabIndex        =   17
         Top             =   420
         Width           =   795
      End
      Begin VB.Label ValorEmpresa 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4440
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame FrameProduto 
      Caption         =   "Produto"
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   6375
      Begin MSMask.MaskEdBox Produto 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label ProdutoLabel 
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
         Height          =   165
         Left            =   435
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   13
         Top             =   435
         Width           =   735
      End
      Begin VB.Label DescricaoProduto 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   1035
         Width           =   4515
      End
      Begin VB.Label Label9 
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
         Height          =   165
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label UnidadeMedida 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4275
         TabIndex        =   10
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Unidade:"
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
         Left            =   3360
         TabIndex        =   9
         Top             =   420
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Novo Preço"
      Height          =   780
      Left            =   120
      TabIndex        =   5
      Top             =   5040
      Width           =   6360
      Begin MSMask.MaskEdBox NovoPreco 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   300
         Width           =   1215
         _ExtentX        =   2143
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
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataVigencia 
         Height          =   300
         Left            =   4440
         TabIndex        =   23
         Top             =   300
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataVigencia 
         Height          =   300
         Left            =   5640
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data de Vigência:"
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
         Left            =   2880
         TabIndex        =   25
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label LabelNovoPreco 
         AutoSize        =   -1  'True
         Caption         =   "Novo Preço:"
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
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1080
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4845
      ScaleHeight     =   495
      ScaleWidth      =   1590
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1650
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1080
         Picture         =   "CustoTabelaPrecoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   600
         Picture         =   "CustoTabelaPrecoOcx.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "CustoTabelaPrecoOcx.ctx":06B0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
End
Attribute VB_Name = "CustoTabelaPrecoOcx"
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
Dim iDataReferenciaAlterada As Integer '???
Dim bCarregandoTela As Boolean

Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1
Private WithEvents objEventoPrecoCalculado As AdmEvento
Attribute objEventoPrecoCalculado.VB_VarHelpID = -1

Private Sub BotaoGravar_Click()
'Aciona Rotinas de gravação

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 116828

    Call Limpa_Tela_TabelaPrecoItem_Grava

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 116828
            'Erro tratado na rotina chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158737)

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

    Set objEventoPrecoCalculado = Nothing
    Set objEventoProduto = Nothing
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Private Function Carrega_Preco_Calculado() As Long
'Funçao que carrega os valores da tabela preço calculado

Dim lErro As Long
Dim objPrecoCalculado As New ClassPrecoCalculado
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Carrega_Preco_Calculado
    
    'Preenche objPrecoCalculado com o código do produto
    If Len(Trim(Produto.ClipText)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 116856
    
        objPrecoCalculado.sCodProduto = sProdutoFormatado
        
    End If
    
    'Preenche objPrecoCalculado com o codigo da tabela
    objPrecoCalculado.iCodTabela = StrParaInt(TabelaPreco.Text)
      
    objPrecoCalculado.dtDataReferencia = StrParaDate(DataReferencia.Text)
    objPrecoCalculado.iFilialEmpresa = giFilialEmpresa
    
    'Lê as Data de Vigência para o par (Produto e Tabela)
    lErro = CF("PrecoCalculado_Le", objPrecoCalculado)
    If lErro <> SUCESSO And lErro <> 116826 Then gError 116833

    'Se nao encontrou ---> sai
    If lErro = 116826 Then gError 116834
   
    Call DateParaMasked(DataReferencia, objPrecoCalculado.dtDataReferencia)
    ValorCalculado.Caption = Format(objPrecoCalculado.dPrecoCalculado, "Fixed")

    If bCarregandoTela = False Then
    
        If StrParaDbl(ValorCalculado.Caption) > StrParaDbl(ValorFilial.Caption) Then
        
            NovoPreco = Format(objPrecoCalculado.dPrecoCalculado, "Fixed")
    
        Else
        
            NovoPreco.PromptInclude = False
            NovoPreco.Text = ""
            NovoPreco.PromptInclude = True
        
        End If
    
    Else
    
        NovoPreco = Format(objPrecoCalculado.dPrecoCalculado, "Fixed")
        Call DateParaMasked(DataVigencia, objPrecoCalculado.dtDataVigencia)
    
    End If
    
    Carrega_Preco_Calculado = SUCESSO
    
    Exit Function
    
Erro_Carrega_Preco_Calculado:
    
    Carrega_Preco_Calculado = gErr
    
    Select Case gErr
        
        Case 116833, 116856
        
        Case 116834
            ValorCalculado.Caption = ""
            DataReferencia.PromptInclude = False
            DataReferencia.Text = ""
            DataReferencia.PromptInclude = True
            NovoPreco.PromptInclude = False
            NovoPreco.Text = ""
            NovoPreco.PromptInclude = True
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158738)

    End Select

    Exit Function

End Function

Private Sub Tabela_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoPrecoCalculado_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objPrecoCalculado As ClassPrecoCalculado

On Error GoTo Erro_objEventoItemTabela_evSelecao

    Set objPrecoCalculado = obj1

    'Traz item de Tabela de Preço para a tela
    lErro = Traz_PrecoCalculado_Tela(objPrecoCalculado)
    If lErro <> SUCESSO Then gError 116902
        
    Me.Show

    Exit Sub

Erro_objEventoItemTabela_evSelecao:

    Select Case gErr

        Case 116902
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 158739)

    End Select

    Exit Sub

End Sub

Private Sub LabelTabela_Click()

Dim lErro As Long
Dim objPrecoCalculado As New ClassPrecoCalculado
Dim colSelecao As New Collection
Dim sProduto As String
Dim iPreenchido As Integer
Dim sSelecaoSQL As String

On Error GoTo Erro_LabelTabela_Click
        
    If Trim(Produto.ClipText) <> "" Then
        
        'Passa para o formato do BD
        lErro = CF("Produto_Formata", Produto.Text, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 116835

        'Testa se o codigo está preenchido
        If iPreenchido = PRODUTO_PREENCHIDO Then
            objPrecoCalculado.sCodProduto = sProduto
        Else
            objPrecoCalculado.sCodProduto = ""
        End If
        
    Else
        objPrecoCalculado.sCodProduto = ""
    End If
    
    If Len(Trim(TabelaPreco.Text)) > 0 Then
    
        colSelecao.Add StrParaInt(TabelaPreco.Text)
        
      sSelecaoSQL = "CodTabela=?"
            
    End If
    
    'Abre a Tela que Lista as Tabelas de Preço
    Call Chama_Tela("PrecoCalculadoLista", colSelecao, objPrecoCalculado, objEventoPrecoCalculado, sSelecaoSQL)

    Exit Sub

Erro_LabelTabela_Click:

    Select Case gErr

        Case 116835

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 158740)

    End Select

    Exit Sub

End Sub

Sub Limpa_Tela_TabelaPrecoItem_Grava()
'Limpa a Tela TabelaPrecoItem quando esta na gravação

Dim lErro As Long

    'Limpa campos tipo Label
    DescricaoProduto.Caption = ""
    UnidadeMedida.Caption = ""
    ValorEmpresa.Caption = ""
    ValorFilial.Caption = ""
    ValorCalculado.Caption = ""

'    TabelaDefault.Value = vbUnchecked
    
    'Funcao generica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

End Sub

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
'Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim objPrecoCalculado As New ClassPrecoCalculado
Dim colCodFilial As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "PrecoCalculado"

    'Lê os dados da Tela Tabela Preço Item
    lErro = Move_Tela_Memoria2(objPrecoCalculado)
    If lErro <> SUCESSO Then gError 116836

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "FilialEmpresa", giFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "CodTabela", objPrecoCalculado.iCodTabela, 0, "CodTabela"
    colCampoValor.Add "CodProduto", objPrecoCalculado.sCodProduto, STRING_PRODUTO, "CodProduto"
    colCampoValor.Add "DataReferencia", objPrecoCalculado.dtDataReferencia, 0, "DataReferencia"

    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa
    
    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 116836

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158741)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
'Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim objPrecoCalculado As New ClassPrecoCalculado

On Error GoTo Erro_Tela_Preenche

    objPrecoCalculado.iFilialEmpresa = giFilialEmpresa
    objPrecoCalculado.iCodTabela = colCampoValor.Item("CodTabela").vValor
    objPrecoCalculado.sCodProduto = colCampoValor.Item("CodProduto").vValor
    objPrecoCalculado.dtDataReferencia = colCampoValor.Item("DataReferencia").vValor
    If (objPrecoCalculado.iCodTabela <> 0) And (objPrecoCalculado.sCodProduto <> "") Then

       'Traz os dados para tela
       lErro = Traz_PrecoCalculado_Tela(objPrecoCalculado)
       If lErro <> SUCESSO Then gError 116837

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 116837

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158742)

    End Select

    Exit Sub

End Sub

Private Function Carrega_Dados_TabelaPreco(objProduto As ClassProduto) As Long

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim objTabelaPrecoItem1 As New ClassTabelaPrecoItem
Dim objProdutosFilial As New ClassProdutoFilial

On Error GoTo Erro_Carrega_Dados_TabelaPreco

    'Preenche objTabelaPrecoItem
    objTabelaPrecoItem.iCodTabela = CInt(TabelaPreco.Text)
    objTabelaPrecoItem.sCodProduto = objProduto.sCodigo
    objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
    
    'Lê TabelaPrecoItem
    lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem)
    If lErro <> SUCESSO And lErro <> 28014 Then gError 116840
    
    'Se encontrou o Ítem mostra o Valor na tela
    If lErro <> 28014 Then
        
        If objTabelaPrecoItem.dPreco <> 0 Then
            ValorFilial.Caption = Format(objTabelaPrecoItem.dPreco, "Fixed")
        Else
            ValorFilial.Caption = ""
        End If
                        
    Else
        
        'Se não encontrou o Ítem limpa o campo Valor
        ValorFilial.Caption = ""
                
    End If
    
    'Preenche objTabelaPrecoItem1
    objTabelaPrecoItem1.iCodTabela = CInt(TabelaPreco.Text)
    objTabelaPrecoItem1.sCodProduto = objProduto.sCodigo
    objTabelaPrecoItem1.iFilialEmpresa = EMPRESA_TODA

    'Lê TabelaPrecoItem
    lErro = CF("TabelaPrecoItem_Le", objTabelaPrecoItem1)
    If lErro <> SUCESSO And lErro <> 28014 Then gError 116841

    'Se encontrou o Ítem mostra o ValorEmpresa na tela
    If lErro <> 28014 Then
        If objTabelaPrecoItem1.dPreco <> 0 Then
            ValorEmpresa.Caption = Format(objTabelaPrecoItem1.dPreco, "Fixed")
        Else
            ValorEmpresa.Caption = ""
        End If
    Else
        'Limpa campo ValorEmpresa
        ValorEmpresa.Caption = ""
    End If

    objProdutosFilial.iFilialEmpresa = giFilialEmpresa
    objProdutosFilial.sProduto = objProduto.sCodigo
    
    lErro = CF("ProdutoFilial_Le", objProdutosFilial)
    If lErro <> SUCESSO And lErro <> 28261 Then gError 116842
    If lErro <> SUCESSO Then gError 116843
        
    Carrega_Dados_TabelaPreco = SUCESSO
    
    Exit Function
    
Erro_Carrega_Dados_TabelaPreco:

    Carrega_Dados_TabelaPreco = gErr
    
    Select Case gErr
    
        Case 116840 To 116843
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158743)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long
Dim iIndice As Integer
Dim colCodigo As New Collection

On Error GoTo Erro_Form_Load
    
    bCarregandoTela = False
    
    'Set objEventoItemTabela = New AdmEvento
    Set objEventoProduto = New AdmEvento
    Set objEventoPrecoCalculado = New AdmEvento
    
    'Inicializa a mascara de produto
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", Produto)
    If lErro <> SUCESSO Then gError 116820

    'Chama Carrega_TabelaPreco
    lErro = Carrega_TabelaPreco(colCodigo)
    If lErro <> SUCESSO Then gError 116821

    If TabelaPreco.ListCount > 0 Then TabelaPreco.ListIndex = 0

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 116820, 116821

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158744)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub


Private Function Carrega_TabelaPreco(colCodigo As Collection) As Long
'Carrega a ComboBox Tabela

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Carrega_TabelaPreco

    'Preenche a ComboBox com  os Tipos de Documentos existentes no BD
    lErro = CF("TabelasPreco_Le_Codigos", colCodigo)
    If lErro <> SUCESSO Then gError 116846

    For iIndice = 1 To colCodigo.Count

        'Preenche a ComboBox Tabela com os objetos da colecao colTabelaPreco
        TabelaPreco.AddItem colCodigo(iIndice)
        TabelaPreco.ItemData(TabelaPreco.NewIndex) = colCodigo(iIndice)
    Next

    Carrega_TabelaPreco = SUCESSO

    Exit Function

Erro_Carrega_TabelaPreco:

    Carrega_TabelaPreco = gErr

    Select Case gErr

        Case 116846

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158745)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objPrecoCalculado As New ClassPrecoCalculado

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica preenchimento de tabela
    If Len(Trim(TabelaPreco.Text)) = 0 Then gError 116847

    'Verifica preenchimento de produto
    If Len(Trim(Produto.ClipText)) = 0 Then gError 116848

    'Verifica preenchimento de valor
    If Len(Trim(NovoPreco.ClipText)) = 0 Then gError 116849
    
    'Verifica valor do novo preco
    If CDbl(NovoPreco.Text) <= 0 Then gError 116900
    
    'Verifica o Preenchimento da Data de Vigencia
    If Len(Trim(DataVigencia.ClipText)) = 0 Then gError 116850
    
    'Verifica se a data de Vigencia é maior que date
    If CDate(DataVigencia) < Date Then gError 116851
        
    'Chama Move_Tela_Memoria
    lErro = Move_Tela_Memoria(objTabelaPrecoItem)
    If lErro <> SUCESSO Then gError 116852

    lErro = Trata_Alteracao(objTabelaPrecoItem, objTabelaPrecoItem.iCodTabela, objTabelaPrecoItem.iFilialEmpresa, objTabelaPrecoItem.sCodProduto, objTabelaPrecoItem.dtDataVigencia)
    If lErro <> SUCESSO Then gError 116853

    'Preenche objPrecoCalculado com dados que serão atualizados na tabela Preco Calculado
    objPrecoCalculado.sCodProduto = objTabelaPrecoItem.sCodProduto
    
    'Preenche o objPrecoCalculado
    lErro = Move_Tela_Memoria2(objPrecoCalculado)
    If lErro <> SUCESSO Then gError 116854
    
    lErro = CF("PrecoCalculado_Grava", objPrecoCalculado)
    If lErro <> SUCESSO Then gError 116855

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 116847
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_NAO_PREENCHIDA", gErr)

        Case 116848
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)

        Case 116849
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOVOPRECO_NAO_PREENCHIDO", gErr)

        Case 116900
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOVOPRECO_NAO_POSITIVO", gErr)
        
        Case 116852 To 116855
  
        Case 116850
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_VIGENCIA_NAO_PREENCHIDA", gErr)
        
        Case 116851
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_VIGENCIA_MENOR_DATA_ATUAL", gErr, Date)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158746)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(objTabelaPrecoItem As ClassTabelaPrecoItem) As Long
'Move os dados da Tela para objTabelaPrecoItem

Dim lErro As Long
Dim objTabelaPreco As New ClassTabelaPreco
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    objTabelaPrecoItem.iFilialEmpresa = giFilialEmpresa
    
    If Len(Trim(TabelaPreco.Text)) > 0 Then objTabelaPrecoItem.iCodTabela = CInt(TabelaPreco.Text)

    If Len(Trim(Produto.ClipText)) > 0 Then

        'Critica o formato do Produto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 116856
        objTabelaPrecoItem.sCodProduto = sProdutoFormatado
    
    End If
   
    'Preenche objTabelaPrecoItem com preço
    objTabelaPrecoItem.dPreco = StrParaDbl(NovoPreco.Text)
    
    'Move a Data de Vigência para a Memoria
    If Len(Trim(DataVigencia.ClipText)) > 0 Then objTabelaPrecoItem.dtDataVigencia = StrParaDate(DataVigencia.Text)
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case 116856
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158747)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria2(objPrecoCalculado As ClassPrecoCalculado) As Long
'Move os dados da Tela para objPrecoCalculado

Dim lErro As Long

On Error GoTo Erro_Move_Tela_Memoria2

    'Preenche objPrecoCalculado
    objPrecoCalculado.iFilialEmpresa = giFilialEmpresa
    objPrecoCalculado.dtDataVigencia = StrParaDate(DataVigencia.Text)
    If Len(Trim(TabelaPreco.Text)) > 0 Then objPrecoCalculado.iCodTabela = CInt(TabelaPreco.Text)
    
    'verifica se datareferencia foi preenchida
    objPrecoCalculado.dtDataReferencia = StrParaDate(DataReferencia.Text)
        
    objPrecoCalculado.dPrecoInformado = StrParaDbl(NovoPreco.Text)
    
    Move_Tela_Memoria2 = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria2:

    Move_Tela_Memoria2 = gErr

    Select Case gErr
       
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158748)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()
Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 116859

    'Limpa a tela
    Call Limpa_Tela_TabelaPrecoItem

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 116859

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158749)

    End Select

    Exit Sub
End Sub

Sub Limpa_Tela_TabelaPrecoItem()
'Limpa a Tela TabelarecoItem

Dim lErro As Long

    'Limpa campos do tipo LABEL
    DescricaoProduto.Caption = ""
    UnidadeMedida.Caption = ""
    ValorFilial.Caption = ""
    ValorCalculado.Caption = ""
    DescricaoTabela.Caption = ""
    
    'Limpa checkbox - tabela default
    'TabelaDefault.Value = vbUnchecked
    
    'Funcao generica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    TabelaPreco.ListIndex = -1
    
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    iAlterado = 0
    iDataReferenciaAlterada = 0
    iProdutoAlterado = 0

End Sub

Private Sub DataReferencia_Change()

    iAlterado = REGISTRO_ALTERADO
    iDataReferenciaAlterada = REGISTRO_ALTERADO

End Sub

Private Sub DataReferencia_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_DataReferencia_Validate

    'se a data referencia nao foi alterada => sai da funcao
    If iDataReferenciaAlterada <> 1 Then Exit Sub
    
    'Verifica se a Data de Referência foi digitada
    If Len(Trim(DataReferencia.ClipText)) = 0 Or Len(Trim(Produto.ClipText)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataReferencia.Text)
    If lErro <> SUCESSO Then Error 116860
    
    'Preenche o código de objProduto
    lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 116861

    objProduto.sCodigo = sProdutoFormatado
    
    lErro = Carrega_Preco_Calculado()
    If lErro <> SUCESSO And lErro <> 116834 Then gError 116862
    
    Exit Sub

Erro_DataReferencia_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 116860, 116861
        
        'Preco Calculado não encontrado
        Case 116862

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158750)

    End Select

    Exit Sub

End Sub

Private Sub DataVigencia_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataVigencia_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataVigencia_Validate

    'Verifica se a Data de Emissao foi digitada
    If Len(Trim(DataVigencia.Text)) = 0 Then Exit Sub

    'Critica a data digitada
    lErro = Data_Critica(DataVigencia.Text)
    If lErro <> SUCESSO Then gError 116863
    
    Exit Sub

Erro_DataVigencia_Validate:

    Cancel = True

    Select Case gErr

        'se houve erro de crítica, segura o foco
        Case 116863

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158751)

    End Select

    Exit Sub

End Sub

Private Sub Produto_Change()
    
    iProdutoAlterado = REGISTRO_ALTERADO
    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim objTabelaPrecoItem1 As New ClassTabelaPrecoItem
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Erro_Produto_Validate

    If iProdutoAlterado = REGISTRO_ALTERADO Then
    
        If Len(Trim(Produto.ClipText)) > 0 Then
        
            sProduto = Produto.Text
    
            'Critica o formato do Produto e se existe no BD
            lErro = CF("Produto_Critica_Filial", sProduto, objProduto, iProdutoPreenchido)
            If lErro <> SUCESSO And lErro <> 51381 Then gError 116864
    
            If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 116865
            
            If lErro = 51381 Then gError 116866
            
            If objProduto.iFaturamento = PRODUTO_NAO_VENDAVEL Then gError 116867
    
            'Preenche ProdutoDescricao com Descrição do Produto
            DescricaoProduto.Caption = objProduto.sDescricao
    
            'Mostra a unidade de medida na tela
            UnidadeMedida.Caption = objProduto.sSiglaUMVenda
            
            'Limpa campos de valores
            ValorFilial.Caption = ""
            ValorCalculado.Caption = ""
            
            'Limpa campo NovoPreco
            NovoPreco.PromptInclude = False
            NovoPreco.Text = ""
            NovoPreco.PromptInclude = True
            
            'Limpa a Data de Vigência
            DataVigencia.PromptInclude = False
            DataVigencia.Text = ""
            DataVigencia.PromptInclude = True
                       
            'Verifica se Tabela está preenchida
            If Len(Trim(TabelaPreco.Text)) > 0 Then
                                
                'Carrega os dados na Tela
                lErro = Carrega_Dados_TabelaPreco(objProduto)
                If lErro <> SUCESSO Then gError 116868
                    
                'Carrega PrecoCalculado
                lErro = Carrega_Preco_Calculado()
                If lErro <> SUCESSO And lErro <> 116834 Then gError 116827
                
                If lErro = 116834 Then gError 116905
                    
            End If
            
        Else
            'Limpa DescricaoProduto
            DescricaoProduto.Caption = ""
            
            'Limpa unidade de medida do produto
            UnidadeMedida.Caption = ""
    
        End If
        iProdutoAlterado = 0
        
    End If
    
    Exit Sub

Erro_Produto_Validate:

    Cancel = True

    Select Case gErr

        Case 116905
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRECOCALCULADO_INEXISTENTE ", gErr, objProduto.sCodigo, giFilialEmpresa)

        Case 116865, 116866 'Não encontrou Produto no BD

            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", objProduto.sCodigo)

            If vbMsgRes = vbYes Then
                'Chama a tela de Produtos
                Call Chama_Tela("Produto", objProduto)

            Else
                'Limpa DescricaoProduto
                DescricaoProduto.Caption = ""
                'Segura o foco
            End If

        Case 116868, 116827, 116864
        
        Case 116867
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PODE_SER_VENDIDO", gErr, Produto.Text)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158752)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataVigencia_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVigencia_DownClick

    'Diminui a DataDe em 1 dia
    lErro = Data_Up_Down_Click(DataVigencia, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 116869

    Exit Sub

Erro_UpDownDataVigencia_DownClick:

    Select Case gErr

        Case 116869

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158753)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataVigencia_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDataVigencia_UpClick

    'Aumenta a DataDe em 1 dia
    lErro = Data_Up_Down_Click(DataVigencia, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 116870

    Exit Sub

Erro_UpDownDataVigencia_UpClick:

    Select Case gErr

        Case 116870

        Case Else
             lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158754)

    End Select

    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto

On Error GoTo Erro_objEventoProduto_evSelecao

    Set objProduto = obj1

    'Lê o Produto
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 116903

    'Se não achou o Produto --> erro
    If lErro = 28030 Then gError 116904
    
    Produto.PromptInclude = False
    Produto.Text = objProduto.sCodigo
    Produto.PromptInclude = True
        
    Me.Show
    
    Call Produto_Validate(bSGECancelDummy)
    
    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 116903

        Case 116904
            lErro = Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158755)

    End Select

    Exit Sub

End Sub

Private Sub ProdutoLabel_Click()

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As New Collection

On Error GoTo Erro_ProdutoLabel_Click

    'Verifica se o produto foi preenchido
    If Len(Produto.ClipText) <> 0 Then

        'Preenche o código de objProduto
        lErro = CF("Produto_Formata", Produto.Text, sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 116901

        objProduto.sCodigo = sProdutoFormatado

    End If

    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub

Erro_ProdutoLabel_Click:

    Select Case gErr

        Case 116901

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158756)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objPrecoCalculado As ClassPrecoCalculado) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se há um item selecionado, exibir seus dados
    If Not (objPrecoCalculado Is Nothing) Then

        lErro = Traz_PrecoCalculado_Tela(objPrecoCalculado)
        If lErro <> SUCESSO Then gError 116871

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 116871

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 158757)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Sub TabelaPreco_Click()
Dim lErro As Long
Dim objTabelaPreco As New ClassTabelaPreco
Dim objTabelaPrecoItem As New ClassTabelaPrecoItem
Dim objTabelaPrecoItem1 As New ClassTabelaPrecoItem
Dim objProduto As New ClassProduto
Dim sProduto As String
Dim iProdutoPreenchido As Integer

On Error GoTo Error_Tabela_Click

    'Verifica se foi preenchida a ComboBox Tabela
    If TabelaPreco.ListIndex <> -1 Then

        objTabelaPreco.iCodigo = CInt(TabelaPreco.Text)

        lErro = CF("TabelaPreco_Le", objTabelaPreco)
        If lErro <> SUCESSO And lErro <> 28004 Then gError 116872

        If lErro = 28004 Then gError 116873

        DescricaoTabela.Caption = objTabelaPreco.sDescricao

        If Len(Trim(Produto.ClipText)) > 0 Then
         
            sProduto = Produto.Text
        
            'Critica o Produto
            lErro = CF("Produto_Critica", sProduto, objProduto, iProdutoPreenchido)
            If lErro <> SUCESSO And lErro <> 25041 Then gError 116874
            
            If iProdutoPreenchido <> PRODUTO_PREENCHIDO Then gError 116875
              
            'Preeche os Dados da Tabela
            lErro = Carrega_Dados_TabelaPreco(objProduto)
            If lErro <> SUCESSO Then gError 116876
              
            Call DataReferencia_Validate(bSGECancelDummy)
              
        End If
    
    End If
    
    Exit Sub

Error_Tabela_Click:

    Select Case gErr

        Case 116872, 116874, 116876, 116875

        Case 116873
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECO_INEXISTENTE", gErr, objTabelaPreco.iCodigo)
            TabelaPreco.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 158758)

    End Select

    Exit Sub
    
End Sub

Private Function Traz_PrecoCalculado_Tela(objPrecoCalculado As ClassPrecoCalculado) As Long
'Traz os dados da Tabela Preço Item para a Tela

Dim lErro As Long
Dim sCodigo As String
Dim sProdutoEnxuto As String
Dim objTabelaPreco As New ClassTabelaPreco
Dim objPrecoCalculado1 As New ClassPrecoCalculado

On Error GoTo Erro_Traz_PrecoCalculado_Tela
    
    bCarregandoTela = True
    
    'Seleciona a Tabela na Combo
    Call Combo_Seleciona_ItemData(TabelaPreco, objPrecoCalculado.iCodTabela)
    
    If TabelaPreco.ListIndex = -1 Then gError 116838
    
    'Preenche o Produto
    lErro = CF("Traz_Produto_MaskEd", objPrecoCalculado.sCodProduto, Produto, DescricaoProduto)
    If lErro <> SUCESSO Then gError 116839
    
    Call DateParaMasked(DataReferencia, objPrecoCalculado.dtDataReferencia)
    
    Call Produto_Validate(bSGECancelDummy)

    bCarregandoTela = False
    
    iAlterado = 0

    Traz_PrecoCalculado_Tela = SUCESSO

    Exit Function

Erro_Traz_PrecoCalculado_Tela:

    bCarregandoTela = False
    
    Traz_PrecoCalculado_Tela = gErr

    Select Case gErr
    
        Case 116839

        Case 116838
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TABELAPRECOITEM_INEXISTENTE", gErr, objPrecoCalculado.iCodTabela)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 158759)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_ITENS_TABELA_PRECO
    Set Form_Load_Ocx = Me
    Caption = "Ajuste de Preços"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "CustoTabelaPreco"
    
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
        
        If Me.ActiveControl Is TabelaPreco Then
            Call LabelTabela_Click
        ElseIf Me.ActiveControl Is Produto Then
            Call ProdutoLabel_Click
        End If
    
    End If

End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub ProdutoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ProdutoLabel, Source, X, Y)
End Sub

Private Sub ProdutoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ProdutoLabel, Button, Shift, X, Y)
End Sub

Private Sub DescricaoProduto_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescricaoProduto, Source, X, Y)
End Sub

Private Sub DescricaoProduto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescricaoProduto, Button, Shift, X, Y)
End Sub

Private Sub DescricaoTabela_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DescricaoTabela, Source, X, Y)
End Sub

Private Sub DescricaoTabela_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DescricaoTabela, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub LabelTabela_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelTabela, Source, X, Y)
End Sub

Private Sub LabelTabela_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelTabela, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub ValorEmpresa_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorEmpresa, Source, X, Y)
End Sub

Private Sub ValorEmpresa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorEmpresa, Button, Shift, X, Y)
End Sub

Private Sub UnidadeMedida_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(UnidadeMedida, Source, X, Y)
End Sub

Private Sub UnidadeMedida_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(UnidadeMedida, Button, Shift, X, Y)
End Sub
