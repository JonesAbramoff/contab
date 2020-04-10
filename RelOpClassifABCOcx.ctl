VERSION 5.00
Begin VB.UserControl RelOpClassifABCOcx 
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8910
   LockControls    =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   8910
   Begin VB.ComboBox ComboOrdena 
      Height          =   315
      ItemData        =   "RelOpClassifABCOcx.ctx":0000
      Left            =   1410
      List            =   "RelOpClassifABCOcx.ctx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   105
      Width           =   2520
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   7560
      ScaleHeight     =   495
      ScaleWidth      =   1125
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1185
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   120
         Picture         =   "RelOpClassifABCOcx.ctx":0020
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "RelOpClassifABCOcx.ctx":0552
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5820
      Picture         =   "RelOpClassifABCOcx.ctx":06D0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox Classificacoes 
      Height          =   4155
      Left            =   5880
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1140
      Width           =   2775
   End
   Begin VB.Frame Frame4 
      Caption         =   "Identificação"
      Height          =   1395
      Left            =   120
      TabIndex        =   29
      Top             =   660
      Width           =   5505
      Begin VB.CheckBox AtualizaProdutosFilial 
         Caption         =   "Atualiza Classe ABC na tabela de Produtos"
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
         Height          =   225
         Left            =   420
         TabIndex        =   1
         Top             =   1065
         Width           =   4005
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   35
         Top             =   690
         Width           =   930
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3225
         TabIndex        =   34
         Top             =   285
         Width           =   480
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   630
         TabIndex        =   33
         Top             =   285
         Width           =   660
      End
      Begin VB.Label Codigo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1380
         TabIndex        =   32
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Data 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3750
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Descricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1365
         TabIndex        =   30
         Top             =   645
         Width           =   3510
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Faixa de Tempo"
      Height          =   1065
      Left            =   120
      TabIndex        =   20
      Top             =   3435
      Width           =   5475
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Mês Inicial:"
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
         Left            =   735
         TabIndex        =   28
         Top             =   330
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Mês Final:"
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
         Left            =   840
         TabIndex        =   27
         Top             =   705
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Ano Inicial:"
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
         Left            =   2805
         TabIndex        =   26
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ano Final:"
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
         Left            =   2910
         TabIndex        =   25
         Top             =   705
         Width           =   870
      End
      Begin VB.Label MesInicial 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1785
         TabIndex        =   24
         Top             =   285
         Width           =   375
      End
      Begin VB.Label MesFinal 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1785
         TabIndex        =   23
         Top             =   675
         Width           =   375
      End
      Begin VB.Label AnoInicial 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3855
         TabIndex        =   22
         Top             =   270
         Width           =   645
      End
      Begin VB.Label AnoFinal 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3855
         TabIndex        =   21
         Top             =   660
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Produto"
      Height          =   1185
      Left            =   120
      TabIndex        =   15
      Top             =   2145
      Width           =   5505
      Begin VB.CheckBox TodosTipos 
         Caption         =   "Todos os tipos"
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
         Height          =   285
         Left            =   210
         TabIndex        =   2
         Top             =   300
         Value           =   1  'Checked
         Width           =   1605
      End
      Begin VB.Label TipoDescricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2580
         TabIndex        =   19
         Top             =   735
         Width           =   2715
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1605
         TabIndex        =   18
         Top             =   765
         Width           =   930
      End
      Begin VB.Label TipoLabel 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
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
         Left            =   2070
         TabIndex        =   17
         Top             =   330
         Width           =   450
      End
      Begin VB.Label Tipo 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2580
         TabIndex        =   16
         Top             =   270
         Width           =   555
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Faixas de Classificação"
      Height          =   765
      Left            =   135
      TabIndex        =   8
      Top             =   4605
      Width           =   5475
      Begin VB.Label FaixaC 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4650
         TabIndex        =   14
         Top             =   330
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Faixa C:"
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
         Left            =   3885
         TabIndex        =   13
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Faixa A:"
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
         Left            =   375
         TabIndex        =   12
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Faixa B:"
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
         Left            =   2160
         TabIndex        =   11
         Top             =   360
         Width           =   705
      End
      Begin VB.Label FaixaA 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1155
         TabIndex        =   10
         Top             =   330
         Width           =   555
      End
      Begin VB.Label FaixaB 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2925
         TabIndex        =   9
         Top             =   330
         Width           =   555
      End
   End
   Begin VB.Label Label12 
      Caption         =   "Ordena por:"
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
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   135
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Classificações"
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
      Left            =   5925
      TabIndex        =   36
      Top             =   915
      Width           =   1230
   End
End
Attribute VB_Name = "RelOpClassifABCOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub Form_Activate()

    Call TelaIndice_Preenche(Me)
    
End Sub

Private Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Private Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Carrega todas as Classificações existentes para a Lista
    lErro = Carrega_ClassificacaoABC()
    If lErro <> SUCESSO Then Error 52098

    ComboOrdena.ListIndex = 1
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 52098

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 167550)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then Error 52099

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case Err

        Case 52099

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167551)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim sOrdena As String

On Error GoTo Erro_PreencherRelOp
        
    If Len(Trim(Codigo.Caption)) = 0 Then Error 52100
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then Error 52101
    
    'preenche com o codigo oqual será feito o relatorio
    lErro = objRelOpcoes.IncluirParametro("TCODIGO", Codigo.Caption)
    If lErro <> AD_BOOL_TRUE Then Error 52102
    
    sOrdena = CStr(ComboOrdena.ListIndex)
    
    lErro = objRelOpcoes.IncluirParametro("NORDENA", sOrdena)
    If lErro <> AD_BOOL_TRUE Then Error 59633
  
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = Err

    Select Case Err

        Case 52100
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FALTA_CODIGO_CLASSIFABC", Err)
        
        Case 52101, 52102, 59633
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167552)

    End Select

    Exit Function

End Function

Private Function Carrega_ClassificacaoABC() As Long
'Carrega a Lista com as Classificações existentes no BD

Dim lErro As Long
Dim iIndice As Integer
Dim colNumIntCodigo As New Collection
Dim objClassABC As New ClassClassificacaoABC

On Error GoTo Erro_Carrega_ClassificacaoABC

    'Lê todas as Classificações existentes para esta FilialEmpresa
    lErro = CF("ClassificacoesABC_Le",colNumIntCodigo)
    If lErro <> SUCESSO Then Error 52103

    'Preenche a Lista das Classificações com os objetos da coleção colCodigos
    For Each objClassABC In colNumIntCodigo
        Classificacoes.AddItem objClassABC.sCodigo
        Classificacoes.ItemData(Classificacoes.NewIndex) = objClassABC.lNumInt
    Next

    Carrega_ClassificacaoABC = SUCESSO

    Exit Function

Erro_Carrega_ClassificacaoABC:

    Carrega_ClassificacaoABC = Err

    Select Case Err

        Case 52103

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167553)

    End Select

    Exit Function

End Function

Private Sub Classificacoes_DblClick()

Dim lErro As Long
Dim objClassABC As New ClassClassificacaoABC

On Error GoTo Erro_Classificacoes_DblClick

    'Guarda o valor do código da Classificação selecionada na ListBox Classificacoes
    objClassABC.lNumInt = Classificacoes.ItemData(Classificacoes.ListIndex)
    
    'Lê a ClassificacaoABC no BD
    lErro = CF("ClassificacaoABC_Le_NumInt",objClassABC)
    If lErro <> SUCESSO And lErro <> 43500 Then Error 52104

    'Se não encontrou a ClassificacaoABC --> Erro
    If lErro <> SUCESSO Then Error 52105

    'Exibe os dados da ClassificacaoABC
    lErro = Traz_ClassificacaoABC_Tela(objClassABC)
    If lErro <> SUCESSO Then Error 52106

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_Classificacoes_DblClick:

    Select Case Err

        Case 52104, 52106
    
        Case 52105
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLASSIFICACAOABC_INEXISTENTE2", Err, objClassABC.lNumInt)
            Classificacoes.RemoveItem (Classificacoes.ListIndex)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 167554)
    
    End Select

    Exit Sub

End Sub

Private Function Traz_ClassificacaoABC_Tela(objClassABC As ClassClassificacaoABC) As Long
'Traz os dados da ClassificacaoABC passada em objClassABC

Dim lErro As Long
Dim objCurvaABC As New ClassCurvaABC
Dim objTipoDeProduto As New ClassTipoDeProduto

On Error GoTo Erro_Traz_ClassificacaoABC_Tela

    'Limpa a tela ClassificaoABC
    Call Limpa_Tela_ClassificacaoABC
    
    Codigo.Caption = objClassABC.sCodigo
    
    Data.Caption = CStr(objClassABC.dtData)
    
    Descricao.Caption = objClassABC.sDescricao
    
    If objClassABC.iAtualizaProdutosFilial = CLASSABC_ATUALIZA_PRODFILIAL Then
        AtualizaProdutosFilial.Value = CLASSABC_ATUALIZA_PRODFILIAL
    Else
        AtualizaProdutosFilial.Value = vbUnchecked
    End If
    
    If objClassABC.iTipoProduto <> 0 Then
        
        TodosTipos.Value = vbUnchecked
        Tipo.Caption = CStr(objClassABC.iTipoProduto)
        
        objTipoDeProduto.iTipo = CInt(Tipo.Caption)
        
        'Lê o Tipo de Produto para que preencha a descricao do tipo
        lErro = CF("TipoDeProduto_Le",objTipoDeProduto)
        If lErro <> SUCESSO And lErro <> 22531 Then Error 52107
        
        'Se não encontrou o Tipo de Produto --> Erro
        If lErro <> SUCESSO Then Error 52108
        
        TipoDescricao.Caption = objTipoDeProduto.sDescricao
        
    Else
        TodosTipos.Value = 1
        TipoDescricao.Caption = ""
        Tipo.Caption = ""
    End If
    
    MesInicial.Caption = CStr(objClassABC.iMesInicial)
    MesFinal.Caption = CStr(objClassABC.iMesFinal)
    AnoInicial.Caption = CStr(objClassABC.iAnoInicial)
    AnoFinal.Caption = CStr(objClassABC.iAnoFinal)
    FaixaA.Caption = CStr(objClassABC.iFaixaA)
    FaixaB.Caption = CStr(objClassABC.iFaixaB)
    FaixaC.Caption = CStr(100 - objClassABC.iFaixaA - objClassABC.iFaixaB) & "%"
    
    Traz_ClassificacaoABC_Tela = SUCESSO

    Exit Function

Erro_Traz_ClassificacaoABC_Tela:

    Traz_ClassificacaoABC_Tela = Err

    Select Case Err
        
        Case 52107
        
        Case 52108
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", Err, objTipoDeProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167555)

    End Select

    Exit Function

End Function

'""""""""""""""""""""""""""""""""""""""""""""""
'"  ROTINAS RELACIONADAS AS SETAS DO SISTEMA "'
'""""""""""""""""""""""""""""""""""""""""""""""

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)

Dim lErro As Long
Dim objClassABC As New ClassClassificacaoABC

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ClassificacaoABC"

    'Lê os dados da Tela ClassificacaoABC
    lErro = Move_Tela_Memoria(objClassABC)
    If lErro <> SUCESSO Then Error 52109

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumInt", CLng(0), 0, "NumInt"
    colCampoValor.Add "FilialEmpresa", objClassABC.iFilialEmpresa, 0, "FilialEmpresa"
    colCampoValor.Add "Codigo", objClassABC.sCodigo, STRING_CLASSABC_CODIGO, "Codigo"
    colCampoValor.Add "Descricao", objClassABC.sDescricao, STRING_CLASSABC_DESCRICAO, "Descricao"
    colCampoValor.Add "Data", objClassABC.dtData, 0, "Data"
    colCampoValor.Add "MesInicial", objClassABC.iMesInicial, 0, "MesInicial"
    colCampoValor.Add "AnoInicial", objClassABC.iAnoInicial, 0, "AnoInicial"
    colCampoValor.Add "MesFinal", objClassABC.iMesFinal, 0, "MesFinal"
    colCampoValor.Add "AnoFinal", objClassABC.iAnoFinal, 0, "AnoFinal"
    colCampoValor.Add "FaixaA", objClassABC.iFaixaA, 0, "FaixaA"
    colCampoValor.Add "FaixaB", objClassABC.iFaixaB, 0, "FaixaB"
    colCampoValor.Add "TipoProduto", objClassABC.iTipoProduto, 0, "TipoProduto"
    colCampoValor.Add "DemandaTotal", objClassABC.dDemandaTotal, 0, "DesmandaTotal"
    colCampoValor.Add "AtualizaProdutosFilial", objClassABC.iAtualizaProdutosFilial, 0, "AtualizaProdutosFilial"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, objClassABC.iFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case Err

        Case 52109

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 167556)

    End Select

    Exit Sub

End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)

Dim lErro As Long
Dim objClassABC As New ClassClassificacaoABC

On Error GoTo Erro_Tela_Preenche

    objClassABC.lNumInt = colCampoValor.Item("NumInt").vValor

    If objClassABC.lNumInt <> 0 Then

        'Carrega objClassABC com os dados passados em colCampoValor
        objClassABC.iFilialEmpresa = colCampoValor.Item("FilialEmpresa").vValor
        objClassABC.sCodigo = colCampoValor.Item("Codigo").vValor
        objClassABC.sDescricao = colCampoValor.Item("Descricao").vValor
        objClassABC.dtData = colCampoValor.Item("Data").vValor
        objClassABC.iMesInicial = colCampoValor.Item("MesInicial").vValor
        objClassABC.iAnoInicial = colCampoValor.Item("AnoInicial").vValor
        objClassABC.iMesFinal = colCampoValor.Item("MesFinal").vValor
        objClassABC.iAnoFinal = colCampoValor.Item("AnoFinal").vValor
        objClassABC.iFaixaA = colCampoValor.Item("FaixaA").vValor
        objClassABC.iFaixaB = colCampoValor.Item("FaixaB").vValor
        objClassABC.iTipoProduto = colCampoValor.Item("TipoProduto").vValor
        objClassABC.dDemandaTotal = colCampoValor.Item("DemandaTotal").vValor
        objClassABC.iAtualizaProdutosFilial = colCampoValor.Item("AtualizaProdutosFilial").vValor

        'Traz dados do Almoxarifado para a Tela
        lErro = Traz_ClassificacaoABC_Tela(objClassABC)
        If lErro <> SUCESSO Then Error 52110

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 52110

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 167557)

    End Select

    Exit Sub

End Sub
'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'"           Termina Aqui as Rotinas relacionadas as Setas do sistema                 ""
'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

Private Function Move_Tela_Memoria(objClassABC As ClassClassificacaoABC) As Long

Dim lErro As Long
Dim objTipoProduto As New ClassTipoDeProduto

On Error GoTo Erro_Move_Tela_Memoria:

    objClassABC.iFilialEmpresa = giFilialEmpresa

    If Len(Trim(Codigo.Caption)) <> 0 Then
        objClassABC.sCodigo = Codigo.Caption
    End If
    
    If Len(Trim(Data.Caption)) <> 0 Then
        objClassABC.dtData = CDate(Data.Caption)
    End If
    
    objClassABC.sDescricao = Descricao.Caption
    
    If AtualizaProdutosFilial.Value = CLASSABC_ATUALIZA_PRODFILIAL Then
        objClassABC.iAtualizaProdutosFilial = CLASSABC_ATUALIZA_PRODFILIAL
    Else
        objClassABC.iAtualizaProdutosFilial = vbUnchecked
    End If
    
    'Verifica se algum Tipo de Produto foi selecionado
    If Len(Trim(Tipo.Caption)) <> 0 Then
        'Preenche objTipoProduto
        objTipoProduto.iTipo = CInt(Tipo.Caption)
        'Lê o Tipo de Produto
        lErro = CF("TipoDeProduto_Le",objTipoProduto)
        If lErro <> SUCESSO And lErro <> 1 Then Error 52111
        
        'Se não encontrou o Tipo de Produto --> Erro
        If lErro <> SUCESSO Then Error 52112
        
        objClassABC.iTipoProduto = objTipoProduto.iTipo
        
    End If
    
    If Len(Trim(MesInicial.Caption)) <> 0 Then objClassABC.iMesInicial = CInt(MesInicial.Caption)
    If Len(Trim(AnoInicial.Caption)) <> 0 Then objClassABC.iAnoInicial = CInt(AnoInicial.Caption)
    If Len(Trim(MesFinal.Caption)) <> 0 Then objClassABC.iMesFinal = CInt(MesFinal.Caption)
    If Len(Trim(AnoFinal.Caption)) <> 0 Then objClassABC.iAnoFinal = CInt(AnoFinal.Caption)
    If Len(Trim(FaixaA.Caption)) <> 0 Then objClassABC.iFaixaA = CInt(FaixaA.Caption)
    If Len(Trim(FaixaB.Caption)) <> 0 Then objClassABC.iFaixaB = CInt(FaixaB.Caption)
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    Select Case Err
        
        Case 52111
        
        Case 52112
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOPRODUTO_INEXISTENTE", Err, objTipoProduto.iTipo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167558)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'Limpa a Tela
    Call Limpa_Tela_ClassificacaoABC

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 167559)

    End Select

    Exit Sub

End Sub

Private Sub Limpa_Tela_ClassificacaoABC()
'Limpa a Tela ClassificacaoABC
    
Dim lErro As Long

    'Limpa os campos que não são limpos pelo Limpa_Tela
    Codigo.Caption = ""
    Descricao.Caption = ""
    AtualizaProdutosFilial.Value = vbUnchecked
    TodosTipos.Value = vbChecked
    Tipo.Caption = ""
    TipoDescricao.Caption = ""
    FaixaA.Caption = ""
    FaixaB.Caption = ""
    FaixaC.Caption = ""
    AnoInicial.Caption = ""
    AnoFinal.Caption = ""
    MesInicial.Caption = ""
    MesFinal.Caption = ""
    Data.Caption = ""
    
    'Fecha o Comando de Setas
    lErro = ComandoSeta_Fechar(Me.Name)
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then Error 52113
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 52113
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 167560)

    End Select

    Exit Function

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_CLASSIFICACAO_ABC
    Set Form_Load_Ocx = Me
    Caption = "Relatório de Classificação ABC"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpClassifABC"
    
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



Private Sub Label2_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label2(Index), Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2(Index), Button, Shift, X, Y)
End Sub


Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Codigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Codigo, Source, X, Y)
End Sub

Private Sub Codigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Codigo, Button, Shift, X, Y)
End Sub

Private Sub Data_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Data, Source, X, Y)
End Sub

Private Sub Data_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Data, Button, Shift, X, Y)
End Sub

Private Sub Descricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Descricao, Source, X, Y)
End Sub

Private Sub Descricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Descricao, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub MesInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MesInicial, Source, X, Y)
End Sub

Private Sub MesInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MesInicial, Button, Shift, X, Y)
End Sub

Private Sub MesFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(MesFinal, Source, X, Y)
End Sub

Private Sub MesFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(MesFinal, Button, Shift, X, Y)
End Sub

Private Sub AnoInicial_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(AnoInicial, Source, X, Y)
End Sub

Private Sub AnoInicial_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(AnoInicial, Button, Shift, X, Y)
End Sub

Private Sub AnoFinal_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(AnoFinal, Source, X, Y)
End Sub

Private Sub AnoFinal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(AnoFinal, Button, Shift, X, Y)
End Sub

Private Sub TipoDescricao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoDescricao, Source, X, Y)
End Sub

Private Sub TipoDescricao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoDescricao, Button, Shift, X, Y)
End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label8, Source, X, Y)
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label8, Button, Shift, X, Y)
End Sub

Private Sub TipoLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TipoLabel, Source, X, Y)
End Sub

Private Sub TipoLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TipoLabel, Button, Shift, X, Y)
End Sub

Private Sub Tipo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Tipo, Source, X, Y)
End Sub

Private Sub Tipo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Tipo, Button, Shift, X, Y)
End Sub

Private Sub FaixaC_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FaixaC, Source, X, Y)
End Sub

Private Sub FaixaC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FaixaC, Button, Shift, X, Y)
End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label11, Source, X, Y)
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label11, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub FaixaA_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FaixaA, Source, X, Y)
End Sub

Private Sub FaixaA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FaixaA, Button, Shift, X, Y)
End Sub

Private Sub FaixaB_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(FaixaB, Source, X, Y)
End Sub

Private Sub FaixaB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(FaixaB, Button, Shift, X, Y)
End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label12, Source, X, Y)
End Sub

Private Sub Label12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label12, Button, Shift, X, Y)
End Sub

