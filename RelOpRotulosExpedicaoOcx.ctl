VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl RelOpRotulosExpedicaoOcx 
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9420
   ScaleHeight     =   5310
   ScaleWidth      =   9420
   Begin VB.CommandButton BotaoEmbalagens 
      Caption         =   "Produto  X  Embalagens"
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
      Left            =   5040
      TabIndex        =   21
      Top             =   4815
      Width           =   2355
   End
   Begin VB.Frame Frame6 
      Caption         =   "Nota Fiscal"
      Height          =   765
      Index           =   0
      Left            =   135
      TabIndex        =   14
      Top             =   105
      Width           =   5280
      Begin VB.CommandButton BotaoExibirDados 
         Caption         =   "Exibir Dados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3555
         TabIndex        =   16
         Top             =   285
         Width           =   1425
      End
      Begin VB.ComboBox Serie 
         Height          =   315
         Left            =   765
         TabIndex        =   15
         Top             =   285
         Width           =   765
      End
      Begin MSMask.MaskEdBox NFiscal 
         Height          =   300
         Left            =   2520
         TabIndex        =   17
         Top             =   315
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin VB.Label NFLabel 
         AutoSize        =   -1  'True
         Caption         =   "Número:"
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
         Left            =   1695
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   345
         Width           =   720
      End
      Begin VB.Label SerieLabel 
         AutoSize        =   -1  'True
         Caption         =   "Série:"
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
         Left            =   165
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   18
         Top             =   330
         Width           =   510
      End
   End
   Begin VB.CommandButton BotaoLotes 
      Caption         =   "Lotes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7575
      TabIndex        =   9
      Top             =   4800
      Width           =   1665
   End
   Begin VB.Frame Frame18 
      Caption         =   "Definição das Etiquetas"
      Height          =   3795
      Left            =   135
      TabIndex        =   4
      Top             =   900
      Width           =   9150
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
         Height          =   570
         Left            =   3270
         TabIndex        =   29
         Top             =   3150
         Width           =   1425
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   570
         Left            =   135
         Picture         =   "RelOpRotulosExpedicaoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3150
         Width           =   1425
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   570
         Left            =   1740
         Picture         =   "RelOpRotulosExpedicaoOcx.ctx":101A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3150
         Width           =   1425
      End
      Begin VB.CheckBox Imprimir 
         Caption         =   "Imprimir"
         Height          =   225
         Left            =   1860
         TabIndex        =   22
         Top             =   2340
         Width           =   1035
      End
      Begin MSMask.MaskEdBox ItemNFRastro 
         Height          =   225
         Left            =   135
         TabIndex        =   5
         Top             =   720
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox LoteRastro 
         Height          =   225
         Left            =   1845
         TabIndex        =   6
         Top             =   705
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox FilialOPRastro 
         Height          =   225
         Left            =   2910
         TabIndex        =   7
         Top             =   795
         Width           =   1365
         _ExtentX        =   2408
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Embalagem 
         Height          =   225
         Left            =   4950
         TabIndex        =   10
         Top             =   690
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QuantEmb 
         Height          =   225
         Left            =   2625
         TabIndex        =   11
         Top             =   1605
         Width           =   1080
         _ExtentX        =   1905
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PesoLiq 
         Height          =   225
         Left            =   4590
         TabIndex        =   12
         Top             =   1170
         Width           =   1050
         _ExtentX        =   1852
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox PesoBruto 
         Height          =   225
         Left            =   5430
         TabIndex        =   13
         Top             =   1305
         Width           =   1170
         _ExtentX        =   2064
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
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ProdutoEmb 
         Height          =   225
         Left            =   975
         TabIndex        =   20
         Top             =   1065
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridEtiquetas 
         Height          =   2655
         Left            =   105
         TabIndex        =   8
         Top             =   225
         Width           =   8940
         _ExtentX        =   15769
         _ExtentY        =   4683
         _Version        =   393216
         Rows            =   51
         Cols            =   7
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
      Begin VB.Label PesoLTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7260
         TabIndex        =   28
         Top             =   3405
         Width           =   1470
      End
      Begin VB.Label PesoBTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7260
         TabIndex        =   27
         Top             =   3045
         Width           =   1470
      End
      Begin VB.Label Label1 
         Caption         =   "Peso Líquido:"
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
         Left            =   5985
         TabIndex        =   26
         Top             =   3420
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Peso Bruto:"
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
         Left            =   6150
         TabIndex        =   25
         Top             =   3075
         Width           =   1080
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8100
      ScaleHeight     =   495
      ScaleMode       =   0  'User
      ScaleWidth      =   1065
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   150
      Width           =   1125
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   75
         Picture         =   "RelOpRotulosExpedicaoOcx.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   585
         Picture         =   "RelOpRotulosExpedicaoOcx.ctx":272E
         Style           =   1  'Graphical
         TabIndex        =   2
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
      Height          =   555
      Left            =   6000
      Picture         =   "RelOpRotulosExpedicaoOcx.ctx":28AC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   165
      Width           =   1845
   End
End
Attribute VB_Name = "RelOpRotulosExpedicaoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'pendencias:
'no browse de lotes filtrar por lotes que estejam associados a nf
'na saida de celula do lote tb validar
'qdo definir o produto, se só houver 1 lote informado no rastreamento entao coloca-lo automaticamente

'Property Variables:
Dim m_Caption As String
Event Unload()

Public iAlterado As Integer
Public iAlteradoNF As Integer
Public iAlteradoSerie As Integer

Private glNumIntNF As Long

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private WithEvents objEventoSerie As AdmEvento
Attribute objEventoSerie.VB_VarHelpID = -1
Private WithEvents objEventoNFiscal As AdmEvento
Attribute objEventoNFiscal.VB_VarHelpID = -1
Private WithEvents objEventoEmbalagens As AdmEvento
Attribute objEventoEmbalagens.VB_VarHelpID = -1
Private WithEvents objEventoLoteRastro As AdmEvento
Attribute objEventoLoteRastro.VB_VarHelpID = -1

Public objGridEtiquetas As AdmGrid

Private iGrid_Imprimir_Col As Integer 'Inserido por Wagner
Private iGrid_ItemNFRastro_Col As Integer
Private iGrid_ProdutoEmb_Col As Integer
Private iGrid_LoteRastro_Col As Integer
Private iGrid_FilialOPRastro_Col As Integer
Private iGrid_Embalagem_Col As Integer
Private iGrid_QuantEmb_Col As Integer
Private iGrid_PesoLiq_Col As Integer
Private iGrid_PesoBruto_Col As Integer

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    glNumIntNF = 0
    
    Set objEventoSerie = New AdmEvento
    Set objEventoNFiscal = New AdmEvento
    Set objEventoEmbalagens = New AdmEvento
    Set objEventoLoteRastro = New AdmEvento
    
    'Carrega as Séries
    lErro = Carrega_Serie()
    If lErro <> SUCESSO Then gError 130320
    
    Set objGridEtiquetas = New AdmGrid
    
    'Inicializa Máscara de ProdutoEmb
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoEmb)
    If lErro <> SUCESSO Then gError 78154
    
    lErro = Inicializa_Grid_Etiquetas(objGridEtiquetas)
    If lErro <> SUCESSO Then gError 130319
    
    iAlteradoNF = 0
    iAlterado = 0
    iAlteradoSerie = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 130319, 130320
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173188)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
  
    Set objEventoSerie = Nothing
    Set objEventoNFiscal = Nothing
    Set objEventoEmbalagens = Nothing
    Set objEventoLoteRastro = Nothing

    Set objGridEtiquetas = Nothing
    
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes, Optional objNFiscal As ClassNFiscal) As Long

Dim lErro As Long
Dim bEncontrou As Boolean
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 122625
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    If Not (objNFiscal Is Nothing) Then
            
        '##############################
        'Inserido por Wagner
        lErro = Preenche_Tela_NF(objNFiscal)
        If lErro <> SUCESSO Then gError 136973
        '##############################
    
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 122625
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 136973
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173189)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Critica_Parametros() As Long
'Critica os parâmetros que serão passados para o relatório

Dim lErro As Long

On Error GoTo Erro_Critica_Parametros
          
    'Verifica se a Série Foi Preenchida
    If Len(Trim(Serie.Text)) = 0 Then gError 122629
    
    'Verifica se a Nota Fiscal Foi Preenchida
    If Len(Trim(NFiscal.Text)) = 0 Then gError 122630
    
    Critica_Parametros = SUCESSO

    Exit Function

Erro_Critica_Parametros:

    Critica_Parametros = gErr

    Select Case gErr

        Case 122629
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
            Serie.SetFocus
        
        Case 122630
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NF_NAO_PREENCHIDO", gErr)
            NFiscal.SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173190)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

   Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    Call Limpa_Tela_RelOpRotulosExpedicao
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 122633 'Tratado na Rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173191)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional bExecutar As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long, colRotuloExpedicao As New Collection, iLinha As Integer
Dim iIndice As Integer, lNumIntRel As Long, objRotuloExpedicao As ClassRelRotuloExpedicao
Dim iPreenchido As Integer, sProdutoFormatado As String, sProdutoMascarado As String

On Error GoTo Erro_PreencherRelOp

    lErro = Critica_Parametros()
    If lErro <> SUCESSO Then gError 122634
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 122635
    
    lErro = objRelOpcoes.IncluirParametro("NNFISCAL", NFiscal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 122636

    lErro = objRelOpcoes.IncluirParametro("TSERIE", Serie.Text)
    If lErro <> AD_BOOL_TRUE Then gError 122638
   
    If bExecutar Then
    
        For iLinha = 1 To objGridEtiquetas.iLinhasExistentes
        
            'validar linha
            '??? obrigar o preenchimento de pesos, lote e qtde de etiquetas
            
            Set objRotuloExpedicao = New ClassRelRotuloExpedicao
            
            With objRotuloExpedicao
                            
                .dPesoLiquido = StrParaDbl(GridEtiquetas.TextMatrix(iLinha, iGrid_PesoLiq_Col))
                .dPesoBruto = StrParaDbl(GridEtiquetas.TextMatrix(iLinha, iGrid_PesoBruto_Col))
                .iItem = StrParaInt(GridEtiquetas.TextMatrix(iLinha, iGrid_ItemNFRastro_Col))
                .sSerie = Serie.Text
                .lNumNotaFiscal = StrParaLong(NFiscal.Text)
                .iFilialOP = Codigo_Extrai(GridEtiquetas.TextMatrix(iLinha, iGrid_FilialOPRastro_Col))
                .sLote = GridEtiquetas.TextMatrix(iLinha, iGrid_LoteRastro_Col)
                .iQtdeEmb = StrParaInt(GridEtiquetas.TextMatrix(iLinha, iGrid_QuantEmb_Col))
                
                '################################################
                'Inserido por Wagner
                .iImprimir = StrParaInt(GridEtiquetas.TextMatrix(iLinha, iGrid_Imprimir_Col))
                '################################################
            
            End With
            
            sProdutoMascarado = GridEtiquetas.TextMatrix(iLinha, iGrid_ProdutoEmb_Col)
            If Len(Trim(sProdutoMascarado)) = 0 Then gError 130333
            
            'Formata o produto
            lErro = CF("Produto_Formata", sProdutoMascarado, sProdutoFormatado, iPreenchido)
            If lErro <> SUCESSO Then gError 130332

            objRotuloExpedicao.sProduto = sProdutoFormatado
                                                    
            colRotuloExpedicao.Add objRotuloExpedicao
            
        Next
        
        '###################################
        'Inserido por Wagner
        lErro = Critica_GridEtiquetas(colRotuloExpedicao)
        If lErro <> SUCESSO Then gError 136985
        '###################################
        
        '??? colocar como CF
        'obter numintrel e criar registros em tabela auxiliar
        lErro = RelRotulosExpedicao_Prepara(colRotuloExpedicao, lNumIntRel)
        If lErro <> SUCESSO Then gError 130331
        
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 122638
   
    End If
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 122634 To 122638, 130331, 130332, 136985 'Tratado na Rotina chamada

        Case 130333
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173192)

    End Select

    Exit Function

End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long
Dim objSerie As New ClassSerie
Dim lFaixaFinal As Long
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 122648

    lErro = gobjRelatorio.Executar_Prossegue
    If lErro <> SUCESSO And lErro <> 7072 Then gError 122653

    'Cancelou o relatório
    If lErro = 7072 Then gError 122654

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 122648, 122649, 122651, 122652, 122653 'Tratado na Rotina chamada

        Case 122650
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", gErr, objSerie.sSerie)

        Case 122654 'Cancelou o relatório
            Unload Me

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173193)

    End Select

    Exit Sub

End Sub

Private Sub LabelSerie_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As Collection

    'Recolhe a Série da tela
    objSerie.sSerie = Serie.Text

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerie)

    Exit Sub

End Sub

Private Sub BotaoLimparGrid_Click()
    Call Grid_Limpa(objGridEtiquetas)
End Sub

Private Sub NFiscal_Change()
    iAlteradoNF = REGISTRO_ALTERADO
End Sub

Private Sub NFiscal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim sNumero As String

On Error GoTo Erro_NFiscal_Validate

    If iAlteradoNF = REGISTRO_ALTERADO Then
    
        If Len(Trim(NFiscal.Text)) > 0 Then
            sNumero = NFiscal.Text
        End If
    
        lErro = Critica_Numero(sNumero)
        If lErro <> SUCESSO Then gError 122656

        Call Grid_Limpa(objGridEtiquetas)
        
        iAlteradoNF = 0

    End If

    Exit Sub

Erro_NFiscal_Validate:

    Cancel = True

    Select Case gErr

        Case 122656

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173194)

    End Select

    Exit Sub

End Sub


'William
'copiada da tela RelOpNotasFiscais
Private Function Critica_Numero(sNumero As String) As Long

Dim lErro As Long

On Error GoTo Erro_Critica_Numero

    If Len(Trim(sNumero)) > 0 Then

        lErro = Long_Critica(sNumero)
        If lErro <> SUCESSO Then gError 122661

        If CLng(sNumero) < 0 Then gError 122662

    End If

    Critica_Numero = SUCESSO

    Exit Function

Erro_Critica_Numero:

    Critica_Numero = gErr

    Select Case gErr

        Case 122661

        Case 122662
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_POSITIVO", gErr, sNumero)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173195)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_EMISSAO_NF
    Set Form_Load_Ocx = Me
    Caption = "Rótulos de Expedição para Notas Fiscais"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRotulosExpedicao"
    
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

Private Sub Serie_Change()
    iAlteradoSerie = REGISTRO_ALTERADO
End Sub

Private Sub Serie_Click()
    iAlteradoSerie = REGISTRO_ALTERADO
End Sub

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
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Serie Then
            Call LabelSerie_Click
        ElseIf Me.ActiveControl Is NFiscal Then
            Call NFLabel_Click
            
        End If
    
    End If

End Sub

Public Sub Serie_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Serie_Validate

    If iAlteradoSerie = REGISTRO_ALTERADO Then
    
        'Verififca se está preenchida
        If Len(Trim(Serie.Text)) = 0 Then Exit Sub
    
        'Verifica se foi alguma Série selecionada
        If Serie.Text = Serie.List(Serie.ListIndex) Then Exit Sub
    
        'Tenta achar a Série na combo
        lErro = Combo_Item_Igual(Serie)
        If lErro <> SUCESSO And lErro <> 12253 Then Error 35174
    
        'Não encontrou a Série
        If lErro = 12253 Then Error 35175
        
        Call Grid_Limpa(objGridEtiquetas)
        
        iAlteradoSerie = 0
        
    End If

    Exit Sub

Erro_Serie_Validate:

    Cancel = True

    Select Case Err

        Case 35174

        Case 35175
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_CADASTRADA", Err, Serie.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173196)

    End Select

    Exit Sub

End Sub

Public Sub objEventoSerie_evSelecao(obj1 As Object)

Dim objSerie As ClassSerie

    Set objSerie = obj1

    'Coloca a Série da Nota Fiscal na tela
    Serie.Text = objSerie.sSerie

    Call Grid_Limpa(objGridEtiquetas)

    Me.Show

End Sub

Public Sub NFLabel_Click()

Dim objNFiscal As New ClassNFiscal
Dim colSelecao As Collection

    'Guarda a Serie e o Número da Nota Fiscal  da Tela
    objNFiscal.sSerie = Serie.Text
    If Len(Trim(NFiscal.ClipText)) > 0 Then
        objNFiscal.lNumNotaFiscal = CLng(NFiscal.Text)
    Else
        objNFiscal.lNumNotaFiscal = 0
    End If

    'Chama a Tela NFiscalInternaSaidalLista
    Call Chama_Tela("NFiscalInternaSaidaLista", colSelecao, objNFiscal, objEventoNFiscal)

End Sub

Public Sub objEventoNFiscal_evSelecao(obj1 As Object)

Dim objNFiscal As ClassNFiscal

    Set objNFiscal = obj1

    'Preenche a Série e o Número da Nota Fiscal
    Serie.Text = objNFiscal.sSerie
    NFiscal.Text = objNFiscal.lNumNotaFiscal

    Call Grid_Limpa(objGridEtiquetas)

    Me.Show

End Sub

Public Sub BotaoExibirDados_Click()

Dim lErro As Long
Dim objNFiscal As New ClassNFiscal

On Error GoTo Erro_BotaoExibirDados_Click

    'Verifica se a Serie e o Número da Nota Fiscal original estão preenchidos
    If Len(Trim(Serie.Text)) = 0 Or Len(Trim(NFiscal.ClipText)) = 0 Then Error 35256

    objNFiscal.sSerie = Serie.Text
    objNFiscal.lNumNotaFiscal = NFiscal.Text
    objNFiscal.iFilialEmpresa = giFilialEmpresa
    
    '#############################
    'Comentado por Wagner
'    'Tenta lêr a Nota com esses dados
'    lErro = CF("NFiscalInternaSaida_Le_Numero", objNFiscal)
'    If lErro <> SUCESSO And lErro <> 30765 Then Error 35257
'    If lErro = 30765 Then Error 35258 'Não encontrou
    '#############################

    'Coloca na tela os dados encontrados
    lErro = Preenche_Tela_NF(objNFiscal)
    If lErro <> SUCESSO Then Error 35259

    '#############################
    'Comentado por Wagner
    'glNumIntNF = objNFiscal.lNumIntDoc
    
    'Serie.Enabled = False
    'NFiscal.Enabled = False
    '#############################
    
    Exit Sub

Erro_BotaoExibirDados_Click:

    Select Case Err

        Case 35256
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NUMERO_FALTANDO3", Err)

        Case 35257, 35259

        Case 35258
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_INTERNA_SAIDA_NAO_CADASTRADA", Err, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 173197)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_Grid_Etiquetas(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Imprimir") 'Inserido por Wagner
    objGridInt.colColuna.Add ("Item NF")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Lote")
    objGridInt.colColuna.Add ("Filial OP")
    objGridInt.colColuna.Add ("Embalagem")
    objGridInt.colColuna.Add ("Qtde.")
    objGridInt.colColuna.Add ("Peso Líq.")
    objGridInt.colColuna.Add ("Peso Bruto")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Imprimir.Name) 'Inserido por Wagner
    objGridInt.colCampo.Add (ItemNFRastro.Name)
    objGridInt.colCampo.Add (ProdutoEmb.Name)
    objGridInt.colCampo.Add (LoteRastro.Name)
    objGridInt.colCampo.Add (FilialOPRastro.Name)
    objGridInt.colCampo.Add (Embalagem.Name)
    objGridInt.colCampo.Add (QuantEmb.Name)
    objGridInt.colCampo.Add (PesoLiq.Name)
    objGridInt.colCampo.Add (PesoBruto.Name)

    'Colunas do Grid
    iGrid_Imprimir_Col = 1 'Inserido por Wagner
    iGrid_ItemNFRastro_Col = 2
    iGrid_ProdutoEmb_Col = 3
    iGrid_LoteRastro_Col = 4
    iGrid_FilialOPRastro_Col = 5
    iGrid_Embalagem_Col = 6
    iGrid_QuantEmb_Col = 7
    iGrid_PesoLiq_Col = 8
    iGrid_PesoBruto_Col = 9

    'Grid do GridInterno
    objGridInt.objGrid = GridEtiquetas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 21

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 9

    'Largura da primeira coluna
    GridEtiquetas.ColWidth(0) = 300

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Etiquetas = SUCESSO

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
    
        Select Case objGridInt.objGrid.Col
        
            Case iGrid_ItemNFRastro_Col
                lErro = Saida_Celula_ItemNFRastro(objGridInt)
                If lErro <> SUCESSO Then gError 83202
        
            Case iGrid_LoteRastro_Col
                lErro = Saida_Celula_LoteRastro(objGridInt)
                If lErro <> SUCESSO Then gError 83162
        
            Case iGrid_FilialOPRastro_Col
                lErro = Saida_Celula_FilialOPRastro(objGridInt)
                If lErro <> SUCESSO Then gError 83163
            
            Case iGrid_PesoBruto_Col
                lErro = Saida_Celula_PesoBruto(objGridInt)
                If lErro <> SUCESSO Then gError 83163
            
            Case iGrid_PesoLiq_Col
                lErro = Saida_Celula_PesoLiq(objGridInt)
                If lErro <> SUCESSO Then gError 83163
            
            Case iGrid_Embalagem_Col
                lErro = Saida_Celula_Embalagem(objGridInt)
                If lErro <> SUCESSO Then gError 83202
        
            Case iGrid_QuantEmb_Col
                lErro = Saida_Celula_QuantEmb(objGridInt)
                If lErro <> SUCESSO Then gError 83202
        
        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 26068
    
    End If
    
    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 83160, 83161, 83162, 83163, 83164, 83202, 83313

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173198)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_ItemNFRastro(objGridInt As AdmGrid) As Long
'Rastreamento

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim sProduto As String
Dim iLinha As Integer
Dim iAchou As Integer
Dim iItem As Integer
Dim iItem_Atual As Integer
Dim sProdutoMascarado As String
Dim objNFiscal As New ClassNFiscal
Dim dQuantidade As Double, objItemNF As New ClassItemNF
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objTipoDocInfo As New ClassTipoDocInfo
'##############################
'Inserido por Wagner
Dim objProdutoEmbalagem As ClassProdutoEmbalagem
Dim objEmbalagem As ClassEmbalagem
Dim objRastreamento As ClassRastreamentoMovto
Dim bAchou As Boolean
Dim objItemMovEstoque As ClassItemMovEstoque
Dim colItemMovEstoque As Collection
Dim colRastreamentoMovto As Collection
Dim objFilialEmpresa As New AdmFiliais
'##############################

On Error GoTo Erro_Saida_Celula_ItemNFRastro

    Set objGridInt.objControle = ItemNFRastro

    'Verifica se o Produto esta preenchido
    If Len(Trim(ItemNFRastro.Text)) = 0 And Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col))) > 0 Then gError 83225
    
    If Len(Trim(ItemNFRastro.Text)) > 0 Then

        iItem = CInt(ItemNFRastro.Text)
        
        objNFiscal.lNumNotaFiscal = StrParaLong(NFiscal.Text)
        objNFiscal.sSerie = Serie.Text
                
        'Tenta lêr a Nota com esses dados
        lErro = CF("NFiscalInternaSaida_Le_Numero", objNFiscal)
        If lErro <> SUCESSO And lErro <> 30765 Then gError 136974
        If lErro = 30765 Then gError 136975 'Não encontrou
    
        'Lê os Ítens da Nota Fiscal
        lErro = CF("NFiscalItens_Le", objNFiscal)
        If lErro <> SUCESSO Then gError 130321
        
        'Acha o item preenchdo
        bAchou = False
        For Each objItemNF In objNFiscal.ColItensNF
            If objItemNF.iItem = iItem Then
                bAchou = True
                Exit For
            End If
        Next
        
        'Se não achou = > Erro
        If Not bAchou Then gError 83201
    
        'Lê os demais atributos do Produto
        objProduto.sCodigo = objItemNF.sProduto
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 83251
            
        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 83252
                
        'se não for um produto rastreavel ==> erro
        If objProduto.iRastro = PRODUTO_RASTRO_NENHUM Then gError 83253

        'descobre qual o item atual
        If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col))) > 0 Then
            iItem_Atual = CInt(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col))
        End If
                
        'se o item que se está preenchendo é diferente do item atual, ==> limpa os campos de lote
        If iItem <> iItem_Atual Then
        
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col) = ""
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantEmb_Col) = ""
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col) = ""
            
            lErro = Mascara_RetornaProdutoTela(objItemNF.sProduto, sProdutoMascarado)
            If lErro <> SUCESSO Then gError 130322
            objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoEmb_Col) = sProdutoMascarado
            
        End If
        
        Set objProdutoEmbalagem = New ClassProdutoEmbalagem
        Set objEmbalagem = New ClassEmbalagem
        
        objProdutoEmbalagem.sProduto = objItemNF.sProduto
            
        'Seleciona embalagem padrao
        lErro = CF("ProdutoEmbalagem_Le_Padrao", objProdutoEmbalagem)
        If lErro <> SUCESSO Then gError 136970
    
        objEmbalagem.iCodigo = objProdutoEmbalagem.iEmbalagem
    
        lErro = CF("Embalagem_Le", objEmbalagem)
        If lErro <> SUCESSO Then gError 136971
    
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Embalagem_Col) = objEmbalagem.sSigla
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_PesoBruto_Col) = Formata_Estoque(objProdutoEmbalagem.dPesoBruto)
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_PesoLiq_Col) = Formata_Estoque(objProdutoEmbalagem.dPesoLiqTotal)
        
        '#####################################
        'Inserido por Wagner
        Set objItemMovEstoque = New ClassItemMovEstoque
        Set colItemMovEstoque = New Collection
        
        'Lê item de movimento de estoque
        objItemMovEstoque.lNumIntDocOrigem = objItemNF.lNumIntDoc
        objItemMovEstoque.iTipoNumIntDocOrigem = TIPO_ORIGEM_ITEMNF
        objItemMovEstoque.iFilialEmpresa = objNFiscal.iFilialEmpresa
                
        lErro = CF("MovEstoque_Le_ItemNF", objItemMovEstoque, colItemMovEstoque)
        If lErro <> SUCESSO Then gError 136978
        
        'Se só tiver um item
        If colItemMovEstoque.Count = 1 Then
        
            Set objItemMovEstoque = colItemMovEstoque.Item(1)
                    
            'se forem os movimentos de transferencia de material consignado ==> não trata-os, pois a venda de material consignado está sendo tratada
            If objItemMovEstoque.iTipoMov <> MOV_EST_SAIDA_TRANSF_CONSIG_TERC And _
               objItemMovEstoque.iTipoMov <> MOV_EST_ENTRADA_TRANSF_DISP1 Then
        
                Set colRastreamentoMovto = New Collection
                        
                'Lê movimentos de rastreamento vinculados ao itemNF passado ao ItemNF
                lErro = CF("RastreamentoMovto_Le_DocOrigem", objItemMovEstoque.lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, colRastreamentoMovto)
                If lErro <> SUCESSO Then gError 136981
                   
                If colRastreamentoMovto.Count = 1 Then
                
                    Set objRastreamento = colRastreamentoMovto.Item(1)
        
                    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col) = objRastreamento.sLote
                                
                    FilialOPRastro.Text = objRastreamento.iFilialOP
                                
                    'Valida a Filial
                    lErro = TP_FilialEmpresa_Le(FilialOPRastro.Text, objFilialEmpresa)
                    If lErro <> SUCESSO And lErro <> 71971 And lErro <> 71972 Then gError 139991
                    
                    FilialOPRastro.Text = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
                                
                    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col) = FilialOPRastro.Text
                                
                End If
                
            End If
                                
        End If
        '#####################################
        
        'Se necessário cria uma nova linha no Grid
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    Else

        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col) = ""
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantEmb_Col) = ""
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col) = ""
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoEmb_Col) = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 83169

    Saida_Celula_ItemNFRastro = SUCESSO

    Exit Function

Erro_Saida_Celula_ItemNFRastro:

    Saida_Celula_ItemNFRastro = gErr

    Select Case gErr
        
        Case 136975
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_INTERNA_SAIDA_NAO_CADASTRADA", gErr, objNFiscal.sSerie, objNFiscal.lNumNotaFiscal)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 130321
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMRASTRO_NAO_ITEMNF", gErr, iItem)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83169, 83207, 83250, 83251, 83346, 83349, 83351, 89515, 130322, 136978, 136981, 139991
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 83201, 83225
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMRASTRO_NAO_ITEMNF", gErr, iItem)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 83249
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_GRID_ITENS", gErr, iItem)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 83252
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 83253
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_RASTRO", gErr, objProduto.sCodigo, iItem)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83352
            Call Rotina_Erro(vbOKOnly, "ERRO_NFISCAL_MOVESTOQUE_INEXISTENTE", gErr, iItem, objAlmoxarifado.sNomeReduzido)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83353
            Call Rotina_Erro(vbOKOnly, "ERRO_SERIE_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 83354
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAEMISSAO_NAO_PREENCHIDA", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 83355
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMNOTAFISCAL_NAO_PREENCHIDO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 89516
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_NFISCAL_NAO_CADASTRADO", gErr, objTipoDocInfo.iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173199)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Sub BotaoEmbalagens_Click()

Dim objProdutoEmbalagem As New ClassProdutoEmbalagem
Dim colSelecao As New Collection
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim lErro As Long
Dim sSQL As String

On Error GoTo Erro_BotaoEmbalagens_Click

    'Verifica se o produto está preenchido
    If Len(Trim(GridEtiquetas.TextMatrix(GridEtiquetas.Row, iGrid_ProdutoEmb_Col))) < 0 Then gError 96144

    'Verifica se há uma linha selecionada
    If GridEtiquetas.Row = 0 Then gError 96145
        
    lErro = CF("Produto_Formata", GridEtiquetas.TextMatrix(GridEtiquetas.Row, iGrid_ProdutoEmb_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 96160

    sSQL = "Produto = '" & sProdutoFormatado & "'"
    
   'chama a tela de browser
    Call Chama_Tela("ProdutoEmbalagemLista", colSelecao, objProdutoEmbalagem, objEventoEmbalagens, sSQL)
    
    Exit Sub
    
Erro_BotaoEmbalagens_Click:
    
    Select Case gErr
    
        Case 96145
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 96144
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 96160
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173200)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoEmbalagens_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objEmbalagem As New ClassEmbalagem
Dim objProdutoEmbalagem As ClassProdutoEmbalagem
Dim iLinha As Integer
Dim iCodigoAux As Integer
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer, dQuantProduto As Double, dQuantEmb As Double

On Error GoTo Erro_objEventoEmbalagens_evSelecao
           
    'Define o tipo de obj recebido (Tipo ProdutoEmbalagem)
    Set objProdutoEmbalagem = obj1

    'Verifica se há alguma relacao Produto X Embalagem repetida no grid
    For iLinha = 1 To objGridEtiquetas.iLinhasExistentes
        
        If iLinha <> GridEtiquetas.Row Then
                            
            'Alteracao Daniel: devido ao fato de nao se ter mais o codigo na tela e sim a sigla _
            faz uma nova leitura em busca do codigo
            objEmbalagem.sSigla = GridEtiquetas.TextMatrix(iLinha, iGrid_Embalagem_Col)
            lErro = CF("Embalagem_Le_Sigla", objEmbalagem)
            If lErro <> SUCESSO And lErro <> 95088 Then gError 95468
            
            iCodigoAux = objEmbalagem.iCodigo
            
            lErro = CF("Produto_Formata", GridEtiquetas.TextMatrix(iLinha, iGrid_ProdutoEmb_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 96152

        End If
                       
    Next
    
    objEmbalagem.iCodigo = objProdutoEmbalagem.iEmbalagem
     
    'Lê a embalagem à partir do Codigo
    lErro = CF("Embalagem_Le", objEmbalagem)
    If lErro <> SUCESSO And lErro <> 25060 Then gError 96155

    lErro = CF("Produto_Formata", GridEtiquetas.TextMatrix(GridEtiquetas.Row, iGrid_ProdutoEmb_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 96152

    objProdutoEmbalagem.sProduto = sProdutoFormatado
            
    lErro = CF("ProdutoEmbalagem_Le", objProdutoEmbalagem)
    If lErro <> SUCESSO And lErro <> 96156 Then gError 96151
        
    'Preenche o grid de embalagens
    GridEtiquetas.TextMatrix(GridEtiquetas.Row, iGrid_Embalagem_Col) = objEmbalagem.sSigla
    
    Embalagem.Text = objEmbalagem.sSigla
    
    '####################################
    'Inserido por Wagner
    GridEtiquetas.TextMatrix(GridEtiquetas.Row, iGrid_PesoBruto_Col) = Formata_Estoque(objProdutoEmbalagem.dPesoBruto)
    GridEtiquetas.TextMatrix(GridEtiquetas.Row, iGrid_PesoLiq_Col) = Formata_Estoque(objProdutoEmbalagem.dPesoLiqTotal)
       
    Call Calcula_Pesos
    '####################################

    'Cria mais uma linha no grid
    If GridEtiquetas.Row - GridEtiquetas.FixedRows = objGridEtiquetas.iLinhasExistentes Then objGridEtiquetas.iLinhasExistentes = objGridEtiquetas.iLinhasExistentes + 1
    
    Me.Show
    
    Exit Sub
    
Erro_objEventoEmbalagens_evSelecao:

    Select Case gErr

        Case 81727, 96151, 96152, 96155
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173201)
              
    End Select
    
    Exit Sub

End Sub

Public Sub Embalagem_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridEtiquetas.objControle = Embalagem
    lErro = Grid_Campo_Libera_Foco(objGridEtiquetas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub Embalagem_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Embalagem_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridEtiquetas)

End Sub

Public Sub Embalagem_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridEtiquetas)

End Sub


Public Sub QuantEmb_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridEtiquetas.objControle = QuantEmb
    lErro = Grid_Campo_Libera_Foco(objGridEtiquetas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub QuantEmb_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub QuantEmb_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridEtiquetas)

End Sub

Public Sub QuantEmb_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridEtiquetas)

End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Rotina_Grid_Enable

    Select Case objControl.Name
        
        '??? completar
        
        Case ItemNFRastro.Name
    
            lErro = Testa_Control_Enable(objControl)
            If lErro <> SUCESSO And lErro <> 1 Then gError 83334
    
            If lErro = SUCESSO Then ItemNFRastro.Enabled = True
    
        Case LoteRastro.Name
            
            lErro = Testa_Control_Enable(objControl)
            If lErro <> SUCESSO And lErro <> 1 Then gError 83336
        
            If lErro = SUCESSO Then
            
                If Len(Trim(objGridEtiquetas.objGrid.TextMatrix(objGridEtiquetas.objGrid.Row, iGrid_ItemNFRastro_Col))) > 0 Then
                    LoteRastro.Enabled = True
                Else
                    LoteRastro.Enabled = False
                End If
            
            End If
            
        Case FilialOPRastro.Name

            lErro = Testa_Control_Enable(objControl)
            If lErro <> SUCESSO And lErro <> 1 Then gError 83338
        
            If lErro = SUCESSO Then

                lErro = CF("Produto_Formata", objGridEtiquetas.objGrid.TextMatrix(objGridEtiquetas.objGrid.Row, iGrid_ProdutoEmb_Col), sProdutoFormatado, iProdutoPreenchido)
                If lErro <> SUCESSO Then gError 83193
            
                If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
                    objProduto.sCodigo = sProdutoFormatado
            
                    'Lê o Produto
                    lErro = CF("Produto_Le", objProduto)
                    If lErro <> SUCESSO And lErro <> 28030 Then gError 83194
            
                    'Não achou o Produto
                    If lErro = 28030 Then gError 83195
            
                    If objProduto.iRastro = PRODUTO_RASTRO_OP Then
                        FilialOPRastro.Enabled = True
                    Else
                        FilialOPRastro.Enabled = False
                    End If
            
                End If

            End If
    
        Case Embalagem.Name
            
            lErro = Testa_Control_Enable(objControl)
            If lErro <> SUCESSO And lErro <> 1 Then gError 83336
        
            If lErro = SUCESSO Then
            
                If Len(Trim(objGridEtiquetas.objGrid.TextMatrix(objGridEtiquetas.objGrid.Row, iGrid_ProdutoEmb_Col))) = 0 Then
                    Embalagem.Enabled = False
                Else
                    Embalagem.Enabled = True
                End If
                
            End If
            
        Case QuantEmb.Name
            
            lErro = Testa_Control_Enable(objControl)
            If lErro <> SUCESSO And lErro <> 1 Then gError 83336
        
            If lErro = SUCESSO Then
            
                If Len(Trim(objGridEtiquetas.objGrid.TextMatrix(objGridEtiquetas.objGrid.Row, iGrid_Embalagem_Col))) > 0 Then
                    QuantEmb.Enabled = True
                Else
                    QuantEmb.Enabled = False
                End If
        
            End If
            
        Case PesoLiq.Name
            
            lErro = Testa_Control_Enable(objControl)
            If lErro <> SUCESSO And lErro <> 1 Then gError 83336
        
            If lErro = SUCESSO Then
            
                If Len(Trim(objGridEtiquetas.objGrid.TextMatrix(objGridEtiquetas.objGrid.Row, iGrid_ProdutoEmb_Col))) = 0 Then
                    PesoLiq.Enabled = False
                Else
                    PesoLiq.Enabled = True
                End If
                
            End If
            
        Case PesoBruto.Name
            
            lErro = Testa_Control_Enable(objControl)
            If lErro <> SUCESSO And lErro <> 1 Then gError 83336
        
            If lErro = SUCESSO Then
            
                If Len(Trim(objGridEtiquetas.objGrid.TextMatrix(objGridEtiquetas.objGrid.Row, iGrid_ProdutoEmb_Col))) = 0 Then
                    PesoBruto.Enabled = False
                Else
                    PesoBruto.Enabled = True
                End If
                
            End If
            
    End Select
    
    Exit Sub
     
Erro_Rotina_Grid_Enable:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173202)
     
    End Select
     
    Exit Sub

End Sub

Private Function Saida_Celula_LoteRastro(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_LoteRastro

    Set objGridInt.objControle = LoteRastro
        
    'Se o lote foi preenchido
    If Len(Trim(LoteRastro.Text)) > 0 Then
        
        'Se o produto não está preenchido ==> erro
        If Len(Trim(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoEmb_Col))) = 0 Then gError 83178
        
        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoEmb_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 83180
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        objProduto.sCodigo = sProdutoFormatado
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 83181
            
        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 83182
                
        'Se o Produto foi preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            objRastroLote.dtDataEntrada = DATA_NULA
            
            'Se o produto possuir rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
                
                For iLinha = 1 To objGridInt.iLinhasExistentes
                    'se se tratar do mesmo par itemnf/lote
                    If iLinha <> objGridInt.objGrid.Row Then
                        If objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col) And _
                           objGridInt.objGrid.TextMatrix(iLinha, iGrid_LoteRastro_Col) = LoteRastro.Text Then gError 83183
                    End If
                Next
                
                objRastroLote.sCodigo = LoteRastro.Text
                objRastroLote.sProduto = sProdutoFormatado
                
                'Lê o Rastreamento do Lote vinculado ao produto
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 83184
                
                'Se não encontrou --> Erro
                If lErro = 75710 Then gError 83185
                
            'Se o produto possuir rastro por OP
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                                                
                objRastroLote.sCodigo = LoteRastro.Text
                objRastroLote.sProduto = sProdutoFormatado
                If Len(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col)) = 0 Then
                    objRastroLote.iFilialOP = giFilialEmpresa
                Else
                    objRastroLote.iFilialOP = Codigo_Extrai(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col))
                End If
                
'                For iLinha = 1 To objGridInt.iLinhasExistentes
'                    If iLinha <> objGridInt.objGrid.Row Then
'                        'se o par itemnf/lote/op já está no grid
'                        If objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col) And _
'                           objGridInt.objGrid.TextMatrix(iLinha, iGrid_LoteRastro_Col) = LoteRastro.Text And Codigo_Extrai(objGridInt.objGrid.TextMatrix(iLinha, iGrid_FilialOPRastro_Col)) = objRastroLote.iFilialOP Then
'                            'se a OP ainda não estava preenchida
'                            If Len(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col)) = 0 Then
'                                objRastroLote.iFilialOP = 0
'                                objRastroLote.dtDataEntrada = DATA_NULA
'                            Else
'                                gError 83186
'                            End If
'                            Exit For
'                        End If
'                    End If
'                Next
                
                If objRastroLote.iFilialOP <> 0 Then
                
                    'Se o produto e Lote estão preenchidos verifica se o Produto pertence ao Lote
                    lErro = CF("RastreamentoLote_Le", objRastroLote)
                    If lErro <> SUCESSO And lErro <> 75710 Then gError 83187
                
                    'Se não encontrou --> Erro
                    If lErro = 75710 Then
                    
                        If Len(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col)) = 0 Then
                            objRastroLote.iFilialOP = 0
                        Else
                            gError 83188
                        End If
                    End If
                
'                    objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteDataRastro_Col) = Format(objRastroLote.dtDataEntrada, "dd/mm/yyyy")
        
                End If
        
            End If
    
            'Preenche campos do lote
            lErro = Lote_Saida_Celula(objGridInt, objRastroLote)
            If lErro <> SUCESSO Then gError 83189
    
        End If
    
    Else
    
'        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteDataRastro_Col) = ""
        
    End If
                                    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 83190

    Saida_Celula_LoteRastro = SUCESSO

    Exit Function

Erro_Saida_Celula_LoteRastro:

    Saida_Celula_LoteRastro = gErr

    Select Case gErr

        Case 83180, 83181, 83184, 83187, 83189, 83190
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83178
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_GRID", gErr, objGridInt.objGrid.Row)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83179
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_PREENCHIDO_GRID", gErr, objGridInt.objGrid.Row)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83182
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83185
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case 83188
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_FILIALOP_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto, objRastroLote.iFilialOP)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173203)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_FilialOPRastro(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim vbMsgRes As VbMsgBoxResult
Dim objOrdemProducao As New ClassOrdemDeProducao
Dim objFilialEmpresa As New AdmFiliais
Dim iLinha As Integer

On Error GoTo Erro_Saida_Celula_FilialOPRastro

    Set objGridInt.objControle = FilialOPRastro
        
    'Se a filial foi preenchida
    If Len(Trim(FilialOPRastro.Text)) > 0 Then
        
        'Valida a Filial
        lErro = TP_FilialEmpresa_Le(FilialOPRastro.Text, objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 71971 And lErro <> 71972 Then gError 83030

        'Se não for encontrado --> Erro
        If lErro = 71971 Then gError 83031
        If lErro = 71972 Then gError 83032
        
        If Len(objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col)) <> 0 Then
        
'            For iLinha = 1 To objGridInt.iLinhasExistentes
'                If iLinha <> objGridInt.objGrid.Row Then
'                    If objGridInt.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ItemNFRastro_Col) And _
'                       objGridInt.objGrid.TextMatrix(iLinha, iGrid_LoteRastro_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col) And Codigo_Extrai(objGridInt.objGrid.TextMatrix(iLinha, iGrid_FilialOPRastro_Col)) = objFilialEmpresa.iCodFilial Then gError 83033
'                End If
'            Next
        
            'Formata o produto
            lErro = CF("Produto_Formata", objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoEmb_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 83034
        
            objRastroLote.sCodigo = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col)
            objRastroLote.sProduto = sProdutoFormatado
            objRastroLote.iFilialOP = objFilialEmpresa.iCodFilial
                
            'Lê o Rastreamento do Lote vinculado ao produto
            lErro = CF("RastreamentoLote_Le", objRastroLote)
            If lErro <> SUCESSO And lErro <> 75710 Then gError 83035
                
            If lErro = 75710 Then gError 83264
                
            FilialOPRastro.Text = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
            
        Else
            
            FilialOPRastro.Text = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
            
        End If
        
    Else
    
        'GridEtiquetas.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteDataRastro_Col) = ""
        
    End If
                                    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 83037

    Saida_Celula_FilialOPRastro = SUCESSO

    Exit Function

Erro_Saida_Celula_FilialOPRastro:

    Saida_Celula_FilialOPRastro = gErr

    Select Case gErr

        Case 83030, 83034, 83035, 83037
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 83031, 83032
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, FilialOPRastro.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
'        Case 83033
'            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_FILIALOP_JA_UTILIZADO_GRID", gErr, objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col), objFilialEmpresa.iCodFilial)
'            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 83041
            Call Rotina_Erro(vbOKOnly, "ERRO_ITEMNF_NAO_SELECIONADO", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 83264
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_LOTE_PRODUTO_FILIALOP_INEXISTENTE", objRastroLote.sCodigo, objRastroLote.sProduto, objRastroLote.iFilialOP)

            If vbMsgRes = vbYes Then
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                Call Chama_Tela("RastreamentoLote", objRastroLote)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If
            
        Case 83374
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_FILIALOP_JA_UTILIZADO_GRID", gErr, objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col), objFilialEmpresa.iCodFilial)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173204)

    End Select

    Exit Function

End Function

Private Function Lote_Saida_Celula(objGridInt As AdmGrid, objRastroLote As ClassRastreamentoLote) As Long
'Executa a saida de celula do campo lote, o tratamento dos erros do Grid é feita na rotina chamadora

Dim lErro As Long
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Lote_Saida_Celula

    'Se a filial empresa foi preenchida
    If objRastroLote.iFilialOP <> 0 Then
        
        objFilialEmpresa.iCodFilial = objRastroLote.iFilialOP
        lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 27378 Then gError 83152

        'Se não encontrou a FilialEmpresa
        If lErro = 27378 Then gError 83153

        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
    
    End If
    
    If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
        objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
    End If
    
    Lote_Saida_Celula = SUCESSO
    
    Exit Function
        
Erro_Lote_Saida_Celula:

    Lote_Saida_Celula = gErr
    
    Select Case gErr
        
        Case 83152
        
        Case 83153
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALEMPRESA_NAO_CADASTRADA", gErr, objFilialEmpresa.iCodFilial)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173205)
    
    End Select
    
    Exit Function
    
End Function

Private Function Testa_Control_Enable(objControl As Object) As Long

Dim lErro As Long
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objProduto As New ClassProduto

On Error GoTo Erro_Testa_Control_Enable

    Testa_Control_Enable = SUCESSO

    If glNumIntNF = 0 Then
        objControl.Enabled = False
        Testa_Control_Enable = 1
    End If
    
    If Testa_Control_Enable = SUCESSO Then
    
        lErro = CF("Produto_Formata", objGridEtiquetas.objGrid.TextMatrix(objGridEtiquetas.objGrid.Row, iGrid_ProdutoEmb_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 92248
    
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
    
            objProduto.sCodigo = sProdutoFormatado
    
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 92246
    
            'Não achou o Produto
            If lErro = 28030 Then gError 92247
    
            If objProduto.iRastro = PRODUTO_RASTRO_NENHUM Then
                objControl.Enabled = False
                Testa_Control_Enable = 1
            End If
            
        End If
    
    End If
    
    Exit Function

Erro_Testa_Control_Enable:

    Testa_Control_Enable = gErr

    Select Case gErr

        Case 92246, 92248

        Case 92247
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173206)

    End Select

    Exit Function

End Function

'Tratamento do Grid
Private Sub GridEtiquetas_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridEtiquetas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridEtiquetas, iAlterado)
    End If

End Sub

Private Sub GridEtiquetas_EnterCell()

    Call Grid_Entrada_Celula(objGridEtiquetas, iAlterado)

End Sub

Private Sub GridEtiquetas_GotFocus()

    Call Grid_Recebe_Foco(objGridEtiquetas)

End Sub

Private Sub GridEtiquetas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridEtiquetas, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridEtiquetas, iAlterado)
    End If

End Sub

Private Sub GridEtiquetas_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridEtiquetas)

End Sub

Private Sub GridEtiquetas_RowColChange()

    Call Grid_RowColChange(objGridEtiquetas)

End Sub

Private Sub GridEtiquetas_Scroll()

    Call Grid_Scroll(objGridEtiquetas)

End Sub

Private Sub PesoBruto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub PesoBruto_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub PesoBruto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridEtiquetas)
End Sub
Private Sub PesoBruto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridEtiquetas)
End Sub
Private Sub PesoBruto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridEtiquetas.objControle = PesoBruto
    lErro = Grid_Campo_Libera_Foco(objGridEtiquetas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub PesoLiq_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub PesoLiq_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub
Private Sub PesoLiq_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridEtiquetas)
End Sub
Private Sub PesoLiq_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridEtiquetas)
End Sub
Private Sub PesoLiq_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridEtiquetas.objControle = PesoLiq
    lErro = Grid_Campo_Libera_Foco(objGridEtiquetas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function Saida_Celula_PesoBruto(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PesoBruto

    Set objGridInt.objControle = PesoBruto
    
    'Se estiver preenchida
    If Len(Trim(PesoBruto.Text)) > 0 Then
    
        'Critica o valor
        lErro = Valor_Positivo_Critica(PesoBruto.Text)
        If lErro <> SUCESSO Then gError 95117

        'Coloca o valor Formatado na tela
        PesoBruto.Text = Formata_Estoque(CDbl(PesoBruto.Text))

    End If
    
    'Abandona a celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 95118
    
    '#########################
    'Inserido por Wagner
    Call Calcula_Pesos
    '#########################
    
    Saida_Celula_PesoBruto = SUCESSO
    
    Exit Function

Erro_Saida_Celula_PesoBruto:

    Saida_Celula_PesoBruto = gErr
    
    Select Case gErr
    
        Case 95118, 95117
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173207)
    
    End Select
    
    Exit Function

End Function

Function Saida_Celula_PesoLiq(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_PesoLiq

    Set objGridInt.objControle = PesoLiq
    
    'Se estiver preenchida
    If Len(Trim(PesoLiq.Text)) > 0 Then
    
        'Critica o valor
        lErro = Valor_Positivo_Critica(PesoLiq.Text)
        If lErro <> SUCESSO Then gError 95117

        'Coloca o valor Formatado na tela
        PesoLiq.Text = Formata_Estoque(CDbl(PesoLiq.Text))

    End If
    
    'Abandona a celula
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 95118
    
    '#########################
    'Inserido por Wagner
    Call Calcula_Pesos
    '#########################
    
    Saida_Celula_PesoLiq = SUCESSO
    
    Exit Function

Erro_Saida_Celula_PesoLiq:

    Saida_Celula_PesoLiq = gErr
    
    Select Case gErr
    
        Case 95118, 95117
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173208)
    
    End Select
    
    Exit Function

End Function

Public Sub FilialOPRastro_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub FilialOPRastro_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridEtiquetas)

End Sub

Public Sub FilialOPRastro_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridEtiquetas)

End Sub

Public Sub FilialOPRastro_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridEtiquetas.objControle = FilialOPRastro
    lErro = Grid_Campo_Libera_Foco(objGridEtiquetas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub LoteRastro_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub LoteRastro_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridEtiquetas)

End Sub

Public Sub LoteRastro_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridEtiquetas)

End Sub

Public Sub LoteRastro_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridEtiquetas.objControle = LoteRastro
    lErro = Grid_Campo_Libera_Foco(objGridEtiquetas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ItemNFRastro_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ItemNFRastro_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridEtiquetas)

End Sub

Public Sub ItemNFRastro_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridEtiquetas)

End Sub

Public Sub ItemNFRastro_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridEtiquetas.objControle = ItemNFRastro
    lErro = Grid_Campo_Libera_Foco(objGridEtiquetas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub ProdutoEmb_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub ProdutoEmb_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridEtiquetas)

End Sub

Public Sub ProdutoEmb_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridEtiquetas)

End Sub

Public Sub ProdutoEmb_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridEtiquetas.objControle = ProdutoEmb
    lErro = Grid_Campo_Libera_Foco(objGridEtiquetas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Function Saida_Celula_Embalagem(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iProdutoPreenchido As Integer
Dim sProdutoFormatado As String
Dim objEmbalagem As New ClassEmbalagem
Dim objProdutoEmbalagem As New ClassProdutoEmbalagem
Dim vbMsg As VbMsgBoxResult, dQuantProduto As Double, dQuantEmb As Double
Dim iLinha As Integer, embText As Object

On Error GoTo Erro_Saida_Celula_Embalagem

    Set objGridInt.objControle = Embalagem

    'Se a Embalagem está preenchido
    If Len(Trim(Embalagem.Text)) > 0 Then
       
        If Embalagem.Text <> GridEtiquetas.TextMatrix(GridEtiquetas.Row, iGrid_ProdutoEmb_Col) Then
        
            'Le os dados da embalagem
            Set embText = Embalagem
            lErro = CF("TP_Embalagem_Le_Grid", embText, objEmbalagem)
            If lErro <> SUCESSO Then gError 96123
    
             objProdutoEmbalagem.iEmbalagem = objEmbalagem.iCodigo
                    
            lErro = CF("Produto_Formata", GridEtiquetas.TextMatrix(GridEtiquetas.Row, iGrid_ProdutoEmb_Col), sProdutoFormatado, iProdutoPreenchido)
            If lErro <> SUCESSO Then gError 96153
    
            objProdutoEmbalagem.sProduto = sProdutoFormatado
            
            lErro = CF("ProdutoEmbalagem_Le", objProdutoEmbalagem)
            If lErro <> SUCESSO And lErro <> 96156 Then gError 81728
            
            'Alteracao Daniel em 13/12/2001
            'Se nao está associado => limpa e gera erro
            If lErro = 96156 Then gError 96154
            
            'Preenche o grid
            Embalagem.Text = objEmbalagem.sSigla
            GridEtiquetas.TextMatrix(GridEtiquetas.Row, iGrid_Embalagem_Col) = objEmbalagem.sSigla
            
            'Se necessário cria uma nova linha no Grid
            If GridEtiquetas.Row - GridEtiquetas.FixedRows = objGridEtiquetas.iLinhasExistentes Then objGridEtiquetas.iLinhasExistentes = objGridEtiquetas.iLinhasExistentes + 1
            
        End If
    
    End If
        
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 96124

    Saida_Celula_Embalagem = SUCESSO

    Exit Function
    
Erro_Saida_Celula_Embalagem:

    Saida_Celula_Embalagem = gErr

    Select Case gErr
       
        Case 96123, 96124, 81727
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case 96153, 81728
        
        Case 96154
            Call Rotina_Erro(vbOKOnly, "AVISO_PRODUTOEMBALAGEM_INEXISTENTE", gErr)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173209)

    End Select

    Exit Function

End Function

Public Function Saida_Celula_QuantEmb(objGridInt As AdmGrid) As Long

Dim lErro As Long, dQuantEmb As Double, dQuantProduto As Double

On Error GoTo Erro_Saida_Celula_QuantEmb
    
    Set objGridInt.objControle = QuantEmb

    'Se a quantidade de embalagens foi preenchida
    If Len(Trim(QuantEmb.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Inteiro_Critica(QuantEmb.Text)
        If lErro <> SUCESSO Then gError 89485
       
        dQuantEmb = StrParaDbl(QuantEmb.Text)
            
        'Janaina
        'se alterar a quantidade de embalagens, desmarca a Check
        If dQuantEmb <> StrParaDbl(GridEtiquetas.TextMatrix(GridEtiquetas.Row, iGrid_QuantEmb_Col)) Then
            
            If dQuantEmb <> 0 Then
            
                GridEtiquetas.TextMatrix(GridEtiquetas.Row, iGrid_QuantEmb_Col) = Formata_Estoque(dQuantEmb)
                
            
            End If
            
        End If
        'Janaina
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 89487

    '#########################
    'Inserido por Wagner
    Call Calcula_Pesos
    '#########################
    
    Saida_Celula_QuantEmb = SUCESSO

    Exit Function

Erro_Saida_Celula_QuantEmb:

    Saida_Celula_QuantEmb = gErr

    Select Case gErr

        Case 89485, 89487, 81727
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 89486
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTDIST_MAIOR_QUANTITEMNF", gErr, QuantEmb.Text, objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantEmb_Col))
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173210)

    End Select

    Exit Function

End Function

Private Function Carrega_Serie() As Long
'Carrega a combo de Séries com as séries lidas do BD

Dim lErro As Long
Dim objSerie As ClassSerie
Dim colSerie As New colSerie

On Error GoTo Erro_Carrega_Serie

    'Lê as séries
    lErro = CF("Series_Le", colSerie)
    If lErro <> SUCESSO Then gError 42121

    'Carrega na combo
    For Each objSerie In colSerie
        Serie.AddItem objSerie.sSerie
    Next

    Carrega_Serie = SUCESSO

    Exit Function

Erro_Carrega_Serie:

    Carrega_Serie = gErr

    Select Case gErr

        Case 42121, 500115, 500116
        
        Case 500117

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173211)

    End Select

    Exit Function

End Function

Public Sub SerieLabel_Click()

Dim objSerie As New ClassSerie
Dim colSelecao As Collection

    'recolhe a serie da tela
    objSerie.sSerie = Serie.Text

    'Chama a Tela de Browse SerieLista
    Call Chama_Tela("SerieLista", colSelecao, objSerie, objEventoSerie)

    Exit Sub

End Sub

Sub Limpa_Tela_RelOpRotulosExpedicao()

    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridEtiquetas)
    
    Serie.Enabled = True
    NFiscal.Enabled = True
    
    Serie.ListIndex = -1
    Serie.SetFocus
    
    iAlterado = 0
    glNumIntNF = 0
    
End Sub

Public Sub BotaoLotes_Click()
'Chama a tela de Lote de Rastreamento

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim colSelecao As New Collection
Dim sSelecao As String
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim objRastroLoteSaldo As New ClassRastroLoteSaldo

On Error GoTo Erro_BotaoLotes_Click
    
    'Verifica se tem alguma linha selecionada no Grid
    If objGridEtiquetas.objGrid.Row = 0 Then gError 83146
        
    'Se o produto não foi preenchido, erro
    If Len(Trim(objGridEtiquetas.objGrid.TextMatrix(objGridEtiquetas.objGrid.Row, iGrid_ProdutoEmb_Col))) = 0 Then gError 83147
        
    'Formata o produto
    lErro = CF("Produto_Formata", objGridEtiquetas.objGrid.TextMatrix(objGridEtiquetas.objGrid.Row, iGrid_ProdutoEmb_Col), sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 83148
    
    'Lê o produto
    objProduto.sCodigo = sProdutoFormatado
    
    lErro = CF("Produto_Le", objProduto)
    If lErro <> SUCESSO And lErro <> 28030 Then gError 83149
    
    'Produto não cadastrado
    If lErro = 28030 Then gError 83150
        
    'Verifica o tipo de rastreamento do produto
    If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
        sSelecao = " FilialOP = ? AND Produto = ?"
    ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
        sSelecao = " FilialOP <> ? AND Produto = ?"
    End If
    
    'Adiciona filtros
    colSelecao.Add 0
    colSelecao.Add sProdutoFormatado
    
    objRastroLote.sProduto = sProdutoFormatado
    objRastroLote.iFilialOP = Codigo_Extrai(objGridEtiquetas.objGrid.TextMatrix(objGridEtiquetas.objGrid.Row, iGrid_FilialOPRastro_Col))
    objRastroLote.sCodigo = objGridEtiquetas.objGrid.TextMatrix(objGridEtiquetas.objGrid.Row, iGrid_LoteRastro_Col)
    
    'Chama a tela de browse RastroLoteLista passando como parâmetro a seleção do Filtro (sSelecao)
    Call Chama_Tela("RastroLoteLista", colSelecao, objRastroLote, objEventoLoteRastro, sSelecao)
                    
    Exit Sub

Erro_BotaoLotes_Click:

    Select Case gErr
        
        Case 83146
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
            
        Case 83147
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO", gErr)
        
        Case 83148, 83149
        
        Case 83150
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_CADASTRADO", gErr, objProduto.sCodigo)
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173212)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoLoteRastro_evSelecao(obj1 As Object)
'Rastreamento

Dim lErro As Long
Dim objRastroLote As New ClassRastreamentoLote
Dim objProduto As New ClassProduto
Dim sProdutoFormatado As String
Dim iProdutoPreenchido As Integer
Dim iLinha As Integer
Dim objRastroLoteSaldo As New ClassRastroLoteSaldo

On Error GoTo Erro_objEventoLoteRastro_evSelecao

    Set objRastroLote = obj1
    
    'Se a Linha corrente for diferente da Linha fixa
    If objGridEtiquetas.objGrid.Row <> 0 Then

        'Formata o Produto para o BD
        lErro = CF("Produto_Formata", objGridEtiquetas.objGrid.TextMatrix(objGridEtiquetas.objGrid.Row, iGrid_ProdutoEmb_Col), sProdutoFormatado, iProdutoPreenchido)
        If lErro <> SUCESSO Then gError 83154
            
        objProduto.sCodigo = sProdutoFormatado
                
        'Lê os demais atributos do Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 83155
            
        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 83156
                
        'Se o Produto foi preenchido
        If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
            
            'Se o produto possuir rastro por lote
            If objProduto.iRastro = PRODUTO_RASTRO_LOTE Then
                
                For iLinha = 1 To objGridEtiquetas.iLinhasExistentes
                    If iLinha <> objGridEtiquetas.objGrid.Row Then
                        If objGridEtiquetas.objGrid.TextMatrix(iLinha, iGrid_LoteRastro_Col) = objRastroLote.sCodigo And _
                           objGridEtiquetas.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col) = objGridEtiquetas.objGrid.TextMatrix(objGridEtiquetas.objGrid.Row, iGrid_ItemNFRastro_Col) Then gError 83157
                    End If
                Next
                
            'Se o produto possuir rastro por OP
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                                                
                For iLinha = 1 To objGridEtiquetas.iLinhasExistentes
                    If iLinha <> objGridEtiquetas.objGrid.Row Then
                        If objGridEtiquetas.objGrid.TextMatrix(iLinha, iGrid_LoteRastro_Col) = objRastroLote.sCodigo And _
                           Codigo_Extrai(objGridEtiquetas.objGrid.TextMatrix(iLinha, iGrid_FilialOPRastro_Col)) = objRastroLote.iFilialOP And objGridEtiquetas.objGrid.TextMatrix(iLinha, iGrid_ItemNFRastro_Col) = objGridEtiquetas.objGrid.TextMatrix(objGridEtiquetas.objGrid.Row, iGrid_ItemNFRastro_Col) Then gError 83158
                    End If
                Next
                
            End If

        End If

        'Coloca o Lote na tela
        objGridEtiquetas.objGrid.TextMatrix(objGridEtiquetas.objGrid.Row, iGrid_LoteRastro_Col) = objRastroLote.sCodigo
        LoteRastro.Text = objRastroLote.sCodigo
           
        'Preenche campos do lote
        lErro = Lote_Saida_Celula(objGridEtiquetas, objRastroLote)
        If lErro <> SUCESSO Then gError 83189
        
        
    End If

    Me.Show

    Exit Sub

Erro_objEventoLoteRastro_evSelecao:

    Select Case gErr

        Case 83154, 83155, 83159, 83189
        
        Case 83156
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_INEXISTENTE", gErr, objProduto.sCodigo)
        
        Case 83157
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_JA_UTILIZADO_GRID", gErr, objRastroLote.sCodigo)
            
        Case 83158
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_FILIALOP_JA_UTILIZADO_GRID", gErr, objRastroLote.sCodigo, objRastroLote.iFilialOP)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173213)

    End Select

    Exit Sub

End Sub

Public Sub GridEtiquetas_KeyDown(KeyCode As Integer, Shift As Integer)
'Rastreamento

Dim lErro As Long

On Error GoTo Erro_GridEtiquetas_KeyDown

    Call Grid_Trata_Tecla1(KeyCode, objGridEtiquetas)

    '#########################
    'Inserido por Wagner
    If KeyCode = vbKeyDelete Then
        Call Calcula_Pesos
    End If
    '#########################

    Exit Sub

Erro_GridEtiquetas_KeyDown:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173214)

    End Select
    
    Exit Sub

End Sub

Public Sub GridEtiquetas_LeaveCell()

    Call Saida_Celula(objGridEtiquetas)

End Sub

Public Sub NFiscal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(NFiscal, iAlterado)

End Sub

Function RelRotulosExpedicao_Prepara(ByVal colRotuloExpedicao As Collection, lNumIntRel As Long) As Long
'grava registros na tabela temporaria para a emissao dos rotulos de expedicao

Dim lErro As Long, lTransacao As Long, objRotuloExpedicao As ClassRelRotuloExpedicao
Dim objItemNF As ClassItemNF, iSeq As Integer, iEtiquetas As Integer
Dim lComando As Long, objRastroLote As ClassRastreamentoLote

On Error GoTo Erro_RelRotulosExpedicao_Prepara

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130323

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130324

    'obtem numintrel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_REL_ROTULOEXPEDICAO", lNumIntRel)
    If lErro <> SUCESSO Then gError 130325
    
    For Each objRotuloExpedicao In colRotuloExpedicao
    
        '###############################
        'Inserido por Wagner
        If objRotuloExpedicao.iImprimir = MARCADO Then
        '###############################
        
            Set objItemNF = New ClassItemNF
                
            With objRotuloExpedicao
                objItemNF.iItem = .iItem
                objItemNF.sSerieNFOrig = .sSerie
                objItemNF.lNumNFOrig = .lNumNotaFiscal
            End With
            
            lErro = CF("ItemNFiscalSaida_Le_NumNFItem", objItemNF)
            If lErro <> SUCESSO Then gError 130326
            
            Set objRastroLote = New ClassRastreamentoLote
            With objRastroLote
                .iFilialOP = objRotuloExpedicao.iFilialOP
                .sProduto = objRotuloExpedicao.sProduto
                .sCodigo = objRotuloExpedicao.sLote
            End With
            
            'Lê o Rastreamento do Lote vinculado ao produto
            lErro = CF("RastreamentoLote_Le", objRastroLote)
            If lErro <> SUCESSO And lErro <> 75710 Then gError 130327
            
            'Se não encontrou --> Erro
            If lErro = 75710 Then gError 130328
            
            For iEtiquetas = 1 To objRotuloExpedicao.iQtdeEmb
            
                iSeq = iSeq + 1
                
                lErro = Comando_Executar(lComando, "INSERT INTO RelRotuloExpedicao (NumIntRel, Seq, NumIntRastreamentoLote, PesoLiquido, PesoBruto, NumIntItemNF) VALUES (?,?,?,?,?,?)", _
                    lNumIntRel, iSeq, objRastroLote.lNumIntDoc, objRotuloExpedicao.dPesoLiquido, objRotuloExpedicao.dPesoBruto, objItemNF.lNumIntDoc)
                If lErro <> AD_SQL_SUCESSO Then gError 130329
            
            Next
        
        '###############################
        'Inserido por Wagner
        End If
        '###############################
        
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130030
    
    Call Comando_Fechar(lComando)
    
    RelRotulosExpedicao_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelRotulosExpedicao_Prepara:

    RelRotulosExpedicao_Prepara = gErr
     
    Select Case gErr
          
        Case 130325, 130326, 130327
        
        Case 130329
            Call Rotina_Erro(vbOKOnly, "ERRO_GRAVACAO_RELROTULOEXPEDICAO", gErr)
        
        Case 130328
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_RASTRO_INEXISTENTE", gErr)
        
        Case 130323
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130024
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130030
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173215)
     
    End Select
     
    Call Transacao_Rollback
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'#####################################################
'Inserido por Wagner
Function Preenche_Tela_NF(ByVal objNF As ClassNFiscal) As Long
'Preenche o Grid de Itens

Dim lErro As Long
Dim objItemNF As ClassItemNF
Dim objRastreamento As ClassRastreamentoMovto
Dim sProdutoMascarado As String
Dim objProdutoEmbalagem As ClassProdutoEmbalagem
Dim objEmbalagem As ClassEmbalagem
Dim iQtdEmb As Integer
Dim iLinha As Integer
Dim objItemMovEstoque As ClassItemMovEstoque
Dim colRastreamentoMovto As Collection
Dim colMovEstoque As Collection
Dim objAlmoxarifado As ClassAlmoxarifado
Dim objFilialEmpresa As AdmFiliais

On Error GoTo Erro_Preenche_Tela_NF
    
    'Tenta lêr a Nota com esses dados
    lErro = CF("NFiscalInternaSaida_Le_Numero", objNF)
    If lErro <> SUCESSO And lErro <> 30765 Then gError 136974
    If lErro = 30765 Then gError 136975 'Não encontrou
    
    'Lê os Ítens da Nota Fiscal
    lErro = CF("NFiscalItens_Le", objNF)
    If lErro <> SUCESSO Then gError 136977
    
    glNumIntNF = objNF.lNumIntDoc

    Call Grid_Limpa(objGridEtiquetas)

    Serie.Text = objNF.sSerie
    NFiscal.Text = CStr(objNF.lNumNotaFiscal)

    For Each objItemNF In objNF.ColItensNF
               
        Set objItemMovEstoque = New ClassItemMovEstoque
        Set colMovEstoque = New Collection
        
        lErro = Mascara_RetornaProdutoTela(objItemNF.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 136972
           
        'Lê item de movimento de estoque
        objItemMovEstoque.lNumIntDocOrigem = objItemNF.lNumIntDoc
        objItemMovEstoque.iTipoNumIntDocOrigem = TIPO_ORIGEM_ITEMNF
        objItemMovEstoque.iFilialEmpresa = objNF.iFilialEmpresa
                
        lErro = CF("MovEstoque_Le_ItemNF", objItemMovEstoque, colMovEstoque)
        If lErro <> SUCESSO Then gError 136978
        
        For Each objItemMovEstoque In colMovEstoque
        
            'se forem os movimentos de transferencia de material consignado ==> não trata-os, pois a venda de material consignado está sendo tratada
            If objItemMovEstoque.iTipoMov <> MOV_EST_SAIDA_TRANSF_CONSIG_TERC And _
               objItemMovEstoque.iTipoMov <> MOV_EST_ENTRADA_TRANSF_DISP1 Then
        
                Set colRastreamentoMovto = New Collection
                Set objAlmoxarifado = New ClassAlmoxarifado
        
                'Lê o Almoxarifado
                objAlmoxarifado.iCodigo = objItemMovEstoque.iAlmoxarifado
                lErro = CF("Almoxarifado_Le", objAlmoxarifado)
                If lErro <> SUCESSO And lErro <> 25056 Then gError 136979
        
                'Se não encontrou Almoxarifado --> Erro
                If lErro = 25056 Then gError 136980
                        
                'Lê movimentos de rastreamento vinculados ao itemNF passado ao ItemNF
                lErro = CF("RastreamentoMovto_Le_DocOrigem", objItemMovEstoque.lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, colRastreamentoMovto)
                If lErro <> SUCESSO Then gError 136981
                   
                For Each objRastreamento In colRastreamentoMovto
                
                    iLinha = iLinha + 1

                    Set objFilialEmpresa = New AdmFiliais
                    Set objProdutoEmbalagem = New ClassProdutoEmbalagem
                    Set objEmbalagem = New ClassEmbalagem
                                
                    GridEtiquetas.TextMatrix(iLinha, iGrid_ItemNFRastro_Col) = objItemNF.iItem
                    GridEtiquetas.TextMatrix(iLinha, iGrid_LoteRastro_Col) = objRastreamento.sLote
                    GridEtiquetas.TextMatrix(iLinha, iGrid_ProdutoEmb_Col) = sProdutoMascarado
                                
                    If objRastreamento.iFilialOP <> 0 Then
                    
                        FilialOPRastro.Text = objRastreamento.iFilialOP
                                            
                        'Valida a Filial
                        lErro = TP_FilialEmpresa_Le(FilialOPRastro.Text, objFilialEmpresa)
                        If lErro <> SUCESSO And lErro <> 71971 And lErro <> 71972 Then gError 139991
                        
                        FilialOPRastro.Text = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
                        
                    Else
                    
                        FilialOPRastro.Text = ""
                        
                    End If
                    GridEtiquetas.TextMatrix(iLinha, iGrid_FilialOPRastro_Col) = FilialOPRastro.Text
                    
                    objProdutoEmbalagem.sProduto = objItemNF.sProduto
                        
                    'Seleciona embalagem padrao
                    lErro = CF("ProdutoEmbalagem_Le_Padrao", objProdutoEmbalagem)
                    If lErro <> SUCESSO Then gError 136970
                
                    objEmbalagem.iCodigo = objProdutoEmbalagem.iEmbalagem
                
                    lErro = CF("Embalagem_Le", objEmbalagem)
                    If lErro <> SUCESSO Then gError 136971
                
                    GridEtiquetas.TextMatrix(iLinha, iGrid_Embalagem_Col) = objEmbalagem.sSigla
                    GridEtiquetas.TextMatrix(iLinha, iGrid_PesoBruto_Col) = Formata_Estoque(objProdutoEmbalagem.dPesoBruto)
                    GridEtiquetas.TextMatrix(iLinha, iGrid_PesoLiq_Col) = Formata_Estoque(objProdutoEmbalagem.dPesoLiqTotal)
                    GridEtiquetas.TextMatrix(iLinha, iGrid_Imprimir_Col) = MARCADO
                
                Next
                
            End If
            
        Next
    
    Next
    
    Call Grid_Refresh_Checkbox(objGridEtiquetas)

    objGridEtiquetas.iLinhasExistentes = iLinha
    
    Preenche_Tela_NF = SUCESSO

    Exit Function

Erro_Preenche_Tela_NF:

    Preenche_Tela_NF = gErr

    Select Case gErr
    
        Case 136970 To 136972, 136974, 136979, 136981
        
        Case 136975
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_INTERNA_SAIDA_NAO_CADASTRADA", gErr, objNF.sSerie, objNF.lNumNotaFiscal)
        
        Case 136980
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", gErr, objAlmoxarifado.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173216)

    End Select

    Exit Function

End Function

Private Function Critica_GridEtiquetas(ByVal colRotuloExpedicao As Collection) As Long
'Critica os parâmetros que serão passados para o relatório

Dim lErro As Long
Dim objRotuloExpedicao As ClassRelRotuloExpedicao
Dim iLinha As Integer
Dim objNF As New ClassNFiscal
Dim vbMsg As VbMsgBoxResult
Dim iContImprimir As Integer

On Error GoTo Erro_Critica_GridEtiquetas

    If colRotuloExpedicao.Count = 0 Then gError 136993

    objNF.lNumNotaFiscal = StrParaLong(NFiscal.Text)
    objNF.sSerie = Serie.Text

    'Tenta lêr a Nota com esses dados
    lErro = CF("NFiscalInternaSaida_Le_Numero", objNF)
    If lErro <> SUCESSO And lErro <> 30765 Then gError 136988
    If lErro = 30765 Then gError 136989 'Não encontrou
    
    iContImprimir = 0
    
    For Each objRotuloExpedicao In colRotuloExpedicao
    
        iLinha = iLinha + 1
        
        If objRotuloExpedicao.iQtdeEmb = 0 Then gError 136992
        If objRotuloExpedicao.sLote = "" Then gError 136994
        If objRotuloExpedicao.sProduto = "" Then gError 136995
        
        If objRotuloExpedicao.dPesoBruto < objRotuloExpedicao.dPesoLiquido Then gError 136986

        If objRotuloExpedicao.iImprimir = MARCADO Then iContImprimir = iContImprimir + 1

    Next
       
    'Se o peso líquido está preenchido na nota
    If objNF.dPesoLiq > QTDE_ESTOQUE_DELTA Then
        If Abs(objNF.dPesoLiq - StrParaDbl(PesoLTotal.Caption)) > QTDE_ESTOQUE_DELTA Then
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_PESO_LIQUIDO_DIFERENTE", PesoLTotal.Caption, Formata_Estoque(objNF.dPesoLiq))
            If vbMsg = vbNo Then gError 136987
        End If
    End If
    
    'Se o peso bruto está preenchido na Nota
    If objNF.dPesoBruto > QTDE_ESTOQUE_DELTA Then
        If Abs(objNF.dPesoBruto - StrParaDbl(PesoBTotal.Caption)) > QTDE_ESTOQUE_DELTA Then
            vbMsg = Rotina_Aviso(vbYesNo, "AVISO_PESO_BRUTO_DIFERENTE", PesoBTotal.Caption, Formata_Estoque(objNF.dPesoBruto))
            If vbMsg = vbNo Then gError 136990
        End If
    End If
    
    'Se não tem nenhum item a imprimir
    If iContImprimir = 0 Then gError 136997

    Critica_GridEtiquetas = SUCESSO

    Exit Function

Erro_Critica_GridEtiquetas:

    Critica_GridEtiquetas = gErr

    Select Case gErr
    
        Case 136986
            Call Rotina_Erro(vbOKOnly, "ERRO_PESO_LIQUIDO_MAIOR_BRUTO_GRID", gErr, Formata_Estoque(objRotuloExpedicao.dPesoLiquido), Formata_Estoque(objRotuloExpedicao.dPesoBruto), iLinha)
        
        Case 136987, 136988, 136990
        
        Case 136989
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_INTERNA_SAIDA_NAO_CADASTRADA", gErr, objNF.sSerie, objNF.lNumNotaFiscal)

        Case 136992
            Call Rotina_Erro(vbOKOnly, "ERRO_QTD_EMBALAGEM_NAO_PREENCHIDO", gErr, iLinha)
        
        Case 136993
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_NAO_PREENCHIDO1", gErr, iLinha)
        
        Case 136994
            Call Rotina_Erro(vbOKOnly, "ERRO_GRID_LOTE_NAO_PREENCHIDO", gErr, iLinha)
        
        Case 136995
            Call Rotina_Erro(vbOKOnly, "ERRO_FALTA_PRODUTO_GRID", gErr, iLinha)

        Case 136997
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPRESSAO_SEM_ITENS_MARCADOS", gErr, iLinha)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173217)

    End Select

    Exit Function

End Function

Public Sub Imprimir_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Imprimir_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridEtiquetas)

End Sub

Public Sub Imprimir_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridEtiquetas)

End Sub

Public Sub Imprimir_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridEtiquetas.objControle = Imprimir
    lErro = Grid_Campo_Libera_Foco(objGridEtiquetas)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Calcula_Pesos()

Dim iLinha As Integer
Dim dPesoL As Double
Dim dPesoB As Double

    For iLinha = 1 To objGridEtiquetas.iLinhasExistentes
    
        dPesoL = dPesoL + (StrParaDbl(GridEtiquetas.TextMatrix(iLinha, iGrid_PesoLiq_Col)) * StrParaInt(GridEtiquetas.TextMatrix(iLinha, iGrid_QuantEmb_Col)))
        dPesoB = dPesoB + (StrParaDbl(GridEtiquetas.TextMatrix(iLinha, iGrid_PesoBruto_Col)) * StrParaInt(GridEtiquetas.TextMatrix(iLinha, iGrid_QuantEmb_Col)))
        
    Next
    
    PesoBTotal.Caption = Formata_Estoque(dPesoB)
    PesoLTotal.Caption = Formata_Estoque(dPesoL)

End Sub

Private Sub BotaoMarcarTodos_Click()
    Call Imprimir_Marca_Desmarca(MARCADO)
End Sub

Private Sub BotaoDesmarcarTodos_Click()
    Call Imprimir_Marca_Desmarca(DESMARCADO)
End Sub

Private Sub Imprimir_Marca_Desmarca(ByVal iFlag As Integer)

Dim iLinha As Integer

    For iLinha = 1 To objGridEtiquetas.iLinhasExistentes
        GridEtiquetas.TextMatrix(iLinha, iGrid_Imprimir_Col) = iFlag
    Next

    Call Grid_Refresh_Checkbox(objGridEtiquetas)

End Sub
'#####################################################

