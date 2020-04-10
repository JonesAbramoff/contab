VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl RelOpRotulosProducaoOcx 
   ClientHeight    =   5595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9555
   KeyPreview      =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   9555
   Begin VB.Frame FrameOP 
      Caption         =   "Informações Complementares"
      Height          =   765
      Left            =   75
      TabIndex        =   35
      Top             =   4725
      Visible         =   0   'False
      Width           =   5790
      Begin MSComCtl2.UpDown UpDownValidade 
         Height          =   300
         Left            =   5460
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   300
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataValidade 
         Height          =   300
         Left            =   4365
         TabIndex        =   10
         Top             =   300
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownFabricacao 
         Height          =   300
         Left            =   2700
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFabricacao 
         Height          =   300
         Left            =   1605
         TabIndex        =   8
         Top             =   315
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Data Fabricação:"
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
         Left            =   105
         TabIndex        =   37
         Top             =   375
         Width           =   1485
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Data Validade:"
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
         Left            =   3075
         TabIndex        =   36
         Top             =   345
         Width           =   1275
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Ordem de Produção"
      Height          =   765
      Index           =   1
      Left            =   3210
      TabIndex        =   33
      Top             =   90
      Width           =   3060
      Begin VB.CommandButton BotaoExibirDadosOP 
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
         Left            =   1680
         TabIndex        =   3
         Top             =   270
         Width           =   1290
      End
      Begin MSMask.MaskEdBox CodigoOP 
         Height          =   300
         Left            =   855
         TabIndex        =   2
         Top             =   285
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin VB.Label CodigoOPLabel 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   165
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   34
         Top             =   315
         Width           =   660
      End
   End
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
      Height          =   390
      Left            =   5940
      TabIndex        =   12
      Top             =   5070
      Width           =   2265
   End
   Begin VB.Frame Frame6 
      Caption         =   "Produção Entrada"
      Height          =   765
      Index           =   0
      Left            =   75
      TabIndex        =   25
      Top             =   90
      Width           =   3060
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
         Left            =   1665
         TabIndex        =   1
         Top             =   270
         Width           =   1290
      End
      Begin MSMask.MaskEdBox Codigo 
         Height          =   300
         Left            =   855
         TabIndex        =   0
         Top             =   285
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   6
         Mask            =   "######"
         PromptChar      =   " "
      End
      Begin VB.Label CodigoLabel 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   150
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   26
         Top             =   330
         Width           =   660
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
      Height          =   375
      Left            =   8370
      TabIndex        =   13
      Top             =   5085
      Width           =   1080
   End
   Begin VB.Frame Frame18 
      Caption         =   "Definição das Etiquetas"
      Height          =   3795
      Left            =   75
      TabIndex        =   18
      Top             =   900
      Width           =   9390
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
         Left            =   3345
         TabIndex        =   7
         Top             =   3135
         Width           =   1425
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   570
         Left            =   135
         Picture         =   "RelOPRotulosProducaoOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3135
         Width           =   1425
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   570
         Left            =   1740
         Picture         =   "RelOPRotulosProducaoOcx.ctx":101A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3135
         Width           =   1425
      End
      Begin VB.CheckBox Imprimir 
         Caption         =   "Imprimir"
         Height          =   225
         Left            =   1860
         TabIndex        =   28
         Top             =   2340
         Width           =   1035
      End
      Begin MSMask.MaskEdBox LoteRastro 
         Height          =   225
         Left            =   1845
         TabIndex        =   19
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
         TabIndex        =   20
         Top             =   780
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
         TabIndex        =   21
         Top             =   675
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
         TabIndex        =   22
         Top             =   1590
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
         Left            =   4635
         TabIndex        =   23
         Top             =   1155
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
         TabIndex        =   24
         Top             =   1320
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
         TabIndex        =   27
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
         Left            =   120
         TabIndex        =   4
         Top             =   210
         Width           =   9120
         _ExtentX        =   16087
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
         TabIndex        =   32
         Top             =   3405
         Width           =   1470
      End
      Begin VB.Label PesoBTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7260
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   3450
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
         TabIndex        =   29
         Top             =   3075
         Width           =   1080
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8325
      ScaleHeight     =   495
      ScaleMode       =   0  'User
      ScaleWidth      =   1065
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   150
      Width           =   1125
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   75
         Picture         =   "RelOPRotulosProducaoOcx.ctx":21FC
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   585
         Picture         =   "RelOPRotulosProducaoOcx.ctx":272E
         Style           =   1  'Graphical
         TabIndex        =   16
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
      Left            =   6435
      Picture         =   "RelOPRotulosProducaoOcx.ctx":28AC
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   165
      Width           =   1845
   End
End
Attribute VB_Name = "RelOpRotulosProducaoOcx"
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

Private glCodigo As Long
Dim lCodigoAntigo As Long
Private gsCodigoOP As String
Dim sCodigoOPAntigo As String

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1
Private WithEvents objEventoCodigoOP As AdmEvento
Attribute objEventoCodigoOP.VB_VarHelpID = -1
Private WithEvents objEventoEmbalagens As AdmEvento
Attribute objEventoEmbalagens.VB_VarHelpID = -1
Private WithEvents objEventoLoteRastro As AdmEvento
Attribute objEventoLoteRastro.VB_VarHelpID = -1

Public objGridEtiquetas As AdmGrid

Private iGrid_Imprimir_Col As Integer 'Inserido por Wagner
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

    glCodigo = 0
    gsCodigoOP = 0
    
    Set objEventoCodigo = New AdmEvento
    Set objEventoCodigoOP = New AdmEvento
    Set objEventoEmbalagens = New AdmEvento
    Set objEventoLoteRastro = New AdmEvento
    
    Set objGridEtiquetas = New AdmGrid
    
    'Inicializa Máscara de ProdutoEmb
    lErro = CF("Inicializa_Mascara_Produto_MaskEd", ProdutoEmb)
    If lErro <> SUCESSO Then gError 78154
    
    lErro = Inicializa_Grid_Etiquetas(objGridEtiquetas)
    If lErro <> SUCESSO Then gError 130319
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 130319, 78154
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173218)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)
  
    Set objEventoCodigo = Nothing
    Set objEventoCodigoOP = Nothing
    Set objEventoEmbalagens = Nothing
    Set objEventoLoteRastro = Nothing

    Set objGridEtiquetas = Nothing
    
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes, Optional obj1 As Object) As Long

Dim lErro As Long
Dim bEncontrou As Boolean
Dim iIndice As Integer
Dim objOp As ClassOrdemDeProducao
Dim objMovEstoque As ClassMovEstoque

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 122625
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes
    
    '#######################################
    'Inserido por Wagner 16/01/2006
    If TypeName(obj1) = "ClassMovEstoque" Then
        Set objMovEstoque = obj1
    End If
    If TypeName(obj1) = "ClassOrdemDeProducao" Then
        Set objOp = obj1
    End If
    '#######################################

    If Not (objMovEstoque Is Nothing) Then
            
        '##############################
        'Inserido por Wagner
        lCodigoAntigo = objMovEstoque.lCodigo
        
        lErro = Preenche_Tela_MovEstoque(objMovEstoque)
        If lErro <> SUCESSO Then gError 136973
        '##############################
    
    '##############################
    'Inserido por Wagner 16/01/2006
        FrameOP.Visible = False
    Else
   
        If Not (objOp Is Nothing) Then
        
            sCodigoOPAntigo = objOp.sCodigo
        
            lErro = Preenche_Tela_OP(objOp)
            If lErro <> SUCESSO Then gError 141505
            '##############################
            
            FrameOP.Visible = True
        End If
    '##############################
    
    End If
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 122625
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case 136973, 141505
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173219)

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
    'If Len(Trim(Codigo.Text)) = 0 Then gError 122629
    
    Critica_Parametros = SUCESSO

    Exit Function

Erro_Critica_Parametros:

    Critica_Parametros = gErr

    Select Case gErr

        Case 122629
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173220)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

   Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    Call Limpa_Tela_RelOpRotulosProducao
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173221)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional bExecutar As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long, colRotuloProducao As New Collection, iLinha As Integer
Dim iIndice As Integer, lNumIntRel As Long, objRotuloProducao As ClassRelRotuloProducao
Dim iPreenchido As Integer, sProdutoFormatado As String, sProdutoMascarado As String
Dim objRastroLote As ClassRastreamentoLote

On Error GoTo Erro_PreencherRelOp

    lErro = Critica_Parametros()
    If lErro <> SUCESSO Then gError 122634
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 122635
    
    lErro = objRelOpcoes.IncluirParametro("NCODIGO", Codigo.Text)
    If lErro <> AD_BOOL_TRUE Then gError 122636
   
    If bExecutar Then
    
        For iLinha = 1 To objGridEtiquetas.iLinhasExistentes
        
            'validar linha
            '??? obrigar o preenchimento de pesos, lote e qtde de etiquetas
            
            Set objRotuloProducao = New ClassRelRotuloProducao
            
            With objRotuloProducao
                            
                .dPesoLiquido = StrParaDbl(GridEtiquetas.TextMatrix(iLinha, iGrid_PesoLiq_Col))
                .dPesoBruto = StrParaDbl(GridEtiquetas.TextMatrix(iLinha, iGrid_PesoBruto_Col))
                .iFilialOP = Codigo_Extrai(GridEtiquetas.TextMatrix(iLinha, iGrid_FilialOPRastro_Col))
                .sLote = GridEtiquetas.TextMatrix(iLinha, iGrid_LoteRastro_Col)
                .iQtdeEmb = StrParaInt(GridEtiquetas.TextMatrix(iLinha, iGrid_QuantEmb_Col))
                .lCodigo = StrParaLong(Codigo.Text)
                
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

            objRotuloProducao.sProduto = sProdutoFormatado
                                                    
                                                    
            '################################################
            'Inserido por Wagner 16/09/2006
            'Se os dados são referentes a uma OP
            If FrameOP.Visible = True Then
                objRotuloProducao.dtDataFabricacao = StrParaDate(DataFabricacao.Text)
                objRotuloProducao.dtDataValidade = StrParaDate(DataValidade.Text)
                objRotuloProducao.lNumIntRastreamentoLote = 0
            Else
                Set objRastroLote = New ClassRastreamentoLote
                
                objRastroLote.sProduto = objRotuloProducao.sProduto
                objRastroLote.iFilialOP = objRotuloProducao.iFilialOP
                objRastroLote.sCodigo = objRotuloProducao.sLote
                
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 141500
            
                objRotuloProducao.dtDataFabricacao = objRastroLote.dtDataFabricacao
                objRotuloProducao.dtDataValidade = objRastroLote.dtDataValidade
                objRotuloProducao.lNumIntRastreamentoLote = objRastroLote.lNumIntDoc

            End If
            '################################################
                                                    
            colRotuloProducao.Add objRotuloProducao
            
        Next
        
        '###################################
        'Inserido por Wagner
        lErro = Critica_GridEtiquetas(colRotuloProducao)
        If lErro <> SUCESSO Then gError 136985
        '###################################
        
        '??? colocar como CF
        'obter numintrel e criar registros em tabela auxiliar
        lErro = RelRotulosProducao_Prepara(colRotuloProducao, lNumIntRel)
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173222)

    End Select

    Exit Function

End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long

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

        Case 122648, 122653 'Tratado na Rotina chamada

        Case 122654 'Cancelou o relatório
            Unload Me

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173223)

    End Select

    Exit Sub

End Sub

Private Sub CodigoLabel_Click()

Dim objMovEstoque As New ClassMovEstoque
Dim colSelecao As New Collection

    If Len(Trim(Codigo.ClipText)) <> 0 Then objMovEstoque.lCodigo = CLng(Codigo.Text)

    colSelecao.Add MOV_EST_PRODUCAO
    colSelecao.Add MOV_EST_PRODUCAO_BENEF3
    
    Call Chama_Tela("MovEstoqueOPLista", colSelecao, objMovEstoque, objEventoCodigo)
    
End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long, iIndice As Integer
Dim objMovEstoque As New ClassMovEstoque
Dim vbMsg As VbMsgBoxResult

On Error GoTo Erro_Codigo_Validate

    'se o codigo foi trocado
    If lCodigoAntigo <> StrParaLong(Trim(Codigo.Text)) Then

        'se o codigo novo está preenchido
        If Len(Trim(Codigo.ClipText)) > 0 Then
                    
            objMovEstoque.lCodigo = Codigo.Text
            
            'Le o Movimento de Estoque e Verifica se ele já foi estornado
            lErro = CF("MovEstoqueItens_Le_Verifica_Estorno", objMovEstoque, MOV_EST_PRODUCAO)
            If lErro <> SUCESSO And lErro <> 78883 And lErro <> 78885 Then gError 34776
            
            'Se todos os Itens do Movimento foram estornados
            If lErro = 78885 Then gError 78890
            
            If lErro = SUCESSO Then
            
                If objMovEstoque.iTipoMov <> MOV_EST_PRODUCAO Then gError 34898
                        
            End If
        
        End If

        lCodigoAntigo = StrParaLong(Trim(Codigo.Text))
        
        Call Grid_Limpa(objGridEtiquetas)
        
   End If

   Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 34776, 34907
    
        Case 34898
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INCOMPATIVEL_PENTRADA", gErr, objMovEstoque.lCodigo)
            lCodigoAntigo = 0

        Case 55324
                
        Case 78890
            Call Rotina_Erro(vbOKOnly, "ERRO_MOVIMENTOESTOQUE_ESTORNADO", gErr, giFilialEmpresa, objMovEstoque.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173224)

    End Select

    Exit Sub


End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_EMISSAO_NF
    Set Form_Load_Ocx = Me
    Caption = "Rótulos de Produção"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpRotulosProducao"
    
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
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Codigo Then
            Call CodigoLabel_Click
        ElseIf Me.ActiveControl Is CodigoOP Then
            Call CodigoOPLabel_Click
        ElseIf Me.ActiveControl Is Embalagem Then
            Call BotaoEmbalagens_Click
        ElseIf Me.ActiveControl Is LoteRastro Then
            Call BotaoLotes_Click
        End If
        
    End If

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim objMovEstoque As ClassMovEstoque
Dim lErro As Long

On Error GoTo Erro_objCodigoEvento_evSelecao

    Set objMovEstoque = obj1
    
    FrameOP.Visible = False 'Inserido por Wagner

    Codigo.Text = objMovEstoque.lCodigo
    Call BotaoExibirDados_Click
        
    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)
    
    Me.Show

    Exit Sub

Erro_objCodigoEvento_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173225)

    End Select

    Exit Sub

End Sub

Public Sub BotaoExibirDados_Click()

Dim lErro As Long
Dim objMovEstoque As New ClassMovEstoque

On Error GoTo Erro_BotaoExibirDados_Click

    'Verifica se a Serie e o Número da Nota Fiscal original estão preenchidos
    If Len(Trim(Codigo.Text)) = 0 Then gError 35256

    FrameOP.Visible = False 'Inserido por Wagner

    objMovEstoque.lCodigo = StrParaLong(Codigo.Text)
    objMovEstoque.iFilialEmpresa = giFilialEmpresa
    
    lErro = CF("MovEstoque_Le", objMovEstoque)
    If lErro <> SUCESSO And lErro <> 30128 Then gError 30981

    'Lê os ítens do Movimento de Estoque
    lErro = CF("MovEstoqueItens_Le1", objMovEstoque, MOV_EST_PRODUCAO)
    If lErro <> SUCESSO And lErro <> 55387 Then gError 30981

    'Coloca na tela os dados encontrados
    lErro = Preenche_Tela_MovEstoque(objMovEstoque)
    If lErro <> SUCESSO Then gError 35259
    
    Exit Sub

Erro_BotaoExibirDados_Click:

    Select Case gErr

        Case 35256
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 35259, 30981
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173226)

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
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Lote")
    objGridInt.colColuna.Add ("Filial OP")
    objGridInt.colColuna.Add ("Embalagem")
    objGridInt.colColuna.Add ("Qtde.")
    objGridInt.colColuna.Add ("Peso Líq.")
    objGridInt.colColuna.Add ("Peso Bruto")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Imprimir.Name) 'Inserido por Wagner
    objGridInt.colCampo.Add (ProdutoEmb.Name)
    objGridInt.colCampo.Add (LoteRastro.Name)
    objGridInt.colCampo.Add (FilialOPRastro.Name)
    objGridInt.colCampo.Add (Embalagem.Name)
    objGridInt.colCampo.Add (QuantEmb.Name)
    objGridInt.colCampo.Add (PesoLiq.Name)
    objGridInt.colCampo.Add (PesoBruto.Name)

    'Colunas do Grid
    iGrid_Imprimir_Col = 1 'Inserido por Wagner
    iGrid_ProdutoEmb_Col = 2
    iGrid_LoteRastro_Col = 3
    iGrid_FilialOPRastro_Col = 4
    iGrid_Embalagem_Col = 5
    iGrid_QuantEmb_Col = 6
    iGrid_PesoLiq_Col = 7
    iGrid_PesoBruto_Col = 8

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
        
            Case iGrid_ProdutoEmb_Col
                lErro = Saida_Celula_Produto(objGridInt)
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173227)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
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
Dim objMovEstoque As New ClassMovEstoque
Dim objItemMovEstoque As ClassItemMovEstoque
Dim iCont As Integer
Dim dQuantidade As Double, objItemNF As New ClassItemNF
Dim objAlmoxarifado As New ClassAlmoxarifado
Dim objTipoDocInfo As New ClassTipoDocInfo
Dim objProdutoEmbalagem As ClassProdutoEmbalagem
Dim objEmbalagem As ClassEmbalagem
Dim objRastreamento As ClassRastreamentoMovto
Dim sProdutoFormatado As String
Dim iIndice As Integer
Dim iPosicao As Integer
Dim colRastreamentoMovto As New Collection
Dim objFilialEmpresa As New AdmFiliais

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = ProdutoEmb
    
    lErro = CF("Produto_Formata", ProdutoEmb.Text, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 96160
    
    If iProdutoPreenchido <> 0 Then

        'Lê os demais atributos do Produto
        objProduto.sCodigo = sProdutoFormatado
        
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 83251
            
        'Se o produto não está cadastrado, erro
        If lErro = 28030 Then gError 83252
                
        'se não for um produto rastreavel ==> erro
        If objProduto.iRastro = PRODUTO_RASTRO_NENHUM Then gError 83253
               
        objMovEstoque.lCodigo = StrParaLong(Codigo.Text)
    
        lErro = CF("MovEstoque_Le", objMovEstoque)
        If lErro <> SUCESSO And lErro <> 30128 Then gError 30981
        
        iCont = 0
        iIndice = 0
        For Each objItemMovEstoque In objMovEstoque.colItens
            iIndice = iIndice + 1
            If objItemMovEstoque.sProduto = sProdutoFormatado Then
                iCont = iCont + 1
                iPosicao = iIndice
            End If
        Next
        
        Set objProdutoEmbalagem = New ClassProdutoEmbalagem
        Set objEmbalagem = New ClassEmbalagem
        
        objProdutoEmbalagem.sProduto = sProdutoFormatado
            
        'Seleciona embalagem padrao
        lErro = CF("ProdutoEmbalagem_Le_Padrao", objProdutoEmbalagem)
        If lErro <> SUCESSO Then gError 136970
    
        objEmbalagem.iCodigo = objProdutoEmbalagem.iEmbalagem
    
        lErro = CF("Embalagem_Le", objEmbalagem)
        If lErro <> SUCESSO Then gError 136971
    
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_Embalagem_Col) = objEmbalagem.sSigla
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_PesoBruto_Col) = Formata_Estoque(objProdutoEmbalagem.dPesoBruto)
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_PesoLiq_Col) = Formata_Estoque(objProdutoEmbalagem.dPesoLiqTotal)
        
        If iCont = 1 Then
        
            Set objItemMovEstoque = objMovEstoque.colItens.Item(iPosicao)
                    
            'Lê movimentos de rastreamento vinculados ao itemNF passado ao ItemNF
            lErro = CF("RastreamentoMovto_Le_DocOrigem", objItemMovEstoque.lNumIntDoc, TIPO_RASTREAMENTO_MOVTO_MOVTO_ESTOQUE, colRastreamentoMovto)
            If lErro <> SUCESSO Then gError 136981
               
            If colRastreamentoMovto.Count = 1 Then
                
                Set objRastreamento = colRastreamentoMovto.Item(1)
                
                GridEtiquetas.TextMatrix(iLinha, iGrid_LoteRastro_Col) = objRastreamento.sLote
                
                FilialOPRastro.Text = objRastreamento.iFilialOP
                                    
                'Valida a Filial
                lErro = TP_FilialEmpresa_Le(FilialOPRastro.Text, objFilialEmpresa)
                If lErro <> SUCESSO And lErro <> 71971 And lErro <> 71972 Then gError 136991
                
                FilialOPRastro.Text = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
                
            End If
             
        End If
        
        'Se necessário cria uma nova linha no Grid
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    Else

        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_LoteRastro_Col) = ""
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_QuantEmb_Col) = ""
        objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_FilialOPRastro_Col) = ""

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 83169

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr
                
        Case 83169, 83207, 83250, 83251, 83346, 83349, 83351, 89515, 130322, 136970, 136971, 136981, 136991
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173228)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173229)
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173230)
              
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
Dim sCodProduto As String

On Error GoTo Erro_Rotina_Grid_Enable

    'Verifica se produto está preenchido
    sCodProduto = GridEtiquetas.TextMatrix(GridEtiquetas.Row, iGrid_ProdutoEmb_Col)

    lErro = CF("Produto_Formata", sCodProduto, sProdutoFormatado, iProdutoPreenchido)
    If lErro <> SUCESSO Then gError 30143

    Select Case objControl.Name
               
        Case ProdutoEmb.Name

            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
               objControl.Enabled = False
        
            Else
                objControl.Enabled = True
            End If
    
        Case LoteRastro.Name
            
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
               objControl.Enabled = True
        
            Else
                objControl.Enabled = False
            End If
            
        Case FilialOPRastro.Name
           
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
            Else
                objControl.Enabled = False
            End If
    
        Case Embalagem.Name
            
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
               objControl.Enabled = True
        
            Else
                objControl.Enabled = False
            End If
            
        Case QuantEmb.Name
            
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
               objControl.Enabled = True
        
            Else
                objControl.Enabled = False
            End If
            
        Case PesoLiq.Name
            
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
               objControl.Enabled = True
        
            Else
                objControl.Enabled = False
            End If
            
        Case PesoBruto.Name
            
            If iProdutoPreenchido = PRODUTO_PREENCHIDO Then
               objControl.Enabled = True
        
            Else
                objControl.Enabled = False
            End If
            
    End Select
    
    Exit Sub
     
Erro_Rotina_Grid_Enable:

    Select Case gErr
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173231)
     
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
                        If objGridInt.objGrid.TextMatrix(iLinha, iGrid_ProdutoEmb_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoEmb_Col) And _
                           objGridInt.objGrid.TextMatrix(iLinha, iGrid_LoteRastro_Col) = LoteRastro.Text Then gError 83183
                    End If
                Next
                
                objRastroLote.sCodigo = LoteRastro.Text
                objRastroLote.sProduto = sProdutoFormatado
                
                'Lê o Rastreamento do Lote vinculado ao produto
                lErro = CF("RastreamentoLote_Le", objRastroLote)
                If lErro <> SUCESSO And lErro <> 75710 Then gError 83184
                
                'Se não encontrou --> Erro
                If FrameOP.Visible = False Then If lErro = 75710 Then gError 83185
                
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
'                        If objGridInt.objGrid.TextMatrix(iLinha, iGrid_ProdutoEmb_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoEmb_Col) And _
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
                
                If objRastroLote.iFilialOP <> 0 And FrameOP.Visible = False Then
                
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
                
                End If
        
            End If
    
            'Preenche campos do lote
            lErro = Lote_Saida_Celula(objGridInt, objRastroLote)
            If lErro <> SUCESSO Then gError 83189
    
        End If
   
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173232)

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
'                    If objGridInt.objGrid.TextMatrix(iLinha, iGrid_ProdutoEmb_Col) = objGridInt.objGrid.TextMatrix(objGridInt.objGrid.Row, iGrid_ProdutoEmb_Col) And _
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
                
            If FrameOP.Visible = False Then If lErro = 75710 Then gError 83264
                
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173233)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173234)
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173235)
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173236)
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173237)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173238)

    End Select

    Exit Function

End Function

Sub Limpa_Tela_RelOpRotulosProducao()

    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridEtiquetas)
    
    PesoBTotal.Caption = ""
    PesoLTotal.Caption = ""
    
    glCodigo = 0
    lCodigoAntigo = 0
    gsCodigoOP = 0
    sCodigoOPAntigo = 0
    
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173239)
    
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
                           objGridEtiquetas.objGrid.TextMatrix(iLinha, iGrid_ProdutoEmb_Col) = objGridEtiquetas.objGrid.TextMatrix(objGridEtiquetas.objGrid.Row, iGrid_ProdutoEmb_Col) Then gError 83157
                    End If
                Next
                
            'Se o produto possuir rastro por OP
            ElseIf objProduto.iRastro = PRODUTO_RASTRO_OP Then
                                                
                For iLinha = 1 To objGridEtiquetas.iLinhasExistentes
                    If iLinha <> objGridEtiquetas.objGrid.Row Then
                        If objGridEtiquetas.objGrid.TextMatrix(iLinha, iGrid_LoteRastro_Col) = objRastroLote.sCodigo And _
                           Codigo_Extrai(objGridEtiquetas.objGrid.TextMatrix(iLinha, iGrid_FilialOPRastro_Col)) = objRastroLote.iFilialOP And objGridEtiquetas.objGrid.TextMatrix(iLinha, iGrid_ProdutoEmb_Col) = objGridEtiquetas.objGrid.TextMatrix(objGridEtiquetas.objGrid.Row, iGrid_ProdutoEmb_Col) Then gError 83158
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173240)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173241)

    End Select
    
    Exit Sub

End Sub

Public Sub GridEtiquetas_LeaveCell()

    Call Saida_Celula(objGridEtiquetas)

End Sub

Function RelRotulosProducao_Prepara(ByVal colRotuloProducao As Collection, lNumIntRel As Long) As Long
'grava registros na tabela temporaria para a emissao dos rotulos de expedicao

Dim lErro As Long, lTransacao As Long, objRotuloProducao As ClassRelRotuloProducao
Dim iSeq As Integer, iEtiquetas As Integer
Dim lComando As Long

On Error GoTo Erro_RelRotulosProducao_Prepara

    lComando = Comando_Abrir()
    If lComando = 0 Then gError 130323

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then gError 130324

    'obtem numintrel
    lErro = CF("Config_ObterNumInt", "FATConfig", "NUM_PROX_REL_ROTULOPRODUCAO", lNumIntRel)
    If lErro <> SUCESSO Then gError 130325
    
    For Each objRotuloProducao In colRotuloProducao
    
        If objRotuloProducao.iImprimir = MARCADO Then
            
            For iEtiquetas = 1 To objRotuloProducao.iQtdeEmb
            
                iSeq = iSeq + 1
                
                '####################################################
                'Alterado por Wagner
                lErro = Comando_Executar(lComando, "INSERT INTO RelRotuloProducao (NumIntRel, Seq, NumIntRastreamentoLote, PesoLiquido, PesoBruto, Lote, Produto, DataValidade, DataFabricacao) VALUES (?,?,?,?,?,?,?,?,?)", _
                    lNumIntRel, iSeq, objRotuloProducao.lNumIntRastreamentoLote, objRotuloProducao.dPesoLiquido, objRotuloProducao.dPesoBruto, objRotuloProducao.sLote, objRotuloProducao.sProduto, objRotuloProducao.dtDataValidade, objRotuloProducao.dtDataFabricacao)
                If lErro <> AD_SQL_SUCESSO Then gError 130329
                '####################################################
            
            Next
        
        End If
        
    Next
    
    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then gError 130030
    
    Call Comando_Fechar(lComando)
    
    RelRotulosProducao_Prepara = SUCESSO
     
    Exit Function
    
Erro_RelRotulosProducao_Prepara:

    RelRotulosProducao_Prepara = gErr
     
    Select Case gErr
          
        Case 130325, 130326, 130327
        
        Case 130329
            Call Rotina_Erro(vbOKOnly, "ERRO_GRAVACAO_RELROTULOPRODUCAO", gErr)
        
        Case 130328
            Call Rotina_Erro(vbOKOnly, "ERRO_LOTE_RASTRO_INEXISTENTE", gErr)
        
        Case 130323
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case 130024
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", gErr)
        
        Case 130030
            Call Rotina_Erro(vbOKOnly, "ERRO_COMMIT", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173242)
     
    End Select
     
    Call Transacao_Rollback
    
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function

'#####################################################
'Inserido por Wagner
Function Preenche_Tela_MovEstoque(ByVal objMovEstoque As ClassMovEstoque) As Long
'Preenche o Grid de Itens

Dim lErro As Long
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

On Error GoTo Erro_Preenche_Tela_MovEstoque

    glCodigo = objMovEstoque.lCodigo
    
    Codigo.Text = CStr(objMovEstoque.lCodigo)
    sCodigoOPAntigo = ""
    CodigoOP.Text = ""

    Call Grid_Limpa(objGridEtiquetas)

    PesoBTotal.Caption = ""
    PesoLTotal.Caption = ""

    For Each objItemMovEstoque In objMovEstoque.colItens
                           
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
                            
                GridEtiquetas.TextMatrix(iLinha, iGrid_LoteRastro_Col) = objRastreamento.sLote
                
                lErro = Mascara_RetornaProdutoTela(objItemMovEstoque.sProduto, sProdutoMascarado)
                If lErro <> SUCESSO Then gError 130322
                
                GridEtiquetas.TextMatrix(iLinha, iGrid_ProdutoEmb_Col) = sProdutoMascarado
                            
                FilialOPRastro.Text = objRastreamento.iFilialOP
                                    
                'Valida a Filial
                lErro = TP_FilialEmpresa_Le(FilialOPRastro.Text, objFilialEmpresa)
                If lErro <> SUCESSO And lErro <> 71971 And lErro <> 71972 Then gError 136991
                
                FilialOPRastro.Text = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
                
                GridEtiquetas.TextMatrix(iLinha, iGrid_FilialOPRastro_Col) = FilialOPRastro.Text
                
                objProdutoEmbalagem.sProduto = objItemMovEstoque.sProduto
                    
                'Seleciona embalagem padrao
                lErro = CF("ProdutoEmbalagem_Le_Padrao", objProdutoEmbalagem)
                If lErro <> SUCESSO And lErro <> 100000 Then gError 136970
            
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
    
    Call Grid_Refresh_Checkbox(objGridEtiquetas)

    objGridEtiquetas.iLinhasExistentes = iLinha
    
    Preenche_Tela_MovEstoque = SUCESSO

    Exit Function

Erro_Preenche_Tela_MovEstoque:

    Preenche_Tela_MovEstoque = gErr

    Select Case gErr
    
        Case 136970 To 136972, 136974, 136979, 136981, 136991
        
        Case 136975
            Call Rotina_Erro(vbOKOnly, "ERRO_NOTA_FISCAL_INTERNA_SAIDA_NAO_CADASTRADA", gErr, objMovEstoque)
        
        Case 136980
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", gErr, objAlmoxarifado.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173243)

    End Select

    Exit Function

End Function

Private Function Critica_GridEtiquetas(ByVal colRotuloProducao As Collection) As Long
'Critica os parâmetros que serão passados para o relatório

Dim lErro As Long
Dim objRotuloProducao As ClassRelRotuloProducao
Dim iLinha As Integer
Dim vbMsg As VbMsgBoxResult
Dim iContImprimir As Integer

On Error GoTo Erro_Critica_GridEtiquetas

    If colRotuloProducao.Count = 0 Then gError 136993
    
    For Each objRotuloProducao In colRotuloProducao
    
        iLinha = iLinha + 1
        
        If objRotuloProducao.iQtdeEmb = 0 Then gError 136992
        If objRotuloProducao.sLote = "" Then gError 136994
        If objRotuloProducao.sProduto = "" Then gError 136995
        
        If objRotuloProducao.dPesoBruto < objRotuloProducao.dPesoLiquido Then gError 136986

        If objRotuloProducao.iImprimir = MARCADO Then iContImprimir = iContImprimir + 1

    Next
       
    'Se não tem nenhum item a imprimir
    If iContImprimir = 0 Then gError 136997
       
    Critica_GridEtiquetas = SUCESSO

    Exit Function

Erro_Critica_GridEtiquetas:

    Critica_GridEtiquetas = gErr

    Select Case gErr
    
        Case 136986
            Call Rotina_Erro(vbOKOnly, "ERRO_PESO_LIQUIDO_MAIOR_BRUTO_GRID", gErr, Formata_Estoque(objRotuloProducao.dPesoLiquido), Formata_Estoque(objRotuloProducao.dPesoBruto), iLinha)
        
        Case 136987, 136988, 136990
        
        Case 136989

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173244)

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

Private Sub BotaoLimparGrid_Click()
    Call Grid_Limpa(objGridEtiquetas)
End Sub
'#####################################################




'#####################################################
'Inserido por Wagner 16/01/2006
Private Sub UpDownValidade_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownValidade_DownClick

    'Verifica se a Data da Validade foi preenchida
    If Len(Trim(DataValidade.ClipText)) = 0 Then Exit Sub

    'Diminui a Data em um dia
    lErro = Data_Up_Down_Click(DataValidade, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 141506

    Exit Sub

Erro_UpDownValidade_DownClick:

    Select Case gErr

        Case 141506

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173245)

    End Select

    Exit Sub

End Sub

Private Sub UpDownValidade_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownValidade_UpClick

    'Verifica se a Data de Validade foi preenchida
    If Len(Trim(DataValidade.ClipText)) = 0 Then Exit Sub

    'Aumenta a Data em um dia
    lErro = Data_Up_Down_Click(DataValidade, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 141507

    Exit Sub

Erro_UpDownValidade_UpClick:

    Select Case gErr

        Case 141507

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173246)

    End Select

    Exit Sub

End Sub

Private Sub UpDownFabricacao_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownFabricacao_DownClick

    'Verifica se a Data da Fabricacao foi preenchida
    If Len(Trim(DataFabricacao.ClipText)) = 0 Then Exit Sub

    'Diminui a Data em um dia
    lErro = Data_Up_Down_Click(DataFabricacao, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 141508

    Exit Sub

Erro_UpDownFabricacao_DownClick:

    Select Case gErr

        Case 141508

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173247)

    End Select

    Exit Sub

End Sub

Private Sub UpDownFabricacao_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownFabricacao_UpClick

    'Verifica se a Data de Fabricacao foi preenchida
    If Len(Trim(DataFabricacao.ClipText)) = 0 Then Exit Sub

    'Aumenta a Data em um dia
    lErro = lErro = Data_Up_Down_Click(DataFabricacao, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 141509

    Exit Sub

Erro_UpDownFabricacao_UpClick:

    Select Case gErr

        Case 141509

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173248)

    End Select

    Exit Sub

End Sub

Private Sub DataFabricacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataFabricacao_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataFabricacao, iAlterado)

End Sub

Private Sub DataFabricacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFabricacao_Validate

    'Se a data não foi preenchida, sai da rotina
    If Len(DataFabricacao.ClipText) = 0 Then Exit Sub

    'Critica a data
    lErro = Data_Critica(DataFabricacao.Text)
    If lErro <> SUCESSO Then gError 141510

    Exit Sub

Erro_DataFabricacao_Validate:

    Cancel = True

    Select Case gErr

        Case 141510

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173249)

    End Select

    Exit Sub

End Sub

Private Sub DataValidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataValidade_GotFocus()

    Call MaskEdBox_TrataGotFocus(DataValidade, iAlterado)

End Sub

Private Sub DataValidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataValidade_Validate

    'Se a data não foi preenchida, sai da rotina
    If Len(DataValidade.ClipText) = 0 Then Exit Sub

    'Critica a data
    lErro = Data_Critica(DataValidade.Text)
    If lErro <> SUCESSO Then gError 141511

    Exit Sub

Erro_DataValidade_Validate:

    Cancel = True

    Select Case gErr

        Case 141511

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 173250)

    End Select

    Exit Sub

End Sub

Public Sub BotaoExibirDadosOP_Click()

Dim lErro As Long
Dim objOp As New ClassOrdemDeProducao

On Error GoTo Erro_BotaoExibirDadosOP_Click

    'Verifica se a Serie e o Número da Nota Fiscal original estão preenchidos
    If Len(Trim(CodigoOP.Text)) = 0 Then gError 141512

    FrameOP.Visible = True

    objOp.sCodigo = CodigoOP.Text
    objOp.iFilialEmpresa = giFilialEmpresa

    'tenta ler a OP desejada
    lErro = CF("OrdemProducao_Le", objOp)
    If lErro <> SUCESSO And lErro <> 30368 And lErro <> 55316 Then gError 141513

    lErro = CF("OrdemDeProducao_Le_ComItens", objOp)
    If lErro <> SUCESSO And lErro <> 21960 Then gError 141514

    If lErro = 21960 Then

        lErro = CF("OrdemDeProducaoBaixada_Le_ComItens", objOp)
        If lErro <> SUCESSO And lErro <> 82797 Then gError 141515
        
        If lErro = 82797 Then gError 141516
                
    End If
    
    'Coloca na tela os dados encontrados
    lErro = Preenche_Tela_OP(objOp)
    If lErro <> SUCESSO Then gError 141517
    
    Exit Sub

Erro_BotaoExibirDadosOP_Click:

    Select Case gErr

        Case 141512
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            CodigoOP.SetFocus
            
        Case 141513 To 141517
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173251)

    End Select

    Exit Sub

End Sub

Private Sub CodigoOPLabel_Click()

Dim objOrdemDeProducao As New ClassOrdemDeProducao
Dim colSelecao As New Collection

    'preenche o objOrdemDeProducao com o código da tela , se estiver preenchido
    If Len(Trim(CodigoOP.Text)) <> 0 Then objOrdemDeProducao.sCodigo = CodigoOP.Text
    
    'lista as OP's
    Call Chama_Tela("OrdemProducaoLista", colSelecao, objOrdemDeProducao, objEventoCodigoOP)

End Sub

Private Sub objEventoCodigoOP_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOrdemDeProducao As ClassOrdemDeProducao

On Error GoTo Erro_objEventoCodigoOP_evSelecao

    Set objOrdemDeProducao = obj1

    FrameOP.Visible = True 'Inserido por Wagner

    'traz OP para a tela
    CodigoOP.Text = objOrdemDeProducao.sCodigo
    Call BotaoExibirDadosOP_Click

    lErro = ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCodigoOP_evSelecao:

    Select Case gErr

        Case 141518

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173252)
            
    End Select

    Exit Sub

End Sub

Function Preenche_Tela_OP(ByVal objOp As ClassOrdemDeProducao) As Long
'Preenche o Grid de Itens

Dim lErro As Long
Dim objRastreamento As ClassRastreamentoMovto
Dim sProdutoMascarado As String
Dim objProdutoEmbalagem As ClassProdutoEmbalagem
Dim objEmbalagem As ClassEmbalagem
Dim iQtdEmb As Integer
Dim iLinha As Integer
Dim objItemOP As ClassItemOP
Dim objAlmoxarifado As ClassAlmoxarifado
Dim objFilialEmpresa As AdmFiliais

On Error GoTo Erro_Preenche_Tela_OP

    gsCodigoOP = objOp.sCodigo
    
    CodigoOP.Text = CStr(objOp.sCodigo)
    lCodigoAntigo = 0
    Codigo.Text = ""

    Call Grid_Limpa(objGridEtiquetas)
    
    PesoBTotal.Caption = ""
    PesoLTotal.Caption = ""

    For Each objItemOP In objOp.colItens
                       
        Set objAlmoxarifado = New ClassAlmoxarifado
        
        iLinha = iLinha + 1

        'Lê o Almoxarifado
        objAlmoxarifado.iCodigo = objItemOP.iAlmoxarifado
        lErro = CF("Almoxarifado_Le", objAlmoxarifado)
        If lErro <> SUCESSO And lErro <> 25056 Then gError 141519

        'Se não encontrou Almoxarifado --> Erro
        If lErro = 25056 Then gError 141520
                
        Set objFilialEmpresa = New AdmFiliais
        Set objProdutoEmbalagem = New ClassProdutoEmbalagem
        Set objEmbalagem = New ClassEmbalagem
                    
        GridEtiquetas.TextMatrix(iLinha, iGrid_LoteRastro_Col) = objOp.sCodigo
        
        lErro = Mascara_RetornaProdutoTela(objItemOP.sProduto, sProdutoMascarado)
        If lErro <> SUCESSO Then gError 141521
        
        GridEtiquetas.TextMatrix(iLinha, iGrid_ProdutoEmb_Col) = sProdutoMascarado
                    
        FilialOPRastro.Text = objOp.iFilialEmpresa
                            
        'Valida a Filial
        lErro = TP_FilialEmpresa_Le(FilialOPRastro.Text, objFilialEmpresa)
        If lErro <> SUCESSO And lErro <> 71971 And lErro <> 71972 Then gError 141522
        
        FilialOPRastro.Text = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
        
        GridEtiquetas.TextMatrix(iLinha, iGrid_FilialOPRastro_Col) = FilialOPRastro.Text
        
        objProdutoEmbalagem.sProduto = objItemOP.sProduto
            
        'Seleciona embalagem padrao
        lErro = CF("ProdutoEmbalagem_Le_Padrao", objProdutoEmbalagem)
        If lErro <> SUCESSO And lErro <> 100000 Then gError 141523
    
        objEmbalagem.iCodigo = objProdutoEmbalagem.iEmbalagem
    
        lErro = CF("Embalagem_Le", objEmbalagem)
        If lErro <> SUCESSO And lErro <> 82763 Then gError 141524
    
        GridEtiquetas.TextMatrix(iLinha, iGrid_Embalagem_Col) = objEmbalagem.sSigla
        GridEtiquetas.TextMatrix(iLinha, iGrid_PesoBruto_Col) = Formata_Estoque(objProdutoEmbalagem.dPesoBruto)
        GridEtiquetas.TextMatrix(iLinha, iGrid_PesoLiq_Col) = Formata_Estoque(objProdutoEmbalagem.dPesoLiqTotal)
        GridEtiquetas.TextMatrix(iLinha, iGrid_Imprimir_Col) = MARCADO
    
    Next
    
    Call Grid_Refresh_Checkbox(objGridEtiquetas)

    objGridEtiquetas.iLinhasExistentes = iLinha
    
    Preenche_Tela_OP = SUCESSO

    Exit Function

Erro_Preenche_Tela_OP:

    Preenche_Tela_OP = gErr

    Select Case gErr
    
        Case 141519, 141521 To 141524
        
        Case 141520
            Call Rotina_Erro(vbOKOnly, "ERRO_ALMOXARIFADO_NAO_CADASTRADO", gErr, objAlmoxarifado.iCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173253)

    End Select

    Exit Function

End Function

Private Sub CodigoOP_Validate(Cancel As Boolean)

Dim lErro As Long, iIndice As Integer

On Error GoTo Erro_CodigoOP_Validate

    'se o codigo foi trocado
    If sCodigoOPAntigo <> CodigoOP.Text Then

        sCodigoOPAntigo = CodigoOP.Text
        
        Call Grid_Limpa(objGridEtiquetas)
        
   End If

   Exit Sub

Erro_CodigoOP_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173254)

    End Select

    Exit Sub


End Sub

'#################################################################
