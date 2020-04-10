VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ChequesPag2Ocx 
   ClientHeight    =   5790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9405
   ScaleHeight     =   5790
   ScaleWidth      =   9405
   Begin VB.CommandButton BotaoDesmarcar 
      Caption         =   "Desmarcar Todas"
      Height          =   690
      Left            =   4915
      Picture         =   "ChequesPag2Ocx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3945
      Width           =   1650
   End
   Begin MSMask.MaskEdBox FilialEmpresa 
      Height          =   225
      Left            =   3825
      TabIndex        =   36
      Top             =   3330
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      Enabled         =   0   'False
      PromptChar      =   " "
   End
   Begin VB.TextBox Tipo 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   4140
      TabIndex        =   18
      Top             =   2250
      Width           =   795
   End
   Begin VB.CommandButton BotaoDocOriginal 
      Height          =   690
      Left            =   7275
      Picture         =   "ChequesPag2Ocx.ctx":11E2
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3945
      Width           =   1650
   End
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   6150
      ScaleHeight     =   495
      ScaleWidth      =   2685
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   150
      Width           =   2745
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2160
         Picture         =   "ChequesPag2Ocx.ctx":40F8
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoVoltar 
         Height          =   345
         Left            =   135
         Picture         =   "ChequesPag2Ocx.ctx":4276
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   90
         Width           =   885
      End
      Begin VB.CommandButton BotaoSeguir 
         Height          =   345
         Left            =   1117
         Picture         =   "ChequesPag2Ocx.ctx":49D4
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   90
         Width           =   930
      End
   End
   Begin VB.CommandButton BotaoAgrupar 
      Caption         =   "Agrupar"
      Height          =   690
      Left            =   195
      Picture         =   "ChequesPag2Ocx.ctx":5166
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3945
      Width           =   1650
   End
   Begin VB.CommandButton BotaoDesagrupar 
      Caption         =   "Desagrupar"
      Height          =   690
      Left            =   2555
      Picture         =   "ChequesPag2Ocx.ctx":578C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3945
      Width           =   1650
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleção Atual"
      Height          =   975
      Left            =   1245
      TabIndex        =   20
      Top             =   4695
      Width           =   6915
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade Títulos:"
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
         Left            =   255
         TabIndex        =   26
         Top             =   315
         Width           =   1710
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total dos Cheques:"
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
         Left            =   3270
         TabIndex        =   27
         Top             =   630
         Width           =   2160
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade Cheques:"
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
         Left            =   3270
         TabIndex        =   28
         Top             =   315
         Width           =   1845
      End
      Begin VB.Label QtdTitulos 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   2115
         TabIndex        =   29
         Top             =   285
         Width           =   1005
      End
      Begin VB.Label QtdCheques 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   5475
         TabIndex        =   30
         Top             =   285
         Width           =   1200
      End
      Begin VB.Label ValorTotalCheques 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   5475
         TabIndex        =   31
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label Label6 
         Caption         =   "Valor Total Títulos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   255
         TabIndex        =   32
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label ValorTotalTitulos 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   2115
         TabIndex        =   33
         Top             =   600
         Width           =   1005
      End
   End
   Begin VB.CommandButton BotaoSubir 
      Height          =   285
      Left            =   8985
      Picture         =   "ChequesPag2Ocx.ctx":5DB2
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1890
      Width           =   300
   End
   Begin VB.CommandButton BotaoDescer 
      Height          =   285
      Left            =   8985
      Picture         =   "ChequesPag2Ocx.ctx":5F74
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2475
      Width           =   300
   End
   Begin VB.CheckBox CheckSelecionar 
      Height          =   225
      Left            =   2955
      TabIndex        =   9
      Top             =   2895
      Width           =   900
   End
   Begin VB.CheckBox CheckEmitir 
      Height          =   225
      Left            =   2070
      TabIndex        =   10
      Top             =   2910
      Width           =   700
   End
   Begin MSMask.MaskEdBox Cheque 
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   2505
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      Enabled         =   0   'False
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fornecedor 
      Height          =   225
      Left            =   870
      TabIndex        =   3
      Top             =   2535
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox Filial 
      Height          =   225
      Left            =   2745
      TabIndex        =   4
      Top             =   2520
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox NumTitulo 
      Height          =   225
      Left            =   3300
      TabIndex        =   5
      Top             =   2520
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#########"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Parcela 
      Height          =   225
      Left            =   4125
      TabIndex        =   6
      Top             =   2520
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   225
      Left            =   4980
      TabIndex        =   7
      Top             =   2505
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox Portador 
      Height          =   225
      Left            =   5040
      TabIndex        =   8
      Top             =   2925
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
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
   Begin MSMask.MaskEdBox TipoCobranca 
      Height          =   225
      Left            =   6015
      TabIndex        =   11
      Top             =   2895
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      AllowPrompt     =   -1  'True
      Enabled         =   0   'False
      MaxLength       =   30
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
   Begin MSMask.MaskEdBox DataVencto 
      Height          =   225
      Left            =   3915
      TabIndex        =   21
      Top             =   2925
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Enabled         =   0   'False
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Juros 
      Height          =   225
      Left            =   570
      TabIndex        =   12
      Top             =   3360
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Desconto 
      Height          =   225
      Left            =   2700
      TabIndex        =   14
      Top             =   3375
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   0
      PromptInclude   =   0   'False
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Multa 
      Height          =   225
      Left            =   1560
      TabIndex        =   13
      Top             =   3330
      Width           =   720
      _ExtentX        =   1270
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
   Begin MSFlexGridLib.MSFlexGrid GridChequesPag2 
      Height          =   2955
      Left            =   150
      TabIndex        =   15
      Top             =   840
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5212
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      AllowBigSelection=   0   'False
      FocusRect       =   2
   End
   Begin VB.Label Label1 
      Caption         =   "Conta :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   225
      TabIndex        =   34
      Top             =   315
      Width           =   735
   End
   Begin VB.Label Conta 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Left            =   960
      TabIndex        =   35
      Top             =   300
      Width           =   2040
   End
End
Attribute VB_Name = "ChequesPag2Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer
Private gobjChequesPag As ClassChequesPag
Private gcolTiposDeCobranca As New AdmColCodigoNome

Dim objGrid As AdmGrid
Dim iGrid_Cheque_Col As Integer
Dim iGrid_FilialEmpresa_Col As Integer
Dim iGrid_Fornecedor_Col As Integer
Dim iGrid_Filial_Col As Integer
Dim iGrid_Tipo_Col As Integer
Dim iGrid_NumTitulo_Col As Integer
Dim iGrid_Parcela_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Emitir_Col As Integer
Dim iGrid_Selecionar_Col As Integer
Dim iGrid_Vencimento As Integer
Dim iGrid_Portador_Col As Integer
Dim iGrid_Juros_Col As Integer
Dim iGrid_Multa_Col As Integer
Dim iGrid_Desconto_Col As Integer
Dim iGrid_TipoCobranca_Col As Integer

Private Type typeAnalise

    iLinhasMarcadas As Integer
    iMesmoPortador As Integer
    sPortadorNome As String
    iMesmoCheque As Integer
    iSeqCheque As Integer
    iMesmoTipoCobranca As Integer
    iTipoCobranca As Integer
    iComSeqZero As Integer
    iMesmoFornFil As Integer
    
End Type

Private gtAnalise As typeAnalise

Private Sub Verifica_Se_Linha_De_Grupo(iLinhaTeste As Integer, iLinhaDeGrupo As Integer)
'retorna em iLinhaDeGrupo 1 ou zero p/indicar se a linha do grid iLinhaTeste faz parte de um grupo de parcelas de um cheque

Dim iLinha As Integer, iSeqCheque As Integer

    iLinhaDeGrupo = 0
    
    iSeqCheque = Linha_ObtemSeqCheque(iLinhaTeste)
    
    If iLinhaTeste > 1 Then
        
        If iSeqCheque = Linha_ObtemSeqCheque(iLinhaTeste - 1) Then iLinhaDeGrupo = 1
    
    End If
    
    If iLinhaTeste < objGrid.iLinhasExistentes Then
    
        If iSeqCheque = Linha_ObtemSeqCheque(iLinhaTeste + 1) Then iLinhaDeGrupo = 1
        
    End If

End Sub

Private Sub Analisa_Grid()
'percorre o grid obtendo informacoes
'para zero ou uma linhas selecionadas iMesmoPortador e iMesmoCheque retornam 1

Dim iLinha As Integer, iSeqChequeLido As Integer
Dim sPortadorNome As String
Dim iLinhasMarcadas As Integer, iMesmoPortador As Integer, iMesmoCheque As Integer
Dim iPortador As Integer, iSeqCheque As Integer, iComSeqZero As Integer
Dim iTipoCobranca As Integer, iMesmoTipoCobranca As Integer
Dim iMesmoFornFil As Integer, sFornecedor As String, iFilialForn As Integer

    iLinhasMarcadas = 0
    iMesmoPortador = 1
    iMesmoCheque = 1
    iMesmoTipoCobranca = 1
    iComSeqZero = 0
    iMesmoFornFil = 1
    
    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes
    
        'Se a Parcela está marcada
        If GridChequesPag2.TextMatrix(iLinha, iGrid_Selecionar_Col) = SELECIONAR_CHECADO Then
        
            iSeqChequeLido = Linha_ObtemSeqCheque(iLinha)
            
            If iLinhasMarcadas = 0 Then
            
                sPortadorNome = GridChequesPag2.TextMatrix(iLinha, iGrid_Portador_Col)
                iSeqCheque = iSeqChequeLido
                iTipoCobranca = Codigo_Extrai(GridChequesPag2.TextMatrix(iLinha, iGrid_TipoCobranca_Col))
                sFornecedor = GridChequesPag2.TextMatrix(iLinha, iGrid_Fornecedor_Col)
                iFilialForn = Codigo_Extrai(GridChequesPag2.TextMatrix(iLinha, iGrid_Filial_Col))
                
            Else
                
                If iMesmoPortador = 1 And sPortadorNome <> GridChequesPag2.TextMatrix(iLinha, iGrid_Portador_Col) Then iMesmoPortador = 0
                If iMesmoCheque = 1 And iSeqCheque <> iSeqChequeLido Then iMesmoCheque = 0
                If iMesmoTipoCobranca = 1 And iTipoCobranca <> Codigo_Extrai(GridChequesPag2.TextMatrix(iLinha, iGrid_TipoCobranca_Col)) Then iMesmoTipoCobranca = 0
                If iMesmoFornFil = 1 And (sFornecedor <> GridChequesPag2.TextMatrix(iLinha, iGrid_Fornecedor_Col) Or iFilialForn <> Codigo_Extrai(GridChequesPag2.TextMatrix(iLinha, iGrid_Filial_Col))) Then iMesmoFornFil = 0
                
            End If
            
            If iSeqChequeLido = 0 Then iComSeqZero = 1
            
            iLinhasMarcadas = iLinhasMarcadas + 1
            
        End If
        
    Next

    With gtAnalise
        .iLinhasMarcadas = iLinhasMarcadas
        .iMesmoPortador = iMesmoPortador
        .sPortadorNome = sPortadorNome
        .iMesmoCheque = iMesmoCheque
        .iSeqCheque = iSeqCheque
        .iMesmoTipoCobranca = iMesmoTipoCobranca
        .iTipoCobranca = iTipoCobranca
        .iComSeqZero = iComSeqZero
        .iMesmoFornFil = iMesmoFornFil
    End With
    
End Sub

Private Sub BotaoAgrupar_Click()
'Agrupa as Parcelas selecionadas, que estiverem marcadas para emissão

Dim lErro As Long
Dim iLinhasMarcadas As Integer

On Error GoTo Erro_BotaoAgrupar_Click
       
    'obter o # de linhas marcadas e se todas tem o mesmo portador
    Call Analisa_Grid
        
    'Verifica se há mais de um Portador nas Parcelas selecionadas
    If gtAnalise.iMesmoPortador <> 1 Then Error 15856
    
    'Se não houverem pelo menos 2 Parcelas selecionadas
    If gtAnalise.iLinhasMarcadas < 2 Then Error 15857
           
    'se selecionou titulos com tipo de cobranca diferente
    If gtAnalise.iMesmoTipoCobranca <> 1 Then Error 19374
    
    'se selecionou titulo desmarcado p/emitir
    If gtAnalise.iComSeqZero = 1 Then Error 19375
    
    'se titulos de fornecedor/filial diferentes e cobranca nao for bancaria
    If gtAnalise.iMesmoFornFil = 0 And (gtAnalise.iTipoCobranca <> TIPO_COBRANCA_BANCARIA Or gtAnalise.sPortadorNome = "") Then Error 19436
    
    'Reposiciona no Obj (agrupando) as Parcelas marcadas
    Call Titulos_Reposiciona_Agrupar
    
    'Atualiza as parcelas na tela
    Call Traz_Dados_Tela
    
    Exit Sub
    
Erro_BotaoAgrupar_Click:

    Select Case Err
    
        Case 19375
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NAO_PODE_AGRUPAR_SEQ_ZERO", Err)
            
        Case 19374
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOS_TIPO_COBR_DIFERENTE", Err)
        
        Case 15856
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOS_PORTADOR_DIFERENTE", Err)
            
        Case 15857
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOS_NAO_MARCADOS", Err)
            
        Case 19436
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOS_FORN_DIF_NAO_COBRBANC", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144559)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoDesagrupar_Click()
'Desagrupa as Parcelas selecionadas, que estiverem marcadas para emissão

Dim lErro As Long

On Error GoTo Erro_BotaoDesagrupar_Click
       
    'obter o # de linhas marcadas e se todas sao do mesmo cheque
    Call Analisa_Grid
    
    'Verifica se há mais de um Cheque nas Parcelas selecionadas
    If gtAnalise.iMesmoCheque <> 1 Then Error 15862
            
    If gtAnalise.iSeqCheque = 0 Then Error 19373
    
    'Se não houver pelo menos 1 Parcela selecionada
    If gtAnalise.iLinhasMarcadas < 1 Then Error 15863
                   
    'Reposiciona no Obj (desagrupando) as Parcelas marcadas
    Call Titulos_Reposiciona_Desagrupar
    
    'Atualiza as Parcelas na tela
    Call Traz_Dados_Tela
    
    Exit Sub
    
Erro_BotaoDesagrupar_Click:

    Select Case Err
    
        Case 19373
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOS_DE_UM_CHEQUE", Err)
        
        Case 15862
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOS_CHEQUE_DIFERENTE", Err)
            
        Case 15863
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_NAO_MARCADO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144560)
    
    End Select
    
    Exit Sub

End Sub

Private Sub BotaoDesmarcar_Click()

Dim iLinha As Integer
Dim objInfoParcPag As ClassInfoParcPag

    For Each objInfoParcPag In gobjChequesPag.colInfoParcPag
    
        'Zera o número do cheque relacionado à Parcela
        objInfoParcPag.iSeqCheque = 0
    
    Next

    'Atualiza as parcelas na tela
    Call Traz_Dados_Tela

End Sub

Private Sub BotaoDocOriginal_Click()

Dim lErro As Long
Dim objInfoParcPag As New ClassInfoParcPag
Dim objTituloPagar As New ClassTituloPagar
Dim objParcelaPagar As New ClassParcelaPagar

On Error GoTo Erro_BotaoDocOriginal_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridChequesPag2.Row = 0 Then Error 60491
        
    'Se foi selecionada uma linha que está preenchida
    If GridChequesPag2.Row <= objGrid.iLinhasExistentes Then
        
        Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(GridChequesPag2.Row)
               
        objParcelaPagar.lNumIntDoc = objInfoParcPag.lNumIntParc
        
        'Le o NumInterno do Titulo para passar no objTituloPag
        lErro = CF("ParcelaPagar_Le", objParcelaPagar)
        If lErro <> SUCESSO And lErro <> 60479 Then Error 60492
        
        'Se não encontrou a Parcela --> ERRO
        If lErro = 60479 Then Error 60493
        
        objTituloPagar.lNumIntDoc = objParcelaPagar.lNumIntTitulo
        
        'Le os Dados do Titulo
        lErro = CF("TituloPagar_Le", objTituloPagar)
        If lErro <> SUCESSO And lErro <> 18372 Then Error 60494
        
        If lErro = 18372 Then Error 60495

        'Abre a tela de títulos a pagar
        Call Chama_Tela("TituloPagar_Consulta", objTituloPagar)
    
    End If
        
    Exit Sub
    
Erro_BotaoDocOriginal_Click:

    Select Case Err
    
        Case 60491
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", Err)
         
        Case 60492, 60494 'Tratado na rotina chamada
        
        Case 60493
            Call Rotina_Erro(vbOKOnly, "ERRO_PARCELA_PAGAR_INEXISTENTE", Err)
        
        Case 60495
            Call Rotina_Erro(vbOKOnly, "ERRO_TITULO_PAGAR_INEXISTENTE", Err)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 144561)

    End Select

    Exit Sub
    
End Sub

Private Sub BotaoFechar_Click()

    'Fecha a tela
    Unload Me
        
End Sub

Private Sub BotaoSeguir_Click()
        
Dim objInfoParcPag As ClassInfoParcPag
Dim iVerificaMarcado As Integer
Dim lErro As Long
    
On Error GoTo Erro_BotaoSeguir_Click

    'Percorre todas as Parcelas da Coleção passada por parâmetro
    For Each objInfoParcPag In gobjChequesPag.colInfoParcPag

        'Se a Parcela está marcada
        If objInfoParcPag.iSeqCheque <> 0 Then
                                
            lErro = CF("TituloPagar_Verifica_Adiantamento", objInfoParcPag.lFornecedor, objInfoParcPag.iFilialForn)
            If lErro <> SUCESSO Then gError 59471
    
            iVerificaMarcado = iVerificaMarcado + 1
            
        End If
        
    Next
    
    'Verifica se nenhuma Parcela está marcada
    If iVerificaMarcado = 0 Then gError 15855
    
    'Preenche a coleção com os cheques selecionados para emissão
    lErro = CF("ChequesPag_ChequesSelecionados", gobjChequesPag)
    If lErro <> SUCESSO Then gError 15873
    
    'Chama a tela do passo seguinte
    Call Chama_Tela("ChequesPag3", gobjChequesPag)
    
    'Fecha a tela
    Unload Me
    
    Exit Sub
    
Erro_BotaoSeguir_Click:

    Select Case gErr
    
        Case 15855
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULO_NAO_MARCADO", Err)
                        
        Case 15873, 59471
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144562)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub BotaoSubir_Click()

Dim iNumCheque As Integer
Dim objInfoParcPag As ClassInfoParcPag
    
    'se alguma linha valida está selecionada
    If (GridChequesPag2.Row > 0) Then
    
        'Passa a linha selecionada do Grid para o Obj
        Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(GridChequesPag2.Row)
    
        'Guarda o número de cheque do grupo selecionado
        iNumCheque = objInfoParcPag.iSeqCheque
        
        'se nao for o 1o cheque
        If iNumCheque <> 1 Then
        
            Call Sobe_Grupo(iNumCheque)
        
            'Atualiza as Parcelas na tela
            Call Traz_Dados_Tela

        End If
        
    End If
    
End Sub

Private Sub BotaoDescer_Click()

Dim iNumCheque As Integer
Dim objInfoParcPag As ClassInfoParcPag, objInfoParcPagUltimo As ClassInfoParcPag

    'se alguma linha valida está selecionada
    If (GridChequesPag2.Row > 0) Then
    
        'Passa a linha selecionada do Grid para o Obj
        Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(GridChequesPag2.Row)
    
        'Guarda o número de cheque do grupo selecionado
        iNumCheque = objInfoParcPag.iSeqCheque
        
        Set objInfoParcPagUltimo = gobjChequesPag.colInfoParcPag.Item(objGrid.iLinhasExistentes)
        
        'se nao for o ultimo cheque
        If iNumCheque <> objInfoParcPagUltimo.iSeqCheque Then
        
            Call Sobe_Grupo(iNumCheque + 1)
        
            'Atualiza as Parcelas na tela
            Call Traz_Dados_Tela

        End If
        
    End If
    
End Sub

Private Sub BotaoVoltar_Click()
    
    'Chama a tela do passo anterior
    Call Chama_Tela("ChequesPag", gobjChequesPag)

    'Fecha a tela
    Unload Me
    
End Sub

Private Sub CheckEmitir_Click()

Dim iLinha As Integer, iLinha2 As Integer, iLinhaDeGrupo As Integer
Dim iCheque As Integer
Dim objInfoParcPag As ClassInfoParcPag

    'Passa para iLinha o número da linha em questão
    iLinha = GridChequesPag2.Row
    
    'Passa a linha do Grid para o Obj
    Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinha)
        
    'Se a Parcela foi desmarcada
    If GridChequesPag2.TextMatrix(iLinha, iGrid_Emitir_Col) = EMITIR_NAO_CHECADO Then
                
        'Lê o número de cheque da Parcela
        iCheque = objInfoParcPag.iSeqCheque
        
        Call Verifica_Se_Linha_De_Grupo(iLinha, iLinhaDeGrupo)
        
        'Zera o número do cheque relacionado à Parcela
        objInfoParcPag.iSeqCheque = 0
                
        If gobjChequesPag.colInfoParcPag.Count <> 1 Then
        
            'Remove a Parcela da colecao
            gobjChequesPag.colInfoParcPag.Remove Index:=iLinha
            
            'Insere a Parcela no início do Grid
            gobjChequesPag.colInfoParcPag.Add Item:=objInfoParcPag, Before:=1
            
        End If
        
        'Se a Parcela desmarcada não era a última e nao fazia parte de um grupo de parcelas de um cheque
        If iLinha <> objGrid.iLinhasExistentes And iLinhaDeGrupo <> 1 Then
        
            'Reordena a numeração dos próximos cheques
            For iLinha2 = iLinha + 1 To objGrid.iLinhasExistentes
                                                    
                'Passa a linha do Grid para o Obj
                Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinha2)
        
                If objInfoParcPag.iSeqCheque > 1 Then
                    
                    objInfoParcPag.iSeqCheque = objInfoParcPag.iSeqCheque - 1
                                
                End If
                
            Next
            
        End If
                              
    'Se a Parcela foi marcada
    Else
            
        'Insere número de cheque 1 para a Parcela
        objInfoParcPag.iSeqCheque = 1
        
        If gobjChequesPag.colInfoParcPag.Count > 1 Then
        
            Call Titulos_Reposiciona_Emitir(iLinha)
         
        End If
        
    End If
    
    'Atualiza as parcelas na tela
    Call Traz_Dados_Tela
        
End Sub

Private Sub Titulos_Reposiciona_Emitir(iLinha As Integer)
'ajusta gobjChequesPag.colInfoParcPag para tratar uma parcela que estava desmarcada e passou a estar marcada p/emissao

Dim iRemovidos As Integer, iLinhaChequeUm As Integer
Dim iLinha1 As Integer
Dim iLinha2 As Integer
Dim objInfoParcPag As ClassInfoParcPag
Dim iLinhasExistentes As Integer

    iLinhasExistentes = gobjChequesPag.colInfoParcPag.Count
    
    'procura primeira parcela que já estava marcada p/imprimir
    iLinhaChequeUm = 0
    For iLinha1 = iLinha + 1 To iLinhasExistentes
                
        'Passa a linha do Grid para o Obj
        Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinha1)
        
        'Se a Parcela está desmarcada
        If objInfoParcPag.iSeqCheque <> 0 Then
            iLinhaChequeUm = iLinha1
            Exit For
        End If
                       
    Next
            
    If iLinhaChequeUm > 0 Then
    
        'jogar o obj de iLinha p/antes do iLinhaChequeUm
            
        If iLinha <> iLinhaChequeUm - 1 Then
        
            Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinha)
            
            'Remove a Parcela da colecao
            gobjChequesPag.colInfoParcPag.Remove Index:=iLinha
    
            'insere a Parcela como antes de iLinhaChequeUm
            gobjChequesPag.colInfoParcPag.Add Item:=objInfoParcPag, Before:=iLinhaChequeUm - 1
            
        End If
    
        'Reordena a numeração dos cheques que já haviam
        For iLinha2 = iLinhaChequeUm To iLinhasExistentes
                                                
            'Passa a linha do Grid para o Obj
            Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinha2)
        
            objInfoParcPag.iSeqCheque = objInfoParcPag.iSeqCheque + 1
            
        Next
    
    Else
    
        'jogar o obj de iLinha p/ultimo item da colecao
            
        If iLinha <> gobjChequesPag.colInfoParcPag.Count Then
        
            Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinha)
            
            'Remove a Parcela da colecao
            gobjChequesPag.colInfoParcPag.Remove Index:=iLinha
    
            'insere a Parcela como ultimo item
            gobjChequesPag.colInfoParcPag.Add objInfoParcPag
            
        End If
    
    End If
            
End Sub

Private Sub CheckSelecionar_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)
    
End Sub

Private Sub CheckSelecionar_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
    
End Sub

Private Sub CheckSelecionar_Validate(Cancel As Boolean)
    
Dim lErro As Long

    Set objGrid.objControle = CheckSelecionar
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True
        
End Sub

Private Sub CheckEmitir_GotFocus()
    
    Call Grid_Campo_Recebe_Foco(objGrid)
    
End Sub

Private Sub CheckEmitir_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)
    
End Sub

Private Sub CheckEmitir_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = CheckEmitir
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Visibilidade para versão LIGHT
    If giTipoVersao = VERSAO_FULL Then
        Juros.left = POSICAO_FORA_TELA
        Juros.TabStop = False
        Multa.left = POSICAO_FORA_TELA
        Multa.TabStop = False
        Desconto.left = POSICAO_FORA_TELA
        Desconto.TabStop = False
    End If

    If giTipoVersao = VERSAO_LIGHT Then
        FilialEmpresa.left = POSICAO_FORA_TELA
        FilialEmpresa.TabStop = False
    End If
    
    Set objGrid = New AdmGrid

    'Lê o código e a descrição de todos os Tipos de Cobrança
    lErro = CF("Cod_Nomes_Le", "TiposDeCobranca", "Codigo", "Descricao", STRING_TIPOSDECOBRANCA_DESCRICAO, gcolTiposDeCobranca)
    If lErro <> SUCESSO Then Error 28217

    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case 28217
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144563)
    
    End Select
    
     iAlterado = 0
    
    Exit Sub

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'Chama rotina de inicialização da saída de célula
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then

        Select Case objGridInt.objGrid.Col
                
            'Se a célula for o campo Multa
            Case iGrid_Multa_Col
                
                Set objGridInt.objControle = Multa
                
                'Chama função de tratamento de saída da célula Multa
                lErro = Saida_Celula_Multa(objGridInt)
                If lErro <> SUCESSO Then Error 57830
                
            'Se a célula for o campo Juros
            Case iGrid_Juros_Col
                
                Set objGridInt.objControle = Juros
                
                'Chama função de tratamento de saída da célula Juros
                lErro = Saida_Celula_Juros(objGridInt)
                If lErro <> SUCESSO Then Error 57831
            
            'Se a célula for o campo Desconto
            Case iGrid_Desconto_Col
                
                Set objGridInt.objControle = Desconto
                
                'Chama função de tratamento de saída da célula Desconto
                lErro = Saida_Celula_Desconto(objGridInt)
                If lErro <> SUCESSO Then Error 57832

        End Select
        
        'Chama função de finalização da saída de célula
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then Error 57833

        
    End If

    Saida_Celula = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula:

    Saida_Celula = Err

    Select Case Err

        Case 57830, 57831, 57832, 57833
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144564)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Multa(objGridInt As AdmGrid) As Long
'Rotina de saída da célula Multa

Dim lErro As Long
Dim iLinha As Integer
Dim objInfoParcPag As ClassInfoParcPag

On Error GoTo Erro_Saida_Celula_Multa

    'Formata o valor da multa na tela
    Multa.Text = Format(Multa.Text, "Standard")
    
    If Len(Trim(Multa.Text)) = 0 Then
        Multa.Text = Format(0, "Standard")
    Else
        'Critica se o valor é positivo
        lErro = Valor_NaoNegativo_Critica(Multa.Text)
        If lErro <> SUCESSO Then Error 57840
    End If
    
    'Passa para iLinha o número da linha em questão
    iLinha = GridChequesPag2.Row

    'Passa os dados da linha do Grid para o Obj
    Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinha)
        
    'Passa para o Obj o valor da multa que está na tela
    If Len(Trim(Multa.Text)) <> 0 Then
        objInfoParcPag.dValorMulta = CDbl(Multa.Text)
    Else
        objInfoParcPag.dValorMulta = 0
    End If
    
    'Calcula o Valor Total
    Call Calcula_Total(objInfoParcPag)
        
    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 57834

    Saida_Celula_Multa = SUCESSO

    Exit Function
    
Erro_Saida_Celula_Multa:

    Saida_Celula_Multa = Err
    
    Select Case Err

        Case 57834, 57840
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144565)
            
    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_Juros(objGridInt As AdmGrid) As Long
'Rotina de saída da célula Juros

Dim lErro As Long
Dim iLinha As Integer
Dim objInfoParcPag As ClassInfoParcPag

On Error GoTo Erro_Saida_Celula_Juros

    'Formata o valor dos Juros na tela
    Juros.Text = Format(Juros.Text, "Standard")
    
    If Len(Trim(Juros.Text)) = 0 Then
        Juros.Text = Format(0, "Standard")
    Else
        'Critica se o valor é positivo
        lErro = Valor_NaoNegativo_Critica(Juros.Text)
        If lErro <> SUCESSO Then Error 57839
    End If
    
    'Passa para iLinha o número da linha em questão
    iLinha = GridChequesPag2.Row

    'Passa os dados da linha do Grid para o Obj
    Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinha)
        
    'Passa para o Obj o valor dos Juros que está na tela
    If Len(Trim(Juros.Text)) <> 0 Then
        objInfoParcPag.dValorJuros = CDbl(Juros.Text)
    Else
        objInfoParcPag.dValorJuros = 0
    End If
    
    'Calcula o Valor Total
    Call Calcula_Total(objInfoParcPag)
        
    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 57835

    Saida_Celula_Juros = SUCESSO

    Exit Function
    
Erro_Saida_Celula_Juros:

    Saida_Celula_Juros = Err
    
    Select Case Err

        Case 57835, 57839
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144566)
            
    End Select
    
    Exit Function
    
End Function

Private Function Saida_Celula_Desconto(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim objInfoParcPag As ClassInfoParcPag

On Error GoTo Erro_Saida_Celula_Desconto

    'Formata o valor do desconto na tela
    Desconto.Text = Format(Desconto.Text, "Standard")
    
    If Len(Trim(Desconto.Text)) = 0 Then
        Desconto.Text = Format(0, "Standard")
    Else
        'Critica se valor é positivo
        lErro = Valor_NaoNegativo_Critica(Desconto.Text)
        If lErro <> SUCESSO Then Error 57838
    End If
        
    'Passa para iLinha o número da linha em questão
    iLinha = GridChequesPag2.Row

    'Passa os dados da linha do Grid para o Obj
    Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinha)
        
    'Passa para o Obj o valor do desconto que está na tela
    If Len(Trim(Desconto.Text)) <> 0 Then
        objInfoParcPag.dValorDesconto = CDbl(Desconto.Text)
    Else
        objInfoParcPag.dValorDesconto = 0
    End If
    
    'Verifica se o Desconto não é maior que o valor Total da parcela, com Juros e Multa
    If objInfoParcPag.dValorDesconto > (objInfoParcPag.dValor + objInfoParcPag.dValorJuros + objInfoParcPag.dValorMulta) Then Error 57836
    
    'Calcula o Valor Total
    Call Calcula_Total(objInfoParcPag)
        
    'Chama função de saída de célula no Grid
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then Error 57837

    Saida_Celula_Desconto = SUCESSO

    Exit Function
    
Erro_Saida_Celula_Desconto:

    Saida_Celula_Desconto = Err
    
    Select Case Err

        Case 57836
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VALORPARCELA_MENOR_DESCONTO", Err)
            Desconto.SetFocus
        
        Case 57837, 57838
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                                            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144567)
            
    End Select
    
    Exit Function
    
End Function

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGrid = Nothing
    
    Set gobjChequesPag = Nothing
    Set gcolTiposDeCobranca = Nothing

End Sub

Private Sub Juros_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Juros_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Juros_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Juros
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Multa_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)
    
End Sub

Private Sub Multa_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Multa_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Multa
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub Desconto_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGrid)

End Sub

Private Sub Desconto_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGrid)

End Sub

Private Sub Desconto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGrid.objControle = Desconto
    lErro = Grid_Campo_Libera_Foco(objGrid)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Private Sub GridChequesPag2_Click()
    
Dim iExecutaEntradaCelula As Integer
    
    Call Grid_Click(objGrid, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If
    
End Sub

Private Sub GridChequesPag2_GotFocus()
    
    Call Grid_Recebe_Foco(objGrid)

End Sub

Private Sub GridChequesPag2_EnterCell()
    
    Call Grid_Entrada_Celula(objGrid, iAlterado)
    
End Sub

Private Sub GridChequesPag2_LeaveCell()
    
    Call Saida_Celula(objGrid)
    
End Sub

Private Sub GridChequesPag2_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGrid)
    
End Sub

Private Sub GridChequesPag2_KeyPress(KeyAscii As Integer)
    
Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGrid, iExecutaEntradaCelula)
    
    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGrid, iAlterado)
    End If

End Sub

Private Sub GridChequesPag2_Validate(Cancel As Boolean)
    
    Call Grid_Libera_Foco(objGrid)

End Sub

Private Sub GridChequesPag2_RowColChange()

    Call Grid_RowColChange(objGrid)
       
End Sub

Private Sub GridChequesPag2_Scroll()

    Call Grid_Scroll(objGrid)
    
End Sub

Private Function Inicializa_Grid_ChequesPag2(objGridInt As AdmGrid, iNumLinhas As Integer) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Inicializa_Grid_ChequesPag2
    
    'tela em questão
    Set objGrid.objForm = Me
    
    'titulos do grid
    objGridInt.colColuna.Add ("    ")
    objGridInt.colColuna.Add ("Cheque")
    objGridInt.colColuna.Add ("Fornecedor")
    objGridInt.colColuna.Add ("Filial")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Nº Título")
    objGridInt.colColuna.Add ("Parcela")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Emitir")
    objGridInt.colColuna.Add ("Selecionar")
    objGridInt.colColuna.Add ("Vencimento")
    objGridInt.colColuna.Add ("Portador")
    
    If giTipoVersao = VERSAO_LIGHT Then
        objGridInt.colColuna.Add ("Juros")
        objGridInt.colColuna.Add ("Multa")
        objGridInt.colColuna.Add ("Desconto")
    End If
    
    objGridInt.colColuna.Add ("Tipo De Cobrança")
    
    If giTipoVersao = VERSAO_FULL Then
        objGridInt.colColuna.Add ("Filial Empresa")
    End If
    
    'campos de edição do grid
    objGridInt.colCampo.Add (Cheque.Name)
    objGridInt.colCampo.Add (Fornecedor.Name)
    objGridInt.colCampo.Add (Filial.Name)
    objGridInt.colCampo.Add (Tipo.Name)
    objGridInt.colCampo.Add (NumTitulo.Name)
    objGridInt.colCampo.Add (Parcela.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (CheckEmitir.Name)
    objGridInt.colCampo.Add (CheckSelecionar.Name)
    objGridInt.colCampo.Add (DataVencto.Name)
    objGridInt.colCampo.Add (Portador.Name)
        
    If giTipoVersao = VERSAO_LIGHT Then
        objGridInt.colCampo.Add (Juros.Name)
        objGridInt.colCampo.Add (Multa.Name)
        objGridInt.colCampo.Add (Desconto.Name)
    End If
    
    objGridInt.colCampo.Add (TipoCobranca.Name)
    
    If giTipoVersao = VERSAO_FULL Then
        objGridInt.colCampo.Add (FilialEmpresa.Name)
    End If
    
    iGrid_Cheque_Col = 1
    iGrid_Fornecedor_Col = 2
    iGrid_Filial_Col = 3
    iGrid_Tipo_Col = 4
    iGrid_NumTitulo_Col = 5
    iGrid_Parcela_Col = 6
    iGrid_Valor_Col = 7
    iGrid_Emitir_Col = 8
    iGrid_Selecionar_Col = 9
    iGrid_Vencimento = 10
    iGrid_Portador_Col = 11
    
    If giTipoVersao = VERSAO_LIGHT Then
        iGrid_Juros_Col = 12
        iGrid_Multa_Col = 13
        iGrid_Desconto_Col = 14
        iGrid_TipoCobranca_Col = 15
    ElseIf giTipoVersao = VERSAO_FULL Then
        iGrid_TipoCobranca_Col = 12
        iGrid_FilialEmpresa_Col = 13
    End If
        
    objGridInt.objGrid = GridChequesPag2
    
    'todas as linhas do grid
    If iNumLinhas >= 10 Then
        objGridInt.objGrid.Rows = iNumLinhas + 1
    Else
        objGridInt.objGrid.Rows = 11
    End If
    
    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 10
        
    GridChequesPag2.ColWidth(0) = 300
    
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoIncluir = PROIBIDO_INCLUIR
    
    objGridInt.iProibidoExcluir = PROIBIDO_EXCLUIR
    
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_ChequesPag2 = SUCESSO
    
    Exit Function
    
Erro_Inicializa_Grid_ChequesPag2:

    Inicializa_Grid_ChequesPag2 = Err
    
    Select Case Err
    
        Case 14251
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144568)
        
    End Select

    Exit Function
        
End Function

Function Trata_Parametros(Optional objChequesPag As ClassChequesPag) As Long
'Traz os dados das Parcelas a pagar para a Tela

Dim objInfoParcPag As ClassInfoParcPag
Dim iLinha As Integer, lErro As Long
Dim iNumCheques As Integer
Dim iCheque As Integer, objCtaCorrenteInt As New ClassContasCorrentesInternas
Dim dValorTotalCheques As Double
Dim dValorTotalTitulos As Double

On Error GoTo Erro_Trata_Parametros
    
    Set gobjChequesPag = objChequesPag
    
    'Passa a Conta Corrente para a tela
    
    objCtaCorrenteInt.iCodigo = gobjChequesPag.iCta
    'Lê a Conta Corrente
    lErro = CF("ContaCorrenteInt_Le", objCtaCorrenteInt.iCodigo, objCtaCorrenteInt)
    If lErro <> SUCESSO And lErro <> 11807 Then Error 56706
    
    Conta.Caption = CStr(objCtaCorrenteInt.iCodigo) & SEPARADOR & objCtaCorrenteInt.sNomeReduzido
        
    lErro = Inicializa_Grid_ChequesPag2(objGrid, gobjChequesPag.colInfoParcPag.Count)
    If lErro <> SUCESSO Then Error 14250
           
    Call Grid_Preenche
    
    Trata_Parametros = SUCESSO
    
    Exit Function
    
Erro_Trata_Parametros:

    Trata_Parametros = Err
    
    Select Case Err
    
        Case 14250, 56706
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 144569)
    
    End Select
    
    iAlterado = 0
    
    Exit Function
    
End Function

Private Function Titulos_Reposiciona_Agrupar() As Long
'Agrupa no Obj as Parcelas selecionadas, que estiverem marcadas para emissão

Dim objInfoParcPag As ClassInfoParcPag
Dim iNumCheque As Integer
Dim iPrimeiroTitulo As Integer
Dim iLinha As Integer
Dim iLinha2 As Integer
    
    iPrimeiroTitulo = 0
    iNumCheque = 0
    
    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes
        
        'Se a Parcela está marcada
        If GridChequesPag2.TextMatrix(iLinha, iGrid_Selecionar_Col) = SELECIONAR_CHECADO Then
    
            'Desmarca a Parcela
            GridChequesPag2.TextMatrix(iLinha, iGrid_Selecionar_Col) = SELECIONAR_NAO_CHECADO
            
            'Passa a linha do Grid para o Obj
            Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinha)
            
            'Lê o número de cheque da primeira Parcela
            If iPrimeiroTitulo = 0 And objInfoParcPag.iSeqCheque <> 0 Then
                
                iPrimeiroTitulo = iLinha
                iNumCheque = objInfoParcPag.iSeqCheque
                Call Obtem_Ult_Linha_Grupo(iPrimeiroTitulo, objInfoParcPag.iSeqCheque)
            End If
            
            'Se não for a primeira Parcela marcada
            If objInfoParcPag.iSeqCheque > iNumCheque Then
            
                'Agrupa a Parcela de acordo com o número de cheque da primeira Parcela
                objInfoParcPag.iSeqCheque = iNumCheque
                
                'Remove a Parcela do Obj
                gobjChequesPag.colInfoParcPag.Remove Index:=iLinha
                
                'Insere a Parcela após a primeira Parcela selecionada
                gobjChequesPag.colInfoParcPag.Add Item:=objInfoParcPag, After:=iPrimeiroTitulo
                
                'Se não for a última Parcela do Grid
                If iLinha <> objGrid.iLinhasExistentes Then
                                    
                    'Reordena a numeração dos próximos cheques
                    For iLinha2 = iLinha + 1 To objGrid.iLinhasExistentes
                                                
                        'Passa a linha do Grid para o Obj
                        Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinha2)
    
                        'Se o número de cheque da Parcela for maior que o número de cheque do grupo em questão (resultante do agrupamento)
                        If objInfoParcPag.iSeqCheque > iNumCheque Then objInfoParcPag.iSeqCheque = objInfoParcPag.iSeqCheque - 1
                                            
                    Next
                    
                End If
                
            End If
            
        End If
            
    Next
                            
    Exit Function

End Function

Private Function Traz_Dados_Tela() As Long
'Atualiza as parcelas na tela

Dim iLinhaNoTopo As Integer
    
    GridChequesPag2.Redraw = False

    iLinhaNoTopo = GridChequesPag2.TopRow
    
    'Limpa o Grid
    Call Grid_Limpa(objGrid)
    
    Call Grid_Preenche

    GridChequesPag2.TopRow = iLinhaNoTopo
    
    GridChequesPag2.Redraw = True
    
    Exit Function

End Function

Private Function Titulos_Reposiciona_Desagrupar() As Long
'Desagrupa no Obj as Parcelas selecionadas, que estiverem marcadas para emissão

Dim objInfoParcPag As ClassInfoParcPag, iChequeAnt As Integer, iChequesCriados As Integer
Dim iLinha As Integer, iLinha2 As Integer, iSeqCheque As Integer, iChequeOriginal As Integer
    
    iChequesCriados = 0
    
    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes
        
        iChequeOriginal = Linha_ObtemSeqCheque(iLinha)
        
        'Se a Parcela está marcada
        If GridChequesPag2.TextMatrix(iLinha, iGrid_Selecionar_Col) = SELECIONAR_CHECADO And Linha_ObtemSeqCheque(iLinha) <> 0 Then
        
            'Desmarca a Parcela
            GridChequesPag2.TextMatrix(iLinha, iGrid_Selecionar_Col) = SELECIONAR_NAO_CHECADO

            'se a linha anterior era do mesmo cheque
            If iChequeAnt = iChequeOriginal Then
                
                'Passa a linha do Grid para o Obj
                Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinha)
            
                iChequesCriados = 1
                
                objInfoParcPag.iSeqCheque = iChequeAnt + 1
                
            End If
            
            'passa pelas outras as linhas selecionadas criando um cheque p/cada
            For iLinha2 = iLinha + 1 To objGrid.iLinhasExistentes
            
                Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinha2)
                    
                If GridChequesPag2.TextMatrix(iLinha2, iGrid_Selecionar_Col) = SELECIONAR_CHECADO Then
                    
                    'Desmarca a Parcela
                    GridChequesPag2.TextMatrix(iLinha2, iGrid_Selecionar_Col) = SELECIONAR_NAO_CHECADO
                    
                    iChequesCriados = iChequesCriados + 1
                    
                    objInfoParcPag.iSeqCheque = objInfoParcPag.iSeqCheque + iChequesCriados
                
                Else 'achou o 1o deselecionado
                
                    'se era do mesmo cheque que sofreu o "desagrupamento"
                    If objInfoParcPag.iSeqCheque = iChequeOriginal Then
                    
                        iChequesCriados = iChequesCriados + 1
                        
                        objInfoParcPag.iSeqCheque = objInfoParcPag.iSeqCheque + iChequesCriados
                        
                        iLinha2 = iLinha2 + 1
                    
                    End If
                    
                    Exit For
                        
                End If
                
            Next
                
            'Reordena a numeração dos próximos cheques
            Do While iLinha2 <= objGrid.iLinhasExistentes
                                                    
                'Passa a linha do Grid para o Obj
                Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinha2)
            
                objInfoParcPag.iSeqCheque = objInfoParcPag.iSeqCheque + iChequesCriados
                    
                iLinha2 = iLinha2 + 1
                
            Loop
            
            Exit For
            
        End If
        
        iChequeAnt = Linha_ObtemSeqCheque(iLinha)
        
    Next
    
    Exit Function

End Function

Private Function Titulos_Reposiciona_Subir() As Long
'Posiciona os títulos do grupo selecionado acima do grupo anterior

Dim objInfoParcPag As ClassInfoParcPag
Dim iLinha As Integer
Dim iNumCheque As Integer
    
    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes
        
        'Se a Parcela está marcada
        If GridChequesPag2.TextMatrix(iLinha, iGrid_Selecionar_Col) = SELECIONAR_CHECADO Then
                
            'Passa a linha do Grid para o Obj
            Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinha)
            
            'Guarda o número de cheque do grupo selecionado
            iNumCheque = objInfoParcPag.iSeqCheque
        
            'Desmarca a Parcela
            GridChequesPag2.TextMatrix(iLinha, iGrid_Selecionar_Col) = SELECIONAR_NAO_CHECADO
            
        End If
        
    Next
    
    Call Sobe_Grupo(iNumCheque)

End Function

Private Function Titulos_Reposiciona_Descer() As Long
'Posiciona os títulos do grupo selecionado abaixo do grupo posterior

Dim objInfoParcPag As ClassInfoParcPag
Dim iLinha As Integer
Dim iNumCheque As Integer
    
    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes
        
        'Se a Parcela está marcada
        If GridChequesPag2.TextMatrix(iLinha, iGrid_Selecionar_Col) = SELECIONAR_CHECADO Then
                
            'Passa a linha do Grid para o Obj
            Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinha)
            
            'Guarda o número de cheque do grupo selecionado
            iNumCheque = objInfoParcPag.iSeqCheque
        
            'Desmarca a Parcela
            GridChequesPag2.TextMatrix(iLinha, iGrid_Selecionar_Col) = SELECIONAR_NAO_CHECADO
            
        End If
        
    Next
    
    Call Sobe_Grupo(iNumCheque + 1)

End Function

Private Sub Sobe_Grupo(iNumCheque As Integer)

Dim objInfoParcPag As ClassInfoParcPag
Dim iLinha As Integer
Dim iLinhaGrupoAnterior As Integer
Dim iPrimeiroTitulo As Integer
Dim iNumTitulos As Integer

    iLinhaGrupoAnterior = 0
    iPrimeiroTitulo = 0
    iNumTitulos = 0
    
    iLinha = 0
    
    'Troca a numeração de cheque das Parcelas
    For Each objInfoParcPag In gobjChequesPag.colInfoParcPag

        iLinha = iLinha + 1
        
        'Passa a linha do Grid para o Obj
        Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinha)
      
        'Se a Parcela está antes do grupo marcado
        If objInfoParcPag.iSeqCheque = iNumCheque - 1 Then
                                
            'Lê a linha da primeira Parcela do grupo
            If iLinhaGrupoAnterior = 0 Then iLinhaGrupoAnterior = iLinha
            
            'Altera o número de cheque da Parcela
            objInfoParcPag.iSeqCheque = objInfoParcPag.iSeqCheque + 1
            
        'Se a Parcela está no grupo marcado
        ElseIf objInfoParcPag.iSeqCheque = iNumCheque Then
        
            'Lê a linha da primeira Parcela do grupo
            If iPrimeiroTitulo = 0 Then iPrimeiroTitulo = iLinha
            
            'Altera o número de cheque da Parcela
            objInfoParcPag.iSeqCheque = objInfoParcPag.iSeqCheque - 1
            
            'Faz o somatório do número de Parcelas do grupo selecionado
            iNumTitulos = iNumTitulos + 1
            
        End If
        
    Next

    'Reposiciona as Parcelas
    For iLinha = iPrimeiroTitulo To iPrimeiroTitulo - 1 + iNumTitulos
        
        'Passa a linha do Grid para o Obj
        Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinha)
        
        'Remove a Parcela do Obj
        gobjChequesPag.colInfoParcPag.Remove Index:=iLinha

        'Insere a Parcela antes do grupo anterior
        gobjChequesPag.colInfoParcPag.Add Item:=objInfoParcPag, Before:=iLinhaGrupoAnterior
        
        iLinhaGrupoAnterior = iLinhaGrupoAnterior + 1
    
    Next
    
End Sub

Private Sub Obtem_Ult_Linha_Grupo(iLinha As Integer, iSeqCheque As Integer)
'retorna em iLinha a ultima linha do grupo que tem iSeqCheque como sequencial de cheque

Dim objInfoParcPag As ClassInfoParcPag
Dim iLinhaTeste As Integer

    iLinhaTeste = iLinha
    
    Do While iLinhaTeste <= objGrid.iLinhasExistentes
    
        Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinhaTeste)
    
        If objInfoParcPag.iSeqCheque <> iSeqCheque Then
        
            iLinhaTeste = iLinhaTeste - 1
            Exit Do
            
        End If
        
        iLinhaTeste = iLinhaTeste + 1
        
    Loop

    iLinha = iLinhaTeste
    
End Sub

Private Sub Grid_Preenche()

Dim iLinha As Integer
Dim objInfoParcPag As ClassInfoParcPag
Dim iNumCheques As Integer
Dim dValorTotalCheques As Double
Dim dValorTotalTitulos As Double
Dim dTotalJurosTitulos As Double
Dim dTotalMultasTitulos As Double
Dim dTotalDescontosTitulos As Double
Dim dTotalPagarTitulos As Double
Dim objCodDescricao As AdmCodigoNome
Dim iTotalTitulosSelecionados As Integer
Dim colNumTitulo As New Collection
Dim objFilialEmpresa As New AdmFiliais
Dim lErro  As Long, iFilialEmpresaAnt As Integer

On Error GoTo Erro_Grid_Preenche

    iFilialEmpresaAnt = -1
    iLinha = 0
    iNumCheques = 0
    dValorTotalCheques = 0
    dValorTotalTitulos = 0
    
    'Percorre todas as Parcelas da Coleção
    For Each objInfoParcPag In gobjChequesPag.colInfoParcPag

        With objInfoParcPag
        
            iLinha = iLinha + 1
    
            'Passa para a tela os dados da Parcela em questão
            GridChequesPag2.TextMatrix(iLinha, iGrid_Cheque_Col) = IIf(.iSeqCheque <> 0, .iSeqCheque + gobjChequesPag.lNumCheque - 1, "")
            GridChequesPag2.TextMatrix(iLinha, iGrid_Fornecedor_Col) = .sNomeRedForn
            GridChequesPag2.TextMatrix(iLinha, iGrid_Filial_Col) = .iFilialForn
            GridChequesPag2.TextMatrix(iLinha, iGrid_Tipo_Col) = .sSiglaDocumento
            GridChequesPag2.TextMatrix(iLinha, iGrid_NumTitulo_Col) = .lNumTitulo
            GridChequesPag2.TextMatrix(iLinha, iGrid_Parcela_Col) = .iNumParcela
            GridChequesPag2.TextMatrix(iLinha, iGrid_Valor_Col) = Format(.dValor, "Standard")
            GridChequesPag2.TextMatrix(iLinha, iGrid_Vencimento) = Format(.dtDataVencimento, "dd/mm/yyyy")
            GridChequesPag2.TextMatrix(iLinha, iGrid_Portador_Col) = .sNomeRedPortador
                            
            If giTipoVersao = VERSAO_LIGHT Then
                GridChequesPag2.TextMatrix(iLinha, iGrid_Juros_Col) = Format(objInfoParcPag.dValorJuros, "Standard")
                GridChequesPag2.TextMatrix(iLinha, iGrid_Multa_Col) = Format(objInfoParcPag.dValorMulta, "Standard")
                GridChequesPag2.TextMatrix(iLinha, iGrid_Desconto_Col) = Format(objInfoParcPag.dValorDesconto, "Standard")
            End If
            
            If .iTipoCobranca <> 0 Then
                For Each objCodDescricao In gcolTiposDeCobranca
                    If objCodDescricao.iCodigo = .iTipoCobranca Then
                        GridChequesPag2.TextMatrix(iLinha, iGrid_TipoCobranca_Col) = .iTipoCobranca & SEPARADOR & objCodDescricao.sNome
                        Exit For
                    End If
                Next
            End If
            
            If giTipoVersao = VERSAO_FULL Then
            
                If iFilialEmpresaAnt <> .iFilialEmpresa Then
                
                    objFilialEmpresa.iCodFilial = .iFilialEmpresa
                    
                    lErro = CF("FilialEmpresa_Le", objFilialEmpresa)
                    If lErro <> SUCESSO Then gError 82783
                                
                    iFilialEmpresaAnt = .iFilialEmpresa
                                    
                End If
                
                GridChequesPag2.TextMatrix(iLinha, iGrid_FilialEmpresa_Col) = objFilialEmpresa.iCodFilial & SEPARADOR & objFilialEmpresa.sNome
                        
            End If
            
            'Se a Parcela tiver número de cheque associado
            If .iSeqCheque <> 0 Then
                
                'Marca a Parcela para emitir
                GridChequesPag2.TextMatrix(iLinha, iGrid_Emitir_Col) = EMITIR_CHECADO
                
                'Faz o somatório do valor total das Parcelas
                dValorTotalTitulos = dValorTotalTitulos + .dValor
                dTotalMultasTitulos = dTotalMultasTitulos + .dValorMulta
                dTotalJurosTitulos = dTotalJurosTitulos + .dValorJuros
                dTotalDescontosTitulos = dTotalDescontosTitulos + .dValorDesconto
                
                Call Conta_Titulos(objInfoParcPag, colNumTitulo, iTotalTitulosSelecionados)
                
                colNumTitulo.Add objInfoParcPag
                
            'Se a Parcela não tiver número de cheque associado
            Else
            
                'Desmarca a Parcela para emitir
                GridChequesPag2.TextMatrix(iLinha, iGrid_Emitir_Col) = EMITIR_NAO_CHECADO
                
            End If
            
            'Faz o somatório do número de cheques
            If iNumCheques < .iSeqCheque Then
                iNumCheques = .iSeqCheque
            End If
                                    
            'Faz o somatório do valor total dos cheques
            If .iSeqCheque <> 0 Then dValorTotalCheques = dValorTotalCheques + .dValor
                   
        End With
                   
    Next
            
    'Passa para o Obj o número de Parcelas passadas pela Coleção
    objGrid.iLinhasExistentes = iLinha
    
    'Passa para a tela a Qtd de Títulos, Valor Total dos Títulos, Qtd de Cheques e Valor Total dos Cheques
    QtdTitulos.Caption = CStr(iTotalTitulosSelecionados)
    QtdCheques.Caption = CStr(iNumCheques)
    ValorTotalTitulos.Caption = CStr(Format(dValorTotalTitulos, "Standard"))
        
    If giTipoVersao = VERSAO_LIGHT Then
        
        'Passa para a tela o somatório das parcelas com multas, juros e descontos
        dTotalPagarTitulos = dValorTotalTitulos + dTotalJurosTitulos + dTotalMultasTitulos - dTotalDescontosTitulos
        ValorTotalCheques.Caption = CStr(Format(dTotalPagarTitulos, "Standard"))
    
    ElseIf giTipoVersao = VERSAO_FULL Then
        
        'Passa para a tela apenas o somatório das parcelas
        ValorTotalCheques.Caption = CStr(Format(dValorTotalTitulos, "Standard"))
    
    End If
    
    'Atualiza as checkboxes
    Call Grid_Refresh_Checkbox(objGrid)
    
    Exit Sub
    
Erro_Grid_Preenche:

    Select Case gErr
    
        Case 82783
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 144570)
            
    End Select
    
    Exit Sub
    
End Sub

Private Sub Conta_Titulos(objInfoParcPag As ClassInfoParcPag, colNumTitulo As Collection, iTotalTitulosSelecionados As Integer)

Dim objInfoParcPag2 As ClassInfoParcPag
Dim lNumTitulo As Long
    
    'Conta o número de Títulos selecionados que estão no Grid
    For Each objInfoParcPag2 In colNumTitulo
                
        If objInfoParcPag.lNumTitulo = objInfoParcPag2.lNumTitulo And objInfoParcPag.sSiglaDocumento = objInfoParcPag2.sSiglaDocumento And objInfoParcPag.lFornecedor = objInfoParcPag2.lFornecedor And objInfoParcPag.iFilialForn = objInfoParcPag2.iFilialForn Then
            lNumTitulo = objInfoParcPag.lNumTitulo
            Exit For
        End If
        
    Next

    If lNumTitulo = 0 Then
        iTotalTitulosSelecionados = iTotalTitulosSelecionados + 1
    End If

End Sub

Private Sub SelecionaParcelasCheque(iCheque As Integer)

Dim iLinha As Integer, objInfoParcPag As ClassInfoParcPag

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGrid.iLinhasExistentes
        
        'Passa a linha do Grid para o Obj
        Set objInfoParcPag = gobjChequesPag.colInfoParcPag.Item(iLinha)
        
        If objInfoParcPag.iSeqCheque = iCheque Then
        
            'marca a parcela
            GridChequesPag2.TextMatrix(iLinha, iGrid_Selecionar_Col) = SELECIONAR_CHECADO
                
        End If
        
    Next

End Sub

Private Function Calcula_Total(objInfoParcPag As ClassInfoParcPag) As Long
'Atualiza os labels com totais (selecionados) dos titulos

Dim dTotalTitulos As Double
Dim dTotalJuros As Double
Dim dTotalMultas As Double
Dim dTotalDescontos As Double
Dim iIndice As Integer
    
    For Each objInfoParcPag In gobjChequesPag.colInfoParcPag

        iIndice = iIndice + 1
        
        'Se a parcela em questão está checada
        If GridChequesPag2.TextMatrix(iIndice, iGrid_Emitir_Col) = EMITIR_CHECADO Then

            'Faz o somatório do Total dos Títulos
            dTotalTitulos = dTotalTitulos + objInfoParcPag.dValor
            dTotalJuros = dTotalJuros + objInfoParcPag.dValorJuros
            dTotalMultas = dTotalMultas + objInfoParcPag.dValorMulta
            dTotalDescontos = dTotalDescontos + objInfoParcPag.dValorDesconto
    
        End If

    Next
    
    'Atualiza na tela os somatórios das parcelas
    ValorTotalCheques.Caption = CStr(Format(dTotalTitulos + dTotalMultas + dTotalJuros - dTotalDescontos, "Standard"))
    
End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object
    
    Parent.HelpContextID = IDH_IMPRESSAO_CHEQUES_P2
    Set Form_Load_Ocx = Me
    Caption = "Impressão de cheques - Passo 2"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ChequesPag2"
    
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

'***** fim do trecho a ser copiado ******

Private Function Linha_ObtemSeqCheque(iLinha As Integer) As Long
'retorna o sequencial do cheque correspondente a linha selecionada
'O sequencial 1 corresponde ao 1o cheque a ser impresso.

Dim lNumCheque As Long

    lNumCheque = StrParaLong(GridChequesPag2.TextMatrix(iLinha, iGrid_Cheque_Col))

    Linha_ObtemSeqCheque = IIf(lNumCheque = 0, 0, lNumCheque - gobjChequesPag.lNumCheque + 1)
    
End Function

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
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

Private Sub QtdTitulos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QtdTitulos, Source, X, Y)
End Sub

Private Sub QtdTitulos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QtdTitulos, Button, Shift, X, Y)
End Sub

Private Sub QtdCheques_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(QtdCheques, Source, X, Y)
End Sub

Private Sub QtdCheques_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(QtdCheques, Button, Shift, X, Y)
End Sub

Private Sub ValorTotalCheques_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorTotalCheques, Source, X, Y)
End Sub

Private Sub ValorTotalCheques_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorTotalCheques, Button, Shift, X, Y)
End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label6, Source, X, Y)
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label6, Button, Shift, X, Y)
End Sub

Private Sub ValorTotalTitulos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(ValorTotalTitulos, Source, X, Y)
End Sub

Private Sub ValorTotalTitulos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(ValorTotalTitulos, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Conta_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Conta, Source, X, Y)
End Sub

Private Sub Conta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Conta, Button, Shift, X, Y)
End Sub

