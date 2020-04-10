VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl OCArtlux 
   ClientHeight    =   7620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12795
   KeyPreview      =   -1  'True
   ScaleHeight     =   7620
   ScaleMode       =   0  'User
   ScaleWidth      =   12795
   Begin VB.Frame FrameEtapa 
      Height          =   6405
      Index           =   0
      Left            =   30
      TabIndex        =   15
      Top             =   540
      Width           =   12710
      Begin VB.TextBox CTipoCouro 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   2700
         TabIndex        =   40
         Top             =   1170
         Width           =   1200
      End
      Begin VB.TextBox CEstSeguranca 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   2025
         TabIndex        =   24
         Top             =   2415
         Width           =   780
      End
      Begin VB.TextBox CQuantPV 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   990
         TabIndex        =   23
         Top             =   2415
         Width           =   780
      End
      Begin VB.TextBox CQuantEst 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   7815
         TabIndex        =   22
         Top             =   1950
         Width           =   780
      End
      Begin VB.TextBox CPrioridade 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   6975
         TabIndex        =   21
         Top             =   1935
         Width           =   465
      End
      Begin VB.TextBox CForro 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   5310
         TabIndex        =   20
         Top             =   1935
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.TextBox CCorte 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   4245
         TabIndex        =   19
         Top             =   1935
         Width           =   910
      End
      Begin VB.TextBox CQuantidade 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   3345
         TabIndex        =   18
         Top             =   1935
         Width           =   630
      End
      Begin VB.TextBox CDescricao 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   1425
         TabIndex        =   17
         Top             =   2895
         Width           =   4030
      End
      Begin VB.TextBox CProduto 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   285
         TabIndex        =   16
         Top             =   1935
         Width           =   1800
      End
      Begin MSFlexGridLib.MSFlexGrid GridC 
         Height          =   1275
         Left            =   15
         TabIndex        =   8
         Top             =   195
         Width           =   12660
         _ExtentX        =   22331
         _ExtentY        =   2249
         _Version        =   393216
         Rows            =   8
         Cols            =   6
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   2
      End
   End
   Begin VB.Frame FrameEtapa 
      Caption         =   "Montagem"
      Height          =   5820
      Index           =   1
      Left            =   30
      TabIndex        =   25
      Top             =   1110
      Visible         =   0   'False
      Width           =   12710
      Begin VB.TextBox MTipoCouro 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   2700
         TabIndex        =   39
         Top             =   1110
         Width           =   1200
      End
      Begin VB.TextBox MQuantProd 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   1185
         TabIndex        =   38
         Top             =   1965
         Width           =   660
      End
      Begin VB.TextBox MQuantPreProd 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   3405
         TabIndex        =   37
         Top             =   1950
         Width           =   615
      End
      Begin VB.TextBox MMontagem 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   4950
         TabIndex        =   36
         Top             =   3735
         Width           =   810
      End
      Begin VB.TextBox MProduto 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   675
         TabIndex        =   34
         Top             =   2925
         Width           =   1800
      End
      Begin VB.TextBox MDescricao 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   735
         TabIndex        =   33
         Top             =   2535
         Width           =   2130
      End
      Begin VB.TextBox MQuantidade 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   2265
         TabIndex        =   32
         Top             =   1935
         Width           =   555
      End
      Begin VB.TextBox MCorte 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   5055
         TabIndex        =   31
         Top             =   2730
         Width           =   810
      End
      Begin VB.TextBox MForro 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   5025
         TabIndex        =   30
         Top             =   3285
         Width           =   810
      End
      Begin VB.TextBox MPrioridade 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   7425
         TabIndex        =   29
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox MQuantEst 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   3525
         TabIndex        =   28
         Top             =   1620
         Width           =   675
      End
      Begin VB.TextBox MQuantPV 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   1170
         TabIndex        =   27
         Top             =   1545
         Width           =   600
      End
      Begin VB.TextBox MEstSeguranca 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   245
         Left            =   2100
         TabIndex        =   26
         Top             =   1575
         Width           =   600
      End
      Begin MSFlexGridLib.MSFlexGrid GridM 
         Height          =   1275
         Left            =   15
         TabIndex        =   35
         Top             =   195
         Width           =   12665
         _ExtentX        =   22331
         _ExtentY        =   2249
         _Version        =   393216
         Rows            =   8
         Cols            =   6
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   2
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detalhe"
      Height          =   630
      Left            =   30
      TabIndex        =   41
      Top             =   6930
      Width           =   12710
      Begin VB.Label Descricao 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4800
         TabIndex        =   45
         Top             =   195
         Width           =   7830
      End
      Begin VB.Label Label7 
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
         Height          =   225
         Left            =   3795
         TabIndex        =   44
         Top             =   255
         Width           =   975
      End
      Begin VB.Label Produto 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1125
         TabIndex        =   43
         Top             =   210
         Width           =   2415
      End
      Begin VB.Label Label1 
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
         Height          =   225
         Left            =   330
         TabIndex        =   42
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtros"
      Height          =   570
      Index           =   1
      Left            =   15
      TabIndex        =   10
      Top             =   -15
      Width           =   12725
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   12210
         Picture         =   "OCArtlux.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Fechar"
         Top             =   150
         Width           =   420
      End
      Begin VB.CommandButton BotaoAtualizar 
         Height          =   360
         Left            =   11730
         Picture         =   "OCArtlux.ctx":017E
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Atualizar"
         Top             =   150
         Width           =   420
      End
      Begin VB.ComboBox FGrupo 
         Height          =   315
         Left            =   9030
         Sorted          =   -1  'True
         TabIndex        =   46
         Top             =   165
         Width           =   1530
      End
      Begin VB.ComboBox FReferencia 
         Height          =   315
         Left            =   1170
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   1125
      End
      Begin VB.ComboBox FCor 
         Height          =   315
         Left            =   2625
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   1830
      End
      Begin VB.ComboBox FCorte 
         Height          =   315
         Left            =   4950
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   1485
      End
      Begin VB.ComboBox FForro 
         Height          =   315
         Left            =   6930
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   180
         Width           =   1485
      End
      Begin VB.CommandButton BotaoLimparFiltros 
         Caption         =   "Limpar Filtros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   10575
         TabIndex        =   7
         Top             =   120
         Width           =   1080
      End
      Begin VB.Label Label6 
         Caption         =   "Grupo:"
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
         Left            =   8430
         TabIndex        =   47
         Top             =   210
         Width           =   990
      End
      Begin VB.Label Label5 
         Caption         =   "Forro:"
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
         Left            =   6435
         TabIndex        =   14
         Top             =   225
         Width           =   990
      End
      Begin VB.Label Label4 
         Caption         =   "Corte:"
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
         Left            =   4440
         TabIndex        =   13
         Top             =   225
         Width           =   990
      End
      Begin VB.Label Label3 
         Caption         =   "Cor:"
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
         Left            =   2295
         TabIndex        =   12
         Top             =   225
         Width           =   990
      End
      Begin VB.Label Label2 
         Caption         =   "Referência:"
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
         Left            =   150
         TabIndex        =   11
         Top             =   225
         Width           =   990
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Etapa"
      Height          =   600
      Index           =   0
      Left            =   30
      TabIndex        =   9
      Top             =   -30
      Visible         =   0   'False
      Width           =   11460
      Begin VB.Timer Timer1 
         Interval        =   10000
         Left            =   8970
         Top             =   120
      End
      Begin VB.OptionButton Etapa 
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
         Height          =   210
         Index           =   2
         Left            =   6780
         TabIndex        =   1
         Top             =   270
         Width           =   1140
      End
      Begin VB.OptionButton Etapa 
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
         Height          =   210
         Index           =   1
         Left            =   2745
         TabIndex        =   0
         Top             =   255
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.OptionButton Etapa 
         Caption         =   "Montagem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   6780
         TabIndex        =   2
         Top             =   270
         Visible         =   0   'False
         Width           =   1350
      End
   End
End
Attribute VB_Name = "OCArtlux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Public iAlterado As Integer
Dim gcolOrdens As Collection
Dim gcolOrdensVisiveis As Collection

Dim iTimer As Integer

Dim sReferenciaAnt As String
Dim sCorAnt As String
Dim sCorteAnt As String
Dim sForroAnt As String
Dim sGrupoAnt As String

'Grid de Corte de Couro e Forro
Dim objGridC As AdmGrid
'Produto, Descrição, Quantidade, Corte, Forro, Prioridade, Quant. Est, Quant PV, Estoque Segurança
Dim iGrid_CProduto_Col As Integer
Dim iGrid_CDescricao_Col As Integer
Dim iGrid_CTipoCouro_Col As Integer
Dim iGrid_CQuantidade_Col As Integer
Dim iGrid_CCorte_Col As Integer
Dim iGrid_CForro_Col As Integer
Dim iGrid_CPrioridade_Col As Integer
Dim iGrid_CQuantEst_Col As Integer
Dim iGrid_CQuantPV_Col As Integer
Dim iGrid_CEstSeguranca_Col As Integer

'Grid de Montagem
Dim objGridM As AdmGrid
'Produto, Descrição, Quantidade,QuantEnt,QuantSai, Corte, Forro,Montagem, Prioridade, Quant. Est, Quant PV, Estoque Segurança
Dim iGrid_MProduto_Col As Integer
Dim iGrid_MDescricao_Col As Integer
Dim iGrid_MTipoCouro_Col As Integer
Dim iGrid_MQuantidade_Col As Integer
Dim iGrid_MQuantPreProd_Col As Integer
Dim iGrid_MQuantProd_Col As Integer
Dim iGrid_MCorte_Col As Integer
Dim iGrid_MForro_Col As Integer
Dim iGrid_MMontagem_Col As Integer
Dim iGrid_MPrioridade_Col As Integer
Dim iGrid_MQuantEst_Col As Integer
Dim iGrid_MQuantPV_Col As Integer
Dim iGrid_MEstSeguranca_Col As Integer

Function Traz_Ordens_Tela() As Long

Dim lErro As Long
Dim colOrdens As New Collection
Dim iEtapa As Integer
Dim iIndice As Integer
Dim iLinha As Integer
Dim objOC As ClassOCArtlux
Dim sProdutoMask As String

On Error GoTo Erro_Traz_Ordens_Tela

    iTimer = 0

    For iIndice = 1 To 3
        If Etapa(iIndice).Value Then
            iEtapa = iIndice
            Exit For
        End If
    Next
    
    Call Grid_Limpa(objGridC)
    Call Grid_Limpa(objGridM)

    lErro = CF("OrdensDeCorteArtlux_Le", colOrdens, iEtapa)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Set gcolOrdens = colOrdens
    
    lErro = Carrega_Filtros(colOrdens)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Set gcolOrdensVisiveis = New Collection
    
    If iEtapa = ETAPA_MONTAGEM Then
    
        FrameEtapa(1).Visible = True
        FrameEtapa(0).Visible = False
    
        iLinha = 0
        For Each objOC In colOrdens
        
            lErro = Trata_Exibir(objOC)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            If objOC.iExibir = MARCADO Then
            
                gcolOrdensVisiveis.Add objOC
            
                iLinha = iLinha + 1
                
                Call Mascara_RetornaProdutoTela(objOC.sProduto, sProdutoMask)
                
                GridM.TextMatrix(iLinha, iGrid_MProduto_Col) = sProdutoMask
                GridM.TextMatrix(iLinha, iGrid_MDescricao_Col) = objOC.sProdutoDesc
                GridM.TextMatrix(iLinha, iGrid_MQuantidade_Col) = Formata_Estoque(objOC.dQuantidade)
                GridM.TextMatrix(iLinha, iGrid_MQuantPreProd_Col) = Formata_Estoque(objOC.dQuantidadePreProd)
                GridM.TextMatrix(iLinha, iGrid_MQuantProd_Col) = Formata_Estoque(objOC.dQuantidadeProd)
                GridM.TextMatrix(iLinha, iGrid_MCorte_Col) = objOC.sUsuCorte
                GridM.TextMatrix(iLinha, iGrid_MForro_Col) = objOC.sUsuForro
                GridM.TextMatrix(iLinha, iGrid_MMontagem_Col) = objOC.sUsuMontagem
                GridM.TextMatrix(iLinha, iGrid_MPrioridade_Col) = CStr(objOC.iPrioridade)
                GridM.TextMatrix(iLinha, iGrid_MQuantEst_Col) = Formata_Estoque(objOC.dQuantidadeEst)
                GridM.TextMatrix(iLinha, iGrid_MQuantPV_Col) = Formata_Estoque(objOC.dQuantidadePV)
                GridM.TextMatrix(iLinha, iGrid_MEstSeguranca_Col) = Formata_Estoque(objOC.dEstoqueSeguranca)
                GridM.TextMatrix(iLinha, iGrid_MTipoCouro_Col) = objOC.sTipoCouro

            End If

        Next
        
        objGridM.iLinhasExistentes = colOrdens.Count
    Else
    
        FrameEtapa(1).Visible = False
        FrameEtapa(0).Visible = True
        
        If iEtapa = ETAPA_CORTE Then
            FrameEtapa(0).Caption = "Corte"
        Else
            FrameEtapa(0).Caption = "Forro"
        End If
        
        iLinha = 0
        For Each objOC In colOrdens
        
            lErro = Trata_Exibir(objOC)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            If objOC.iExibir = MARCADO Then
        
                iLinha = iLinha + 1
                
                gcolOrdensVisiveis.Add objOC
                
                Call Mascara_RetornaProdutoTela(objOC.sProduto, sProdutoMask)
                
                GridC.TextMatrix(iLinha, iGrid_CProduto_Col) = sProdutoMask
                GridC.TextMatrix(iLinha, iGrid_CDescricao_Col) = objOC.sProdutoDesc
                GridC.TextMatrix(iLinha, iGrid_CQuantidade_Col) = Formata_Estoque(objOC.dQuantidade)
                GridC.TextMatrix(iLinha, iGrid_CCorte_Col) = objOC.sUsuCorte
                GridC.TextMatrix(iLinha, iGrid_CForro_Col) = objOC.sUsuForro
                GridC.TextMatrix(iLinha, iGrid_CPrioridade_Col) = CStr(objOC.iPrioridade)
                GridC.TextMatrix(iLinha, iGrid_CQuantEst_Col) = Formata_Estoque(objOC.dQuantidadeEst)
                GridC.TextMatrix(iLinha, iGrid_CQuantPV_Col) = Formata_Estoque(objOC.dQuantidadePV)
                GridC.TextMatrix(iLinha, iGrid_CEstSeguranca_Col) = Formata_Estoque(objOC.dEstoqueSeguranca)
                GridC.TextMatrix(iLinha, iGrid_CTipoCouro_Col) = objOC.sTipoCouro

            End If

        Next
        
        objGridC.iLinhasExistentes = colOrdens.Count
    
    End If
    
    iTimer = 0

    iAlterado = 0

    Traz_Ordens_Tela = SUCESSO

    Exit Function

Erro_Traz_Ordens_Tela:

    Traz_Ordens_Tela = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206728)

    End Select

    Exit Function

End Function

Function Trata_Parametros() As Long
'Trata os parametros

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206729)

    End Select

    iAlterado = 0

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub BotaoAtualizar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoAtualizar_Click

    'Grava os registros na tabela
    lErro = Traz_Ordens_Tela()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub
    
Erro_BotaoAtualizar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206730)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()
    'Call TelaIndice_Preenche(Me)
End Sub

Public Sub Form_Deactivate()
    'gi_ST_SetaIgnoraClick = 1
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iAlterado = 0

    Set gcolOrdens = New Collection
    Set objGridC = New AdmGrid
    Set objGridM = New AdmGrid
    
    lErro = Inicializa_Grid_C(objGridC)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Inicializa_Grid_M(objGridM)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Traz_Ordens_Tela
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206731)

    End Select

    iAlterado = 0
    
    Exit Sub

End Sub

'Extrai os campos da tela que correspondem aos campos no BD
Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'
End Sub

'Preenche os campos da tela com os correspondentes do BD
Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)
    Set gcolOrdens = Nothing
    Set gcolOrdensVisiveis = Nothing
    Set objGridC = Nothing
    Set objGridM = Nothing
End Sub

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Ordens de Corte"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "OCArtlux"
    
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
'**** fim do trecho a ser copiado *****

Private Sub GridC_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridC, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridC, iAlterado)
    End If

End Sub

Private Sub GridC_GotFocus()
    Call Grid_Recebe_Foco(objGridC)
End Sub

Private Sub GridC_EnterCell()
    Call Grid_Entrada_Celula(objGridC, iAlterado)
End Sub

Private Sub GridC_LeaveCell()
    Call Saida_Celula(objGridC)
End Sub

Private Sub GridC_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridC, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridC, iAlterado)
    End If

End Sub

Private Sub GridC_RowColChange()
    Call Grid_RowColChange(objGridC)
    If Not (gcolOrdensVisiveis Is Nothing) Then If gcolOrdensVisiveis.Count >= GridC.Row And GridC.Row <> 0 Then Call Mostra_Produto(gcolOrdensVisiveis.Item(GridC.Row))
End Sub

Private Sub GridC_Scroll()
    Call Grid_Scroll(objGridC)
End Sub

Private Sub GridC_KeyDown(KeyCode As Integer, Shift As Integer)
Dim iIndice As Integer
Dim iEtapa As Integer
    Call Grid_Trata_Tecla1(KeyCode, objGridC)
    For iIndice = 1 To 3
        If Etapa(iIndice).Value Then
            iEtapa = iIndice
            Exit For
        End If
    Next
    If KeyCode = vbKeyReturn Then
        If GridC.Row > 0 And GridC.Row <= gcolOrdensVisiveis.Count Then
            Timer1.Enabled = False
            Call Chama_Tela_Modal("OCUsuArtlux", gcolOrdensVisiveis.Item(GridC.Row), iEtapa)
            Timer1.Enabled = True
            If giRetornoTela = vbOK Then Call Traz_Ordens_Tela
        End If
    End If
End Sub

Private Sub GridC_LostFocus()
    Call Grid_Libera_Foco(objGridC)
End Sub

Private Sub GridM_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridM, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridM, iAlterado)
    End If

End Sub

Private Sub GridM_GotFocus()
    Call Grid_Recebe_Foco(objGridM)
End Sub

Private Sub GridM_EnterCell()
    Call Grid_Entrada_Celula(objGridM, iAlterado)
End Sub

Private Sub GridM_LeaveCell()
    Call Saida_Celula(objGridM)
End Sub

Private Sub GridM_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridM, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridM, iAlterado)
    End If

End Sub

Private Sub GridM_RowColChange()
    Call Grid_RowColChange(objGridM)
    If Not (gcolOrdensVisiveis Is Nothing) Then If gcolOrdensVisiveis.Count >= GridM.Row And GridM.Row <> 0 Then Call Mostra_Produto(gcolOrdensVisiveis.Item(GridM.Row))
End Sub

Private Sub GridM_Scroll()
    Call Grid_Scroll(objGridM)
End Sub

Private Sub GridM_KeyDown(KeyCode As Integer, Shift As Integer)
   
    Call Grid_Trata_Tecla1(KeyCode, objGridM)
    
    If KeyCode = vbKeyReturn Then
        If GridM.Row > 0 And GridM.Row <= gcolOrdensVisiveis.Count Then
            Timer1.Enabled = False
            Call Chama_Tela_Modal("OCProdArtlux", gcolOrdensVisiveis.Item(GridM.Row))
            Timer1.Enabled = True
            If giRetornoTela = vbOK Then Call Traz_Ordens_Tela
        End If
    End If

End Sub

Private Sub GridM_LostFocus()
    Call Grid_Libera_Foco(objGridM)
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
    
        'Verifica qual é o grid
        If objGridInt.objGrid.Name = GridC.Name Then
                                 
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError ERRO_SEM_MENSAGEM

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 206732)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_C(objGridInt As AdmGrid) As Long
'Executa a Inicialização

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_C

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("T.Couro")
    objGridInt.colColuna.Add ("Qtde")
    objGridInt.colColuna.Add ("Corte")
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("P")
    objGridInt.colColuna.Add ("Estoque")
    objGridInt.colColuna.Add ("Pedido")
    objGridInt.colColuna.Add ("Mínimo")

    'campos de edição do grid
    objGridInt.colCampo.Add (CProduto.Name)
    objGridInt.colCampo.Add (CDescricao.Name)
    objGridInt.colCampo.Add (CTipoCouro.Name)
    objGridInt.colCampo.Add (CQuantidade.Name)
    objGridInt.colCampo.Add (CCorte.Name)
    objGridInt.colCampo.Add (CForro.Name)
    objGridInt.colCampo.Add (CPrioridade.Name)
    objGridInt.colCampo.Add (CQuantEst.Name)
    objGridInt.colCampo.Add (CQuantPV.Name)
    objGridInt.colCampo.Add (CEstSeguranca.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_CProduto_Col = 1
    iGrid_CDescricao_Col = 2
    iGrid_CTipoCouro_Col = 3
    iGrid_CQuantidade_Col = 4
    iGrid_CCorte_Col = 5
    iGrid_CForro_Col = 6
    iGrid_CPrioridade_Col = 7
    iGrid_CQuantEst_Col = 8
    iGrid_CQuantPV_Col = 9
    iGrid_CEstSeguranca_Col = 10
    
    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridC
    
    'Largura da primeira coluna
    GridC.ColWidth(0) = 400

    'Linhas do grid
    objGridInt.objGrid.Rows = 2000 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 22

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_C = SUCESSO

    Exit Function

Erro_Inicializa_Grid_C:

    Inicializa_Grid_C = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206733)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_M(objGridInt As AdmGrid) As Long
'Executa a Inicialização

Dim lErro As Long

On Error GoTo Erro_Inicializa_Grid_M

    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("T.Couro")
    objGridInt.colColuna.Add ("Qtde")
    objGridInt.colColuna.Add ("Saída")
    objGridInt.colColuna.Add ("Entrada")
    objGridInt.colColuna.Add ("Corte")
    objGridInt.colColuna.Add ("Forro")
    objGridInt.colColuna.Add ("Montagem")
    objGridInt.colColuna.Add ("P")
    objGridInt.colColuna.Add ("Estoque")
    objGridInt.colColuna.Add ("Pedido")
    objGridInt.colColuna.Add ("Mínimo")

    'campos de edição do grid
    objGridInt.colCampo.Add (MProduto.Name)
    objGridInt.colCampo.Add (MDescricao.Name)
    objGridInt.colCampo.Add (MTipoCouro.Name)
    objGridInt.colCampo.Add (MQuantidade.Name)
    objGridInt.colCampo.Add (MQuantPreProd.Name)
    objGridInt.colCampo.Add (MQuantProd.Name)
    objGridInt.colCampo.Add (MCorte.Name)
    objGridInt.colCampo.Add (MForro.Name)
    objGridInt.colCampo.Add (MMontagem.Name)
    objGridInt.colCampo.Add (MPrioridade.Name)
    objGridInt.colCampo.Add (MQuantEst.Name)
    objGridInt.colCampo.Add (MQuantPV.Name)
    objGridInt.colCampo.Add (MEstSeguranca.Name)

    'indica onde estao situadas as colunas do grid
    iGrid_MProduto_Col = 1
    iGrid_MDescricao_Col = 2
    iGrid_MTipoCouro_Col = 3
    iGrid_MQuantidade_Col = 4
    iGrid_MQuantPreProd_Col = 5
    iGrid_MQuantProd_Col = 6
    iGrid_MCorte_Col = 7
    iGrid_MForro_Col = 8
    iGrid_MMontagem_Col = 9
    iGrid_MPrioridade_Col = 10
    iGrid_MQuantEst_Col = 11
    iGrid_MQuantPV_Col = 12
    iGrid_MEstSeguranca_Col = 13

    'Relaciona com o grid correspondente na tela
    objGridInt.objGrid = GridM

    'Largura da primeira coluna
    GridM.ColWidth(0) = 400
    
    'Linhas do grid
    objGridInt.objGrid.Rows = 2000 + 1

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 19

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_M = SUCESSO

    Exit Function

Erro_Inicializa_Grid_M:

    Inicializa_Grid_M = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206734)

    End Select

    Exit Function

End Function

Private Sub Etapa_Click(Index As Integer)
    Call Traz_Ordens_Tela
End Sub

Private Sub Timer1_Timer()
    iTimer = iTimer + 1
    If iTimer = 30 Then Call Traz_Ordens_Tela
End Sub

Private Sub GridC_DblClick()
Dim iIndice As Integer
Dim iEtapa As Integer
    For iIndice = 1 To 3
        If Etapa(iIndice).Value Then
            iEtapa = iIndice
            Exit For
        End If
    Next
    If GridC.Row > 0 And GridC.Row <= gcolOrdensVisiveis.Count Then
        Timer1.Enabled = False
        Call Chama_Tela_Modal("OCUsuArtlux", gcolOrdensVisiveis.Item(GridC.Row), iEtapa)
        Timer1.Enabled = True
        If giRetornoTela = vbOK Then Call Traz_Ordens_Tela
    End If
End Sub

Private Sub GridM_DblClick()
    If GridM.Row > 0 And GridM.Row <= gcolOrdensVisiveis.Count Then
        Timer1.Enabled = False
        Call Chama_Tela_Modal("OCProdArtlux", gcolOrdensVisiveis.Item(GridM.Row))
        Timer1.Enabled = True
        If giRetornoTela = vbOK Then Call Traz_Ordens_Tela
    End If
End Sub

Private Function Carrega_Filtros(ByVal colOrdens As Collection) As Long

Dim lErro As Long
Dim objOC As New ClassOCArtlux
Dim iIndice As Integer
Dim bAchou As Boolean
Dim sFiltro As String

On Error GoTo Erro_Carrega_Filtros

    sFiltro = FReferencia.Text
    FReferencia.Clear
    For Each objOC In colOrdens
        bAchou = False
        For iIndice = 0 To FReferencia.ListCount - 1
            If left(objOC.sProduto, 5) = FReferencia.List(iIndice) Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            FReferencia.AddItem left(objOC.sProduto, 5)
        End If
    Next
    FReferencia.Text = sFiltro
    
    sFiltro = FCor.Text
    FCor.Clear
    For Each objOC In colOrdens
        bAchou = False
        For iIndice = 0 To FCor.ListCount - 1
            If Trim(Mid(objOC.sProduto, 6)) = FCor.List(iIndice) Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            FCor.AddItem Trim(Mid(objOC.sProduto, 6))
        End If
    Next
    FCor.Text = sFiltro

    sFiltro = FCorte.Text
    FCorte.Clear
    For Each objOC In colOrdens
        bAchou = False
        For iIndice = 0 To FCorte.ListCount - 1
            If objOC.sUsuCorte = FCorte.List(iIndice) Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            FCorte.AddItem objOC.sUsuCorte
        End If
    Next
    FCorte.Text = sFiltro
    
    sFiltro = FForro.Text
    FForro.Clear
    For Each objOC In colOrdens
        bAchou = False
        For iIndice = 0 To FForro.ListCount - 1
            If objOC.sUsuForro = FForro.List(iIndice) Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            FForro.AddItem objOC.sUsuForro
        End If
    Next
    FForro.Text = sFiltro
    
    sFiltro = FGrupo.Text
    FGrupo.Clear
    For Each objOC In colOrdens
        bAchou = False
        For iIndice = 0 To FGrupo.ListCount - 1
            If objOC.sGrupo = FGrupo.List(iIndice) Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            FGrupo.AddItem objOC.sGrupo
        End If
    Next
    FGrupo.Text = sFiltro
    
    Carrega_Filtros = SUCESSO

    Exit Function

Erro_Carrega_Filtros:

    Carrega_Filtros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206735)

    End Select

    Exit Function

End Function

Private Function Trata_Exibir(ByVal objOC As ClassOCArtlux) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Exibir

    objOC.iExibir = MARCADO
    If Len(Trim(FReferencia.Text)) > 0 Then
        If InStr(1, UCase(left(objOC.sProduto, 5)), UCase(FReferencia.Text)) = 0 Then objOC.iExibir = DESMARCADO
    End If
    If Len(Trim(FCor.Text)) > 0 And objOC.iExibir = MARCADO Then
        If InStr(1, UCase(Mid(objOC.sProduto, 6)), UCase(FCor.Text)) = 0 Then objOC.iExibir = DESMARCADO
    End If
    If Len(Trim(FCorte.Text)) > 0 And objOC.iExibir = MARCADO Then
        If InStr(1, UCase(objOC.sUsuCorte), UCase(FCorte.Text)) = 0 Then objOC.iExibir = DESMARCADO
    End If
    If Len(Trim(FForro.Text)) > 0 And objOC.iExibir = MARCADO Then
        If InStr(1, UCase(objOC.sUsuForro), UCase(FForro.Text)) = 0 Then objOC.iExibir = DESMARCADO
    End If
    If Len(Trim(FGrupo.Text)) > 0 And objOC.iExibir = MARCADO Then
        If InStr(1, UCase(objOC.sGrupo), UCase(FGrupo.Text)) = 0 Then objOC.iExibir = DESMARCADO
    End If
    
    Trata_Exibir = SUCESSO

    Exit Function

Erro_Trata_Exibir:

    Trata_Exibir = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 206736)

    End Select

    Exit Function

End Function

Private Sub FReferencia_Validate(Cancel As Boolean)
    If sReferenciaAnt <> FReferencia.Text Then
        sReferenciaAnt = FReferencia.Text
        Call Traz_Ordens_Tela
    End If
End Sub

Private Sub FCor_Validate(Cancel As Boolean)
    If sCorAnt <> FCor.Text Then
        sCorAnt = FCor.Text
        Call Traz_Ordens_Tela
    End If
End Sub

Private Sub FCorte_Validate(Cancel As Boolean)
    If sCorteAnt <> FCorte.Text Then
        sCorteAnt = FCorte.Text
        Call Traz_Ordens_Tela
    End If
End Sub

Private Sub FForro_Validate(Cancel As Boolean)
    If sForroAnt <> FForro.Text Then
        sForroAnt = FForro.Text
        Call Traz_Ordens_Tela
    End If
End Sub

Private Sub FGrupo_Validate(Cancel As Boolean)
    If sGrupoAnt <> FGrupo.Text Then
        sGrupoAnt = FGrupo.Text
        Call Traz_Ordens_Tela
    End If
End Sub

Private Sub BotaoLimparFiltros_Click()
    FReferencia.Text = ""
    FCor.Text = ""
    FCorte.Text = ""
    FForro.Text = ""
    FGrupo.Text = ""
    sReferenciaAnt = ""
    sCorAnt = ""
    sCorteAnt = ""
    sForroAnt = ""
    sGrupoAnt = ""
    Call Traz_Ordens_Tela
End Sub

Private Sub Mostra_Produto(ByVal objOC As ClassOCArtlux)
Dim sProdutoMask As String
    Call Mascara_RetornaProdutoTela(objOC.sProduto, sProdutoMask)
    Produto.Caption = sProdutoMask
    Descricao.Caption = objOC.sProdutoDesc
End Sub
