VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TRPLiberaOcr 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4920
      Index           =   1
      Left            =   225
      TabIndex        =   15
      Top             =   810
      Width           =   8925
      Begin VB.Frame Frame2 
         Caption         =   "Exibe Bloqueios"
         Height          =   4035
         Left            =   945
         TabIndex        =   24
         Top             =   300
         Width           =   7020
         Begin VB.OptionButton OptVouPagAmbos 
            Caption         =   "Ambos"
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
            Left            =   5520
            TabIndex        =   2
            Top             =   585
            Width           =   1215
         End
         Begin VB.OptionButton OptVouNaoPag 
            Caption         =   "Vouchers Não Pagos"
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
            Left            =   2895
            TabIndex        =   1
            Top             =   585
            Width           =   2295
         End
         Begin VB.OptionButton OptVouPag 
            Caption         =   "Vouchers Pagos"
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
            Left            =   930
            TabIndex        =   0
            Top             =   585
            Value           =   -1  'True
            Width           =   1830
         End
         Begin VB.Frame Frame4 
            Caption         =   "Data de emissão"
            Height          =   1125
            Left            =   945
            TabIndex        =   29
            Top             =   2640
            Width           =   5505
            Begin MSComCtl2.UpDown UpDownEmissaoDe 
               Height          =   300
               Left            =   1920
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   480
               Width           =   225
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataEmissaoDe 
               Height          =   300
               Left            =   795
               TabIndex        =   5
               Top             =   480
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox DataEmissaoAte 
               Height          =   300
               Left            =   3390
               TabIndex        =   7
               Top             =   480
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               _Version        =   393216
               MaxLength       =   8
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/##"
               PromptChar      =   " "
            End
            Begin MSComCtl2.UpDown UpDownEmissaoAte 
               Height          =   300
               Left            =   4545
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   480
               Width           =   225
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Até:"
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
               Left            =   2985
               TabIndex        =   31
               Top             =   540
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "De:"
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
               Index           =   0
               Left            =   345
               TabIndex        =   30
               Top             =   540
               Width           =   315
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Ocorrências"
            Height          =   1125
            Left            =   945
            TabIndex        =   25
            Top             =   1080
            Width           =   5505
            Begin MSMask.MaskEdBox OcorrenciaDe 
               Height          =   300
               Left            =   780
               TabIndex        =   3
               Top             =   465
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox OcorrenciaAte 
               Height          =   300
               Left            =   3435
               TabIndex        =   4
               Top             =   465
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label LabelOCRAte 
               AutoSize        =   -1  'True
               Caption         =   "Até:"
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
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   27
               Top             =   525
               Width           =   360
            End
            Begin VB.Label LabelOCRDe 
               AutoSize        =   -1  'True
               Caption         =   "De:"
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
               TabIndex        =   26
               Top             =   525
               Width           =   315
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4890
      Index           =   2
      Left            =   210
      TabIndex        =   16
      Top             =   795
      Visible         =   0   'False
      Width           =   8895
      Begin VB.ComboBox Forma 
         Height          =   315
         ItemData        =   "TRPLiberaOcr.ctx":0000
         Left            =   1755
         List            =   "TRPLiberaOcr.ctx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1800
         Width           =   1275
      End
      Begin VB.TextBox Voucher 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1095
         TabIndex        =   33
         Text            =   "Voucher"
         Top             =   1275
         Width           =   735
      End
      Begin VB.TextBox Serie 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   6345
         TabIndex        =   32
         Text            =   "Serie"
         Top             =   735
         Width           =   420
      End
      Begin VB.TextBox Observacao 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   2895
         MaxLength       =   250
         TabIndex        =   28
         Text            =   "Observacao"
         Top             =   1230
         Width           =   2625
      End
      Begin VB.CheckBox Libera 
         Height          =   210
         Left            =   180
         TabIndex        =   17
         Top             =   735
         Width           =   690
      End
      Begin VB.TextBox Numero 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1005
         TabIndex        =   18
         Text            =   "Numero"
         Top             =   705
         Width           =   930
      End
      Begin VB.TextBox Tipo 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   5655
         TabIndex        =   19
         Text            =   "Tipo"
         Top             =   765
         Width           =   585
      End
      Begin VB.CommandButton BotaoLibera 
         Caption         =   "Libera os Bloqueios Assinalados"
         Height          =   960
         Left            =   165
         Picture         =   "TRPLiberaOcr.ctx":002C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3885
         Width           =   1590
      End
      Begin VB.CommandButton BotaoOCR 
         Caption         =   "Editar a Ocorrência"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   7080
         TabIndex        =   13
         Top             =   3885
         Width           =   1590
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   570
         Left            =   4545
         Picture         =   "TRPLiberaOcr.ctx":046E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4080
         Width           =   1425
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   570
         Left            =   2925
         Picture         =   "TRPLiberaOcr.ctx":1650
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4080
         Width           =   1425
      End
      Begin VB.TextBox Cliente 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2010
         TabIndex        =   20
         Text            =   "Cliente"
         Top             =   435
         Width           =   1470
      End
      Begin VB.TextBox Valor 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   4635
         TabIndex        =   22
         Text            =   "Valor"
         Top             =   795
         Width           =   840
      End
      Begin MSMask.MaskEdBox DataEmissao 
         Height          =   285
         Left            =   3435
         TabIndex        =   21
         Top             =   795
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid GridBloqueio 
         Height          =   3765
         Left            =   30
         TabIndex        =   9
         Top             =   45
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   6641
         _Version        =   393216
         Rows            =   11
         Cols            =   7
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
      End
   End
   Begin VB.CommandButton BotaoFechar 
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
      Left            =   7965
      Picture         =   "TRPLiberaOcr.ctx":266A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Fechar"
      Top             =   120
      Width           =   1230
   End
   Begin MSComctlLib.TabStrip TabStripOpcao 
      Height          =   5430
      Left            =   150
      TabIndex        =   23
      Top             =   390
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   9578
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bloqueios"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "TRPLiberaOcr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer 'apenas p/uso pela interface c/grid
Dim iFrameAtual As Integer

'Grid Bloqueio:
Dim objGridBloqueio As AdmGrid
Dim iGrid_Libera_Col As Integer
Dim iGrid_Numero_Col As Integer
Dim iGrid_Forma_Col As Integer
Dim iGrid_Cliente_Col As Integer
Dim iGrid_DataEmissao_Col  As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Tipo_Col As Integer
Dim iGrid_Serie_Col As Integer
Dim iGrid_Voucher_Col As Integer
Dim iGrid_Observacao_Col As Integer

Dim gobjTRPLiberaOcrSel As New ClassTRPLiberaOCRSel

'Eventos de Browse
Private WithEvents objEventoOcrDe As AdmEvento
Attribute objEventoOcrDe.VB_VarHelpID = -1
Private WithEvents objEventoOcrAte As AdmEvento
Attribute objEventoOcrAte.VB_VarHelpID = -1

'CONTANTES GLOBAIS DA TELA
Const TAB_Selecao = 1
Const TAB_BLOQUEIOS = 2

Private Function Move_Selecao_Memoria(ByVal objTRPLiberaOcrSel As ClassTRPLiberaOCRSel) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Selecao_Memoria

    If OptVouPag.Value Then
        objTRPLiberaOcrSel.iPago = TRP_VOU_PAGO
    ElseIf OptVouNaoPag.Value Then
        objTRPLiberaOcrSel.iPago = TRP_VOU_NAO_PAGO
    Else
        objTRPLiberaOcrSel.iPago = TRP_VOU_PAGO_E_NAO_PAGO
    End If

    objTRPLiberaOcrSel.dtDataEmissaoAte = StrParaDate(DataEmissaoAte.Text)
    objTRPLiberaOcrSel.dtDataEmissaoDe = StrParaDate(DataEmissaoDe.Text)
    objTRPLiberaOcrSel.lCodigoAte = StrParaLong(OcorrenciaAte.Text)
    objTRPLiberaOcrSel.lCodigoDe = StrParaLong(OcorrenciaDe.Text)
    
    If objTRPLiberaOcrSel.dtDataEmissaoAte <> DATA_NULA And objTRPLiberaOcrSel.dtDataEmissaoDe <> DATA_NULA Then
        If objTRPLiberaOcrSel.dtDataEmissaoAte < objTRPLiberaOcrSel.dtDataEmissaoDe Then gError 190256
    End If
            
    Move_Selecao_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Selecao_Memoria:

    Move_Selecao_Memoria = gErr
    
    Select Case gErr
    
        Case 190256
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190257)

    End Select

End Function

Private Function Traz_Bloqueios_Tela(Optional ByVal bTirandoLiberados As Boolean = False) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objTRPLiberaOcrSel As New ClassTRPLiberaOCRSel

On Error GoTo Erro_Traz_Bloqueios_Tela

    'Limpa o GridBloqueio
    Call Grid_Limpa(objGridBloqueio)
    
    lErro = Move_Selecao_Memoria(objTRPLiberaOcrSel)
    If lErro <> SUCESSO Then gError 190258
  
    'Preenche a Coleção de Bloqueios
    lErro = CF("TRPOcorrencias_Le_Bloqueios", objTRPLiberaOcrSel)
    If lErro <> SUCESSO Then gError 190259
    
    If Not bTirandoLiberados Then
        If objTRPLiberaOcrSel.colOcorrenciais.Count = 0 Then gError 190260
    End If
    
    Set gobjTRPLiberaOcrSel = objTRPLiberaOcrSel
    
    'Preenche o GridBloqueio
    lErro = Grid_Bloqueio_Preenche(objTRPLiberaOcrSel)
    If lErro <> SUCESSO Then gError 190261
            
    Traz_Bloqueios_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Bloqueios_Tela:

    Traz_Bloqueios_Tela = gErr
    
    Select Case gErr

        Case 190258, 190259, 190261
            
        Case 190260 'ERRO_SEM_BLOQUEIOS_PC_SEL
             Call Rotina_Erro(vbOKOnly, "ERRO_SEM_BLOQUEIOS_PC_SEL", gErr)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190262)

    End Select

End Function

Private Sub BotaoDesmarcarTodos_Click()
'Desmarca todos os bloqueios do Grid

Dim iLinha As Integer
Dim objBloqueioLiberacaoInfo As ClassBloqueioLiberacaoInfo

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridBloqueio.iLinhasExistentes

        'Desmarca na tela o bloqueio em questão
        GridBloqueio.TextMatrix(iLinha, iGrid_Libera_Col) = GRID_CHECKBOX_INATIVO
        
    Next
    
    'Atualiza na tela os checkbox desmarcados
    Call Grid_Refresh_Checkbox(objGridBloqueio)
    
End Sub

Private Sub BotaoMarcarTodos_Click()
'Marca todos os bloqueios do Grid

Dim iLinha As Integer
Dim objBloqueioLiberacaoInfo As ClassBloqueioLiberacaoInfo

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridBloqueio.iLinhasExistentes

        'Marca na tela o bloqueio em questão
        GridBloqueio.TextMatrix(iLinha, iGrid_Libera_Col) = GRID_CHECKBOX_ATIVO
        
    Next
    
    'Atualiza na tela os checkbox marcados
    Call Grid_Refresh_Checkbox(objGridBloqueio)
    
End Sub

Private Sub BotaoOCR_Click()
    
Dim lErro As Long
Dim objOcorrencia As New ClassTRPOcorrencias

On Error GoTo Erro_BotaoOCR_Click
    
    'Verifica se alguma linha do Grid está selecionada
    If GridBloqueio.Row = 0 Then gError 190263
    
    'Passa a linha do Grid para o Obj
    objOcorrencia.lCodigo = gobjTRPLiberaOcrSel.colOcorrenciais.Item(GridBloqueio.Row).lCodigo

    'Chama a tela de Pedidos de Venda
    Call Chama_Tela("TRPOcorrencias", objOcorrencia)
    
    Exit Sub

Erro_BotaoOCR_Click:

    Select Case gErr

        Case 190263
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190264)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridBloqueio = Nothing
    Set gobjTRPLiberaOcrSel = Nothing
    
End Sub

Private Sub DataEmissaoAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataEmissaoAte, iAlterado)
End Sub

Private Sub DataEmissaoAte_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataEmissaoAte_Validate

    'Se a DataEmissaoAte está preenchida
    If Len(DataEmissaoAte.ClipText) > 0 Then

        'Verifica se a DataEmissaoAte é válida
        lErro = Data_Critica(DataEmissaoAte.Text)
        If lErro <> SUCESSO Then gError 190265

    End If
    
    Exit Sub

Erro_DataEmissaoAte_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 190265 'Tratado na rotina chamada
         
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190266)

    End Select

    Exit Sub

End Sub

Private Sub DataEmissaoDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataEmissaoDe, iAlterado)

End Sub

Private Sub DataEmissaoDe_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataEmissaoDe_Validate

    'Se a DataEmissaoDe está preenchida
    If Len(DataEmissaoDe.ClipText) > 0 Then

        'Verifica se a DataEmissaoDe é válida
        lErro = Data_Critica(DataEmissaoDe.Text)
        If lErro <> SUCESSO Then gError 190267

    End If

    Exit Sub

Erro_DataEmissaoDe_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 190267

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190268)

    End Select

    Exit Sub

End Sub

Private Sub GridBloqueio_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcolColecoes As New Collection

    Call Grid_Click(objGridBloqueio, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)
    End If
    
    colcolColecoes.Add gobjTRPLiberaOcrSel.colOcorrenciais
    
    Call Ordenacao_ClickGrid(objGridBloqueio, , colcolColecoes)

End Sub

Private Sub GridBloqueio_GotFocus()
    Call Grid_Recebe_Foco(objGridBloqueio)
End Sub

Private Sub GridBloqueio_EnterCell()
    Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)
End Sub

Private Sub GridBloqueio_LeaveCell()
    Call Saida_Celula(objGridBloqueio)
End Sub

Private Sub GridBloqueio_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridBloqueio)
End Sub

Private Sub GridBloqueio_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridBloqueio, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)
    End If

End Sub

Private Sub GridBloqueio_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridBloqueio)
End Sub

Private Sub GridBloqueio_RowColChange()
    Call Grid_RowColChange(objGridBloqueio)
End Sub

Private Sub GridBloqueio_Scroll()
    Call Grid_Scroll(objGridBloqueio)
End Sub

Private Sub LabelOCRAte_Click()

Dim colSelecao As Collection
Dim objOcorrencia As New ClassTRPOcorrencias

    'Preenche PedidoAte com o pedido da tela
    objOcorrencia.lCodigo = StrParaLong(OcorrenciaAte.Text)
    
    'Chama Tela PedidoVendaLista
    Call Chama_Tela("TRPOcorrenciaLista", colSelecao, objOcorrencia, objEventoOcrAte)

End Sub

Private Sub LabelOCRDe_Click()

Dim colSelecao As Collection
Dim objOcorrencia As New ClassTRPOcorrencias

    'Preenche PedidoAte com o pedido da tela
    objOcorrencia.lCodigo = StrParaLong(OcorrenciaDe.Text)
    
    'Chama Tela PedidoVendaLista
    Call Chama_Tela("TRPOcorrenciaLista", colSelecao, objOcorrencia, objEventoOcrDe)

End Sub

Private Sub objEventoOCRAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOcorrencia As ClassTRPOcorrencias
Dim bCancel As Boolean

On Error GoTo Erro_objEventoOCRAte_evSelecao

    Set objOcorrencia = obj1
    
    OcorrenciaAte.PromptInclude = False
    OcorrenciaAte.Text = CStr(objOcorrencia.lCodigo)
    OcorrenciaAte.PromptInclude = True

    'Chama o Validate de OcorrenciaAte
    Call OcorrenciaAte_Validate(bCancel)

    Me.Show

    Exit Sub

Erro_objEventoOCRAte_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190269)

    End Select

    Exit Sub

End Sub

Private Sub objEventoOCRDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOcorrencia As ClassTRPOcorrencias
Dim bCancel As Boolean

On Error GoTo Erro_objEventoOCRDe_evSelecao

    Set objOcorrencia = obj1
    
    OcorrenciaDe.PromptInclude = False
    OcorrenciaDe.Text = CStr(objOcorrencia.lCodigo)
    OcorrenciaDe.PromptInclude = True

    'Chama o Validate de OcorrenciaAte
    Call OcorrenciaAte_Validate(bCancel)

    Me.Show

    Exit Sub

Erro_objEventoOCRDe_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190270)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    'Fecha a tela
    Unload Me

End Sub

Public Sub Form_Load()
    
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load
    
    iFrameAtual = 1
    
    Set objGridBloqueio = New AdmGrid
    
    'Inicializa os Eventos de Browser
    Set objEventoOcrDe = New AdmEvento
    Set objEventoOcrAte = New AdmEvento
    
    'Executa a Inicialização do grid Bloqueio
    lErro = Inicializa_Grid_Bloqueio(objGridBloqueio)
    If lErro <> SUCESSO Then gError 190271
    
    lErro = CF("Carrega_Combo_FormaPagto", Forma)
    If lErro <> SUCESSO Then gError 190744
        
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 190271, 190744

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190272)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub OcorrenciaAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(OcorrenciaAte, iAlterado)

End Sub

Private Sub OcorrenciaDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(OcorrenciaDe, iAlterado)

End Sub

Private Sub OcorrenciaDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOcorrencia As New ClassTRPOcorrencias

On Error GoTo Erro_OcorrenciaDe_Validate

    If Len(Trim(OcorrenciaDe.Text)) > 0 Then
        
        'Critica para ver se é um Long
        lErro = Long_Critica(OcorrenciaDe.Text)
        If lErro <> SUCESSO Then gError 190273
            
        'Se o Pedido Final estiver preenchido então
        If Len(Trim(OcorrenciaAte.Text)) > 0 Then
            'Verifica se o Pedido Inicial é maior que o Pedido Final ---- Erro
            If CLng(OcorrenciaDe.Text) > CLng(OcorrenciaAte.Text) Then gError 190274
        End If
        
        objOcorrencia.lCodigo = CLng(OcorrenciaDe.Text)
        
        'Verifica se o Pedido está cadastrado no BD
        lErro = CF("TRPOcorrencias_Le", objOcorrencia)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190275
            
        'Pedido não está cadastrado
        If lErro <> SUCESSO Then gError 190276
        
    End If
       
    Exit Sub

Erro_OcorrenciaDe_Validate:

    Cancel = True

    Select Case gErr
    
        Case 190273, 190275
        
        Case 190274
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INICIAL_MAIOR_FINAL", gErr)
        
        Case 190276
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPOCORRENCIAS_NAO_CADASTRADO", gErr, objOcorrencia.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190277)

    End Select

    Exit Sub

End Sub

Private Sub OcorrenciaAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOcorrencia As New ClassTRPOcorrencias

On Error GoTo Erro_OcorrenciaAte_Validate

    If Len(Trim(OcorrenciaAte.Text)) > 0 Then
        
        'Critica para ver se é um Long
        lErro = Long_Critica(OcorrenciaAte.Text)
        If lErro <> SUCESSO Then gError 190278
            
        'Se o Pedido Final estiver preenchido então
        If Len(Trim(OcorrenciaDe.Text)) > 0 Then
            'Verifica se o Pedido Inicial é maior que o Pedido Final ---- Erro
            If CLng(OcorrenciaDe.Text) > CLng(OcorrenciaAte.Text) Then gError 190279
        End If
        
        objOcorrencia.lCodigo = CLng(OcorrenciaAte.Text)
        
        'Verifica se o Pedido está cadastrado no BD
        lErro = CF("TRPOcorrencias_Le", objOcorrencia)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190280
            
        'Pedido não está cadastrado
        If lErro <> SUCESSO Then gError 190281
        
    End If
       
    Exit Sub

Erro_OcorrenciaAte_Validate:

    Cancel = True

    Select Case gErr
    
        Case 190278, 190280
        
        Case 190279
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INICIAL_MAIOR_FINAL", gErr)
        
        Case 190281
            Call Rotina_Erro(vbOKOnly, "ERRO_TRPOCORRENCIAS_NAO_CADASTRADO", gErr, objOcorrencia.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190282)
            
    End Select

    Exit Sub

End Sub

Private Sub TabStripOpcao_Click()

Dim lErro As Long

On Error GoTo Erro_TabStripOpcao_Click

    'Se Frame atual não corresponde ao Tab clicado
    If TabStripOpcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, TabStripOpcao, Me) <> SUCESSO Then Exit Sub
       
        'Torna Frame de Bloqueios visível
        Frame1(TabStripOpcao.SelectedItem.Index).Visible = True
        'Torna Frame atual invisível
        Frame1(iFrameAtual).Visible = False
        
        'Armazena novo valor de iFrameAtual
        iFrameAtual = TabStripOpcao.SelectedItem.Index
       
        'Se Frame selecionado foi o de Bloqueios
        If TabStripOpcao.SelectedItem.Index = TAB_BLOQUEIOS Then
            
            lErro = Traz_Bloqueios_Tela
            If lErro <> SUCESSO Then gError 190283
            
        End If
    
    End If

    Exit Sub

Erro_TabStripOpcao_Click:

    Select Case gErr
        
        Case 190283
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190284)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    'Diminui a DataEmissaoAte em 1 dia
    lErro = Data_Up_Down_Click(DataEmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 190285

    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case gErr

        Case 190285

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190286)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    'Aumenta a DataEmissaoAte em 1 dia
    lErro = Data_Up_Down_Click(DataEmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 190287

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case gErr

        Case 190287

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190288)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    'Diminui a DataEmissaoDe em 1 dia
    lErro = Data_Up_Down_Click(DataEmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 190289

    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case gErr

        Case 190289

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190290)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    'Aumenta a DataEmissaoDe em 1 dia
    lErro = Data_Up_Down_Click(DataEmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 190291

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case gErr

        Case 190291

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190292)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objOcorrencia As ClassTRPOcorrencias) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    If Not (objOcorrencia Is Nothing) Then
        
        If objOcorrencia.lCodigo > 0 Then OcorrenciaDe.Text = CStr(objOcorrencia.lCodigo)
        If objOcorrencia.lCodigo > 0 Then OcorrenciaAte.Text = CStr(objOcorrencia.lCodigo)
                
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190293)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    Call Ordenacao_Limpa(objGridBloqueio)

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub

Private Function Grid_Bloqueio_Preenche(ByVal objLiberaOcrSel As ClassTRPLiberaOCRSel) As Long
'Preenche o Grid Bloqueio com os dados de colBloqueioLiberacaoInfo

Dim lErro As Long
Dim iLinha As Integer
Dim objOcorrencia As ClassTRPOcorrencias
Dim objcliente As ClassCliente

On Error GoTo Erro_Grid_Bloqueio_Preenche

    'Se o número de Bloqueios for maior que o número de linhas do Grid
    If objLiberaOcrSel.colOcorrenciais.Count >= objGridBloqueio.objGrid.Rows Then
        Call Refaz_Grid(objGridBloqueio, objLiberaOcrSel.colOcorrenciais.Count)
    End If

    iLinha = 0

    'Percorre todos os Bloqueios da Coleção
    For Each objOcorrencia In objLiberaOcrSel.colOcorrenciais

        iLinha = iLinha + 1
        
        Set objcliente = New ClassCliente

        'Passa para a tela os dados do Bloqueio em questão
        GridBloqueio.TextMatrix(iLinha, iGrid_Libera_Col) = CStr(MARCADO)
        GridBloqueio.TextMatrix(iLinha, iGrid_Numero_Col) = CStr(objOcorrencia.lCodigo)
        
        Call Combo_Seleciona_ItemData(Forma, objOcorrencia.iFormaPagto)
        GridBloqueio.TextMatrix(iLinha, iGrid_Forma_Col) = Forma.Text
        
        objcliente.lCodigo = objOcorrencia.lCliente
        
        lErro = CF("Cliente_Le", objcliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 190294
        
        GridBloqueio.TextMatrix(iLinha, iGrid_Cliente_Col) = CStr(objcliente.lCodigo) & SEPARADOR & objcliente.sNomeReduzido
        
        GridBloqueio.TextMatrix(iLinha, iGrid_DataEmissao_Col) = Format(objOcorrencia.dtDataEmissao, "dd/mm/yyyy")
        GridBloqueio.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objOcorrencia.dValorTotal, "STANDARD")
        GridBloqueio.TextMatrix(iLinha, iGrid_Tipo_Col) = objOcorrencia.sTipoDoc
        GridBloqueio.TextMatrix(iLinha, iGrid_Serie_Col) = objOcorrencia.sSerie
        GridBloqueio.TextMatrix(iLinha, iGrid_Voucher_Col) = CStr(objOcorrencia.lNumVou)
        GridBloqueio.TextMatrix(iLinha, iGrid_Observacao_Col) = objOcorrencia.sObservacao
        
    Next

    'Passa para o Obj o número de Bloqueios passados pela Coleção
    objGridBloqueio.iLinhasExistentes = objLiberaOcrSel.colOcorrenciais.Count
    
    Call Grid_Refresh_Checkbox(objGridBloqueio)

    Grid_Bloqueio_Preenche = SUCESSO
    
    Exit Function

Erro_Grid_Bloqueio_Preenche:
    
    Grid_Bloqueio_Preenche = gErr
    
    Select Case gErr
    
        Case 190294

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190295)
    
    End Select
    
    Exit Function
    
End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da ceélula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
    
        'Verifica qual é o grid
        If objGridInt.objGrid.Name = GridBloqueio.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_Libera_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, Libera)
                    If lErro <> SUCESSO Then gError 190826
                
                Case iGrid_Forma_Col

                    lErro = Saida_Celula_Padrao(objGridInt, Forma)
                    If lErro <> SUCESSO Then gError 190827
                    
            End Select
            
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 190296

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 190296
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 190826, 190827

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190297)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Bloqueio(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Bloqueio
    
    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Libera")
    objGridInt.colColuna.Add ("Ocorrência")
    objGridInt.colColuna.Add ("Forma")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Emissão")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Série")
    objGridInt.colColuna.Add ("Voucher")
    objGridInt.colColuna.Add ("Observação")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Libera.Name)
    objGridInt.colCampo.Add (Numero.Name)
    objGridInt.colCampo.Add (Forma.Name)
    objGridInt.colCampo.Add (Cliente.Name)
    objGridInt.colCampo.Add (DataEmissao.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (Tipo.Name)
    objGridInt.colCampo.Add (Serie.Name)
    objGridInt.colCampo.Add (Voucher.Name)
    objGridInt.colCampo.Add (Observacao.Name)
    
    iGrid_Libera_Col = 1
    iGrid_Numero_Col = 2
    iGrid_Forma_Col = 3
    iGrid_Cliente_Col = 4
    iGrid_DataEmissao_Col = 5
    iGrid_Valor_Col = 6
    iGrid_Tipo_Col = 7
    iGrid_Serie_Col = 8
    iGrid_Voucher_Col = 9
    iGrid_Observacao_Col = 10
    
    objGridInt.objGrid = GridBloqueio

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 9

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'largura da primeira coluna
    GridBloqueio.ColWidth(0) = 400

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Não permite incluir novas linhas
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Exit Function

End Function

Private Sub BotaoLibera_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLibera_Click

    'Chama rotina de Gravação
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 190298
    
    'Descarrega o Grid de Bloqueios
    lErro = Traz_Bloqueios_Tela(True)
    If lErro <> SUCESSO Then gError 190299
    
    Exit Sub

Erro_BotaoLibera_Click:

    Select Case gErr

        Case 190298, 190299

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 190300)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colOcorrencias As New Collection
Dim objOcr As ClassTRPOcorrencias
Dim objVou As ClassTRPVouchers
Dim bAbreTela As Boolean

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Passa os itens do Grid para a colecao
    lErro = Move_Tela_Memoria(colOcorrencias)
    If lErro <> SUCESSO Then gError 190301
    
    'Verifica se deve ser considerado Impostos + Tarifas
'    lErro = Verifica_Inativacao_VouCartao(colOcorrencias)
'    If lErro <> SUCESSO Then gError 190302

    bAbreTela = False

    For Each objOcr In colOcorrencias
    
        'Se é uma inativação
        If objOcr.iOrigem = INATIVACAO_AUTOMATICA_CODIGO Then
        
            Set objVou = New ClassTRPVouchers

            objVou.sTipVou = objOcr.sTipoDoc
            objVou.sSerie = objOcr.sSerie
            objVou.lNumVou = objOcr.lNumVou
    
            lErro = CF("TRPVouchers_Le", objVou)
            If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 192786
            
            'Se é um voucher de cartão
            If objVou.iCartao = MARCADO Then bAbreTela = True
            
        End If
        
    Next

    If bAbreTela Then
        Call Chama_Tela_Modal("TRPLiberaOcrAux", colOcorrencias)
        If giRetornoTela <> vbOK Then gError 192822
    End If
    
    'Libera os Bloqueios selecionados
    lErro = CF("TRPOcorrencias_Libera", colOcorrencias)
    If lErro <> SUCESSO Then gError 190303
  
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 190301 To 190303, 192822

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190304)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(ByVal colOcorrencias As Collection) As Long
'move para colBloqueioPV os bloqueios marcados para liberação  (Só move o pedido e o tipo de bloqueio pois é o suficiente)

Dim lErro As Long
Dim iIndice As Integer
Dim objOcorrencia As ClassTRPOcorrencias

On Error GoTo Erro_Move_Tela_Memoria

    For iIndice = 1 To objGridBloqueio.iLinhasExistentes
        
        'se o elemento está marcado para ser liberado
        If GridBloqueio.TextMatrix(iIndice, iGrid_Libera_Col) = GRID_CHECKBOX_ATIVO Then
        
            Set objOcorrencia = gobjTRPLiberaOcrSel.colOcorrenciais.Item(iIndice)
            
            objOcorrencia.iFormaPagto = Codigo_Extrai(GridBloqueio.TextMatrix(iIndice, iGrid_Forma_Col))
        
            colOcorrencias.Add objOcorrencia
            
        End If
        
    Next
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190305)
            
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_LIBERACAO_BLOQUEIO_SELECAO
    Set Form_Load_Ocx = Me
    Caption = "Liberação de ocorrências"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TRPLiberaOcr"
    
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

Private Sub TabStripOpcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStripOpcao)
End Sub

Public Sub Forma_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Forma_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridBloqueio)
End Sub

Public Sub Forma_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)
End Sub

Public Sub Forma_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridBloqueio.objControle = Forma
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Function Saida_Celula_Padrao(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Padrao

    Set objGridInt.objControle = objControle
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 190822

    Saida_Celula_Padrao = SUCESSO

    Exit Function

Erro_Saida_Celula_Padrao:

    Saida_Celula_Padrao = gErr

    Select Case gErr

        Case 190822
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190823)

    End Select

    Exit Function

End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is OcorrenciaDe Then Call LabelOCRDe_Click
        If Me.ActiveControl Is OcorrenciaAte Then Call LabelOCRAte_Click
    
    End If
    
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub
