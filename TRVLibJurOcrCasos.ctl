VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TRVLibJurOcrCasos 
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10515
   KeyPreview      =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   10515
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5715
      Index           =   1
      Left            =   30
      TabIndex        =   12
      Top             =   465
      Width           =   10410
      Begin VB.Frame Frame2 
         Caption         =   "Exibe Ocorrências da Assistência"
         Height          =   4035
         Left            =   1575
         TabIndex        =   20
         Top             =   660
         Width           =   7020
         Begin VB.Frame Frame4 
            Caption         =   "Data de Finalização do Processo"
            Height          =   1125
            Left            =   825
            TabIndex        =   24
            Top             =   2145
            Width           =   5505
            Begin MSComCtl2.UpDown UpDownDe 
               Height          =   300
               Left            =   1920
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   480
               Width           =   225
               _ExtentX        =   423
               _ExtentY        =   529
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin MSMask.MaskEdBox DataDe 
               Height          =   300
               Left            =   795
               TabIndex        =   2
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
            Begin MSMask.MaskEdBox DataAte 
               Height          =   300
               Left            =   3390
               TabIndex        =   4
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
            Begin MSComCtl2.UpDown UpDownAte 
               Height          =   300
               Left            =   4545
               TabIndex        =   5
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
               TabIndex        =   26
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
               TabIndex        =   25
               Top             =   540
               Width           =   315
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Ocorrências"
            Height          =   1080
            Left            =   825
            TabIndex        =   21
            Top             =   645
            Width           =   5505
            Begin MSMask.MaskEdBox OcorrenciaDe 
               Height          =   300
               Left            =   780
               TabIndex        =   0
               Top             =   465
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox OcorrenciaAte 
               Height          =   300
               Left            =   3435
               TabIndex        =   1
               Top             =   465
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   20
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
               TabIndex        =   23
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
               TabIndex        =   22
               Top             =   525
               Width           =   315
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5730
      Index           =   2
      Left            =   30
      TabIndex        =   13
      Top             =   465
      Visible         =   0   'False
      Width           =   10425
      Begin MSMask.MaskEdBox Ccl 
         Height          =   270
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   0
         AllowPrompt     =   -1  'True
         MaxLength       =   10
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
      Begin VB.CommandButton BotaoCcls 
         Caption         =   "Centros de Custo"
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
         Left            =   1785
         TabIndex        =   31
         Top             =   4965
         Visible         =   0   'False
         Width           =   1680
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   285
         Left            =   2955
         TabIndex        =   30
         Top             =   1275
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.TextBox NumProcesso 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   4545
         TabIndex        =   29
         Text            =   "NumProcesso"
         Top             =   1455
         Width           =   1110
      End
      Begin VB.TextBox Voucher 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1095
         TabIndex        =   28
         Text            =   "Voucher"
         Top             =   1275
         Width           =   1020
      End
      Begin VB.TextBox Serie 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   6345
         TabIndex        =   27
         Text            =   "Serie"
         Top             =   735
         Width           =   300
      End
      Begin VB.CheckBox Libera 
         Height          =   210
         Left            =   180
         TabIndex        =   14
         Top             =   735
         Width           =   510
      End
      Begin VB.TextBox Numero 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1005
         TabIndex        =   15
         Text            =   "Numero"
         Top             =   705
         Width           =   735
      End
      Begin VB.TextBox Tipo 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   5655
         TabIndex        =   16
         Text            =   "Tipo"
         Top             =   765
         Width           =   300
      End
      Begin VB.CommandButton BotaoLibera 
         Caption         =   "Gerar documentos financeiros"
         Height          =   960
         Left            =   165
         Picture         =   "TRVLibJurOcrCasos.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4740
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
         Left            =   8775
         TabIndex        =   10
         Top             =   4740
         Width           =   1590
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   570
         Left            =   5460
         Picture         =   "TRVLibJurOcrCasos.ctx":0442
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4935
         Width           =   1425
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   570
         Left            =   3840
         Picture         =   "TRVLibJurOcrCasos.ctx":1624
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4935
         Width           =   1425
      End
      Begin VB.TextBox Favorecido 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2010
         TabIndex        =   17
         Text            =   "Favorecido"
         Top             =   435
         Width           =   1890
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   285
         Left            =   2370
         TabIndex        =   18
         Top             =   810
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   3765
         Left            =   30
         TabIndex        =   6
         Top             =   45
         Width           =   10395
         _ExtentX        =   18336
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
      Left            =   9270
      Picture         =   "TRVLibJurOcrCasos.ctx":263E
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Fechar"
      Top             =   45
      Width           =   1230
   End
   Begin MSComctlLib.TabStrip TabStripOpcao 
      Height          =   6105
      Left            =   0
      TabIndex        =   19
      Top             =   120
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   10769
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ocorrências"
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
Attribute VB_Name = "TRVLibJurOcrCasos"
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
Dim objGridItens As AdmGrid
Dim iGrid_Libera_Col As Integer
Dim iGrid_Numero_Col As Integer
Dim iGrid_Favorecido_Col As Integer
Dim iGrid_Data_Col  As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_NumProcesso_Col As Integer
Dim iGrid_Ccl_Col As Integer
Dim iGrid_Tipo_Col As Integer
Dim iGrid_Serie_Col As Integer
Dim iGrid_Voucher_Col As Integer

Dim gobjLibOcr As New ClassTRVLibOcrAssist

'Eventos de Browse
Private WithEvents objEventoOcrDe As AdmEvento
Attribute objEventoOcrDe.VB_VarHelpID = -1
Private WithEvents objEventoOcrAte As AdmEvento
Attribute objEventoOcrAte.VB_VarHelpID = -1
Private WithEvents objEventoCcl As AdmEvento
Attribute objEventoCcl.VB_VarHelpID = -1

'CONTANTES GLOBAIS DA TELA
Const TAB_Selecao = 1
Const TAB_BLOQUEIOS = 2

Private Function Move_Selecao_Memoria(ByVal objLibOcr As ClassTRVLibOcrAssist) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Selecao_Memoria

    objLibOcr.iTipo = TRV_OCRCASOS_LIB_JUDICIAL

    objLibOcr.dtDataAte = StrParaDate(DataAte.Text)
    objLibOcr.dtDataDe = StrParaDate(DataDe.Text)
    objLibOcr.sCodigoAte = OcorrenciaAte.Text
    objLibOcr.sCodigoDe = OcorrenciaDe.Text
   
    If objLibOcr.dtDataAte <> DATA_NULA And objLibOcr.dtDataDe <> DATA_NULA Then
        If objLibOcr.dtDataAte < objLibOcr.dtDataDe Then gError 208707
    End If
            
    Move_Selecao_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Selecao_Memoria:

    Move_Selecao_Memoria = gErr
    
    Select Case gErr
    
        Case 208707
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208708)

    End Select

End Function

Private Function Traz_Itens_Tela(Optional ByVal bTirandoLiberados As Boolean = False) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objLibOcr As New ClassTRVLibOcrAssist

On Error GoTo Erro_Traz_Itens_Tela

    'Limpa o GridItens
    Call Grid_Limpa(objGridItens)
    
    lErro = Move_Selecao_Memoria(objLibOcr)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
  
    'Preenche a Coleção de Bloqueios
    lErro = CF("TRVLibOcrCasos_Le", objLibOcr)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    If Not bTirandoLiberados Then
        If objLibOcr.colOcorrenciais.Count = 0 Then gError 208709
    End If
    
    Set gobjLibOcr = objLibOcr
    
    'Preenche o GridItens
    lErro = Grid_Itens_Preenche(objLibOcr)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
    Traz_Itens_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Itens_Tela:

    Traz_Itens_Tela = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            
        Case 208709
             Call Rotina_Erro(vbOKOnly, "ERRO_SEM_TRVOCRCASOS_LIB", gErr)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208710)

    End Select

End Function

Private Sub BotaoDesmarcarTodos_Click()
'Desmarca todos os bloqueios do Grid

Dim iLinha As Integer
Dim objBloqueioLiberacaoInfo As ClassBloqueioLiberacaoInfo

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridItens.iLinhasExistentes

        'Desmarca na tela o bloqueio em questão
        GridItens.TextMatrix(iLinha, iGrid_Libera_Col) = GRID_CHECKBOX_INATIVO
        
    Next
    
    'Atualiza na tela os checkbox desmarcados
    Call Grid_Refresh_Checkbox(objGridItens)
    
End Sub

Private Sub BotaoMarcarTodos_Click()
'Marca todos os bloqueios do Grid

Dim iLinha As Integer
Dim objBloqueioLiberacaoInfo As ClassBloqueioLiberacaoInfo

    'Percorre todas as linhas do Grid
    For iLinha = 1 To objGridItens.iLinhasExistentes

        'Marca na tela o bloqueio em questão
        GridItens.TextMatrix(iLinha, iGrid_Libera_Col) = GRID_CHECKBOX_ATIVO
        
    Next
    
    'Atualiza na tela os checkbox marcados
    Call Grid_Refresh_Checkbox(objGridItens)
    
End Sub

Private Sub BotaoOCR_Click()
    
Dim lErro As Long
Dim objOcorrencia As New ClassTRVOcrCasos

On Error GoTo Erro_BotaoOCR_Click
    
    'Verifica se alguma linha do Grid está selecionada
    If GridItens.Row = 0 Then gError 208711
    
    'Passa a linha do Grid para o Obj
    objOcorrencia.sCodigo = gobjLibOcr.colOcorrenciais.Item(GridItens.Row).sCodigo

    'Chama a tela de Pedidos de Venda
    Call Chama_Tela("TRVOcrCasos", objOcorrencia)
    
    Exit Sub

Erro_BotaoOCR_Click:

    Select Case gErr

        Case 208711
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208712)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridItens = Nothing
    Set gobjLibOcr = Nothing
    Set objEventoOcrAte = Nothing
    Set objEventoOcrDe = Nothing
    Set objEventoCcl = Nothing
    
End Sub

Private Sub DataAte_GotFocus()
    Call MaskEdBox_TrataGotFocus(DataAte, iAlterado)
End Sub

Private Sub DataAte_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataAte_Validate

    'Se a DataAte está preenchida
    If Len(DataAte.ClipText) > 0 Then

        'Verifica se a DataAte é válida
        lErro = Data_Critica(DataAte.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    Exit Sub

Erro_DataAte_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM 'Tratado na rotina chamada
         
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208713)

    End Select

    Exit Sub

End Sub

Private Sub DataDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataDe, iAlterado)

End Sub

Private Sub DataDe_Validate(Cancel As Boolean)
'Critica a Data

Dim lErro As Long

On Error GoTo Erro_DataDe_Validate

    'Se a DataDe está preenchida
    If Len(DataDe.ClipText) > 0 Then

        'Verifica se a DataDe é válida
        lErro = Data_Critica(DataDe.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_DataDe_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208714)

    End Select

    Exit Sub

End Sub

Private Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcolColecoes As New Collection

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If
    
    colcolColecoes.Add gobjLibOcr.colOcorrenciais
    
    Call Ordenacao_ClickGrid(objGridItens, , colcolColecoes)

End Sub

Private Sub GridItens_GotFocus()
    Call Grid_Recebe_Foco(objGridItens)
End Sub

Private Sub GridItens_EnterCell()
    Call Grid_Entrada_Celula(objGridItens, iAlterado)
End Sub

Private Sub GridItens_LeaveCell()
    Call Saida_Celula(objGridItens)
End Sub

Private Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
End Sub

Private Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Private Sub GridItens_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridItens)
End Sub

Private Sub GridItens_RowColChange()
    Call Grid_RowColChange(objGridItens)
End Sub

Private Sub GridItens_Scroll()
    Call Grid_Scroll(objGridItens)
End Sub

Private Sub LabelOCRAte_Click()

Dim colSelecao As Collection
Dim objOcorrencia As New ClassTRVOcrCasos

    'Preenche PedidoAte com o pedido da tela
    objOcorrencia.sCodigo = OcorrenciaAte.Text
    
    'Chama Tela PedidoVendaLista
    Call Chama_Tela("TRVOcrCasosLista", colSelecao, objOcorrencia, objEventoOcrAte)

End Sub

Private Sub LabelOCRDe_Click()

Dim colSelecao As Collection
Dim objOcorrencia As New ClassTRVOcrCasos

    'Preenche PedidoAte com o pedido da tela
    objOcorrencia.sCodigo = OcorrenciaDe.Text
    
    'Chama Tela PedidoVendaLista
    Call Chama_Tela("TRVOcrCasosLista", colSelecao, objOcorrencia, objEventoOcrDe)

End Sub

Private Sub objEventoOCRAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOcorrencia As ClassTRVOcrCasos
Dim bCancel As Boolean

On Error GoTo Erro_objEventoOCRAte_evSelecao

    Set objOcorrencia = obj1
    
    OcorrenciaAte.PromptInclude = False
    OcorrenciaAte.Text = objOcorrencia.sCodigo
    OcorrenciaAte.PromptInclude = True

    'Chama o Validate de OcorrenciaAte
    Call OcorrenciaAte_Validate(bCancel)

    Me.Show

    Exit Sub

Erro_objEventoOCRAte_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208717)

    End Select

    Exit Sub

End Sub

Private Sub objEventoOCRDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objOcorrencia As ClassTRVOcrCasos
Dim bCancel As Boolean

On Error GoTo Erro_objEventoOCRDe_evSelecao

    Set objOcorrencia = obj1
    
    OcorrenciaDe.PromptInclude = False
    OcorrenciaDe.Text = objOcorrencia.sCodigo
    OcorrenciaDe.PromptInclude = True

    'Chama o Validate de OcorrenciaAte
    Call OcorrenciaAte_Validate(bCancel)

    Me.Show

    Exit Sub

Erro_objEventoOCRDe_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208718)

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
Dim sMascaraCclPadrao As String

On Error GoTo Erro_Form_Load
    
    iFrameAtual = 1
    
    Set objGridItens = New AdmGrid
    
    'Inicializa os Eventos de Browser
    Set objEventoOcrDe = New AdmEvento
    Set objEventoOcrAte = New AdmEvento
    Set objEventoCcl = New AdmEvento
    
    'Executa a Inicialização do grid Bloqueio
    lErro = Inicializa_Grid_Itens(objGridItens)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Inicializa Máscara de CclPadrao e Ccl
    sMascaraCclPadrao = String(STRING_CCL, 0)

    lErro = MascaraCcl(sMascaraCclPadrao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Ccl.Mask = sMascaraCclPadrao
        
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208719)

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
Dim objOcorrencia As New ClassTRVOcrCasos

On Error GoTo Erro_OcorrenciaDe_Validate

    If Len(Trim(OcorrenciaDe.Text)) > 0 Then
                    
        'Se o Pedido Final estiver preenchido então
        If Len(Trim(OcorrenciaAte.Text)) > 0 Then
            'Verifica se o Pedido Inicial é maior que o Pedido Final ---- Erro
            If OcorrenciaDe.Text > OcorrenciaAte.Text Then gError 208720
        End If
        
        objOcorrencia.sCodigo = Trim(OcorrenciaDe.Text)
        
        'Verifica se o Pedido está cadastrado no BD
        lErro = CF("TRVOcrCasos_Le", objOcorrencia)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM

        If lErro <> SUCESSO Then gError 208721
        
    End If
       
    Exit Sub

Erro_OcorrenciaDe_Validate:

    Cancel = True

    Select Case gErr

        Case 208720
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INICIAL_MAIOR_FINAL", gErr)
        
        Case 208721
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASOS_NAO_CADASTRADO", gErr, objOcorrencia.sCodigo)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208722)

    End Select

    Exit Sub

End Sub

Private Sub OcorrenciaAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objOcorrencia As New ClassTRVOcrCasos

On Error GoTo Erro_OcorrenciaAte_Validate

    If Len(Trim(OcorrenciaAte.Text)) > 0 Then
            
        'Se o Pedido Final estiver preenchido então
        If Len(Trim(OcorrenciaDe.Text)) > 0 Then
            'Verifica se o Pedido Inicial é maior que o Pedido Final ---- Erro
            If OcorrenciaDe.Text > OcorrenciaAte.Text Then gError 208723
        End If
        
        objOcorrencia.sCodigo = Trim(OcorrenciaAte.Text)
        
        'Verifica se o Pedido está cadastrado no BD
        lErro = CF("TRVOcrCasos_Le", objOcorrencia)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
            
        'Pedido não está cadastrado
        If lErro <> SUCESSO Then gError 208724
        
    End If
       
    Exit Sub

Erro_OcorrenciaAte_Validate:

    Cancel = True

    Select Case gErr
        
        Case 208723
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INICIAL_MAIOR_FINAL", gErr)
        
        Case 208724
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVOCRCASOS_NAO_CADASTRADO", gErr, objOcorrencia.sCodigo)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208725)
            
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
            
            lErro = Traz_Itens_Tela
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
            
        End If
    
    End If

    Exit Sub

Erro_TabStripOpcao_Click:

    Select Case gErr
        
        Case ERRO_SEM_MENSAGEM
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208726)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownAte_DownClick

    'Diminui a DataAte em 1 dia
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownAte_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208727)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownAte_UpClick

    'Aumenta a DataAte em 1 dia
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownAte_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208728)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDe_DownClick

    'Diminui a DataDe em 1 dia
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDe_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208729)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDe_UpClick

    'Aumenta a DataDe em 1 dia
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Exit Sub

Erro_UpDownDe_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208730)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objOcorrencia As ClassTRVOcrCasos) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    If Not (objOcorrencia Is Nothing) Then
        
        OcorrenciaDe.Text = objOcorrencia.sCodigo
        OcorrenciaAte.Text = objOcorrencia.sCodigo
                
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208735)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1

    Call Ordenacao_Limpa(objGridItens)

    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub

Private Function Grid_Itens_Preenche(ByVal objLiberaOcrSel As ClassTRVLibOcrAssist) As Long
'Preenche o Grid Bloqueio com os dados de colBloqueioLiberacaoInfo

Dim lErro As Long
Dim iLinha As Integer
Dim objOcorrencia As ClassTRVOcrCasos
Dim objcliente As ClassCliente

On Error GoTo Erro_Grid_Itens_Preenche

    'Se o número de Bloqueios for maior que o número de linhas do Grid
    If objLiberaOcrSel.colOcorrenciais.Count >= objGridItens.objGrid.Rows Then
        Call Refaz_Grid(objGridItens, objLiberaOcrSel.colOcorrenciais.Count)
    End If

    iLinha = 0

    'Percorre todos os Bloqueios da Coleção
    For Each objOcorrencia In objLiberaOcrSel.colOcorrenciais

        iLinha = iLinha + 1
        
        Set objcliente = New ClassCliente

        'Passa para a tela os dados do Bloqueio em questão
        GridItens.TextMatrix(iLinha, iGrid_Libera_Col) = CStr(MARCADO)
        GridItens.TextMatrix(iLinha, iGrid_Numero_Col) = objOcorrencia.sCodigo
        GridItens.TextMatrix(iLinha, iGrid_Favorecido_Col) = objOcorrencia.sNomeFavorecido
        If objOcorrencia.dtDataFimProcesso <> DATA_NULA Then GridItens.TextMatrix(iLinha, iGrid_Data_Col) = Format(objOcorrencia.dtDataFimProcesso, "dd/mm/yyyy")
        GridItens.TextMatrix(iLinha, iGrid_Valor_Col) = Format(objOcorrencia.dValorCondenacao, "STANDARD")
        GridItens.TextMatrix(iLinha, iGrid_NumProcesso_Col) = objOcorrencia.sNumProcesso
        GridItens.TextMatrix(iLinha, iGrid_Tipo_Col) = objOcorrencia.sTipVou
        GridItens.TextMatrix(iLinha, iGrid_Serie_Col) = objOcorrencia.sSerie
        GridItens.TextMatrix(iLinha, iGrid_Voucher_Col) = CStr(objOcorrencia.lNumVou)
        
    Next

    'Passa para o Obj o número de Bloqueios passados pela Coleção
    objGridItens.iLinhasExistentes = objLiberaOcrSel.colOcorrenciais.Count
    
    Call Grid_Refresh_Checkbox(objGridItens)

    Grid_Itens_Preenche = SUCESSO
    
    Exit Function

Erro_Grid_Itens_Preenche:
    
    Grid_Itens_Preenche = gErr
    
    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208736)
    
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
        If objGridInt.objGrid.Name = GridItens.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_Libera_Col
                
                    lErro = Saida_Celula_Padrao(objGridInt, Libera)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                Case iGrid_Ccl_Col
                
                    lErro = Saida_Celula_Ccl(objGridInt)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
            End Select
            
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 208737

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 208737
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208738)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Itens(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Bloqueio
    
    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Libera")
    objGridInt.colColuna.Add ("Ocr")
    objGridInt.colColuna.Add ("Favorecido")
    objGridInt.colColuna.Add ("Fim Processo")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Processo")
    objGridInt.colColuna.Add ("Ccl")
    objGridInt.colColuna.Add ("T")
    objGridInt.colColuna.Add ("S")
    objGridInt.colColuna.Add ("Voucher")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Libera.Name)
    objGridInt.colCampo.Add (Numero.Name)
    objGridInt.colCampo.Add (Favorecido.Name)
    objGridInt.colCampo.Add (Data.Name)
    objGridInt.colCampo.Add (Valor.Name)
    objGridInt.colCampo.Add (NumProcesso.Name)
    objGridInt.colCampo.Add (Ccl.Name)
    objGridInt.colCampo.Add (Tipo.Name)
    objGridInt.colCampo.Add (Serie.Name)
    objGridInt.colCampo.Add (Voucher.Name)
    
    iGrid_Libera_Col = 1
    iGrid_Numero_Col = 2
    iGrid_Favorecido_Col = 3
    iGrid_Data_Col = 4
    iGrid_Valor_Col = 5
    iGrid_NumProcesso_Col = 6
    iGrid_Ccl_Col = 7
    iGrid_Tipo_Col = 8
    iGrid_Serie_Col = 9
    iGrid_Voucher_Col = 10
    
    objGridInt.objGrid = GridItens

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 13

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'largura da primeira coluna
    GridItens.ColWidth(0) = 400

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
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Descarrega o Grid de Bloqueios
    lErro = Traz_Itens_Tela(True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    iAlterado = 0
    
    Exit Sub

Erro_BotaoLibera_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 208739)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim colOcorrencias As New Collection

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Passa os itens do Grid para a colecao
    lErro = Move_Tela_Memoria(colOcorrencias)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    'Libera os Bloqueios selecionados
    lErro = CF("TRVLibOcrCasos_Libera", TRV_OCRCASOS_LIB_JUDICIAL, colOcorrencias)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
  
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208740)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(ByVal colOcorrencias As Collection) As Long
'move para colBloqueioPV os bloqueios marcados para liberação  (Só move o pedido e o tipo de bloqueio pois é o suficiente)

Dim lErro As Long
Dim iIndice As Integer
Dim objOcorrencia As ClassTRVOcrCasos
Dim sCcl As String, sCclFormatada As String, iCclPreenchida As Integer

On Error GoTo Erro_Move_Tela_Memoria

    For iIndice = 1 To objGridItens.iLinhasExistentes
        
        'se o elemento está marcado para ser liberado
        If GridItens.TextMatrix(iIndice, iGrid_Libera_Col) = GRID_CHECKBOX_ATIVO Then
        
            Set objOcorrencia = gobjLibOcr.colOcorrenciais.Item(iIndice)
                    
            sCcl = GridItens.TextMatrix(iIndice, iGrid_Ccl_Col)

            If Len(Trim(sCcl)) = 0 Then gError 209154
    
            'Formata Ccl para BD
            lErro = CF("Ccl_Formata", sCcl, sCclFormatada, iCclPreenchida)
            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
            objOcorrencia.sCcl = sCcl 'sCclFormatada
                    
            colOcorrencias.Add objOcorrencia
            
        End If
        
    Next
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 209154 'ERRO_CCL_GRID_NAO_PREENCHIDO
            Call Rotina_Erro(vbOKOnly, "ERRO_CCL_GRID_NAO_PREENCHIDO", gErr, iIndice)
    
        Case ERRO_SEM_MENSAGEM
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208741)
            
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_LIBERACAO_BLOQUEIO_SELECAO
    Set Form_Load_Ocx = Me
    Caption = "Liberação de ocorrências de processo"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TRVLibJurOcrCasos"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is OcorrenciaDe Then Call LabelOCRDe_Click
        If Me.ActiveControl Is OcorrenciaAte Then Call LabelOCRAte_Click
        If Me.ActiveControl Is Ccl Then Call BotaoCcls_Click
    
    End If
    
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub

Private Function Saida_Celula_Ccl(objGridInt As AdmGrid) As Long
'faz a critica da celula de produto do grid que está deixando de ser a corrente

Dim lErro As Long, sCclFormatada As String
Dim objCcl As New ClassCcl
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Ccl

    Set objGridInt.objControle = Ccl

    If Len(Ccl.ClipText) > 0 Then

        lErro = CF("Ccl_Critica", Ccl.Text, sCclFormatada, objCcl)
        If lErro <> SUCESSO And lErro <> 5703 Then gError 209150

        If lErro = 5703 Then gError 209151

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 209152

    Saida_Celula_Ccl = SUCESSO

    Exit Function

Erro_Saida_Celula_Ccl:

    Saida_Celula_Ccl = gErr

    Select Case gErr

        Case 209150, 209152
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 209151
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CCL_INEXISTENTE", Ccl.Text)
            
            If vbMsgRes = vbYes Then
            
                objCcl.sCcl = sCclFormatada
                
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("CclTela", objCcl)

            Else
            
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
                
            End If

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 209153)

    End Select

    Exit Function

End Function

Private Sub Ccl_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Ccl_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Private Sub Ccl_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Private Sub Ccl_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = Ccl
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub BotaoCcls_Click()
'chama tela de Lista de Ccl

Dim lErro As Long
Dim objCcls As New ClassCcl
Dim colSelecao As New Collection

On Error GoTo Erro_BotaoCcls_Click

    'Verifica se tem alguma linha selecionada no Grid
    If GridItens.Row = 0 Then gError 209154


    Call Chama_Tela("CclLista", colSelecao, objCcls, objEventoCcl)
    
    Exit Sub
    
Erro_BotaoCcls_Click:

    Select Case gErr
    
        Case 209154
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209155)
    
    End Select
    
    Exit Sub
    
End Sub

Private Sub objEventoCcl_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCcl As New ClassCcl
Dim sCclMascarado As String
Dim sCclFormatada As String

On Error GoTo Erro_objEventoCcl_evSelecao

    Set objCcl = obj1

    sCclMascarado = String(STRING_CCL, 0)

    lErro = Mascara_MascararCcl(objCcl.sCcl, sCclMascarado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Coloca o valor do Ccl na coluna correspondente
    GridItens.TextMatrix(GridItens.Row, iGrid_Ccl_Col) = sCclMascarado

    Ccl.PromptInclude = False
    Ccl.Text = sCclMascarado
    Ccl.PromptInclude = True

    Me.Show

    Exit Sub

Erro_objEventoCcl_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 209156)

    End Select

    Exit Sub

End Sub
