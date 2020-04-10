VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TRVLiberaAporte 
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
      Height          =   5100
      Index           =   2
      Left            =   135
      TabIndex        =   13
      Top             =   735
      Visible         =   0   'False
      Width           =   9270
      Begin MSMask.MaskEdBox Percentual 
         Height          =   315
         Left            =   1890
         TabIndex        =   32
         Top             =   1905
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0%"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Valor 
         Height          =   315
         Left            =   5820
         TabIndex        =   33
         Top             =   1590
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
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
      Begin MSMask.MaskEdBox PrevDataDe 
         Height          =   315
         Left            =   4380
         TabIndex        =   31
         Top             =   960
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox ValorPrev 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Left            =   2625
         TabIndex        =   30
         Text            =   "ValorPrev"
         Top             =   660
         Width           =   1095
      End
      Begin VB.ComboBox Base 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "TRVLiberaAporte.ctx":0000
         Left            =   855
         List            =   "TRVLiberaAporte.ctx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1305
         Width           =   1815
      End
      Begin VB.ComboBox Forma 
         Height          =   315
         ItemData        =   "TRVLiberaAporte.ctx":0037
         Left            =   2985
         List            =   "TRVLiberaAporte.ctx":0041
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   240
         Width           =   1245
      End
      Begin VB.CheckBox Libera 
         Height          =   315
         Left            =   960
         TabIndex        =   14
         Top             =   315
         Width           =   675
      End
      Begin VB.TextBox Numero 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Left            =   1830
         TabIndex        =   15
         Text            =   "Numero"
         Top             =   315
         Width           =   735
      End
      Begin VB.CommandButton BotaoLibera 
         Caption         =   "Libera os Bloqueios Assinalados"
         Height          =   960
         Left            =   15
         Picture         =   "TRVLiberaAporte.ctx":0063
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4065
         Width           =   1590
      End
      Begin VB.CommandButton BotaoAporte 
         Caption         =   "Editar o Aporte"
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
         Left            =   7455
         TabIndex        =   10
         Top             =   4065
         Width           =   1590
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   570
         Left            =   4635
         Picture         =   "TRVLiberaAporte.ctx":04A5
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4260
         Width           =   1425
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   570
         Left            =   3000
         Picture         =   "TRVLiberaAporte.ctx":1687
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4260
         Width           =   1425
      End
      Begin VB.TextBox Cliente 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Left            =   4545
         TabIndex        =   17
         Text            =   "Cliente"
         Top             =   300
         Width           =   1515
      End
      Begin VB.TextBox ValorReal 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Left            =   3555
         TabIndex        =   19
         Text            =   "ValorReal"
         Top             =   630
         Width           =   1095
      End
      Begin MSMask.MaskEdBox PrevDataAte 
         Height          =   315
         Left            =   4530
         TabIndex        =   16
         Top             =   1470
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox DataPagto 
         Height          =   315
         Left            =   5685
         TabIndex        =   18
         Top             =   315
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid GridBloqueio 
         Height          =   3930
         Left            =   0
         TabIndex        =   6
         Top             =   30
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   6932
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5220
      Index           =   1
      Left            =   135
      TabIndex        =   12
      Top             =   600
      Width           =   9210
      Begin VB.Frame Frame2 
         Caption         =   "Filtros"
         Height          =   3840
         Left            =   1440
         TabIndex        =   21
         Top             =   450
         Width           =   6270
         Begin VB.Frame Frame4 
            Caption         =   "Data prevista para pagamento"
            Height          =   1125
            Left            =   405
            TabIndex        =   25
            Top             =   2160
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
               Left            =   780
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
               Top             =   465
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
               Left            =   4560
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
               Left            =   3000
               TabIndex        =   27
               Top             =   525
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
               TabIndex        =   26
               Top             =   540
               Width           =   315
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Aporte"
            Height          =   1125
            Left            =   405
            TabIndex        =   22
            Top             =   540
            Width           =   5505
            Begin MSMask.MaskEdBox AporteDe 
               Height          =   300
               Left            =   780
               TabIndex        =   0
               Top             =   465
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox AporteAte 
               Height          =   300
               Left            =   3450
               TabIndex        =   1
               Top             =   465
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   529
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   6
               Mask            =   "######"
               PromptChar      =   " "
            End
            Begin VB.Label LabelAporteAte 
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
               TabIndex        =   24
               Top             =   525
               Width           =   360
            End
            Begin VB.Label LabelAporteDe 
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
               TabIndex        =   23
               Top             =   525
               Width           =   315
            End
         End
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
      Left            =   8175
      Picture         =   "TRVLiberaAporte.ctx":26A1
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Fechar"
      Top             =   120
      Width           =   1230
   End
   Begin MSComctlLib.TabStrip TabStripOpcao 
      Height          =   5595
      Left            =   60
      TabIndex        =   20
      Top             =   285
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9869
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seleção"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Aportes a serem liberados"
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
Attribute VB_Name = "TRVLiberaAporte"
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
Dim iGrid_DataPagto_Col  As Integer
Dim iGrid_Base_Col  As Integer
Dim iGrid_ValorPrev_Col As Integer
Dim iGrid_ValorReal_Col As Integer
Dim iGrid_DataDe_Col As Integer
Dim iGrid_DataAte_Col As Integer
Dim iGrid_Percentual_Col As Integer
Dim iGrid_Valor_Col As Integer

Dim gobjTRVLiberaAporteSel As New ClassTRVLiberaAporteSel

'Eventos de Browse
Private WithEvents objEventoAporteDe As AdmEvento
Attribute objEventoAporteDe.VB_VarHelpID = -1
Private WithEvents objEventoAporteAte As AdmEvento
Attribute objEventoAporteAte.VB_VarHelpID = -1

'CONTANTES GLOBAIS DA TELA
Const TAB_Selecao = 1
Const TAB_BLOQUEIOS = 2

Private Function Move_Selecao_Memoria(ByVal objTRVLiberaAporteSel As ClassTRVLiberaAporteSel) As Long

Dim lErro As Long

On Error GoTo Erro_Move_Selecao_Memoria

    objTRVLiberaAporteSel.dtDataPagtoAte = StrParaDate(DataAte.Text)
    objTRVLiberaAporteSel.dtDataPagtoDe = StrParaDate(DataDe.Text)
    objTRVLiberaAporteSel.lCodigoDe = StrParaLong(AporteDe.Text)
    objTRVLiberaAporteSel.lCodigoAte = StrParaLong(AporteAte.Text)
    
    If objTRVLiberaAporteSel.dtDataPagtoAte <> DATA_NULA And objTRVLiberaAporteSel.dtDataPagtoDe <> DATA_NULA Then
        If objTRVLiberaAporteSel.dtDataPagtoAte < objTRVLiberaAporteSel.dtDataPagtoDe Then gError 190761
    End If
            
    Move_Selecao_Memoria = SUCESSO
    
    Exit Function
    
Erro_Move_Selecao_Memoria:

    Move_Selecao_Memoria = gErr
    
    Select Case gErr
    
        Case 190761
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
            
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190762)

    End Select

End Function

Private Function Traz_Bloqueios_Tela(Optional ByVal bTirandoLiberados As Boolean = False) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objTRVLiberaAporteSel As New ClassTRVLiberaAporteSel

On Error GoTo Erro_Traz_Bloqueios_Tela

    'Limpa o GridBloqueio
    Call Grid_Limpa(objGridBloqueio)
    
    lErro = Move_Selecao_Memoria(objTRVLiberaAporteSel)
    If lErro <> SUCESSO Then gError 190763
  
    'Preenche a Coleção de Bloqueios
    lErro = CF("TRVAportes_Le_Bloqueios", objTRVLiberaAporteSel)
    If lErro <> SUCESSO Then gError 190764
    
    If Not bTirandoLiberados Then
        If objTRVLiberaAporteSel.colAportes.Count = 0 Then gError 190765
    End If
    
    Set gobjTRVLiberaAporteSel = objTRVLiberaAporteSel
    
    'Preenche o GridBloqueio
    lErro = Grid_Bloqueio_Preenche(objTRVLiberaAporteSel)
    If lErro <> SUCESSO Then gError 190766
            
    Traz_Bloqueios_Tela = SUCESSO
    
    Exit Function
    
Erro_Traz_Bloqueios_Tela:

    Traz_Bloqueios_Tela = gErr
    
    Select Case gErr

        Case 190763, 190764, 190766
            
        Case 190765 'ERRO_SEM_BLOQUEIOS_PC_SEL
             Call Rotina_Erro(vbOKOnly, "ERRO_SEM_BLOQUEIOS_PC_SEL", gErr)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190767)

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

Private Sub BotaoAporte_Click()
    
Dim lErro As Long
Dim objAporte As New ClassTRVAportes

On Error GoTo Erro_BotaoAporte_Click
    
    'Verifica se alguma linha do Grid está selecionada
    If GridBloqueio.Row = 0 Then gError 190768
    
    'Passa a linha do Grid para o Obj
    objAporte.lNumIntDoc = gobjTRVLiberaAporteSel.colAportes.Item(GridBloqueio.Row).lNumIntDocAporte

    'Chama a tela de Pedidos de Venda
    Call Chama_Tela("TRVAporte", objAporte)
    
    Exit Sub

Erro_BotaoAporte_Click:

    Select Case gErr

        Case 190768
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190769)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set objGridBloqueio = Nothing
    Set gobjTRVLiberaAporteSel = Nothing
    
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
        If lErro <> SUCESSO Then gError 190770

    End If
    
    Exit Sub

Erro_DataAte_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 190770 'Tratado na rotina chamada
         
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190771)

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
        If lErro <> SUCESSO Then gError 190772

    End If

    Exit Sub

Erro_DataDe_Validate:
    
    Cancel = True
    
    Select Case gErr

        Case 190772

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190773)

    End Select

    Exit Sub

End Sub

Private Sub GridBloqueio_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridBloqueio, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)
    End If
    
    Call Ordenacao_ClickGrid(objGridBloqueio)

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

Private Sub LabelAporteAte_Click()

Dim colSelecao As Collection
Dim objAporte As New ClassTRVAportes

    'Preenche PedidoAte com o pedido da tela
    objAporte.lCodigo = StrParaLong(AporteAte.Text)
    
    'Chama Tela PedidoVendaLista
    Call Chama_Tela("TRVAportesLista", colSelecao, objAporte, objEventoAporteAte)

End Sub

Private Sub LabelAporteDe_Click()

Dim colSelecao As Collection
Dim objAporte As New ClassTRVAportes

    'Preenche PedidoAte com o pedido da tela
    objAporte.lCodigo = StrParaLong(AporteDe.Text)
    
    'Chama Tela PedidoVendaLista
    Call Chama_Tela("TRVAportesLista", colSelecao, objAporte, objEventoAporteDe)

End Sub

Private Sub objEventoAporteAte_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objAporte As ClassTRVAportes
Dim bCancel As Boolean

On Error GoTo Erro_objEventoAporteAte_evSelecao

    Set objAporte = obj1
    
    AporteAte.PromptInclude = False
    AporteAte.Text = CStr(objAporte.lCodigo)
    AporteAte.PromptInclude = True

    'Chama o Validate de AporteAte
    Call AporteAte_Validate(bCancel)

    Me.Show

    Exit Sub

Erro_objEventoAporteAte_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190774)

    End Select

    Exit Sub

End Sub

Private Sub objEventoAporteDe_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objAporte As ClassTRVAportes
Dim bCancel As Boolean

On Error GoTo Erro_objEventoAporteDe_evSelecao

    Set objAporte = obj1
    
    AporteDe.PromptInclude = False
    AporteDe.Text = CStr(objAporte.lCodigo)
    AporteDe.PromptInclude = True

    'Chama o Validate de AporteAte
    Call AporteAte_Validate(bCancel)

    Me.Show

    Exit Sub

Erro_objEventoAporteDe_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190775)

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
    Set objEventoAporteDe = New AdmEvento
    Set objEventoAporteAte = New AdmEvento
    
    'Executa a Inicialização do grid Bloqueio
    lErro = Inicializa_Grid_Bloqueio(objGridBloqueio)
    If lErro <> SUCESSO Then gError 190776
    
    lErro = CF("Carrega_Combo_Base", Base)
    If lErro <> SUCESSO Then gError 190777
        
    lErro = CF("Carrega_Combo_FormaPagto", Forma)
    If lErro <> SUCESSO Then gError 190778
        
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 190776 To 190778

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190779)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Private Sub AporteAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(AporteAte, iAlterado)

End Sub

Private Sub AporteDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(AporteDe, iAlterado)

End Sub

Private Sub AporteDe_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAporte As New ClassTRVAportes

On Error GoTo Erro_AporteDe_Validate

    If Len(Trim(AporteDe.Text)) > 0 Then
        
        'Critica para ver se é um Long
        lErro = Long_Critica(AporteDe.Text)
        If lErro <> SUCESSO Then gError 190780
            
        'Se o Pedido Final estiver preenchido então
        If Len(Trim(AporteAte.Text)) > 0 Then
            'Verifica se o Pedido Inicial é maior que o Pedido Final ---- Erro
            If CLng(AporteDe.Text) > CLng(AporteAte.Text) Then gError 190781
        End If
        
        objAporte.lCodigo = CLng(AporteDe.Text)
        
        'Verifica se o aporte está cadastrado no BD
        lErro = CF("TRVAportes_Le", objAporte)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190782
            
        'Pedido não está cadastrado
        If lErro <> SUCESSO Then gError 190783
        
    End If
       
    Exit Sub

Erro_AporteDe_Validate:

    Cancel = True

    Select Case gErr
    
        Case 190780, 190782
        
        Case 190781
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INICIAL_MAIOR_FINAL", gErr)
        
        Case 190783
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVAPORTES_NAO_CADASTRADO", gErr, objAporte.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190784)

    End Select

    Exit Sub

End Sub

Private Sub AporteAte_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objAporte As New ClassTRVAportes

On Error GoTo Erro_AporteAte_Validate

    If Len(Trim(AporteAte.Text)) > 0 Then
        
        'Critica para ver se é um Long
        lErro = Long_Critica(AporteAte.Text)
        If lErro <> SUCESSO Then gError 190785
            
        'Se o Pedido Final estiver preenchido então
        If Len(Trim(AporteDe.Text)) > 0 Then
            'Verifica se o Pedido Inicial é maior que o Pedido Final ---- Erro
            If CLng(AporteDe.Text) > CLng(AporteAte.Text) Then gError 190786
        End If
        
        objAporte.lCodigo = CLng(AporteAte.Text)
        
        'Verifica se o aporte está cadastrado no BD
        lErro = CF("TRVAportes_Le", objAporte)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190787
            
        'Pedido não está cadastrado
        If lErro <> SUCESSO Then gError 190788
        
    End If
       
    Exit Sub

Erro_AporteAte_Validate:

    Cancel = True

    Select Case gErr
    
        Case 190785, 190787
        
        Case 190786
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_INICIAL_MAIOR_FINAL", gErr)
        
        Case 190788
            Call Rotina_Erro(vbOKOnly, "ERRO_TRVAPORTES_NAO_CADASTRADO", gErr, objAporte.lCodigo)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190789)
            
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
            If lErro <> SUCESSO Then gError 190790
            
        End If
    
    End If

    Exit Sub

Erro_TabStripOpcao_Click:

    Select Case gErr
        
        Case 190790
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190791)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownAte_DownClick

    'Diminui a DataAte em 1 dia
    lErro = Data_Up_Down_Click(DataAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 190792

    Exit Sub

Erro_UpDownAte_DownClick:

    Select Case gErr

        Case 190792

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190793)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownAte_UpClick

    'Aumenta a DataAte em 1 dia
    lErro = Data_Up_Down_Click(DataAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 190794

    Exit Sub

Erro_UpDownAte_UpClick:

    Select Case gErr

        Case 190794

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190795)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDe_DownClick

    'Diminui a DataDe em 1 dia
    lErro = Data_Up_Down_Click(DataDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 190796

    Exit Sub

Erro_UpDownDe_DownClick:

    Select Case gErr

        Case 190796

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190797)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownDe_UpClick

    'Aumenta a DataDe em 1 dia
    lErro = Data_Up_Down_Click(DataDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 190798

    Exit Sub

Erro_UpDownDe_UpClick:

    Select Case gErr

        Case 190798

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190799)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(Optional objAporte As ClassTRVAportes) As Long

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Trata_Parametros

    If Not (objAporte Is Nothing) Then
        
        If objAporte.lCodigo > 0 Then AporteDe.Text = CStr(objAporte.lCodigo)
        If objAporte.lCodigo > 0 Then AporteAte.Text = CStr(objAporte.lCodigo)
                
    End If
    
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190800)

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

Private Function Grid_Bloqueio_Preenche(ByVal objLiberaAporteSel As ClassTRVLiberaAporteSel) As Long
'Preenche o Grid Bloqueio com os dados de colBloqueioLiberacaoInfo

Dim lErro As Long
Dim iLinha As Integer
Dim objAporte As ClassTRVAportes
Dim objAportePagtoC As ClassTRVAportePagtoCond
Dim objCliente As ClassCliente
Dim dValorReal As Double
Dim dValor As Double
Dim dValorUSS As Double

On Error GoTo Erro_Grid_Bloqueio_Preenche

    'Se o número de Bloqueios for maior que o número de linhas do Grid
    If objLiberaAporteSel.colAportes.Count >= objGridBloqueio.objGrid.Rows Then
        Call Refaz_Grid(objGridBloqueio, objLiberaAporteSel.colAportes.Count)
    End If

    iLinha = 0

    'Percorre todos os Bloqueios da Coleção
    For Each objAportePagtoC In objLiberaAporteSel.colAportes

        iLinha = iLinha + 1
        
        Set objAporte = New ClassTRVAportes
        Set objCliente = New ClassCliente
        
        objAporte.lNumIntDoc = objAportePagtoC.lNumIntDocAporte
        
        lErro = CF("TRVAportes_Le", objAporte)
        If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError 190801

        'Passa para a tela os dados do Bloqueio em questão
        GridBloqueio.TextMatrix(iLinha, iGrid_Libera_Col) = CStr(MARCADO)
        GridBloqueio.TextMatrix(iLinha, iGrid_Numero_Col) = CStr(objAporte.lCodigo)
        
        Call Combo_Seleciona_ItemData(Forma, objAportePagtoC.iFormaPagto)
        GridBloqueio.TextMatrix(iLinha, iGrid_Forma_Col) = Forma.Text
        
        Call Combo_Seleciona_ItemData(Base, objAportePagtoC.iBase)
        GridBloqueio.TextMatrix(iLinha, iGrid_Base_Col) = Base.Text
        
        objCliente.lCodigo = objAporte.lCliente
        
        lErro = CF("Cliente_Le", objCliente)
        If lErro <> SUCESSO And lErro <> 12293 Then gError 190802
        
        GridBloqueio.TextMatrix(iLinha, iGrid_Cliente_Col) = CStr(objCliente.lCodigo) & SEPARADOR & objCliente.sNomeReduzido
        
        GridBloqueio.TextMatrix(iLinha, iGrid_DataPagto_Col) = Format(objAportePagtoC.dtDataPagto, "dd/mm/yyyy")
        GridBloqueio.TextMatrix(iLinha, iGrid_ValorPrev_Col) = Format(objAporte.dPrevValor, "STANDARD")

        lErro = CF("Vouchers_Le_Periodo_Cliente", objCliente.lCodigo, objAporte.dtPrevDataDe, objAporte.dtPrevDataAte, dValorReal, dValorUSS)
        If lErro <> SUCESSO Then gError 190803
        
        GridBloqueio.TextMatrix(iLinha, iGrid_ValorReal_Col) = Format(dValorUSS, "STANDARD")
        
        If objAporte.dtPrevDataDe <> DATA_NULA Then
            GridBloqueio.TextMatrix(iLinha, iGrid_DataDe_Col) = Format(objAporte.dtPrevDataDe, "dd/mm/yyyy")
        Else
            GridBloqueio.TextMatrix(iLinha, iGrid_DataDe_Col) = ""
        End If
        
        If objAporte.dtPrevDataAte <> DATA_NULA Then
            GridBloqueio.TextMatrix(iLinha, iGrid_DataAte_Col) = Format(objAporte.dtPrevDataAte, "dd/mm/yyyy")
        Else
            GridBloqueio.TextMatrix(iLinha, iGrid_DataAte_Col) = ""
        End If
        
        GridBloqueio.TextMatrix(iLinha, iGrid_Percentual_Col) = Format(objAportePagtoC.dPercentual, "PERCENT")
        
        dValor = 0
        If objAportePagtoC.iBase = BASE_TRV_APORTE_REAL Then
            If dValorUSS - objAporte.dPrevValor > 0 Then
                dValor = objAportePagtoC.dPercentual * dValorReal
            End If
        Else
            dValor = objAportePagtoC.dPercentual * (dValorUSS - objAporte.dPrevValor) * (dValorReal / dValorUSS) ' (FALTOU O CAMBIO, COO ELES NÂO ESTÃO USANDO NÃO VAI SER FEITO)
        End If
        
        If dValor > 0 Then
            GridBloqueio.TextMatrix(iLinha, iGrid_Valor_Col) = Format(dValor, "STANDARD")
        Else
            GridBloqueio.TextMatrix(iLinha, iGrid_Valor_Col) = ""
        End If
        
    Next

    'Passa para o Obj o número de Bloqueios passados pela Coleção
    objGridBloqueio.iLinhasExistentes = objLiberaAporteSel.colAportes.Count
    
    Call Grid_Refresh_Checkbox(objGridBloqueio)

    Grid_Bloqueio_Preenche = SUCESSO
    
    Exit Function

Erro_Grid_Bloqueio_Preenche:
    
    Grid_Bloqueio_Preenche = gErr
    
    Select Case gErr
    
        Case 190801 To 190803

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190804)
    
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
                    If lErro <> SUCESSO Then gError 190824
                
                Case iGrid_Forma_Col

                    lErro = Saida_Celula_Padrao(objGridInt, Forma)
                    If lErro <> SUCESSO Then gError 190825

                Case iGrid_Valor_Col
                
                    lErro = Saida_Celula_Valor(objGridInt, Valor)
                    If lErro <> SUCESSO Then gError 190805
                    
            End Select
            
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 190806

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 190805, 190824, 190825

        Case 190806
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190807)

    End Select

    Exit Function

End Function

Function Saida_Celula_Valor(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Data que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Valor

    Set objGridInt.objControle = objControle

    If Len(Trim(objControle.Text)) > 0 Then
    
        'Critica o valor informado
        lErro = Valor_Positivo_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 190808

        objControle.Text = Format(objControle.Text, "STANDARD")
       
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 190809

    Saida_Celula_Valor = SUCESSO

    Exit Function

Erro_Saida_Celula_Valor:

    Saida_Celula_Valor = gErr

    Select Case gErr

        Case 190808 To 190809
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190810)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

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
    objGridInt.colColuna.Add ("Aporte")
    objGridInt.colColuna.Add ("Forma")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Pagto")
    objGridInt.colColuna.Add ("Base")
    objGridInt.colColuna.Add ("Vlr Prev")
    objGridInt.colColuna.Add ("Vlr Real")
    objGridInt.colColuna.Add ("De")
    objGridInt.colColuna.Add ("Até")
    objGridInt.colColuna.Add ("%")
    objGridInt.colColuna.Add ("Valor")
    
   'campos de edição do grid
    objGridInt.colCampo.Add (Libera.Name)
    objGridInt.colCampo.Add (Numero.Name)
    objGridInt.colCampo.Add (Forma.Name)
    objGridInt.colCampo.Add (Cliente.Name)
    objGridInt.colCampo.Add (DataPagto.Name)
    objGridInt.colCampo.Add (Base.Name)
    objGridInt.colCampo.Add (ValorPrev.Name)
    objGridInt.colCampo.Add (ValorReal.Name)
    objGridInt.colCampo.Add (PrevDataDe.Name)
    objGridInt.colCampo.Add (PrevDataAte.Name)
    objGridInt.colCampo.Add (Percentual.Name)
    objGridInt.colCampo.Add (Valor.Name)
    
    iGrid_Libera_Col = 1
    iGrid_Numero_Col = 2
    iGrid_Forma_Col = 3
    iGrid_Cliente_Col = 4
    iGrid_DataPagto_Col = 5
    iGrid_Base_Col = 6
    iGrid_ValorPrev_Col = 7
    iGrid_ValorReal_Col = 8
    iGrid_DataDe_Col = 9
    iGrid_DataAte_Col = 10
    iGrid_Percentual_Col = 11
    iGrid_Valor_Col = 12
    
    objGridInt.objGrid = GridBloqueio

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 9

    'todas as linhas do grid
    objGridInt.objGrid.Rows = 100 + 1

    'largura da primeira coluna
    GridBloqueio.ColWidth(0) = 200

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
    If lErro <> SUCESSO Then gError 190811
    
    'Descarrega o Grid de Bloqueios
    lErro = Traz_Bloqueios_Tela(True)
    If lErro <> SUCESSO Then gError 190812
    
    Exit Sub

Erro_BotaoLibera_Click:

    Select Case gErr

        Case 190811, 190812

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 190813)

    End Select

    Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim iIndice As Integer
Dim colAportes As New Collection

On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    'Passa os itens do Grid para a colecao
    lErro = Move_Tela_Memoria(colAportes)
    If lErro <> SUCESSO Then gError 190814
        
    'Libera os Bloqueios selecionados
    lErro = CF("TRVAportes_Libera", colAportes)
    If lErro <> SUCESSO Then gError 190815
  
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case gErr

        Case 190814, 190815

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190816)

    End Select

    Exit Function

End Function

Private Function Move_Tela_Memoria(ByVal colAportes As Collection) As Long
'move para colBloqueioPV os bloqueios marcados para liberação  (Só move o pedido e o tipo de bloqueio pois é o suficiente)

Dim lErro As Long
Dim iIndice As Integer
Dim objAportePagto As ClassTRVAportePagtoCond

On Error GoTo Erro_Move_Tela_Memoria

    For iIndice = 1 To objGridBloqueio.iLinhasExistentes
        
        'se o elemento está marcado para ser liberado
        If GridBloqueio.TextMatrix(iIndice, iGrid_Libera_Col) = GRID_CHECKBOX_ATIVO Then
        
            Set objAportePagto = gobjTRVLiberaAporteSel.colAportes.Item(iIndice)
            
            objAportePagto.iFormaPagto = Codigo_Extrai(GridBloqueio.TextMatrix(iIndice, iGrid_Forma_Col))
            objAportePagto.dValor = StrParaDbl(GridBloqueio.TextMatrix(iIndice, iGrid_Valor_Col))
        
            If objAportePagto.iFormaPagto = 0 Then gError 192373
            If objAportePagto.dValor <= 0 Then gError 192374
        
            colAportes.Add objAportePagto
            
        End If
        
    Next
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 192373
            Call Rotina_Erro(vbOKOnly, "ERRO_FORMAPAGTO_NAO_PREENCHIDA_GRID", gErr, iIndice)
    
        Case 192374
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_PREENCHIDO_GRID", gErr, iIndice)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190817)
            
    End Select
    
    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_LIBERACAO_BLOQUEIO_SELECAO
    Set Form_Load_Ocx = Me
    Caption = "Liberação de Aportes"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TRVLiberaAporte"
    
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

Public Sub Valor_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Valor_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridBloqueio)
End Sub

Public Sub Valor_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)
End Sub

Public Sub Valor_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridBloqueio.objControle = Valor
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Function Saida_Celula_Padrao(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'faz a critica da celula de quantidade do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Padrao

    Set objGridInt.objControle = objControle
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 190819

    Saida_Celula_Padrao = SUCESSO

    Exit Function

Erro_Saida_Celula_Padrao:

    Saida_Celula_Padrao = gErr

    Select Case gErr

        Case 190819
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190820)

    End Select

    Exit Function

End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is AporteDe Then Call LabelAporteDe_Click
        If Me.ActiveControl Is AporteAte Then Call LabelAporteAte_Click
    
    End If
    
End Sub

