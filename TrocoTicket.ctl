VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl TrocoTicket 
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   KeyPreview      =   -1  'True
   ScaleHeight     =   4320
   ScaleWidth      =   7095
   Begin VB.Frame FrameOutros 
      Caption         =   "Tickets"
      Height          =   2145
      Left            =   120
      TabIndex        =   11
      Top             =   1140
      Width           =   6855
      Begin MSMask.MaskEdBox AdmMeioPagtoGrid 
         Height          =   300
         Left            =   630
         TabIndex        =   6
         Top             =   1035
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox QuantidadeGrid 
         Height          =   300
         Left            =   2700
         TabIndex        =   7
         Top             =   975
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorGrid 
         Height          =   300
         Left            =   4110
         TabIndex        =   8
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox TotalGrid 
         Height          =   300
         Left            =   5430
         TabIndex        =   9
         Top             =   960
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridTicket 
         Height          =   1680
         Left            =   120
         TabIndex        =   5
         Top             =   255
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   2963
         _Version        =   393216
         Rows            =   5
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.ComboBox AdmMeioPagto 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   270
      Width           =   1890
   End
   Begin VB.CommandButton BotaoOk 
      Caption         =   "(F5)   Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1545
      TabIndex        =   10
      Top             =   3795
      Width           =   1740
   End
   Begin VB.CommandButton BotaoCancelar 
      Caption         =   "(Esc)  Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3750
      TabIndex        =   12
      Top             =   3795
      Width           =   1740
   End
   Begin VB.CommandButton BotaoIncluir 
      Caption         =   "(Ins)  Incluir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   285
      Width           =   1365
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   300
      Left            =   1065
      TabIndex        =   1
      Top             =   720
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
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
   Begin MSMask.MaskEdBox Quantidade 
      Height          =   300
      Left            =   4575
      TabIndex        =   2
      Top             =   285
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   529
      _Version        =   393216
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
   Begin VB.Label TotalTroco 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   5370
      TabIndex        =   18
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Total Troco Ticket: "
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
      Left            =   3660
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   17
      Top             =   3405
      Width           =   1725
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ticket:"
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
      Left            =   405
      TabIndex        =   16
      Top             =   330
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
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
      Left            =   495
      TabIndex        =   15
      Top             =   780
      Width           =   525
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade:"
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
      Height          =   210
      Left            =   3510
      TabIndex        =   14
      Top             =   345
      Width           =   1065
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Total: "
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
      Left            =   4020
      TabIndex        =   13
      Top             =   810
      Width           =   555
   End
   Begin VB.Label LabelTotal 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   4560
      TabIndex        =   3
      Top             =   765
      Width           =   1560
   End
End
Attribute VB_Name = "TrocoTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjVenda As ClassVenda
Dim iAlterado As Integer
Dim giTipo As Integer

'Variável que guarda as características do grid da tela
Dim objGridTicket As AdmGrid

'Constantes Relacionadas as Colunas do Grid
Dim iGrid_Ticket_Col As Integer
Dim iGrid_Quantidade_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Total_Col As Integer

Function Trata_Parametros(objVenda As ClassVenda) As Long
    
Dim objMovimento As ClassMovimentoCaixa
Dim iIndice As Integer
    
    Set gobjVenda = objVenda
           
    'Joga na tela todos os Tickets
    For Each objMovimento In gobjVenda.colMovimentosCaixa
        If objMovimento.iTipo = giTipo And objMovimento.iAdmMeioPagto <> 0 Then
        
            objGridTicket.iLinhasExistentes = objGridTicket.iLinhasExistentes + 1
                
            GridTicket.TextMatrix(objGridTicket.iLinhasExistentes, iGrid_Quantidade_Col) = QUANTIDADE_DEFAULT
            GridTicket.TextMatrix(objGridTicket.iLinhasExistentes, iGrid_Valor_Col) = Format(objMovimento.dValor, "standard")
            GridTicket.TextMatrix(objGridTicket.iLinhasExistentes, iGrid_Total_Col) = Format(objMovimento.dValor, "standard")
            
            For iIndice = 0 To AdmMeioPagto.ListCount - 1
                If AdmMeioPagto.ItemData(iIndice) = objMovimento.iAdmMeioPagto Then
                    GridTicket.TextMatrix(objGridTicket.iLinhasExistentes, iGrid_Ticket_Col) = AdmMeioPagto.List(iIndice)
                    Exit For
                End If
            Next
        End If
    Next
        
    'Atualiza o total do troco
    Call Atualiza_Total
        
    Trata_Parametros = SUCESSO

    Exit Function

End Function

Public Sub Form_Load()
    
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim colAdmMeioPagto As New Collection
    
    giTipo = MOVIMENTOCAIXA_TROCO_VALE
    
    'Se o projeto <> SGEECF
    If gsNomePrinc <> "SGEECF" Then
        Call CF("AdmMeioPagto_Le_Todas", colAdmMeioPagto)
        Set gcolAdmMeioPagto = colAdmMeioPagto
        giTipo = MOVIMENTOCAIXA_CARNE_TROCO_TICKET
    End If
        
    'Adiciona na combo de AdmMeioPagto somente os de Ticket
    For Each objAdmMeioPagto In gcolAdmMeioPagto
        If AFRAC_TipoMeioPagtoTicket(objAdmMeioPagto.iTipoMeioPagto) Then
            AdmMeioPagto.AddItem objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome
            AdmMeioPagto.ItemData(AdmMeioPagto.NewIndex) = objAdmMeioPagto.iCodigo
        End If
    Next
    
    Set objGridTicket = New AdmGrid
        
    Call Inicializa_Grid_Ticket(objGridTicket)
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub

End Sub

Function Inicializa_Grid_Ticket(objGridInt As AdmGrid) As Long

   'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Ticket")
    objGridInt.colColuna.Add ("Quantidade")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Total")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (AdmMeioPagtoGrid.Name)
    objGridInt.colCampo.Add (QuantidadeGrid.Name)
    objGridInt.colCampo.Add (ValorGrid.Name)
    objGridInt.colCampo.Add (TotalGrid.Name)
    
    'Colunas do Grid
    iGrid_Ticket_Col = 1
    iGrid_Quantidade_Col = 2
    iGrid_Valor_Col = 3
    iGrid_Total_Col = 4
    
    'Grid do GridInterno
    objGridInt.objGrid = GridTicket

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_TICKET + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 4

    'Largura da primeira coluna
    GridTicket.ColWidth(0) = 900

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)
    
    Inicializa_Grid_Ticket = SUCESSO

    Exit Function

End Function

Private Sub BotaoCancelar_Click()

    Unload Me
    
End Sub

Private Sub BotaoIncluir_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoIncluir_Click
    
    'se AdmMeioPagto não selecionado --> Erro.
    If AdmMeioPagto.ListIndex = -1 Then gError 99636
    'Se quantidade não preenchido --> Erro.
    If Len(Trim(Quantidade.Text)) = 0 Then gError 99637
    'Se valor não preenchido --> Erro.
    If Len(Trim(Valor.Text)) = 0 Then gError 99638
    
    objGridTicket.iLinhasExistentes = objGridTicket.iLinhasExistentes + 1
    
    GridTicket.TextMatrix(objGridTicket.iLinhasExistentes, iGrid_Ticket_Col) = AdmMeioPagto.Text
    GridTicket.TextMatrix(objGridTicket.iLinhasExistentes, iGrid_Quantidade_Col) = Quantidade.Text
    GridTicket.TextMatrix(objGridTicket.iLinhasExistentes, iGrid_Valor_Col) = Format(Valor.Text, "standard")
    GridTicket.TextMatrix(objGridTicket.iLinhasExistentes, iGrid_Total_Col) = Format(LabelTotal.Caption, "standard")
    
    'Atualiza o total do troco
    Call Atualiza_Total
        
    'Limpa os campos da tela
    Quantidade.Text = ""
    Valor.Text = ""
    AdmMeioPagto.ListIndex = -1
    LabelTotal.Caption = ""
    
    Exit Sub

Erro_BotaoIncluir_Click:

    Select Case gErr

        Case 99636
            Call Rotina_ErroECF(vbOKOnly, ERRO_TICKET_NAO_SELECIONADO, gErr)
            
        Case 99637
            Call Rotina_ErroECF(vbOKOnly, ERRO_QUANTIDADE_NAO_PREENCHIDO1, gErr)
            
        Case 99638
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_NAO_PREENCHIDO2, gErr)
            
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175619)

    End Select

    Exit Sub

End Sub

Private Sub Atualiza_Total()
    
Dim iIndice As Integer
    
    TotalTroco.Caption = ""
    
    For iIndice = 1 To objGridTicket.iLinhasExistentes
        TotalTroco.Caption = Format(StrParaDbl(TotalTroco.Caption) + StrParaDbl(GridTicket.TextMatrix(iIndice, iGrid_Total_Col)), "standard")
    Next
    
End Sub

Private Sub BotaoOk_Click()

Dim lErro As Long
Dim objMovimento As New ClassMovimentoCaixa
Dim iIndice As Integer
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_BotaoOk_Click

    'Se o valor do troco do Ticket é maior que o do troco passado --> Erro.
    If StrParaDbl(TotalTroco.Caption) - gobjVenda.objCupomFiscal.dValorTroco > 0.00001 Then gError 99640
    
    'Exclui todos os movimentos de Ticket
    For iIndice = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
        Set objMovimento = gobjVenda.colMovimentosCaixa.Item(iIndice)
        If objMovimento.iTipo = giTipo Then gobjVenda.colMovimentosCaixa.Remove (iIndice)
    Next
    
    'Para cada linha do grid...
    For iIndice = 1 To objGridTicket.iLinhasExistentes
            
        Set objMovimento = New ClassMovimentoCaixa
    
        'Insere um novo movimento
        objMovimento.iFilialEmpresa = giFilialEmpresa
        objMovimento.iCaixa = giCodCaixa
        objMovimento.iCodOperador = giCodOperador
        objMovimento.iTipo = giTipo
        objMovimento.iAdmMeioPagto = Codigo_Extrai(GridTicket.TextMatrix(iIndice, iGrid_Ticket_Col))
        
        For Each objAdmMeioPagtoCondPagto In gcolTicket
            'Se encontrou o ticket/parcelamento
            If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovimento.iAdmMeioPagto And objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO Then
                objMovimento.iParcelamento = objAdmMeioPagtoCondPagto.iParcelamento
                Exit For
            End If
        Next
        
        objMovimento.dtDataMovimento = Date
        objMovimento.dValor = StrParaDbl(GridTicket.TextMatrix(iIndice, iGrid_Total_Col))
        objMovimento.dHora = CDbl(Time)

        gobjVenda.colMovimentosCaixa.Add objMovimento
        
    Next
    
    Unload Me
    
    Exit Sub

Erro_BotaoOk_Click:

    Select Case gErr

        Case 99640
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_TROCO_MAIOR, gErr, Error)
            
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175620)

    End Select

    Exit Sub

End Sub

Private Sub Valor_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_Valor_Validate
    
    If Len(Trim(Valor.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 99634
        
        Valor.Text = Format(Valor.Text, "standard")
        
        'Recalcula o valor total
        If Len(Trim(Quantidade.Text)) > 0 Then
            LabelTotal.Caption = Format(StrParaDbl(Quantidade.Text) * StrParaDbl(Valor.Text), "Standard")
        End If
    
    Else
        LabelTotal.Caption = ""
        
    End If
        
    Exit Sub
    
Erro_Valor_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99634
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175621)

    End Select

    Exit Sub
    
End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_Quantidade_Validate
    
    If Len(Trim(Quantidade.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Quantidade.Text)
        If lErro <> SUCESSO Then gError 99635
        
        'Recalcula o Quantidade total
        If Len(Trim(Quantidade.Text)) > 0 Then
            LabelTotal.Caption = Format(StrParaDbl(Quantidade.Text) * StrParaDbl(Valor.Text), "Standard")
        End If
        
    Else
        LabelTotal.Caption = ""
        
    End If
        
    Exit Sub
    
Erro_Quantidade_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99635
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175622)

    End Select

    Exit Sub
    
End Sub

Private Sub GridTicket_Click()

    Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridTicket, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        'Variavel não definida
        Call Grid_Entrada_Celula(objGridTicket, iAlterado)
    End If

End Sub

Private Sub GridTicket_EnterCell()
    'Parametro não opcional
    Call Grid_Entrada_Celula(objGridTicket, iAlterado)

End Sub

Private Sub GridTicket_GotFocus()

    Call Grid_Recebe_Foco(objGridTicket)

End Sub

Private Sub GridTicket_KeyDown(KeyCode As Integer, Shift As Integer)
    
Dim dValor As Double

    If GridTicket.Row <> 0 Then
        dValor = StrParaDbl(GridTicket.TextMatrix(GridTicket.Row, iGrid_Total_Col))
    End If
    
    Call Grid_Trata_Tecla1(KeyCode, objGridTicket)
    
    If KeyCode = vbKeyDelete Then
        TotalTroco.Caption = Format(StrParaDbl(TotalTroco.Caption) - dValor, "standard")
    End If
    
End Sub

Private Sub GridTicket_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridTicket, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridTicket, iAlterado)
    End If
        
End Sub

Private Sub GridTicket_LeaveCell()

    Call Saida_Celula(objGridTicket)

End Sub

Private Sub GridTicket_LostFocus()

    Call Grid_Libera_Foco(objGridTicket)

End Sub

Private Sub GridTicket_RowColChange()

    Call Grid_RowColChange(objGridTicket)

End Sub

Private Sub GridTicket_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridTicket)
        
End Sub

Private Sub GridTicket_Scroll()

    Call Grid_Scroll(objGridTicket)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 99639

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr
        
    Select Case gErr
        
        Case 99639
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 175623)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

    'Libera a referência da tela
    Set gobjVenda = Nothing
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Clique em f5
    If KeyCode = vbKeyF5 Then
        Call BotaoOk_Click
    End If

    'Clique em esc
    If KeyCode = vbKeyEscape Then
        Call BotaoCancelar_Click
    End If

    'Clique em ins
    If KeyCode = vbKeyInsert Then
        Call BotaoIncluir_Click
    End If
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "TrocoTicket"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "TrocoTicket"
    
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

