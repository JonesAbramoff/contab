VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl PagamentoTicket 
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6645
   KeyPreview      =   -1  'True
   ScaleHeight     =   4260
   ScaleWidth      =   6645
   Begin VB.Frame FrameTicket 
      Caption         =   "Ticket"
      Height          =   2700
      Left            =   45
      TabIndex        =   5
      Top             =   945
      Width           =   6390
      Begin MSMask.MaskEdBox AdmMeioPagtoGrid 
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   945
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox ValorGrid 
         Height          =   255
         Left            =   2565
         TabIndex        =   7
         Top             =   930
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox AutorizacaoGrid 
         Height          =   255
         Left            =   4170
         TabIndex        =   8
         Top             =   915
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSFlexGridLib.MSFlexGrid GridTicket 
         Height          =   1950
         Left            =   195
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   285
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   3440
         _Version        =   393216
         Rows            =   5
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label TotalTicket 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4485
         TabIndex        =   13
         Top             =   2280
         Width           =   1425
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Total Ticket: "
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
         TabIndex        =   12
         Top             =   2340
         Width           =   1170
      End
   End
   Begin VB.ComboBox AdmMeioPagto 
      Height          =   315
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   180
      Width           =   1875
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
      Height          =   375
      Left            =   1350
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1725
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
      Height          =   375
      Left            =   3555
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1725
   End
   Begin VB.CommandButton BotaoIncluir 
      Caption         =   "(F6)  Incluir"
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
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   165
      Width           =   1350
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   300
      Left            =   1530
      TabIndex        =   1
      Top             =   615
      Width           =   1215
      _ExtentX        =   2143
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
   Begin MSMask.MaskEdBox Autorizacao 
      Height          =   300
      Left            =   4800
      TabIndex        =   14
      Top             =   615
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Autorização:"
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
      Left            =   3690
      TabIndex        =   15
      Top             =   660
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Administradora:"
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
      Left            =   135
      TabIndex        =   11
      Top             =   210
      Width           =   1320
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
      Left            =   945
      TabIndex        =   10
      Top             =   660
      Width           =   510
   End
End
Attribute VB_Name = "PagamentoTicket"
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

'Variável que guarda as características do grid da tela
Dim objGridTicket As AdmGrid

'Constantes Relacionadas as Colunas do Grid
Dim iGrid_AdmMeioPagto_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Autorizacao_Col As Integer

Function Trata_Parametros(objVenda As ClassVenda) As Long
    
Dim objMovimento As ClassMovimentoCaixa
Dim iIndice As Integer
    
    Set gobjVenda = objVenda
    
    'Joga na tela todos os dados referentes a Ticket
    For Each objMovimento In gobjVenda.colMovimentosCaixa
        If objMovimento.iTipo = MOVIMENTOCAIXA_RECEB_VALETICKET And objMovimento.iAdmMeioPagto <> 0 Then
        
            objGridTicket.iLinhasExistentes = objGridTicket.iLinhasExistentes + 1
                
            GridTicket.TextMatrix(objGridTicket.iLinhasExistentes, iGrid_Valor_Col) = Format(objMovimento.dValor, "standard")
            GridTicket.TextMatrix(objGridTicket.iLinhasExistentes, iGrid_Autorizacao_Col) = objMovimento.sAutorizacao
                    
            For iIndice = 0 To AdmMeioPagto.ListCount - 1
                If AdmMeioPagto.ItemData(iIndice) = objMovimento.iAdmMeioPagto Then
                    GridTicket.TextMatrix(objGridTicket.iLinhasExistentes, iGrid_AdmMeioPagto_Col) = AdmMeioPagto.List(iIndice)
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

    'Adiciona na combo de AdmMeioPagto todos
    For Each objAdmMeioPagto In gcolAdmMeioPagto
        If AFRAC_TipoMeioPagtoTicket(objAdmMeioPagto.iTipoMeioPagto) And objAdmMeioPagto.iAtivo = ADMMEIOPAGTO_ATIVO Then
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
    objGridInt.colColuna.Add ("Administradora")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Autorização")
    
    'Controles que participam do Grid
    objGridInt.colCampo.Add (AdmMeioPagtoGrid.Name)
    objGridInt.colCampo.Add (ValorGrid.Name)
    objGridInt.colCampo.Add (AutorizacaoGrid.Name)
    
    'Colunas do Grid
    iGrid_AdmMeioPagto_Col = 1
    iGrid_Valor_Col = 2
    iGrid_Autorizacao_Col = 3
    
    'Grid do GridInterno
    objGridInt.objGrid = GridTicket

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_TICKET + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridTicket.ColWidth(0) = 400

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
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_BotaoIncluir_Click
    
    'se AdmMeioPagto não selecionado --> Erro.
    If AdmMeioPagto.ListIndex = -1 Then gError 99754
    
    'Se valor não preenchido --> Erro.
    If Len(Trim(Valor.Text)) = 0 Then gError 99756
    
    For Each objAdmMeioPagto In gcolAdmMeioPagto
        If objAdmMeioPagto.iCodigo = AdmMeioPagto.ItemData(AdmMeioPagto.ListIndex) Then
            For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
                If objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO Then
                    If StrParaDbl(Valor.Text) < objAdmMeioPagtoCondPagto.dValorMinimo Then gError 126816
                    Exit For
                End If
            Next
            Exit For
        End If
    Next
    
    objGridTicket.iLinhasExistentes = objGridTicket.iLinhasExistentes + 1
    
    GridTicket.TextMatrix(objGridTicket.iLinhasExistentes, iGrid_AdmMeioPagto_Col) = AdmMeioPagto.Text
    GridTicket.TextMatrix(objGridTicket.iLinhasExistentes, iGrid_Valor_Col) = Format(Valor.Text, "standard")
    GridTicket.TextMatrix(objGridTicket.iLinhasExistentes, iGrid_Autorizacao_Col) = Autorizacao.Text
    
    'Atualiza o total do troco
    Call Atualiza_Total
        
    'Limpa os campos da tela
    Valor.Text = ""
    AdmMeioPagto.ListIndex = -1
    Autorizacao.Text = ""
        
    Exit Sub

Erro_BotaoIncluir_Click:

    Select Case gErr

        Case 99754
            Call Rotina_ErroECF(vbOKOnly, ERRO_ADMMEIOPAGTO_NAO_SELECIONADO, gErr)
            
        Case 99755
            Call Rotina_ErroECF(vbOKOnly, ERRO_QUANTIDADE_NAO_PREENCHIDO1, gErr)
            
        Case 99756
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_NAO_PREENCHIDO2, gErr)
            
        Case 126816
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORMINIMO_CONDPAGTO, gErr, objAdmMeioPagtoCondPagto.dValorMinimo, Valor.Text)
            
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164203)

    End Select

    Exit Sub

End Sub

Private Sub Atualiza_Total()
    
Dim iIndice As Integer
    
    TotalTicket.Caption = ""
    
    For iIndice = 1 To objGridTicket.iLinhasExistentes
        TotalTicket.Caption = Format(StrParaDbl(TotalTicket.Caption) + StrParaDbl(GridTicket.TextMatrix(iIndice, iGrid_Valor_Col)), "standard")
    Next
    
End Sub

Private Sub BotaoOk_Click()

Dim lErro As Long
Dim objMovimento As New ClassMovimentoCaixa
Dim iIndice As Integer
Dim objAdm As ClassAdmMeioPagto
Dim iTipo As Integer
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto

On Error GoTo Erro_BotaoOk_Click
    
    If Not gobjVenda Is Nothing Then
    
    'Exclui todos os movimentos de Ticket
    For iIndice = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
        Set objMovimento = gobjVenda.colMovimentosCaixa.Item(iIndice)
        If objMovimento.iTipo = MOVIMENTOCAIXA_RECEB_VALETICKET And objMovimento.iAdmMeioPagto <> 0 Then gobjVenda.colMovimentosCaixa.Remove (iIndice)
    Next
    
    'Para cada linha do grid...
    For iIndice = 1 To objGridTicket.iLinhasExistentes
            
        Set objMovimento = New ClassMovimentoCaixa
    
        'Insere um novo movimento
        objMovimento.iFilialEmpresa = giFilialEmpresa
        objMovimento.iCaixa = giCodCaixa
        objMovimento.iCodOperador = giCodOperador
        objMovimento.iAdmMeioPagto = Codigo_Extrai(GridTicket.TextMatrix(iIndice, iGrid_AdmMeioPagto_Col))
        For Each objAdmMeioPagtoCondPagto In gcolTicket
            If objAdmMeioPagtoCondPagto.iAdmMeioPagto = objMovimento.iAdmMeioPagto And objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO Then
                objMovimento.iParcelamento = objAdmMeioPagtoCondPagto.iParcelamento
                Exit For
            End If
        Next
        objMovimento.dtDataMovimento = Date
        objMovimento.dValor = StrParaDbl(GridTicket.TextMatrix(iIndice, iGrid_Valor_Col))
        objMovimento.dHora = CDbl(Time)
        objMovimento.iTipo = MOVIMENTOCAIXA_RECEB_VALETICKET
        objMovimento.lCupomFiscal = gobjVenda.objCupomFiscal.lNumero
        objMovimento.lNumIntExt = gobjVenda.objCupomFiscal.lNumOrcamento
        objMovimento.sAutorizacao = GridTicket.TextMatrix(iIndice, iGrid_Autorizacao_Col)
        
        objMovimento.iTipoCartao = TIPO_POS
        
        gobjVenda.colMovimentosCaixa.Add objMovimento
        
    Next
    
    Unload Me
    
    End If
    
    Exit Sub

Erro_BotaoOk_Click:

    Select Case gErr
            
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164204)

    End Select

    Exit Sub

End Sub

Private Sub Valor_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_Valor_Validate
    
    If Len(Trim(Valor.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 99757
        
    End If
        
    Exit Sub
    
Erro_Valor_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99757
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164205)

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

    Call Grid_Trata_Tecla1(KeyCode, objGridTicket)
    
    Call Atualiza_Total
    
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
    If lErro <> SUCESSO Then gError 99759

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr
        
    Select Case gErr
        
        Case 99759
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164207)

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
    
    If Not gobjVenda Is Nothing Then
    
    Select Case KeyCode
    
        Case vbKeyF5
            If Not TrocaFoco(Me, BotaoOk) Then Exit Sub
            Call BotaoOk_Click

        Case vbKeyEscape
            If Not TrocaFoco(Me, BotaoCancelar) Then Exit Sub
            Call BotaoCancelar_Click

        Case vbKeyF6
            If Not TrocaFoco(Me, BotaoIncluir) Then Exit Sub
            Call BotaoIncluir_Click
    
        Case vbKeyF7
            GridTicket.SetFocus
            
    End Select
    
    End If
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Pagamentos em Ticket"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PagamentoTicket"
    
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



