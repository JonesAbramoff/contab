VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl PagamentoPrazo 
   ClientHeight    =   5115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6705
   KeyPreview      =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   6705
   Begin VB.ComboBox Parcelamento 
      Height          =   315
      Left            =   1425
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1740
   End
   Begin VB.CommandButton BotaoLimpar 
      Height          =   345
      Left            =   5640
      Picture         =   "PagamentoPrazo.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Limpar"
      Top             =   1155
      Width           =   870
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
      Height          =   360
      Left            =   3420
      TabIndex        =   8
      Top             =   4605
      Width           =   1725
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
      Height          =   360
      Left            =   1275
      TabIndex        =   7
      Top             =   4590
      Width           =   1725
   End
   Begin VB.Frame SSFrame3 
      Caption         =   "Parcelas"
      Height          =   2925
      Left            =   90
      TabIndex        =   11
      Top             =   1530
      Width           =   6525
      Begin MSMask.MaskEdBox DataVencimento 
         Height          =   225
         Left            =   630
         TabIndex        =   4
         Top             =   1170
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorParcela 
         Height          =   225
         Left            =   2085
         TabIndex        =   5
         Top             =   1155
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridParcelas 
         Height          =   2400
         Left            =   375
         TabIndex        =   6
         Top             =   330
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   4233
         _Version        =   393216
         Rows            =   9
         Cols            =   3
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
   End
   Begin MSMask.MaskEdBox Autorizacao 
      Height          =   300
      Left            =   -20000
      TabIndex        =   10
      Top             =   630
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Cliente 
      Height          =   300
      Left            =   4920
      TabIndex        =   1
      Top             =   240
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   20
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox DataReferencia 
      Height          =   300
      Left            =   5205
      TabIndex        =   3
      Top             =   690
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown BotaoData 
      Height          =   315
      Left            =   6225
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   675
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   300
      Left            =   1425
      TabIndex        =   2
      Top             =   675
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   529
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   15
      Format          =   "#,##0.00"
      PromptChar      =   " "
   End
   Begin VB.Label CondPagtoLabel 
      AutoSize        =   -1  'True
      Caption         =   "Parcelamento:"
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
      Left            =   120
      TabIndex        =   17
      Top             =   300
      Width           =   1230
   End
   Begin VB.Label Label2 
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
      Left            =   840
      TabIndex        =   16
      Top             =   750
      Width           =   510
   End
   Begin VB.Label Label1 
      Caption         =   "Data de Referência:"
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
      Index           =   3
      Left            =   3330
      TabIndex        =   15
      Top             =   720
      Width           =   1740
   End
   Begin VB.Label LabelCliente 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
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
      Left            =   4155
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   14
      Top             =   300
      Width           =   720
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   -20000
      TabIndex        =   12
      Top             =   690
      Visible         =   0   'False
      Width           =   1080
   End
End
Attribute VB_Name = "PagamentoPrazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjVenda As ClassVenda
Dim gobjAdm As ClassAdmMeioPagto
Dim iAlterado As Integer
Dim gdtData As Date

'Variável que guarda as características do grid da tela
Dim objGridParcelas As AdmGrid

'Constantes Relacionadas as Colunas do Grid
Dim iGrid_Data_Col As Integer
Dim iGrid_Valor_Col As Integer


Function Trata_Parametros(objVenda As ClassVenda) As Long
    
Dim objCarneParcelas As ClassCarneParcelas
Dim objMovimento As ClassMovimentoCaixa
Dim iIndice As Integer
Dim dValor As Double

    Set gobjVenda = objVenda
        
    'Joga na tela todas as parcelas
    For Each objCarneParcelas In gobjVenda.objCarne.colParcelas
        
        objGridParcelas.iLinhasExistentes = objGridParcelas.iLinhasExistentes + 1
            
        GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, iGrid_Data_Col) = objCarneParcelas.dtDataVencimento
        GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, iGrid_Valor_Col) = Format(objCarneParcelas.dValor, "standard")
        dValor = dValor + objCarneParcelas.dValor
        
    Next
    
    '!!!!!RAFAEL
    
    For iIndice = 0 To (Parcelamento.ListCount - 1)
        If Parcelamento.ItemData(iIndice) = objVenda.objCarne.iParcelamento Then
            Parcelamento.ListIndex = iIndice
            Exit For
        End If
    Next
        
    'Verifica se objVenda possui algum valor para cliente. Se possuir, preenche o campo Cliente
    If objVenda.objCarne.lCliente <> 0 Then
        Cliente.Text = objVenda.objCarne.lCliente
    End If
     
    'Verifica se gobjVenda possui algum valor para data. Se possuir, preenche o campo DataReferencia
    If objVenda.objCarne.dtDataReferencia <> 0 Then
        DataReferencia.Text = Format(objVenda.objCarne.dtDataReferencia, "dd/mm/yy")
    Else
        DataReferencia.Text = Format(Date, "dd/mm/yy")
    End If
    
    'Verifica se dValor possui algum valor. Se possuir, preenche o campo Valor
    If dValor <> 0 Then
        Valor.Text = Format(dValor, "standard")
    ElseIf objVenda.dValorTEF <> 0 Then
        Valor.Text = Format(objVenda.dValorTEF, "standard")
    End If
    
    '!!!!!!RAFAEL
            
    gdtData = DATA_NULA
    
    Trata_Parametros = SUCESSO

    Exit Function

End Function

Public Sub Form_Load()
    
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim iIndice As Integer

    Set objGridParcelas = New AdmGrid
        
    'Joga na tela todos os dados referentes ao carne
    For iIndice = 1 To gcolAdmMeioPagto.Count
        Set objAdmMeioPagto = gcolAdmMeioPagto.Item(iIndice)
        If objAdmMeioPagto.iCodigo = MEIO_PAGAMENTO_CARNE Then
            'Carrega a coleção global com os dados do Adm de carne
            Set gobjAdm = gcolAdmMeioPagto.Item(iIndice)
            'Adiciona na combo de Parcelamento
            For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
                If objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO Then
                    Parcelamento.AddItem objAdmMeioPagtoCondPagto.sNomeParcelamento
                    Parcelamento.ItemData(Parcelamento.NewIndex) = objAdmMeioPagtoCondPagto.iParcelamento
                End If
            Next
        End If
    Next
    
    Call Inicializa_Grid_Parcelas(objGridParcelas)
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub
    
Erro_Form_Load:

    Select Case gErr

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164192)

    End Select

    Exit Sub

End Sub

Function Inicializa_Grid_Parcelas(objGridInt As AdmGrid) As Long

   'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Valor")
       
    'Controles que participam do Grid
    objGridInt.colCampo.Add (DataVencimento.Name)
    objGridInt.colCampo.Add (ValorParcela.Name)
        
    'Colunas do Grid
    iGrid_Data_Col = 1
    iGrid_Valor_Col = 2
    
    'Grid do GridInterno
    objGridInt.objGrid = GridParcelas

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_PARC + 1

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 8

    'Largura da primeira coluna
    GridParcelas.ColWidth(0) = 900

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL
    
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)
    
    Inicializa_Grid_Parcelas = SUCESSO

End Function

Private Sub BotaoCancelar_Click()

    Unload Me
    
End Sub

''!!!!!RAFAEL
'Private Sub BotaoIncluir_Click()
'
'Dim lErro As Long
'
'On Error GoTo Erro_BotaoIncluir_Click
'
'    'se AdmMeioPagto não selecionado --> Erro.
'    If Parcelamento.ListIndex = -1 Then gError 115002
'
'    'Se valor não preenchido --> Erro.
'    If Len(Trim(Valor.Text)) = 0 Then gError 115003
'
'    'Se Data de Referência não preenchida --> Erro.
'    If Len(Trim(DataReferencia.ClipText)) = 0 Then gError 115004
'
'    'Se Cliente não preenchido --> Erro.
'    If Len(Trim(Cliente.Text)) = 0 Then gError 115005
'
'    objGridParcelas.iLinhasExistentes = objGridParcelas.iLinhasExistentes + 1
'
'    GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, iGrid_Data_Col) = DataReferencia.Text
'    GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, iGrid_Valor_Col) = Valor.Text
'    GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, iGrid_Cliente_Col) = Cliente.Text
'
'    'Atualiza o total do troco
'    Call Atualiza_Total
'
'    'Limpa os campos da tela
'    Valor.Text = ""
'    Parcelamento.ListIndex = -1
'    Cliente.Text = ""
'    DataReferencia.Text = ""
'
'    Exit Sub
'
'Erro_BotaoIncluir_Click:
'
'    Select Case gErr
'
'        Case 115002
'            Call Rotina_ErroECF(vbOKOnly, ERRO_PARCELAMENTO_NAO_PREENCHIDO, gErr)
'
'        Case 115003
'            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_NAO_PREENCHIDO2, gErr)
'
'        Case 115004
'            Call Rotina_ErroECF(vbOKOnly, ERRO_DATA_NAO_PREENCHIDA1, gErr)
'
'        Case 115005
'            Call Rotina_ErroECF(vbOKOnly, ERRO_CLIENTE_NAO_PREENCHIDO1, gErr)
'
'        Case Else
'            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164193)
'
'    End Select
'
'    Exit Sub
'
'End Sub
''!!!!!RAFAEL

Private Sub BotaoLimpar_Click()
    
    'Limpa os campos da tela
    Call Limpa_Tela(Me)
    Call Grid_Limpa(objGridParcelas)

    Parcelamento.ListIndex = -1
    
    Exit Sub
    
End Sub

Private Sub BotaoOk_Click()

Dim lErro As Long
Dim objMovimento As New ClassMovimentoCaixa
Dim iIndice As Integer
Dim objCarneParcelas As ClassCarneParcelas
Dim iTipo As Integer
Dim dValor As Double

On Error GoTo Erro_BotaoOk_Click
    
    '!!!!!RAFAEL
    
    If (Parcelamento.ListIndex <> -1) Or (Len(Trim(Valor.Text)) <> 0) Or (Len(Trim(Cliente.Text)) <> 0) Or (Len(Trim(DataReferencia.ClipText)) <> 0) Then
    
        If Parcelamento.ListIndex = -1 Then
            gError 115002
        End If
        
        If Len(Trim(Valor.Text)) = 0 Then
            gError 115003
        End If
        
        If Len(Trim(Cliente.Text)) = 0 Then
            gError 115004
        End If
        
        If Len(Trim(DataReferencia.ClipText)) = 0 Then
            gError 115005
        End If
        
        'Insere um novo carne
        gobjVenda.objCarne.dtDataReferencia = CDate(DataReferencia.Text)
        gobjVenda.objCarne.iFilialEmpresa = giFilialEmpresa
        gobjVenda.objCarne.lCupomFiscal = gobjVenda.objCupomFiscal.lNumero
        gobjVenda.objCarne.lNumIntExt = gobjVenda.objCupomFiscal.lNumOrcamento
        gobjVenda.objCarne.iParcelamento = Parcelamento.ItemData(Parcelamento.ListIndex)
        gobjVenda.objCarne.lCliente = StrParaLong(Cliente.Text)
        gobjVenda.objCarne.iStatus = STATUS_LANCADO
    
        dValor = 0
        'Calcula o valor total das parcelas
        For iIndice = 1 To objGridParcelas.iLinhasExistentes
            dValor = dValor + StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))
        Next
        
        If dValor - gobjVenda.objCupomFiscal.dValorTotal > 0.0001 Then gError 99691
    
    
    Else
    
        gobjVenda.objCarne.dtDataReferencia = DATA_NULA
        gobjVenda.objCarne.iFilialEmpresa = 0
        gobjVenda.objCarne.lCupomFiscal = 0
        gobjVenda.objCarne.lNumIntExt = 0
        gobjVenda.objCarne.iParcelamento = 0
        gobjVenda.objCarne.lCliente = 0
        
    End If
    
    
    'Exclui todos os movimentos de Troca
    Set gobjVenda.objCarne.colParcelas = New Collection
            
    'Exclui todos os movimentos de Parcelas
    For iIndice = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
        Set objMovimento = gobjVenda.colMovimentosCaixa.Item(iIndice)
        If objMovimento.iTipo = MOVIMENTOCAIXA_RECEB_CARNE Then gobjVenda.colMovimentosCaixa.Remove (iIndice)
    Next
    
    'Para cada linha do grid...
    For iIndice = 1 To objGridParcelas.iLinhasExistentes
            
        Set objCarneParcelas = New ClassCarneParcelas
        objCarneParcelas.dtDataVencimento = StrParaDate(GridParcelas.TextMatrix(iIndice, iGrid_Data_Col))
        objCarneParcelas.dValor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))
        objCarneParcelas.iFilialEmpresa = giFilialEmpresa
        objCarneParcelas.iParcela = iIndice
        objCarneParcelas.iStatus = STATUS_LANCADO
        
        'Adiciona na coleção
        gobjVenda.objCarne.colParcelas.Add objCarneParcelas
            
        Set objMovimento = New ClassMovimentoCaixa
    
        'Insere um novo movimento
        objMovimento.iFilialEmpresa = giFilialEmpresa
        objMovimento.iCaixa = giCodCaixa
        objMovimento.iCodOperador = giCodOperador
        objMovimento.iAdmMeioPagto = MEIO_PAGAMENTO_CARNE
        objMovimento.iParcelamento = Parcelamento.ItemData(Parcelamento.NewIndex)
        objMovimento.dtDataMovimento = Date
        objMovimento.dValor = StrParaDbl(GridParcelas.TextMatrix(iIndice, iGrid_Valor_Col))
        objMovimento.dHora = CDbl(Time)
        objMovimento.iTipo = MOVIMENTOCAIXA_RECEB_CARNE
        objMovimento.lNumMovto = iIndice
        objMovimento.lCupomFiscal = gobjVenda.objCupomFiscal.lNumero
        objMovimento.lNumIntExt = gobjVenda.objCupomFiscal.lNumOrcamento
        
        gobjVenda.colMovimentosCaixa.Add objMovimento
        
    Next
    
    Unload Me
    
    Exit Sub

Erro_BotaoOk_Click:

    Select Case gErr
                
        Case 99691
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_CARNE_MAIOR, gErr)
            
        Case 115002
            Call Rotina_ErroECF(vbOKOnly, ERRO_PARCELAMENTO_NAO_PREENCHIDO, gErr)

        Case 115003
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_NAO_PREENCHIDO2, gErr)

        Case 115004
            Call Rotina_ErroECF(vbOKOnly, ERRO_CLIENTE_NAO_PREENCHIDO1, gErr)

        Case 115005
            Call Rotina_ErroECF(vbOKOnly, ERRO_DATA_NAO_PREENCHIDA1, gErr)
            
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164194)

    End Select

    Exit Sub

End Sub

Private Sub LabelCliente_Click()

Dim objCliente As New ClassCliente
Dim colSelecao As Collection
    
     'Chama Tela ClienteLista
    Call Chama_TelaECF_Modal("ClienteLista", objCliente)
        
    If giRetornoTela = vbOK Then
        Cliente.Text = objCliente.lCodigo
    End If
            
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCli As ClassCliente
Dim bAchou As Boolean
    
On Error GoTo Erro_Cliente_Validate
    
    bAchou = True
    
    If Len(Trim(Cliente.Text)) <> 0 Then
    
        bAchou = False
    
        lErro = Valor_Positivo_Critica(Cliente.Text)
        If lErro <> SUCESSO Then gError 115006
        
        'verifica se o cliente esta na colecao
        For Each objCli In gcolCliente
        
            If objCli.lCodigo = CLng(Cliente.Text) Then
                
                bAchou = True
                Exit For
                
            End If
                
        Next
        
    End If
    
    If bAchou = False Then gError 120080
    
    Exit Sub
    
Erro_Cliente_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 115006
        
        Case 120080
            Call Rotina_ErroECF(vbOKOnly, ERRO_CLIENTE_NAO_CADASTRADO2, gErr)
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164195)

    End Select

    Exit Sub

End Sub

'Private Sub Parcelamento_Click()
'
'Dim lErro As Long
'Dim objAdmCondPagto As ClassAdmMeioPagtoCondPagto
'Dim objAdmPagtoParc As ClassAdmMeioPagtoParcelas
'Dim iIndice As Integer
'Dim dValor As Double
'Dim dtData As Date
'
'On Error GoTo Erro_Parcelamento_Click
'
'    'se Adm não selecionado --> sai da função
'    If Parcelamento.ListIndex = -1 Then Exit Sub
'
'    'Se valor não preenchido --> Erro.
'    If Len(Trim(Valor.Text)) = 0 Then
'        Parcelamento.ListIndex = -1
'        gError 99681
'    End If
'
'    'Se Data não preenchido --> Erro.
'    If Len(Trim(DataReferencia.ClipText)) = 0 Then
'        Parcelamento.ListIndex = -1
'        gError 99682
'    End If
'
'    'Se Data não preenchido --> Erro.
'    If Len(Trim(Cliente.Text)) = 0 Then
'        Parcelamento.ListIndex = -1
'        gError 99898
'    End If
'
'    'Se Data não preenchido --> Erro.
'    If Len(Trim(Autorizacao.Text)) = 0 Then
'        Parcelamento.ListIndex = -1
'        gError 99899
'    End If
'
'    gobjVenda.objCarne.lCliente = StrParaLong(Cliente.Text)
'    gobjVenda.objCarne.sAutorizacao = Autorizacao.Text
'
'    'Limpa todos os itens do grid
'    lErro = Grid_Limpa(objGridParcelas)
'    If lErro <> SUCESSO Then gError 99683
'
'    'Seta com o item referente ao item selecionado na combo
'    Set objAdmCondPagto = gobjAdm.colCondPagtoLoja.Item(Parcelamento.ListIndex + 1)
'
'    dtData = DataReferencia.Text
'    gdtData = dtData
'
'    'Enquanto houver parcelas...
'    For iIndice = 1 To objAdmCondPagto.iNumParcelas
'        'Percorre a coleção d parcelas
'        For Each objAdmPagtoParc In objAdmCondPagto.colParcelas
'            'Se for a parcela referente a que se quer pesquisar --> recolhe os dados e joga no grid
'            If objAdmPagtoParc.iParcela = iIndice Then
'
'                objGridParcelas.iLinhasExistentes = objGridParcelas.iLinhasExistentes + 1
'
'                dValor = StrParaDbl(Valor.Text) * objAdmPagtoParc.dPercRecebimento
'
'                dtData = dtData + objAdmPagtoParc.iIntervaloRecebimento
'
'                GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, iGrid_Data_Col) = dtData
'                GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, iGrid_Valor_Col) = Format(dValor, "standard")
'
'                Exit For
'
'            End If
'        Next
'    Next
'
'    'Limpa os campos da tela
'    Call Limpa_Tela(Me)
'    Parcelamento.ListIndex = -1
'
'    Exit Sub
'
'Erro_Parcelamento_Click:
'
'    Select Case gErr
'
'        Case 99680
'            Call Rotina_ErroECF(vbOKOnly, ERRO_PARCELAMENTO_NAO_SELECIONADO, gErr)
'
'        Case 99681
'            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_NAO_PREENCHIDO2, gErr)
'
'        Case 99682
'            Call Rotina_ErroECF(vbOKOnly, ERRO_DATA_NAO_PREENCHIDA1, gErr)
'
'        Case 99683
'
'        Case 99898
'            Call Rotina_ErroECF(vbOKOnly, ERRO_CLIENTE_NAO_PREENCHIDO1, gErr)
'
'        Case 99899
'            Call Rotina_ErroECF(vbOKOnly, ERRO_AUTORIZACAO_NAO_PREENCHIDA, gErr)
'
'        Case Else
'            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164196)
'
'    End Select
'
'    Exit Sub
'
'End Sub

Private Sub DataReferencia_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_DataReferencia_Validate
    
    If Len(Trim(DataReferencia.ClipText)) > 0 Then
    
        lErro = Data_Critica(DataReferencia.Text)
        If lErro <> SUCESSO Then gError 99685
        
    End If
    
    'Limpa todos os itens do grid
    lErro = Grid_Limpa(objGridParcelas)
    If lErro <> SUCESSO Then gError 99683
    
    If Parcelamento.ListIndex <> -1 And Len(Trim(DataReferencia.ClipText)) <> 0 And Len(Trim(Valor.Text)) <> 0 Then
    
        Call Preenche_Grid_Parcelas
        
    End If
    
    Exit Sub
    
Erro_DataReferencia_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99685
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164197)

    End Select

    Exit Sub
        
End Sub

Private Sub Parcelamento_Click()

    Call Parcelamento_Validate(False)
End Sub

Private Sub Parcelamento_Validate(Cancel As Boolean)

Dim lErro As Long
    
On Error GoTo Erro_Parcelamento_Validate
    
    'Limpa todos os itens do grid
    lErro = Grid_Limpa(objGridParcelas)
    If lErro <> SUCESSO Then gError 99683
    
    If Parcelamento.ListIndex <> -1 And Len(Trim(DataReferencia.ClipText)) <> 0 And Len(Trim(Valor.Text)) <> 0 Then
    
        Call Preenche_Grid_Parcelas
        
    End If
    
    Exit Sub
    
Erro_Parcelamento_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164198)

    End Select

    Exit Sub

End Sub

Private Sub Valor_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_Valor_Validate
    
    If Len(Trim(Valor.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 99690
        
        Valor.Text = Format(Valor.Text, "standard")
        
    End If
    
    'Limpa todos os itens do grid
    lErro = Grid_Limpa(objGridParcelas)
    If lErro <> SUCESSO Then gError 99683
    
    If Parcelamento.ListIndex <> -1 And Len(Trim(DataReferencia.ClipText)) <> 0 And Len(Trim(Valor.Text)) <> 0 Then
    
        Call Preenche_Grid_Parcelas
        
    End If
    
    Exit Sub
    
Erro_Valor_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99690
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164199)

    End Select

    Exit Sub
    
End Sub


Private Sub Preenche_Grid_Parcelas()

Dim iIndice As Integer
Dim objAdmCondPagto As ClassAdmMeioPagtoCondPagto
Dim objAdmPagtoParc As ClassAdmMeioPagtoParcelas
Dim dtData As Date
Dim dValor As Double

    'Inicializa as variáveis com o valor dos campos na tela
    dtData = DataReferencia.Text
    gdtData = dtData
    dValor = Valor.Text

    'Seta com o item referente ao item selecionado na combo
    Set objAdmCondPagto = gobjAdm.colCondPagtoLoja.Item(Parcelamento.ListIndex + 1)

    'Enquanto houver parcelas...
    For iIndice = 1 To objAdmCondPagto.iNumParcelas
        'Percorre a coleção d parcelas
        For Each objAdmPagtoParc In objAdmCondPagto.colParcelas
            'Se for a parcela referente a que se quer pesquisar --> recolhe os dados e joga no grid
            If objAdmPagtoParc.iParcela = iIndice Then

                objGridParcelas.iLinhasExistentes = objGridParcelas.iLinhasExistentes + 1

                dValor = StrParaDbl(Valor.Text) * objAdmPagtoParc.dPercRecebimento

                dtData = dtData + objAdmPagtoParc.iIntervaloRecebimento

                GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, iGrid_Data_Col) = Format(dtData, "dd/mm/yyyy")
                GridParcelas.TextMatrix(objGridParcelas.iLinhasExistentes, iGrid_Valor_Col) = Format(dValor, "standard")

                Exit For

            End If
        Next
    Next

End Sub

Private Sub BotaoData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_BotaoData_DownClick

    lErro = Data_Up_Down_Click(DataReferencia, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 99686
    
    Exit Sub

Erro_BotaoData_DownClick:

    Select Case gErr

        Case 99686

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164200)

    End Select

    Exit Sub

End Sub

Private Sub BotaoData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_BotaoData_UpClick

    lErro = Data_Up_Down_Click(DataReferencia, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 99687

    Exit Sub

Erro_BotaoData_UpClick:

    Select Case gErr

        Case 99687

        Case Else
             lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164201)

    End Select

    Exit Sub

End Sub

Private Sub GridParcelas_Click()

    Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        'Variavel não definida
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If

End Sub

Private Sub GridParcelas_EnterCell()
    'Parametro não opcional
    Call Grid_Entrada_Celula(objGridParcelas, iAlterado)

End Sub

Private Sub GridParcelas_GotFocus()

    Call Grid_Recebe_Foco(objGridParcelas)

End Sub

Private Sub GridParcelas_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridParcelas)
    
End Sub

Private Sub GridParcelas_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridParcelas, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridParcelas, iAlterado)
    End If
        
End Sub

Private Sub GridParcelas_LeaveCell()

    Call Saida_Celula(objGridParcelas)

End Sub

Private Sub GridParcelas_LostFocus()

    Call Grid_Libera_Foco(objGridParcelas)

End Sub

Private Sub GridParcelas_RowColChange()

    Call Grid_RowColChange(objGridParcelas)

End Sub

Private Sub GridParcelas_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridParcelas)
        
End Sub

Private Sub GridParcelas_Scroll()

    Call Grid_Scroll(objGridParcelas)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 99688

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr
        
    Select Case gErr
        
        Case 99688
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164202)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

    'Libera a referência da tela
    Set gobjVenda = Nothing
    Set gobjAdm = Nothing
       
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'Clique em f5
    If KeyCode = vbKeyF5 Then
        If Not TrocaFoco(Me, BotaoOk) Then Exit Sub
        Call BotaoOk_Click
    End If

    'Clique em esc
    If KeyCode = vbKeyEscape Then
        If Not TrocaFoco(Me, BotaoCancelar) Then Exit Sub
        Call BotaoCancelar_Click
    End If
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Carnê"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PagamentoPrazo"
    
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




