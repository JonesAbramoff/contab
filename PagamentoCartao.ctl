VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl PagamentoCartao 
   ClientHeight    =   4905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   KeyPreview      =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   6630
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
      Left            =   3630
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4410
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
      Left            =   1425
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4410
      Width           =   1725
   End
   Begin VB.ComboBox Adm 
      Height          =   315
      ItemData        =   "PagamentoCartao.ctx":0000
      Left            =   1425
      List            =   "PagamentoCartao.ctx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2760
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
      TabIndex        =   5
      Top             =   150
      Width           =   1350
   End
   Begin VB.ComboBox Parcelamento 
      Height          =   315
      ItemData        =   "PagamentoCartao.ctx":0004
      Left            =   1410
      List            =   "PagamentoCartao.ctx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   615
      Width           =   2760
   End
   Begin VB.Frame FrameCartao 
      Caption         =   "Cart�es de Cr�dito/D�bito"
      Height          =   2715
      Left            =   135
      TabIndex        =   8
      Top             =   1545
      Width           =   6345
      Begin MSMask.MaskEdBox AutorizacaoCartao 
         Height          =   255
         Left            =   4650
         TabIndex        =   18
         Top             =   765
         Width           =   1485
         _ExtentX        =   2619
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox AdmCartao 
         Height          =   255
         Left            =   495
         TabIndex        =   10
         Top             =   795
         Width           =   1350
         _ExtentX        =   2381
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ParcelamentoCartao 
         Height          =   255
         Left            =   1890
         TabIndex        =   11
         Top             =   750
         Width           =   1215
         _ExtentX        =   2143
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ValorCartao 
         Height          =   255
         Left            =   3225
         TabIndex        =   12
         Top             =   735
         Width           =   1365
         _ExtentX        =   2408
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridCartoes 
         Height          =   1995
         Left            =   150
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   300
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   3519
         _Version        =   393216
         Rows            =   5
         Cols            =   5
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Total Cart�o:"
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
         Left            =   3510
         TabIndex        =   14
         Top             =   2370
         Width           =   1125
      End
      Begin VB.Label TotalCartoes 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4695
         TabIndex        =   13
         Top             =   2310
         Width           =   1365
      End
   End
   Begin MSMask.MaskEdBox Valor 
      Height          =   300
      Left            =   1395
      TabIndex        =   3
      Top             =   1110
      Width           =   1260
      _ExtentX        =   2223
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
      Left            =   4890
      TabIndex        =   4
      Top             =   1140
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
      Caption         =   "Autoriza��o:"
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
      Left            =   3780
      TabIndex        =   17
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Cart�o:"
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
      Left            =   720
      TabIndex        =   0
      Top             =   165
      Width           =   630
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
      Left            =   825
      TabIndex        =   7
      Top             =   1170
      Width           =   510
   End
   Begin VB.Label Label2 
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
      Left            =   135
      TabIndex        =   6
      Top             =   660
      Width           =   1230
   End
End
Attribute VB_Name = "PagamentoCartao"
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

'Vari�vel que guarda as caracter�sticas do grid da tela
Dim objGridCartoes As AdmGrid

'Constantes Relacionadas as Colunas do Grid
Dim iGrid_Adm_Col As Integer
'Dim iGrid_Terminal_Col As Integer
Dim iGrid_Parcelamento_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_Autorizacao_Col As Integer


Function Trata_Parametros(objVenda As ClassVenda, iTipo As Integer) As Long
    
Dim objMovimento As ClassMovimentoCaixa
Dim objAdmMeioPagto As ClassAdmMeioPagto
Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim iIndice As Integer
Dim iTipoLoja As Integer
Dim colAdmMeioPagto As New Collection

    Set gobjVenda = objVenda
    
    giTipo = iTipo
    
    If iTipo <> MOVIMENTOCAIXA_RECEB_CARNE_CARTAODEBITO Then
        iTipoLoja = Verifica_Tipo
'        If giBoletoManual = MARCADO And iTipo <> MOVIMENTOCAIXA_RECEB_CARTAODEBITO Then
'           Terminal.AddItem "Manual"
'           Terminal.ItemData(Terminal.NewIndex) = 0
'       End If
        
'        If giPos = MARCADO Then
'            Terminal.AddItem "Pos"
'            Terminal.ItemData(Terminal.NewIndex) = 1
'        End If
    
    Else
        iTipoLoja = TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
        Call CF("AdmMeioPagto_Le_Todas", colAdmMeioPagto)
        Set gcolAdmMeioPagto = colAdmMeioPagto
'        Terminal.AddItem "Pos"
'        Terminal.ItemData(Terminal.NewIndex) = 1
    End If
    
    'Adiciona na combo de AdmMeioPagto todos
    For Each objAdmMeioPagto In gcolAdmMeioPagto
        If objAdmMeioPagto.iTipoMeioPagto = iTipoLoja And objAdmMeioPagto.iAtivo = ADMMEIOPAGTO_ATIVO Then
            Adm.AddItem objAdmMeioPagto.iCodigo & SEPARADOR & objAdmMeioPagto.sNome
            Adm.ItemData(Adm.NewIndex) = objAdmMeioPagto.iCodigo
        End If
    Next
    
    'Joga na tela todos os dados referentes a Contra-vale, Convenio e Vale
    For Each objMovimento In gobjVenda.colMovimentosCaixa
        If objMovimento.iTipo = giTipo And objMovimento.iAdmMeioPagto <> 0 Then
        
            objGridCartoes.iLinhasExistentes = objGridCartoes.iLinhasExistentes + 1
            
            GridCartoes.TextMatrix(objGridCartoes.iLinhasExistentes, iGrid_Valor_Col) = Format(objMovimento.dValor, "standard")
            GridCartoes.TextMatrix(objGridCartoes.iLinhasExistentes, iGrid_Autorizacao_Col) = objMovimento.sAutorizacao
                        
'            'Joga o Terminal na Tela
'            For iIndice = 0 To Terminal.ListCount - 1
'                If Terminal.ItemData(iIndice) = objMovimento.iTipoCartao Then
'                    GridCartoes.TextMatrix(objGridCartoes.iLinhasExistentes, iGrid_Terminal_Col) = Terminal.List(iIndice)
'                    Exit For
'                End If
'            Next
            'Joga o Adm na tela
            For iIndice = 0 To Adm.ListCount - 1
                If Adm.ItemData(iIndice) = objMovimento.iAdmMeioPagto Then
                    GridCartoes.TextMatrix(objGridCartoes.iLinhasExistentes, iGrid_Adm_Col) = Adm.List(iIndice)
                    Exit For
                End If
            Next
            'Joga o Parcelamento na tela
            For iIndice = 1 To gcolAdmMeioPagto.Count
                Set objAdmMeioPagto = gcolAdmMeioPagto.Item(iIndice)
                If objAdmMeioPagto.iCodigo = objMovimento.iAdmMeioPagto And objAdmMeioPagto.iAtivo = ADMMEIOPAGTO_ATIVO Then
                    If iTipo = MOVIMENTOCAIXA_RECEB_CARNE_CARTAODEBITO Then
                        objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
                        Call CF("AdmMeioPagtoCondPagto_Le", objAdmMeioPagto)
                    End If
                    For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
                        If objAdmMeioPagtoCondPagto.iParcelamento = objMovimento.iParcelamento And objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO Then GridCartoes.TextMatrix(objGridCartoes.iLinhasExistentes, iGrid_Parcelamento_Col) = objMovimento.iParcelamento & SEPARADOR & objAdmMeioPagtoCondPagto.sNomeParcelamento
                    Next
                End If
            Next
            
        End If
    Next
    
    'Atualiza o total do troco
    Call Atualiza_Total
        
'    Terminal.ListIndex = 0
        
    Trata_Parametros = SUCESSO

    Exit Function

End Function

Function Verifica_Tipo() As Long

    'Verifica o c�digo que deve ser o MeioPagtoLoja
    Select Case giTipo
    
    Case MOVIMENTOCAIXA_RECEB_CARTAOCREDITO
        Verifica_Tipo = TIPOMEIOPAGTOLOJA_CARTAO_CREDITO
        
    Case MOVIMENTOCAIXA_RECEB_CARTAODEBITO
        Verifica_Tipo = TIPOMEIOPAGTOLOJA_CARTAO_DEBITO
        
    End Select
    
End Function

Public Sub Form_Load()
        
    Set objGridCartoes = New AdmGrid
        
    Call Inicializa_Grid_Cartoes(objGridCartoes)
        
        
    lErro_Chama_Tela = SUCESSO

    Exit Sub

End Sub

Function Inicializa_Grid_Cartoes(objGridInt As AdmGrid) As Long

   'Form do Grid
    Set objGridInt.objForm = Me

    'T�tulos das colunas
    objGridInt.colColuna.Add ("")
    objGridInt.colColuna.Add ("Cart�o")
    objGridInt.colColuna.Add ("Parcelamento")
'    objGridInt.colColuna.Add ("Terminal")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("Autoriza��o")
       
    'Controles que participam do Grid
    objGridInt.colCampo.Add (AdmCartao.Name)
    objGridInt.colCampo.Add (ParcelamentoCartao.Name)
'    objGridInt.colCampo.Add (TerminalCartao.Name)
    objGridInt.colCampo.Add (ValorCartao.Name)
    objGridInt.colCampo.Add (AutorizacaoCartao.Name)
        
    'Colunas do Grid
    iGrid_Adm_Col = 1
    iGrid_Parcelamento_Col = 2
'    iGrid_Terminal_Col = 3
    iGrid_Valor_Col = 3
    iGrid_Autorizacao_Col = 4
    
    'Grid do GridInterno
    objGridInt.objGrid = GridCartoes

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = NUM_MAXIMO_CARTOES + 1

    'Linhas vis�veis do grid
    objGridInt.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridCartoes.ColWidth(0) = 400

    'Largura autom�tica para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA
    
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    
    'Chama fun��o que inicializa o Grid
    Call Grid_Inicializa(objGridInt)
    
    Inicializa_Grid_Cartoes = SUCESSO

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
    
    'se Adm n�o selecionado --> Erro.
    If Adm.ListIndex = -1 Then gError 99674
    'se Adm n�o selecionado --> Erro.
    If Parcelamento.ListIndex = -1 Then gError 99675
    'se Adm n�o selecionado --> Erro.
'    If Terminal.ListIndex = -1 Then gError 99676
    'Se valor n�o preenchido --> Erro.
    If Len(Trim(Valor.Text)) = 0 Then gError 99677
    
    'verifica se o valor pago ultrapassa o valor minimo da condicao de pagto
    For Each objAdmMeioPagto In gcolAdmMeioPagto
        If objAdmMeioPagto.iCodigo = Adm.ItemData(Adm.ListIndex) Then
            For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
                If objAdmMeioPagtoCondPagto.iParcelamento = Parcelamento.ItemData(Parcelamento.ListIndex) Then
                    If StrParaDbl(Valor.Text) < objAdmMeioPagtoCondPagto.dValorMinimo Then gError 126817
                    Exit For
                End If
            Next
            Exit For
        End If
    Next
    
    objGridCartoes.iLinhasExistentes = objGridCartoes.iLinhasExistentes + 1
    
    GridCartoes.TextMatrix(objGridCartoes.iLinhasExistentes, iGrid_Adm_Col) = Adm.Text
    GridCartoes.TextMatrix(objGridCartoes.iLinhasExistentes, iGrid_Parcelamento_Col) = Parcelamento.Text
'    GridCartoes.TextMatrix(objGridCartoes.iLinhasExistentes, iGrid_Terminal_Col) = Terminal.Text
    GridCartoes.TextMatrix(objGridCartoes.iLinhasExistentes, iGrid_Valor_Col) = Format(Valor.Text, "standard")
    GridCartoes.TextMatrix(objGridCartoes.iLinhasExistentes, iGrid_Autorizacao_Col) = Autorizacao.Text
        
    'Atualiza o total do troco
    Call Atualiza_Total
        
    'Limpa os campos da tela
    Valor.Text = ""
    Adm.ListIndex = -1
    Parcelamento.Clear
    Autorizacao.Text = ""
    
    Exit Sub

Erro_BotaoIncluir_Click:

    Select Case gErr

        Case 99674
            Call Rotina_ErroECF(vbOKOnly, ERRO_ADMMEIOPAGTO_NAO_SELECIONADO, gErr)
            
        Case 99675
            Call Rotina_ErroECF(vbOKOnly, ERRO_PARCELAMENTO_NAO_SELECIONADO, gErr)
            
        Case 99676
            Call Rotina_ErroECF(vbOKOnly, ERRO_TERMINAL_NAO_SELECIONADO, gErr)
            
        Case 99677
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALOR_NAO_PREENCHIDO2, gErr)
            
        Case 126817
            Call Rotina_ErroECF(vbOKOnly, ERRO_VALORMINIMO_CONDPAGTO, gErr, objAdmMeioPagtoCondPagto.dValorMinimo, Valor.Text)
            
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164172)

    End Select

    Exit Sub

End Sub

Private Sub Atualiza_Total()
    
Dim iIndice As Integer
    
    TotalCartoes.Caption = ""
    
    For iIndice = 1 To objGridCartoes.iLinhasExistentes
        TotalCartoes.Caption = Format(StrParaDbl(TotalCartoes.Caption) + StrParaDbl(GridCartoes.TextMatrix(iIndice, iGrid_Valor_Col)), "standard")
    Next
    
End Sub

Private Sub BotaoOk_Click()

Dim lErro As Long
Dim objMovimento As New ClassMovimentoCaixa
Dim iIndice As Integer
Dim objAdm As ClassAdmMeioPagto
Dim iTipo As Integer
Dim iIndice2 As Integer

On Error GoTo Erro_BotaoOk_Click
    
    'Exclui todos os movimentos de Cartoes especificados
    For iIndice = gobjVenda.colMovimentosCaixa.Count To 1 Step -1
        Set objMovimento = gobjVenda.colMovimentosCaixa.Item(iIndice)
        If objMovimento.iTipo = giTipo And objMovimento.iAdmMeioPagto <> 0 Then gobjVenda.colMovimentosCaixa.Remove (iIndice)
    Next
    
    'Para cada linha do grid...
    For iIndice = 1 To objGridCartoes.iLinhasExistentes
            
        Set objMovimento = New ClassMovimentoCaixa
    
        'Insere um novo movimento
        objMovimento.iFilialEmpresa = giFilialEmpresa
        objMovimento.iCaixa = giCodCaixa
        objMovimento.iCodOperador = giCodOperador
        objMovimento.iAdmMeioPagto = Codigo_Extrai(GridCartoes.TextMatrix(iIndice, iGrid_Adm_Col))
        objMovimento.iParcelamento = Codigo_Extrai(GridCartoes.TextMatrix(iIndice, iGrid_Parcelamento_Col))
        objMovimento.dtDataMovimento = Date
        objMovimento.dValor = StrParaDbl(GridCartoes.TextMatrix(iIndice, iGrid_Valor_Col))
        objMovimento.dHora = CDbl(Time)
        objMovimento.iTipo = giTipo
        objMovimento.lCupomFiscal = gobjVenda.objCupomFiscal.lNumero
        objMovimento.lNumIntExt = gobjVenda.objCupomFiscal.lNumOrcamento
        objMovimento.sAutorizacao = GridCartoes.TextMatrix(iIndice, iGrid_Autorizacao_Col)
        
'        For iIndice2 = 0 To Terminal.ListCount - 1
'            If Terminal.List(iIndice2) = GridCartoes.TextMatrix(iIndice, iGrid_Terminal_Col) Then
'                objMovimento.iTipoCartao = Terminal.ItemData(iIndice2)
'                Exit For
'            End If
'        Next
        
        objMovimento.iTipoCartao = TIPO_POS
        
        gobjVenda.colMovimentosCaixa.Add objMovimento
        
    Next
    
    Unload Me
    
    Exit Sub

Erro_BotaoOk_Click:

    Select Case gErr
            
        Case Else
            lErro = Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 164173)

    End Select

    Exit Sub

End Sub

Private Sub Adm_Click()

Dim objAdmMeioPagtoCondPagto As ClassAdmMeioPagtoCondPagto
Dim objAdmMeioPagto As New ClassAdmMeioPagto
Dim iIndice As Integer
    
    Parcelamento.Clear
    
    If giTipo = MOVIMENTOCAIXA_RECEB_CARNE_CARTAODEBITO Then
        objAdmMeioPagto.iFilialEmpresa = giFilialEmpresa
        objAdmMeioPagto.iCodigo = Codigo_Extrai(Adm.List(Adm.ListIndex))
        Call CF("AdmMeioPagtoCondPagto_Le", objAdmMeioPagto)
        'Adiciona na combo de Parcelamento
        For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
            Parcelamento.AddItem objAdmMeioPagtoCondPagto.iParcelamento & SEPARADOR & objAdmMeioPagtoCondPagto.sNomeParcelamento
            Parcelamento.ItemData(Parcelamento.NewIndex) = objAdmMeioPagtoCondPagto.iParcelamento
        Next
    Else
        For iIndice = 1 To gcolAdmMeioPagto.Count
            Set objAdmMeioPagto = gcolAdmMeioPagto.Item(iIndice)
            If objAdmMeioPagto.iCodigo = Codigo_Extrai(Adm.List(Adm.ListIndex)) And objAdmMeioPagto.iAtivo = ADMMEIOPAGTO_ATIVO Then
                'Adiciona na combo de Parcelamento
                For Each objAdmMeioPagtoCondPagto In objAdmMeioPagto.colCondPagtoLoja
                    If objAdmMeioPagtoCondPagto.iAtivo = ADMMEIOPAGTOCONDPAGTO_ATIVO Then
                        Parcelamento.AddItem objAdmMeioPagtoCondPagto.iParcelamento & SEPARADOR & objAdmMeioPagtoCondPagto.sNomeParcelamento
                        Parcelamento.ItemData(Parcelamento.NewIndex) = objAdmMeioPagtoCondPagto.iParcelamento
                    End If
                Next
            End If
        Next
    End If
    
End Sub

Private Sub Valor_Validate(Cancel As Boolean)
    
Dim lErro As Long
    
On Error GoTo Erro_Valor_Validate
    
    If Len(Trim(Valor.Text)) > 0 Then
    
        lErro = Valor_Positivo_Critica(Valor.Text)
        If lErro <> SUCESSO Then gError 99678
        
    End If
        
    Exit Sub
    
Erro_Valor_Validate:
    
    Cancel = True
    
    Select Case gErr
        
        Case 99678
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164174)

    End Select

    Exit Sub
    
End Sub

Private Sub GridCartoes_Click()

    Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridCartoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        'Variavel n�o definida
        Call Grid_Entrada_Celula(objGridCartoes, iAlterado)
    End If

End Sub

Private Sub GridCartoes_EnterCell()
    'Parametro n�o opcional
    Call Grid_Entrada_Celula(objGridCartoes, iAlterado)

End Sub

Private Sub GridCartoes_GotFocus()

    Call Grid_Recebe_Foco(objGridCartoes)

End Sub

Private Sub GridCartoes_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridCartoes)
    
    'Atualiza o total dos cart�es
    Call Atualiza_Total
    
End Sub

Private Sub GridCartoes_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridCartoes, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridCartoes, iAlterado)
    End If
        
End Sub

Private Sub GridCartoes_LeaveCell()

    Call Saida_Celula(objGridCartoes)

End Sub

Private Sub GridCartoes_LostFocus()

    Call Grid_Libera_Foco(objGridCartoes)

End Sub

Private Sub GridCartoes_RowColChange()

    Call Grid_RowColChange(objGridCartoes)

End Sub

Private Sub GridCartoes_Validate(Cancel As Boolean)

    Call Grid_Libera_Foco(objGridCartoes)
        
End Sub

Private Sub GridCartoes_Scroll()

    Call Grid_Scroll(objGridCartoes)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da c�lula do grid que est� deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Finaliza_Saida_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 99679

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr
        
    Select Case gErr
        
        Case 99679
        
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error$, 164175)

    End Select

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
 
    Call Tela_QueryUnload(Me, Cancel, UnloadMode, iTelaCorrenteAtiva)
      
End Sub

Public Sub Form_Unload(Cancel As Integer)

    'Libera a refer�ncia da tela
    Set gobjVenda = Nothing
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Not gobjVenda Is Nothing Then
    
    If KeyCode = vbKeyReturn Then
        KeyCode = vbKeyTab
    End If
    
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

    'Clique em ins
    If KeyCode = vbKeyF6 Then
        If Not TrocaFoco(Me, BotaoIncluir) Then Exit Sub
        Call BotaoIncluir_Click
    End If
    
    If KeyCode = vbKeyF7 Then
        GridCartoes.SetFocus
    End If
    
    End If
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Pagamentos em Cart�o"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "PagamentoCartao"
    
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





