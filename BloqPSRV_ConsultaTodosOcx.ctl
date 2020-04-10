VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl BloqPSRV_ConsultaTodosOcx 
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9405
   ScaleHeight     =   4890
   ScaleWidth      =   9405
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   8010
      ScaleHeight     =   450
      ScaleWidth      =   1200
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   195
      Width           =   1260
      Begin VB.CommandButton BotaoGravar 
         Height          =   345
         Left            =   135
         Picture         =   "BloqPSRV_ConsultaTodosOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   675
         Picture         =   "BloqPSRV_ConsultaTodosOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identifica��o"
      Height          =   1185
      Left            =   120
      TabIndex        =   7
      Top             =   105
      Width           =   7560
      Begin VB.Label LabelPedido 
         Caption         =   "Pedido:"
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
         Left            =   1035
         TabIndex        =   15
         Top             =   330
         Width           =   630
      End
      Begin VB.Label LabelCliente 
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
         Height          =   195
         Left            =   1050
         TabIndex        =   14
         Top             =   795
         Width           =   705
      End
      Begin VB.Label LabelData 
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   4065
         TabIndex        =   13
         Top             =   330
         Width           =   630
      End
      Begin VB.Label LabelValor 
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
         Height          =   195
         Left            =   4080
         TabIndex        =   12
         Top             =   780
         Width           =   510
      End
      Begin VB.Label Pedido 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1905
         TabIndex        =   11
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Cliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1905
         TabIndex        =   10
         Top             =   765
         Width           =   1950
      End
      Begin VB.Label Data 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4860
         TabIndex        =   9
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Valor 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4860
         TabIndex        =   8
         Top             =   765
         Width           =   1200
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Bloqueios"
      Height          =   3405
      Left            =   120
      TabIndex        =   0
      Top             =   1335
      Width           =   9165
      Begin VB.ComboBox TipoBloqueio 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "BloqPSRV_ConsultaTodosOcx.ctx":02D8
         Left            =   300
         List            =   "BloqPSRV_ConsultaTodosOcx.ctx":02DA
         TabIndex        =   3
         Top             =   1005
         Width           =   1605
      End
      Begin VB.TextBox ObservacaoBloqueio 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   4365
         MaxLength       =   250
         TabIndex        =   2
         Top             =   1005
         Width           =   4245
      End
      Begin VB.TextBox SeqBloqueio 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   1740
         MaxLength       =   250
         TabIndex        =   1
         Top             =   1455
         Width           =   675
      End
      Begin MSMask.MaskEdBox DataLiberacao 
         Height          =   225
         Left            =   3105
         TabIndex        =   4
         Top             =   1035
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
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
      Begin MSMask.MaskEdBox DataBloqueio 
         Height          =   240
         Left            =   1920
         TabIndex        =   5
         Top             =   1035
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
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
      Begin MSFlexGridLib.MSFlexGrid GridBloqueio 
         Height          =   2010
         Left            =   135
         TabIndex        =   6
         Top             =   270
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   3545
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "BloqPSRV_ConsultaTodosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const STRING_BLOQUEIOSPV_OBSERVACAO = 250

'Grid Bloqueio:
Dim objGridBloqueio As AdmGrid
Dim iGrid_TipoBloqueio_Col As Integer
Dim iGrid_DataBloqueio_Col As Integer
Dim iGrid_DataLiberacao_Col As Integer
Dim iGrid_Observacao_Col As Integer
Dim iGrid_SeqBloqueio_Col As Integer

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

'*** FUN��ES DE INICIALIZA��O DA TELA - IN�CIO ***
Public Sub Form_Load()
    
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load
    
    Set objGridBloqueio = New AdmGrid
    
    'Executa a Inicializa��o do grid Bloqueio
    lErro = Inicializa_Grid_Bloqueio(objGridBloqueio)
    If lErro <> SUCESSO Then gError 191490
    
    'Carrega o Combo TipoBloqueio
    lErro = Carrega_TipoBloqueio
    If lErro <> SUCESSO Then gError 191491
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 191490, 191491

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 191492)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objPedidoDeVenda As ClassPedidoDeVenda) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objPedidoDeVenda Is Nothing) Then
    
        'Chama a fun��o que ir� preencher a tela BloqPV_ConsultaTodos
        lErro = Traz_BloqPSRV_ConsultaTodos_Tela(objPedidoDeVenda)
        If lErro <> SUCESSO Then gError 191493

    End If
        
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 191493
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191493)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Function Inicializa_Grid_Bloqueio(objGridInt As AdmGrid) As Long
'Executa a Inicializa��o do grid Bloqueio
    
    'tela em quest�o
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Data Libera��o")
    objGridInt.colColuna.Add ("Observa��o")
    objGridInt.colColuna.Add ("Seq.")
    
   'campos de edi��o do grid
    objGridInt.colCampo.Add (TipoBloqueio.Name)
    objGridInt.colCampo.Add (DataBloqueio.Name)
    objGridInt.colCampo.Add (DataLiberacao.Name)
    objGridInt.colCampo.Add (ObservacaoBloqueio.Name)
    objGridInt.colCampo.Add (SeqBloqueio.Name)
    
    iGrid_TipoBloqueio_Col = 1
    iGrid_DataBloqueio_Col = 2
    iGrid_DataLiberacao_Col = 3
    iGrid_Observacao_Col = 4
    iGrid_SeqBloqueio_Col = 5
    
    objGridInt.objGrid = GridBloqueio

    'linhas visiveis do grid
    objGridInt.iLinhasVisiveis = 7

    'todas as linhas do grid
    objGridInt.objGrid.Rows = objGridInt.iLinhasVisiveis + 1

    'largura da primeira coluna
    GridBloqueio.ColWidth(0) = 500

    'largura total do grid
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE
    
    'N�o permite incluir novas linhas
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'Chama rotina de Inicializa��o do Grid
    Call Grid_Inicializa(objGridInt)

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da c�lula do grid que est� deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula
    
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual a Coluna em quest�o
        Select Case GridBloqueio.Col

            Case iGrid_TipoBloqueio_Col
                lErro = Saida_Celula_TipoBloqueio(objGridInt)
                If lErro <> SUCESSO Then gError 191494
                
            Case iGrid_DataBloqueio_Col
                lErro = Saida_Celula_DataBloqueio(objGridInt)
                If lErro <> SUCESSO Then gError 191495

            Case iGrid_DataLiberacao_Col
                lErro = Saida_Celula_DataLiberacao(objGridInt)
                If lErro <> SUCESSO Then gError 191496

            Case iGrid_Observacao_Col
                lErro = Saida_Celula_Observacao(objGridInt)
                If lErro <> SUCESSO Then gError 191497

        End Select

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 191498
    
    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 191494 To 191498

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191499)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoBloqueio(objGridInt As AdmGrid) As Long
'Faz a cr�tica da c�lula Tipo Bloqueio que est� deixando de ser a corrente

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer
Dim objTipoDeBloqueio As New ClassTipoDeBloqueio

On Error GoTo Erro_Saida_Celula_TipoBloqueio

    Set objGridInt.objControle = TipoBloqueio

    'Verifica se o Tipo foi preenchido
    If Len(Trim(TipoBloqueio.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If TipoBloqueio.Text <> TipoBloqueio.List(TipoBloqueio.ListIndex) Then

            'Tenta selecion�-lo na combo
            lErro = Combo_Seleciona_Grid(TipoBloqueio, iCodigo)
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 191500

            'N�o foi encontrado
            If lErro = 25085 Then gError 191501
            If lErro = 25086 Then gError 191502

        End If

        'Acrescenta uma linha no Grid se for o caso
        If GridBloqueio.Row - GridBloqueio.FixedRows = objGridInt.iLinhasExistentes Then objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191503

    Saida_Celula_TipoBloqueio = SUCESSO

    Exit Function

Erro_Saida_Celula_TipoBloqueio:

    Saida_Celula_TipoBloqueio = gErr

    Select Case gErr

        Case 191500, 191503
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191501
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_NAO_ENCONTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191502
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_NAO_ENCONTRADO1", gErr, TipoBloqueio.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191504)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataBloqueio(objGridInt As AdmGrid) As Long
'Faz a cr�tica da c�lula DataBloqueio que est� deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataBloqueio

    Set objGridInt.objControle = DataBloqueio
    
    'Critica a data preenchida
    lErro = Data_Critica(DataBloqueio.Text)
    If lErro <> SUCESSO Then gError 191505
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191506
   
    Saida_Celula_DataBloqueio = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_DataBloqueio:

    Saida_Celula_DataBloqueio = gErr
    
    Select Case gErr
    
        Case 191505, 191506
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191507)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataLiberacao(objGridInt As AdmGrid) As Long
'Faz a cr�tica da c�lula DataLiberacao que est� deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_DataLiberacao

    Set objGridInt.objControle = DataLiberacao
    
    'Critica a Data informada
    lErro = Data_Critica(DataLiberacao.Text)
    If lErro <> SUCESSO Then gError 191508
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191509
   
    Saida_Celula_DataLiberacao = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_DataLiberacao:

    Saida_Celula_DataLiberacao = gErr
    
    Select Case gErr
    
        Case 191508, 191509
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191510)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Observacao(objGridInt As AdmGrid) As Long
'Faz a cr�tica da c�lula Observa��o que est� deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Observacao

    Set objGridInt.objControle = ObservacaoBloqueio

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 191511
    
    Saida_Celula_Observacao = SUCESSO
    
    Exit Function
    
Erro_Saida_Celula_Observacao:

    Saida_Celula_Observacao = gErr
    
    Select Case gErr
    
        Case 191511
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191512)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iChamada As Integer)

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa o controle da coluna em quest�o
    Select Case objControl.Name

        Case ObservacaoBloqueio.Name

            'Verifica se o Tipo Bloqueio foi preenchido
            If Len(Trim(GridBloqueio.TextMatrix(GridBloqueio.Row, iGrid_TipoBloqueio_Col))) <> 0 Then
                ObservacaoBloqueio.Enabled = True
            Else
                ObservacaoBloqueio.Enabled = False
            End If

    End Select
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191513)

    End Select

    Exit Sub

End Sub

Private Function Carrega_TipoBloqueio() As Long
'Carrega a Combo TipoBloqueio com as informa��es do BD

Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodDescricao As AdmCodigoNome
Dim lErro As Long

On Error GoTo Erro_Carrega_TipoBloqueio

    'L� o c�digo e a descri��o de todas as Tabelas de Pre�os
    lErro = CF("Cod_Nomes_Le", "TiposDeBloqueio", "Codigo", "NomeReduzido", STRING_TIPO_BLOQUEIO_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then gError 191514

    For Each objCodDescricao In colCodigoDescricao

        'Adiciona o item na Lista de Tabela de Pre�os
        TipoBloqueio.AddItem CInt(objCodDescricao.iCodigo) & SEPARADOR & objCodDescricao.sNome
        TipoBloqueio.ItemData(TipoBloqueio.NewIndex) = objCodDescricao.iCodigo
        
    Next

    Carrega_TipoBloqueio = SUCESSO

    Exit Function

Erro_Carrega_TipoBloqueio:

    Carrega_TipoBloqueio = gErr

    Select Case gErr

        Case 191514

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191515)

    End Select

    Exit Function

End Function

Function Traz_BloqPSRV_ConsultaTodos_Tela(objPedidoVenda As ClassPedidoDeVenda) As Long
'Coloca os dados do Tab de Bloqueio na tela

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer, objPedidoVendaAux As New ClassPedidoDeVenda
Dim objBloqueioPV As New ClassBloqueioPV, objcliente As New ClassCliente
Dim objTipoBloqueio As New ClassTipoDeBloqueio

On Error GoTo Erro_Traz_BloqPSRV_ConsultaTodos_Tela

    'Realiza a leitura dos BloqueioPV
    objPedidoVendaAux.lCodigo = objPedidoVenda.lCodigo
    objPedidoVendaAux.iFilialEmpresa = objPedidoVenda.iFilialEmpresa
    
    lErro = CF("BloqueiosPSRV_Le", objPedidoVendaAux)
    If lErro <> SUCESSO Then gError 191521

    'Preenche a tela com as informa��es do BD
    Pedido.Caption = objPedidoVenda.lCodigo
    
    objcliente.lCodigo = objPedidoVenda.lCliente
    lErro = CF("Cliente_Le", objcliente)
    If lErro <> SUCESSO And lErro <> 12293 Then gError 191522
    
    Cliente.Caption = objcliente.sNomeReduzido
    
    Data.Caption = Format(objPedidoVenda.dtDataEmissao, "dd/mm/yy")
    Valor.Caption = Format(objPedidoVenda.dValorTotal, "Standard")

    'Limpa o Grid de Bloqueios antes de preencher com os dados da cole��o
    Call Grid_Limpa(objGridBloqueio)

    iIndice = 0

    For Each objBloqueioPV In objPedidoVendaAux.colBloqueiosPV

        iIndice = iIndice + 1

        objTipoBloqueio.iCodigo = objBloqueioPV.iTipoDeBloqueio

        'L� o Tipo de bloqueio
        lErro = CF("TipoDeBloqueio_Le", objTipoBloqueio)
        If lErro <> SUCESSO And lErro <> 23666 Then gError 191523
        
        If lErro = 23666 Then gError 191524
        
        'Coloca o bloqueio no Grid de bloqueios
        GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col) = objTipoBloqueio.iCodigo & SEPARADOR & objTipoBloqueio.sNomeReduzido
        If objBloqueioPV.dtData <> DATA_NULA Then GridBloqueio.TextMatrix(iIndice, iGrid_DataBloqueio_Col) = Format(objBloqueioPV.dtData, "dd/mm/yy")
        If objBloqueioPV.dtDataLib <> DATA_NULA Then GridBloqueio.TextMatrix(iIndice, iGrid_DataLiberacao_Col) = Format(objBloqueioPV.dtDataLib, "dd/mm/yy")
        GridBloqueio.TextMatrix(iIndice, iGrid_Observacao_Col) = objBloqueioPV.sObservacao
        GridBloqueio.TextMatrix(iIndice, iGrid_SeqBloqueio_Col) = CStr(objBloqueioPV.iSequencial)

    Next
    
    objGridBloqueio.iLinhasExistentes = iIndice

    Traz_BloqPSRV_ConsultaTodos_Tela = SUCESSO

    Exit Function

Erro_Traz_BloqPSRV_ConsultaTodos_Tela:

    Traz_BloqPSRV_ConsultaTodos_Tela = gErr

    Select Case gErr

        Case 191521 To 191523
        
        Case 191524
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPODEBLOQUEIO_NAO_CADASTRADO", gErr, objTipoBloqueio.iCodigo)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191525)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long
'Realiza a Grava��o da informa��es no banco de dados

Dim lErro As Long
Dim iIndice As Integer
Dim objBloqueioPV As ClassBloqueioPV
Dim objPedidoDeVenda As New ClassPedidoDeVenda

On Error GoTo Erro_Gravar_Registro

    'Preenche o objPedidoDeVenda
    objPedidoDeVenda.lCodigo = Pedido.Caption
    objPedidoDeVenda.iFilialEmpresa = giFilialEmpresa

    For iIndice = 1 To objGridBloqueio.iLinhasExistentes

        Set objBloqueioPV = New ClassBloqueioPV
        
        objBloqueioPV.sObservacao = GridBloqueio.TextMatrix(iIndice, iGrid_Observacao_Col)
        objBloqueioPV.iSequencial = StrParaInt(GridBloqueio.TextMatrix(iIndice, iGrid_SeqBloqueio_Col))

        'Adiciona o bloqueio na cole��o de bloqueios
        With objBloqueioPV
            Call objPedidoDeVenda.colBloqueiosPV.Add(objPedidoDeVenda.iFilialEmpresa, objPedidoDeVenda.lCodigo, objBloqueioPV.iSequencial, .iTipoDeBloqueio, .sCodUsuario, .sResponsavel, .dtData, "", "", DATA_NULA, .sObservacao)
        End With

    Next
    
    lErro = CF("PedidoServico_AtualizaObsBloq", objPedidoDeVenda)
    If lErro <> SUCESSO Then gError 191530
    
    iAlterado = 0 'para evitar que pergunte se deseja salvar
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr
    
    Select Case gErr

        Case 191530
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191531)

    End Select

    Exit Function

End Function

'*** FUN��ES DE APOIO A TELA - FIM ***

'*** EVENTOS DO GRIDBLOQUEIO - IN�CIO ***
Private Sub GridBloqueio_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridBloqueio, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridBloqueio, iAlterado)
    End If

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
'*** EVENTOS DO GRIDBLOQUEIO - FIM ***

'*** EVENTOS DOS CONTROLES DO GRID - IN�CIO ***
Private Sub TipoBloqueio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub TipoBloqueio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBloqueio)

End Sub

Private Sub TipoBloqueio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)

End Sub

Private Sub TipoBloqueio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBloqueio.objControle = TipoBloqueio
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataBloqueio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataBloqueio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBloqueio)

End Sub

Private Sub DataBloqueio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)

End Sub

Private Sub DataBloqueio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBloqueio.objControle = DataBloqueio
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub DataLiberacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub DataLiberacao_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBloqueio)

End Sub

Private Sub DataLiberacao_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)

End Sub

Private Sub DataLiberacao_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBloqueio.objControle = DataLiberacao
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub ObservacaoBloqueio_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub ObservacaoBloqueio_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridBloqueio)

End Sub

Private Sub ObservacaoBloqueio_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridBloqueio)

End Sub

Private Sub ObservacaoBloqueio_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridBloqueio.objControle = ObservacaoBloqueio
    lErro = Grid_Campo_Libera_Foco(objGridBloqueio)
    If lErro <> SUCESSO Then Cancel = True

End Sub
'*** EVENTOS DOS CONTROLES DO GRID - FIM ***

'*** EVENTOS CLICK DOS CONTROLES - IN�CIO ***
Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 191532
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoGravar_Click:
 
    Select Case gErr
    
        Case 191532
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191533)

    End Select

    Exit Sub
    
End Sub
'*** EVENTOS CLICK DOS CONTROLES - FIM ***

Public Sub Form_Unload(Cancel As Integer)
    
    Set objGridBloqueio = Nothing

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_EMPENHO
    Set Form_Load_Ocx = Me
    Caption = "Bloqueios do Pedido de Servi�o"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "BloqPSRV_ConsultaTodos"

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


