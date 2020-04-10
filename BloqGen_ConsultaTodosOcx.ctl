VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl BloqueiosGen_ConsultaTodosOcx 
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9405
   LockControls    =   -1  'True
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
         Picture         =   "BloqGen_ConsultaTodosOcx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Gravar"
         Top             =   60
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   675
         Picture         =   "BloqGen_ConsultaTodosOcx.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   1185
      Left            =   120
      TabIndex        =   7
      Top             =   105
      Width           =   7560
      Begin VB.Label Label1 
         Caption         =   "Código:"
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
         Index           =   0
         Left            =   1035
         TabIndex        =   15
         Top             =   345
         Width           =   630
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   1050
         TabIndex        =   14
         Top             =   795
         Width           =   705
      End
      Begin VB.Label Label1 
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
         Index           =   2
         Left            =   4065
         TabIndex        =   13
         Top             =   330
         Width           =   630
      End
      Begin VB.Label Label1 
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
         Index           =   3
         Left            =   4080
         TabIndex        =   12
         Top             =   780
         Width           =   510
      End
      Begin VB.Label Codigo 
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
         ItemData        =   "BloqGen_ConsultaTodosOcx.ctx":02D8
         Left            =   300
         List            =   "BloqGen_ConsultaTodosOcx.ctx":02DA
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
         Width           =   3825
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
Attribute VB_Name = "BloqueiosGen_ConsultaTodosOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim gobjMapBloqGen As ClassMapeamentoBloqGen

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

'*** FUNÇÕES DE INICIALIZAÇÃO DA TELA - INÍCIO ***
Public Sub Form_Load()
    
Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Form_Load
    
    Set objGridBloqueio = New AdmGrid
    
    'Executa a Inicialização do grid Bloqueio
    lErro = Inicializa_Grid_Bloqueio(objGridBloqueio)
    If lErro <> SUCESSO Then gError 191490
    
'    'Carrega o Combo TipoBloqueio
'    lErro = Carrega_TipoBloqueio
'    If lErro <> SUCESSO Then gError 191491
    
    iAlterado = 0
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 191490

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 191492)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(ByVal objMapBloqGen As ClassMapeamentoBloqGen, ByVal objDocGravado As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjMapBloqGen = objMapBloqGen

    'Carrega o Combo TipoBloqueio
    lErro = Carrega_TipoBloqueio
    If lErro <> SUCESSO Then gError 191491
    
    'Chama a função que irá preencher a tela BloqPV_ConsultaTodos
    lErro = Traz_Bloqueios_Tela(objDocGravado)
    If lErro <> SUCESSO Then gError 191493
        
    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 191491, 191493
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191493)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Private Function Inicializa_Grid_Bloqueio(objGridInt As AdmGrid) As Long
'Executa a Inicialização do grid Bloqueio
    
    'tela em questão
    Set objGridInt.objForm = Me

    'titulos do grid
    objGridInt.colColuna.Add ("  ")
    objGridInt.colColuna.Add ("Tipo")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Data Liberação")
    objGridInt.colColuna.Add ("Observação")
    objGridInt.colColuna.Add ("Seq.")
    
   'campos de edição do grid
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
    
    'Não permite incluir novas linhas
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR
    
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a critica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula
    
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual a Coluna em questão
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191499)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_TipoBloqueio(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Tipo Bloqueio que está deixando de ser a corrente

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

            'Tenta selecioná-lo na combo
            lErro = Combo_Seleciona_Grid(TipoBloqueio, iCodigo)
            If lErro <> SUCESSO And lErro <> 25085 And lErro <> 25086 Then gError 191500

            'Não foi encontrado
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
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_NAO_ENCONTRADO", gErr, iCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 191502
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPOBLOQUEIO_NAO_ENCONTRADO1", gErr, TipoBloqueio.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191504)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataBloqueio(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DataBloqueio que está deixando de ser a corrente

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191507)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_DataLiberacao(objGridInt As AdmGrid) As Long
'Faz a crítica da célula DataLiberacao que está deixando de ser a corrente

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
'Faz a crítica da célula Observação que está deixando de ser a corrente

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 191512)

    End Select

    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iChamada As Integer)

On Error GoTo Erro_Rotina_Grid_Enable

    'Pesquisa o controle da coluna em questão
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
'Carrega a Combo TipoBloqueio com as informações do BD

Dim lErro As Long
Dim colTipoDeBloqueio As New Collection
Dim objTipoDeBloqueio As ClassTiposDeBloqueioGen

On Error GoTo Erro_Carrega_TipoBloqueio

    'Lê o código e a descrição de todas as Tabelas de Preços
    lErro = CF("TiposDeBloqueioGen_Le_TipoTela", gobjMapBloqGen.iTipoTelaBloqueio, colTipoDeBloqueio)
    If lErro <> SUCESSO Then gError 191514

    For Each objTipoDeBloqueio In colTipoDeBloqueio

        'Adiciona o item na Lista de Tabela de Preços
        TipoBloqueio.AddItem CInt(objTipoDeBloqueio.iCodigo) & SEPARADOR & objTipoDeBloqueio.sNomeReduzido
        TipoBloqueio.ItemData(TipoBloqueio.NewIndex) = objTipoDeBloqueio.iCodigo
        
    Next

    Carrega_TipoBloqueio = SUCESSO

    Exit Function

Erro_Carrega_TipoBloqueio:

    Carrega_TipoBloqueio = gErr

    Select Case gErr

        Case 191514

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191515)

    End Select

    Exit Function

End Function

Function Traz_Bloqueios_Tela(objDocBloq As Object) As Long
'Coloca os dados do Tab de Bloqueio na tela

Dim lErro As Long
Dim iIndice As Integer
Dim iIndice2 As Integer
Dim objDocBloqAux As Object
Dim objBloqueioGen As New ClassBloqueioGen
Dim objcliente As New ClassCliente
Dim objTipoBloqueio As New ClassTiposDeBloqueioGen
Dim lCodigo As Long
Dim colBloqueios As Collection

On Error GoTo Erro_Traz_Bloqueios_Tela

    Set objDocBloqAux = CreateObject(gobjMapBloqGen.sProjetoClasseDocBloq & "." & gobjMapBloqGen.sNomeClasseDocBloq)

    'Passa os dados do Bloqueio para o Obj
    If gobjMapBloqGen.iClassePossuiFilEmp = MARCADO Then
        objDocBloqAux.iFilialEmpresa = objDocBloq.iFilialEmpresa
    End If
    
    lCodigo = CallByName(objDocBloq, gobjMapBloqGen.sClasseNomeCampoChave, VbGet)
    
    Call CallByName(objDocBloqAux, gobjMapBloqGen.sClasseNomeCampoChave, VbLet, lCodigo)
    
    lErro = CF("BloqueiosGen_Le", gobjMapBloqGen, objDocBloq)
    If lErro <> SUCESSO Then gError 191521

    'Preenche a tela com as informações do BD
    Codigo.Caption = lCodigo
    
    Set colBloqueios = ColecaoDef_Trans_Collection(CallByName(objDocBloq, gobjMapBloqGen.sNomeColecaoBloqDoc, VbGet))

    'Limpa o Grid de Bloqueios antes de preencher com os dados da coleção
    Call Grid_Limpa(objGridBloqueio)

    iIndice = 0

    For Each objBloqueioGen In colBloqueios

        iIndice = iIndice + 1
        
        If iIndice = 1 Then
            
            Cliente.Caption = objBloqueioGen.lClienteDoc & SEPARADOR & objBloqueioGen.sNomeClienteDoc
            Data.Caption = Format(objBloqueioGen.dtDataEmissaoDoc, "dd/mm/yy")
            Valor.Caption = Format(objBloqueioGen.dValorDoc, "Standard")
        
        End If
        
        'Coloca o bloqueio no Grid de bloqueios
        GridBloqueio.TextMatrix(iIndice, iGrid_TipoBloqueio_Col) = objBloqueioGen.iTipoDeBloqueio & SEPARADOR & objBloqueioGen.sNomeTipoDeBloqueio
        If objBloqueioGen.dtData <> DATA_NULA Then GridBloqueio.TextMatrix(iIndice, iGrid_DataBloqueio_Col) = Format(objBloqueioGen.dtData, "dd/mm/yy")
        If objBloqueioGen.dtDataLib <> DATA_NULA Then GridBloqueio.TextMatrix(iIndice, iGrid_DataLiberacao_Col) = Format(objBloqueioGen.dtDataLib, "dd/mm/yy")
        GridBloqueio.TextMatrix(iIndice, iGrid_Observacao_Col) = objBloqueioGen.sObservacao
        GridBloqueio.TextMatrix(iIndice, iGrid_SeqBloqueio_Col) = CStr(objBloqueioGen.iSequencial)

    Next
    
    objGridBloqueio.iLinhasExistentes = iIndice

    Traz_Bloqueios_Tela = SUCESSO

    Exit Function

Erro_Traz_Bloqueios_Tela:

    Traz_Bloqueios_Tela = gErr

    Select Case gErr

        Case 191521
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191525)

    End Select

    Exit Function

End Function

Public Function Gravar_Registro() As Long
'Realiza a Gravação da informações no banco de dados

Dim lErro As Long
Dim iIndice As Integer
Dim objBloqueioGen As ClassBloqueioGen
Dim colBloqueios As New Collection

On Error GoTo Erro_Gravar_Registro

    For iIndice = 1 To objGridBloqueio.iLinhasExistentes

        Set objBloqueioGen = New ClassBloqueioGen
        
        objBloqueioGen.iFilialEmpresa = giFilialEmpresa
        objBloqueioGen.lCodigo = StrParaLong(Codigo.Caption)
        objBloqueioGen.iTipoTelaBloqueio = gobjMapBloqGen.iTipoTelaBloqueio
        
        objBloqueioGen.sObservacao = GridBloqueio.TextMatrix(iIndice, iGrid_Observacao_Col)
        objBloqueioGen.iSequencial = StrParaInt(GridBloqueio.TextMatrix(iIndice, iGrid_SeqBloqueio_Col))

        'Adiciona o bloqueio na coleção de bloqueios
        colBloqueios.Add objBloqueioGen

    Next
    
    lErro = CF("BloqueiosGen_AtualizaObsBloq", gobjMapBloqGen, colBloqueios)
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

'*** FUNÇÕES DE APOIO A TELA - FIM ***

'*** EVENTOS DO GRIDBLOQUEIO - INÍCIO ***
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


'*** EVENTOS DOS CONTROLES DO GRID - INÍCIO ***
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

'*** EVENTOS CLICK DOS CONTROLES - INÍCIO ***
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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 191533)

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
    Caption = "Bloqueios"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "BloqueiosGen_ConsultaTodos"

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


