VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl RelOpAcomInadTRV 
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7965
   LockControls    =   -1  'True
   ScaleHeight     =   5325
   ScaleWidth      =   7965
   Begin VB.Frame Frame2 
      Caption         =   "Datas de referência"
      Height          =   3120
      Left            =   180
      TabIndex        =   24
      Top             =   1725
      Width           =   2970
      Begin MSMask.MaskEdBox DataRef 
         Height          =   315
         Left            =   1020
         TabIndex        =   26
         Top             =   285
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridData 
         Height          =   2670
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   4710
         _Version        =   393216
      End
   End
   Begin VB.CheckBox OptMes 
      Caption         =   "Detalhar mês a mês"
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
      Left            =   4290
      TabIndex        =   23
      Top             =   4965
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Documento"
      Height          =   1275
      Left            =   3270
      TabIndex        =   19
      Top             =   3570
      Width           =   4530
      Begin VB.ComboBox TipoDocSeleciona 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "RelOpAcomInadTRV.ctx":0000
         Left            =   1155
         List            =   "RelOpAcomInadTRV.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   780
         Width           =   3255
      End
      Begin VB.OptionButton TipoDocTodos 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   75
         TabIndex        =   21
         Top             =   315
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton TipoDocApenas 
         Caption         =   "Apenas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   20
         Top             =   795
         Width           =   1050
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Filiais Empresa"
      Height          =   1785
      Left            =   3270
      TabIndex        =   15
      Top             =   1725
      Width           =   4545
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         Height          =   540
         Left            =   3000
         Picture         =   "RelOpAcomInadTRV.ctx":0004
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   885
         Width           =   1425
      End
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         Height          =   540
         Left            =   3000
         Picture         =   "RelOpAcomInadTRV.ctx":11E6
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   210
         Width           =   1425
      End
      Begin VB.ListBox FilialEmpresa 
         Height          =   1410
         ItemData        =   "RelOpAcomInadTRV.ctx":2200
         Left            =   120
         List            =   "RelOpAcomInadTRV.ctx":2216
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Vencimento"
      Height          =   720
      Left            =   165
      TabIndex        =   7
      Top             =   750
      Width           =   5535
      Begin MSComCtl2.UpDown UpDownEmissaoDe 
         Height          =   315
         Left            =   2385
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox EmissaoDe 
         Height          =   285
         Left            =   1230
         TabIndex        =   9
         Top             =   270
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownEmissaoAte 
         Height          =   315
         Left            =   4485
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox EmissaoAte 
         Height          =   285
         Left            =   3330
         TabIndex        =   11
         Top             =   270
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
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
         Left            =   870
         TabIndex        =   12
         Top             =   315
         Width           =   315
      End
      Begin VB.Label Label3 
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
         Left            =   2940
         TabIndex        =   13
         Top             =   315
         Width           =   360
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpAcomInadTRV.ctx":22B3
      Left            =   1380
      List            =   "RelOpAcomInadTRV.ctx":22B5
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   255
      Width           =   2730
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6015
      Picture         =   "RelOpAcomInadTRV.ctx":22B7
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   825
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5670
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpAcomInadTRV.ctx":23B9
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpAcomInadTRV.ctx":2537
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpAcomInadTRV.ctx":2A69
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpAcomInadTRV.ctx":2BF3
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Opção:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   660
      TabIndex        =   14
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpAcomInadTRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim objGridData As AdmGrid
Dim iGrid_Data_Col As Integer

Dim iAlterado As Integer

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 190462
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 190463
        
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 190462
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
            
        Case 190463
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190464)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    'Limpa os Campos
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError 190465
    
    ComboOpcoes.Text = ""
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 190466
    
    Call Grid_Limpa(objGridData)
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case 190465, 190466
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190467)

    End Select

    Exit Sub
   
End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Define os Campos
    lErro = Define_Padrao()
    If lErro <> SUCESSO Then gError 190468
    
    lErro = Carrega_TipoDocumento(TipoDocSeleciona)
    If lErro <> SUCESSO Then gError 190469
    
    lErro = Carrega_FilialEmpresa
    If lErro <> SUCESSO Then gError 190470
    
    Set objGridData = New AdmGrid
    
    lErro = Inicializa_GridData(objGridData)
    If lErro <> SUCESSO Then gError 190471
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case 190468 To 190471
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190472)

    End Select

    Exit Sub

End Sub

Private Sub EmissaoAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(EmissaoAte)

End Sub

Private Sub EmissaoDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(EmissaoDe)

End Sub

Private Sub BotaoGravar_Click()
'Grava a opção de relatório com os parâmetros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 190473

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 190474

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 190475
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError 190476
    
    Call BotaoLimpar_Click
               
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 190473
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 190474 To 190476
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190477)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 190478

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 190479

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        Call BotaoLimpar_Click
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 190478
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 190479

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190480)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError 190481

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 190481

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190482)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutar As Boolean = False) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim aiFilial() As Integer
Dim adtDataRef() As Date
Dim sTipoDoc As String
Dim iIndice As Integer
Dim iNumFiliais As Integer
Dim iNumDatas As Integer
Dim iDetalharMes As Integer
Dim lNumIntRel As Long
Dim objRelAcomInadTRV As New ClassRelAcomInadTRV

On Error GoTo Erro_PreencherRelOp

    If FilialEmpresa.ListCount >= 6 Then
        iNumFiliais = FilialEmpresa.ListCount
    Else
        iNumFiliais = 6
    End If
    ReDim aiFilial(1 To iNumFiliais)
    iNumDatas = objGridData.iLinhasExistentes
    If iNumDatas <> 0 Then
        ReDim adtDataRef(1 To iNumDatas)
    End If
            
    'Faz a Critica se o Inicial é Maior que o Final, se tudo está preenchido correto
    lErro = Formata_E_Critica_Parametros(sTipoDoc, aiFilial, iDetalharMes, adtDataRef)
    If lErro <> SUCESSO Then gError 190483

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 190484

    lErro = objRelOpcoes.IncluirParametro("DDATAINI", CStr(StrParaDate(EmissaoDe.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 190485

    lErro = objRelOpcoes.IncluirParametro("DDATAFIM", CStr(StrParaDate(EmissaoAte.Text)))
    If lErro <> AD_BOOL_TRUE Then gError 190486
    
    lErro = objRelOpcoes.IncluirParametro("TTIPODOC", sTipoDoc)
    If lErro <> AD_BOOL_TRUE Then gError 190487
    
    For iIndice = 1 To UBound(aiFilial)
    
        lErro = objRelOpcoes.IncluirParametro("NFILIAL" & CStr(iIndice), CStr(aiFilial(iIndice)))
        If lErro <> AD_BOOL_TRUE Then gError 190488
    
    Next

    lErro = objRelOpcoes.IncluirParametro("NNUMFILIAIS", CStr(iNumFiliais))
    If lErro <> AD_BOOL_TRUE Then gError 190489
    
    For iIndice = 1 To iNumDatas
    
        lErro = objRelOpcoes.IncluirParametro("DDATAREF" & CStr(iIndice), CStr(adtDataRef(iIndice)))
        If lErro <> AD_BOOL_TRUE Then gError 190490
    
    Next

    lErro = objRelOpcoes.IncluirParametro("NNUMDATAS", CStr(iNumDatas))
    If lErro <> AD_BOOL_TRUE Then gError 190491
    
    lErro = objRelOpcoes.IncluirParametro("NDETALHARMES", CStr(iDetalharMes))
    If lErro <> AD_BOOL_TRUE Then gError 190492
    
    'Faz a selecao
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sTipoDoc)
    If lErro <> SUCESSO Then gError 190493
    
    If bExecutar Then
    
        objRelAcomInadTRV.iNumDatas = iNumDatas
        objRelAcomInadTRV.iNumFiliais = iNumFiliais
        objRelAcomInadTRV.sTipoDoc = sTipoDoc
        objRelAcomInadTRV.dtDataVencDe = StrParaDate(EmissaoDe.Text)
        objRelAcomInadTRV.dtDataVencAte = StrParaDate(EmissaoAte.Text)
        
        For iIndice = 1 To iNumDatas
            objRelAcomInadTRV.adtDatasRef(iIndice) = adtDataRef(iIndice)
        Next
        
        For iIndice = 1 To iNumFiliais
            objRelAcomInadTRV.aiFiliais(iIndice) = aiFilial(iIndice)
        Next
        
        lErro = CF("RelOpAcomInadTRV_Prepara", objRelAcomInadTRV, lNumIntRel)
        If lErro <> SUCESSO Then gError 190494
        
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError 190495
        
    End If

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 190483 To 190495
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190496)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(sTipoDoc As String, aiFilial() As Integer, iDetalharMes As Integer, adtDataRef() As Date) As Long
'Verifica se os parâmetros iniciais são maiores que os finais
'E critica o Tipocliente e Cobrador

Dim lErro As Long
Dim iIndice As Integer
Dim iIndiceAux As Integer

On Error GoTo Erro_Formata_E_Critica_Parametros
                
    'data inicial não pode ser maior que a data final
    If Trim(EmissaoDe.ClipText) <> "" And Trim(EmissaoAte.ClipText) <> "" Then
    
         If StrParaDate(EmissaoDe.Text) > StrParaDate(EmissaoAte.Text) Then gError 190497
    
    End If
    
    iIndiceAux = 0
    For iIndice = 0 To FilialEmpresa.ListCount - 1
        If FilialEmpresa.Selected(iIndice) Then
            iIndiceAux = iIndiceAux + 1
            aiFilial(iIndiceAux) = Codigo_Extrai(FilialEmpresa.List(iIndice))
        End If
    Next
    For iIndice = FilialEmpresa.ListCount + 1 To 6
        aiFilial(iIndice) = 0
    Next
    
    If TipoDocApenas.Value = True Then
        sTipoDoc = SCodigo_Extrai(TipoDocSeleciona.Text)
    Else
        sTipoDoc = ""
    End If
    
    If OptMes.Value = vbChecked Then
        iDetalharMes = MARCADO
    Else
        iDetalharMes = DESMARCADO
    End If
    
    For iIndice = 1 To objGridData.iLinhasExistentes
        adtDataRef(iIndice) = StrParaDate(GridData.TextMatrix(iIndice, iGrid_Data_Col))
    Next
    
    For iIndice = 1 To objGridData.iLinhasExistentes
        For iIndiceAux = 1 To objGridData.iLinhasExistentes
            If iIndice <> iIndiceAux Then
                If adtDataRef(iIndice) = adtDataRef(iIndiceAux) Then gError 190548
            End If
        Next
    Next
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
                
        Case 190497
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIO_MAIOR_DATAFIM", gErr)
            EmissaoDe.SetFocus
            
        Case 190548
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_REPETIDA_GRID", gErr, iIndice, iIndiceAux)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190498)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, ByVal sTipoDoc As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao
              
'    If Trim(EmissaoDe.ClipText) <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Emissao >= " & Forprint_ConvData(CDate(EmissaoDe.Text))
'
'    End If
'
'    If Trim(EmissaoAte.ClipText) <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "Emissao <= " & Forprint_ConvData(CDate(EmissaoAte.Text))
'
'    End If
'
'    If Trim(sTipoDoc) <> "" Then
'
'        If sExpressao <> "" Then sExpressao = sExpressao & " E "
'        sExpressao = sExpressao & "TipoDoc <= " & Forprint_ConvTexto(sTipoDoc)
'
'    End If

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190499)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros armazenados no bd e exibe na tela

Dim lErro As Long
Dim sParam As String
Dim sTipoCliente As String
Dim iIndice As Integer
Dim iIndiceAux As Integer
Dim iNumFiliais As Integer
Dim iNumDatas As Integer

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError 190500
   
    'pega data inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAINI", sParam)
    If lErro <> SUCESSO Then gError 190501

    Call DateParaMasked(EmissaoDe, StrParaDate(sParam))
    
    'pega data final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAFIM", sParam)
    If lErro <> SUCESSO Then gError 190502

    Call DateParaMasked(EmissaoAte, StrParaDate(sParam))
    
    'pega o tipo de documento
    lErro = objRelOpcoes.ObterParametro("TTIPODOC", sParam)
    If lErro <> SUCESSO Then gError 190503

    If Len(Trim(sParam)) > 0 Then
        TipoDocTodos.Value = False
        TipoDocApenas.Value = True
        For iIndice = 0 To TipoDocSeleciona.ListCount - 1
            If SCodigo_Extrai(TipoDocSeleciona.List(iIndice)) = sParam Then
                TipoDocSeleciona.ListIndex = iIndice
                Exit For
            End If
        Next
    Else
        TipoDocTodos.Value = True
        TipoDocApenas.Value = False
        TipoDocSeleciona.ListIndex = -1
    End If
    
    'pega o número de filiais
    lErro = objRelOpcoes.ObterParametro("NDETALHARMES", sParam)
    If lErro <> SUCESSO Then gError 190504

    If StrParaInt(sParam) = MARCADO Then
        OptMes.Value = vbChecked
    Else
        OptMes.Value = vbUnchecked
    End If
    
    'pega o número de filiais
    lErro = objRelOpcoes.ObterParametro("NNUMFILIAIS", sParam)
    If lErro <> SUCESSO Then gError 190505

    iNumFiliais = StrParaInt(sParam)
    
    For iIndiceAux = 0 To FilialEmpresa.ListCount - 1
        FilialEmpresa.Selected(iIndiceAux) = False
    Next
    
    For iIndice = 1 To iNumFiliais
    
        'pega as filiais que foram marcadas
        lErro = objRelOpcoes.ObterParametro("NFILIAL" & CStr(iIndice), sParam)
        If lErro <> SUCESSO Then gError 190506
    
        For iIndiceAux = 0 To FilialEmpresa.ListCount - 1
            If Codigo_Extrai(FilialEmpresa.List(iIndiceAux)) = StrParaInt(sParam) Then
                FilialEmpresa.Selected(iIndiceAux) = True
            End If
        Next
        
    Next
    
    'pega o número de datas de referência
    lErro = objRelOpcoes.ObterParametro("NNUMDATAS", sParam)
    If lErro <> SUCESSO Then gError 190507

    iNumDatas = StrParaInt(sParam)
    
    Call Grid_Limpa(objGridData)
    
    For iIndice = 1 To iNumDatas
    
        'pega as filiais que foram marcadas
        lErro = objRelOpcoes.ObterParametro("DDATAREF" & CStr(iIndice), sParam)
        If lErro <> SUCESSO Then gError 190508
    
        GridData.TextMatrix(iIndice, iGrid_Data_Col) = Format(StrParaDate(sParam), "DD/MM/YYYY")
        
    Next

    objGridData.iLinhasExistentes = iNumDatas
        
    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 189612 To 190508
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190509)

    End Select

    Exit Function

End Function

Function Define_Padrao() As Long

Dim lErro As Long

On Error GoTo Erro_Define_Padrao
    
    Call MarcaDesmarca(True)
    
    OptMes.Value = vbUnchecked
    
    Define_Padrao = SUCESSO
    
    Exit Function
    
Erro_Define_Padrao:

    Define_Padrao = gErr
    
    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190510)
    
    End Select
    
    Exit Function
    
End Function

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub EmissaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoAte_Validate

    If Len(EmissaoAte.ClipText) > 0 Then
        
        lErro = Data_Critica(EmissaoAte.Text)
        If lErro <> SUCESSO Then gError 190511

    End If

    Exit Sub

Erro_EmissaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 190511

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190512)

    End Select

    Exit Sub

End Sub

Private Sub EmissaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoDe_Validate

    If Len(EmissaoDe.ClipText) > 0 Then

        lErro = Data_Critica(EmissaoDe.Text)
        If lErro <> SUCESSO Then gError 190513

    End If

    Exit Sub

Erro_EmissaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 190513

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190514)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objGridData = Nothing
    
End Sub
    
Private Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    lErro = Data_Up_Down_Click(EmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 190515

    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case gErr

        Case 190515
            EmissaoDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190516)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    lErro = Data_Up_Down_Click(EmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 190517

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case gErr

        Case 190517
            EmissaoDe.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190518)

    End Select

    Exit Sub
    
End Sub

Private Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    lErro = Data_Up_Down_Click(EmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 190519

    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case gErr

        Case 190519
            EmissaoAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190520)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    lErro = Data_Up_Down_Click(EmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 190521

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case gErr

        Case 190521
            EmissaoAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190522)

    End Select

    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_TITREC_L
    Set Form_Load_Ocx = Me
    Caption = "Acompanhamento de Inadimplência"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpAcomInadTRV"
    
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

Public Sub Unload(objme As Object)
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then

    
    End If

End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Public Sub TipoDocApenas_Click()

    'Habilita a combo para a seleção da conta corrente
    TipoDocSeleciona.Enabled = True

End Sub

Public Sub TipoDocTodos_Click()

    'Desabilita a combo para a seleção da conta corrente
    TipoDocSeleciona.Enabled = False

    'Limpa a combo de seleção de conta corrente
    TipoDocSeleciona.ListIndex = COMBO_INDICE

End Sub

Private Function Carrega_TipoDocumento(ByVal objComboBox As ComboBox)
'Carrega os Tipos de Documento

Dim lErro As Long
Dim iIndice As Integer
Dim colTipoDocumento As New Collection
Dim objTipoDocumento As ClassTipoDocumento

On Error GoTo Erro_Carrega_TipoDocumento

    'Le os Tipos de Documentos utilizados em Titulos a Receber
    lErro = CF("TiposDocumento_Le_TituloRec", colTipoDocumento)
    If lErro <> SUCESSO Then gError 190523
    
    'Carrega a combobox com as Siglas  - DescricaoReduzida lidas
    For iIndice = 1 To colTipoDocumento.Count
        Set objTipoDocumento = colTipoDocumento.Item(iIndice)
                    
        objComboBox.AddItem objTipoDocumento.sSigla & SEPARADOR & objTipoDocumento.sDescricaoReduzida
    
    Next

    Carrega_TipoDocumento = SUCESSO

    Exit Function

Erro_Carrega_TipoDocumento:

    Carrega_TipoDocumento = gErr

    Select Case gErr

        Case 190523

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190524)

    End Select

    Exit Function

End Function

Private Function Carrega_FilialEmpresa() As Long

Dim lErro As Long
Dim objCodigoNome As New AdmCodigoNome
Dim colCodigoNome As New AdmColCodigoNome
Dim iIndice As Integer

On Error GoTo Erro_Carrega_FilialEmpresa

    FilialEmpresa.Clear

    'Lê o Código e o NOme de Toda FilialEmpresa do BD
    lErro = CF("Cod_Nomes_Le_FilEmp", colCodigoNome)
    If lErro <> SUCESSO Then gError 190525

    iIndice = 0
    'Carrega a combo de Filial Empresa
    For Each objCodigoNome In colCodigoNome
    
        If objCodigoNome.iCodigo < Abs(giFilialAuxiliar) Then
            FilialEmpresa.AddItem CStr(objCodigoNome.iCodigo) & SEPARADOR & objCodigoNome.sNome
            FilialEmpresa.ItemData(FilialEmpresa.NewIndex) = objCodigoNome.iCodigo
            FilialEmpresa.Selected(iIndice) = True
        
            iIndice = iIndice + 1
        End If
    
    Next

    Carrega_FilialEmpresa = SUCESSO

    Exit Function

Erro_Carrega_FilialEmpresa:

    Carrega_FilialEmpresa = gErr

    Select Case gErr

        Case 190525

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190526)

    End Select

    Exit Function

End Function

Private Sub BotaoMarcarTodos_Click()
    Call MarcaDesmarca(True)
End Sub

Private Sub BotaoDesmarcarTodos_Click()
    Call MarcaDesmarca(False)
End Sub

Private Sub MarcaDesmarca(ByVal bFlag As Boolean)

Dim iIndice As Integer

    For iIndice = 0 To FilialEmpresa.ListCount - 1
    
        FilialEmpresa.Selected(iIndice) = bFlag
        
    Next

End Sub

Private Function Inicializa_GridData(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Data")
    
    'Controles que participam do Grid
    objGrid.colCampo.Add (DataRef.Name)

    'Colunas do Grid
    iGrid_Data_Col = 1

    objGrid.objGrid = GridData

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iLinhasVisiveis = 6

    'Largura da primeira coluna
    GridData.ColWidth(0) = 400

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridData = SUCESSO

End Function

Private Sub GridData_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridData, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridData, iAlterado)
    End If

End Sub

Private Sub GridData_GotFocus()
    Call Grid_Recebe_Foco(objGridData)
End Sub

Private Sub GridData_EnterCell()
    Call Grid_Entrada_Celula(objGridData, iAlterado)
End Sub

Private Sub GridData_LeaveCell()
    Call Saida_Celula(objGridData)
End Sub

Private Sub GridData_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridData, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridData, iAlterado)
    End If

End Sub

Private Sub GridData_RowColChange()
    Call Grid_RowColChange(objGridData)
End Sub

Private Sub GridData_Scroll()
    Call Grid_Scroll(objGridData)
End Sub

Private Sub GridData_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridData)
End Sub

Private Sub GridData_LostFocus()
    Call Grid_Libera_Foco(objGridData)
End Sub

Private Sub DataRef_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataRef_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridData)
End Sub

Private Sub DataRef_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridData)
End Sub

Private Sub DataRef_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridData.objControle = DataRef
    lErro = Grid_Campo_Libera_Foco(objGridData)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Function Saida_Celula_DataRef(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Data Entrega que está deixando de ser a corrente

Dim lErro As Long
Dim dtDataRef As Date
Dim dtDataEmissao As Date

On Error GoTo Erro_Saida_Celula_DataRef

    Set objGridInt.objControle = DataRef

    If Len(Trim(DataRef.ClipText)) > 0 Then
    
        'Critica a Data informada
        lErro = Data_Critica(DataRef.Text)
        If lErro <> SUCESSO Then gError 190527

        If GridData.Row - GridData.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
        
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 190528

    Saida_Celula_DataRef = SUCESSO

    Exit Function

Erro_Saida_Celula_DataRef:

    Saida_Celula_DataRef = gErr

    Select Case gErr

        Case 190527, 190528
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190529)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)
    If lErro = SUCESSO Then
    
        'Verifica qual é o grid
        If objGridInt.objGrid.Name = GridData.Name Then
        
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col

                Case iGrid_Data_Col
                
                    lErro = Saida_Celula_DataRef(objGridInt)
                    If lErro <> SUCESSO Then gError 190530
                    
            End Select
                         
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 190531

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 190530
            'erros tratatos nas rotinas chamadas
        
        Case 190531
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190532)

    End Select

    Exit Function

End Function

