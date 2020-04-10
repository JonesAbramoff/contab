VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpFlCxTotal1Ocx 
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
   ScaleHeight     =   2190
   ScaleWidth      =   5985
   Begin VB.ComboBox Identificacao 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   255
      Width           =   2790
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
      Left            =   4665
      Picture         =   "RelOpFlCxTotal1Ocx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1140
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   4665
      ScaleHeight     =   495
      ScaleWidth      =   1080
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   1140
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   585
         Picture         =   "RelOpFlCxTotal1Ocx.ctx":0102
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   75
         Picture         =   "RelOpFlCxTotal1Ocx.ctx":0280
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
   End
   Begin MSComCtl2.UpDown UpDownEmissaoAte 
      Height          =   315
      Left            =   4140
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1665
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox EmissaoAte 
      Height          =   315
      Left            =   2970
      TabIndex        =   13
      Top             =   1665
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownEmissaoDe 
      Height          =   315
      Left            =   1935
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1665
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox EmissaoDe 
      Height          =   315
      Left            =   765
      TabIndex        =   15
      Top             =   1665
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   390
      TabIndex        =   17
      Top             =   1725
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2550
      TabIndex        =   16
      Top             =   1725
      Width           =   360
   End
   Begin VB.Label DataInicial 
      Height          =   240
      Left            =   1440
      TabIndex        =   11
      Top             =   1215
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Data Base:"
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
      Left            =   390
      TabIndex        =   10
      Top             =   1230
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Identificação:"
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
      Left            =   165
      TabIndex        =   9
      Top             =   285
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
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
      Left            =   420
      TabIndex        =   8
      Top             =   795
      Width           =   930
   End
   Begin VB.Label DataFinal 
      Height          =   240
      Left            =   3480
      TabIndex        =   7
      Top             =   1215
      Width           =   930
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Data Final:"
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
      Left            =   2430
      TabIndex        =   6
      Top             =   1230
      Width           =   945
   End
   Begin VB.Label Descricao 
      Height          =   240
      Left            =   1515
      TabIndex        =   5
      Top             =   795
      Width           =   2820
   End
End
Attribute VB_Name = "RelOpFlCxTotal1Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelatorio As AdmRelatorio
Dim gobjRelOpcoes As AdmRelOpcoes

Private Sub Identificacao_Click()

Dim lErro As Long
Dim iIndice As Integer, sFluxo As String

On Error GoTo Erro_Identificacao_Click

    If Identificacao.ListIndex = -1 Then Exit Sub

    'Pega o nome do fluxo atual
    sFluxo = Identificacao.Text

    'Exibe na tela os dados do fluxo
    lErro = Traz_Fluxo_Tela(sFluxo)
    If lErro <> SUCESSO Then gError 195483

    Exit Sub

Erro_Identificacao_Click:

    Select Case gErr

        Case 195483

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195484)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados fornecidos pelo usuário

Dim lErro As Long
Dim objFluxo As New ClassFluxo
Dim lNumIntRel As Long
Dim iIndice As Integer
Dim lFluxo As Long
Dim dAcumAnt As Double
Dim lComando As Long

On Error GoTo Erro_PreencherRelOp

    'Faz Critica se data inicial é maior que data Final
    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then gError 195485

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 195486
    
    For iIndice = 0 To Identificacao.ListCount - 1
    
        If Identificacao.List(iIndice) = Identificacao.Text Then
            lFluxo = Identificacao.ItemData(iIndice)
            Exit For
            
        End If
    Next
    
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 195510
    
    lErro = Comando_Executar(lComando, "SELECT AcumuladoAnterior FROM FluxoSintetico WHERE FluxoID = ? AND Data >= ? ORDER BY Data", _
                            dAcumAnt, lFluxo, StrParaDate(EmissaoDe.Text))
    If lErro <> AD_SQL_SUCESSO Then gError 195508
    
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 195509
    
    lErro = objRelOpcoes.IncluirParametro("NFLUXO", CStr(lFluxo))
    If lErro <> AD_BOOL_TRUE Then gError 195487

    lErro = objRelOpcoes.IncluirParametro("DINIC", EmissaoDe.Text)
    If lErro <> AD_BOOL_TRUE Then gError 195488

    lErro = objRelOpcoes.IncluirParametro("DFIM", EmissaoAte.Text)
    If lErro <> AD_BOOL_TRUE Then gError 195489

    lErro = objRelOpcoes.IncluirParametro("NSLDINI", Format(dAcumAnt, "standard"))
    If lErro <> AD_BOOL_TRUE Then gError 195510

    Call Comando_Fechar(lComando)

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 195485 To 195489
        
        Case 195508, 195509
            Call Rotina_Erro(vbOKOnly, "ERRO_LEITURA_RELFLUXOCXTOT1", gErr)

        Case 195510
            Call Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195490)

    End Select

    Call Comando_Fechar(lComando)

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing

End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 195491

    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 195491
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195492)

    End Select

    Exit Function

End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 195493

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 195493

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195494)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

    Call Limpa_Tela(Me)

    Identificacao.ListIndex = -1
    Descricao.Caption = ""
    DataInicial.Caption = ""
    DataFinal.Caption = ""

End Sub

Private Sub EmissaoDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoDe_Validate

    'Verifica se a data foi preenchida
    If Len(Trim(EmissaoDe.ClipText)) = 0 Then Exit Sub

    'Verifica se é uma data válida
    lErro = Data_Critica(EmissaoDe.Text)
    If lErro <> SUCESSO Then gError 195495

    Exit Sub

Erro_EmissaoDe_Validate:

    Cancel = True

    Select Case gErr

        Case 195495

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195496)

    End Select

    Exit Sub

End Sub

Private Sub EmissaoAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_EmissaoAte_Validate

    'Verifica se a data foi preenchida
    If Len(Trim(EmissaoAte.ClipText)) = 0 Then Exit Sub

    'Verifica se é uma data válida
    lErro = Data_Critica(EmissaoAte.Text)
    If lErro <> SUCESSO Then gError 195497

    Exit Sub

Erro_EmissaoAte_Validate:

    Cancel = True

    Select Case gErr

        Case 195497

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195498)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long
Dim colFluxo As New Collection
Dim objFluxo As ClassFluxo

On Error GoTo Erro_Form_Load

    lErro = CF("Fluxo_Le_Todos", colFluxo)
    If lErro <> SUCESSO Then gError 195499

    For Each objFluxo In colFluxo

        Identificacao.AddItem objFluxo.sFluxo
        Identificacao.ItemData(Identificacao.NewIndex) = objFluxo.lFluxoId

    Next

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 195499

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195500)

    End Select
    
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_POSCLI
    Set Form_Load_Ocx = Me
    Caption = "Fluxo de Caixa Analítico Total 1"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "RelOpFlCxAnTotal1"

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

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Function Formata_E_Critica_Parametros() As Long
'Verifica se os parâmetros iniciais são maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    'Verificar se o cliente foi preenchido
    If Len(Trim(Identificacao.Text)) = 0 Then gError 195501

    If Len(Trim(EmissaoDe.ClipText)) = 0 Then gError 195502

    If Len(Trim(EmissaoAte.ClipText)) = 0 Then gError 195503

    'data inicial nao pode ser menor que a data base
    If StrParaDate(EmissaoDe.Text) < StrParaDate(DataInicial.Caption) Then gError 195504

    'data final não pode ser maior que a data final
    If StrParaDate(EmissaoDe.Text) > CDate(EmissaoAte.Text) Then gError 195505

    'data final não pode ser maior que a data final do fluxo
    If StrParaDate(EmissaoAte.Text) > StrParaDate(DataFinal.Caption) Then gError 195506

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 195501
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_FLUXO_VAZIO", gErr)

        Case 195502
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAO_PREENCHIDA", gErr)

        Case 195503
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAO_PREENCHIDA", gErr)

        Case 195436
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MENOR_DATAATUAL", gErr)

        Case 195504
            Call Rotina_Erro(vbOKOnly, "ERRO_DATADE_MENOR_DATABASE", gErr)

        Case 195505
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_EMISSAO_INICIAL_MAIOR", gErr)

        Case 195506
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAATE_MAIOR_DATAFINALFLUXO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 195507)

    End Select

    Exit Function

End Function

Private Function Traz_Fluxo_Tela(sFluxo As String) As Long
'Coloca na Tela os dados do Fluxo passado como parametro

Dim lErro As Long
Dim objFluxo As New ClassFluxo

On Error GoTo Erro_Traz_Fluxo_Tela

    objFluxo.sFluxo = sFluxo
    objFluxo.iFilialEmpresa = giFilialEmpresa

    'Le o fluxo passado como parametro
    lErro = CF("Fluxo_Le", objFluxo)
    If lErro <> SUCESSO And lErro <> 20104 Then gError 195508

    If lErro = 20104 Then gError 195509

    'passa os dados para a Tela
    Descricao.Caption = objFluxo.sDescricao
    DataInicial.Caption = Format(objFluxo.dtDataBase, "dd/mm/yyyy")
    DataFinal.Caption = Format(objFluxo.dtDataFinal, "dd/mm/yyyy")
    
    EmissaoDe.Text = Format(objFluxo.dtDataBase, "dd/mm/yy")
    EmissaoAte.Text = Format(objFluxo.dtDataFinal, "dd/mm/yy")

    Traz_Fluxo_Tela = SUCESSO

    Exit Function

Erro_Traz_Fluxo_Tela:

    Traz_Fluxo_Tela = gErr

    Select Case gErr

        Case 195509

        Case 195510
            Call Rotina_Erro(vbOKOnly, "ERRO_FLUXO_NAO_CADASTRADO", gErr, objFluxo.sFluxo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 195511)

    End Select

    Exit Function

End Function


Private Sub UpDownEmissaoDe_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_DownClick

    lErro = Data_Up_Down_Click(EmissaoDe, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 195457
    Exit Sub

Erro_UpDownEmissaoDe_DownClick:

    Select Case gErr

        Case 195457

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195458)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoDe_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoDe_UpClick

    lErro = Data_Up_Down_Click(EmissaoDe, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 195459

    Exit Sub

Erro_UpDownEmissaoDe_UpClick:

    Select Case gErr

        Case 195459

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195460)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_DownClick

    lErro = Data_Up_Down_Click(EmissaoAte, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 195453
    Exit Sub

Erro_UpDownEmissaoAte_DownClick:

    Select Case gErr

        Case 195453

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195454)

    End Select

    Exit Sub

End Sub

Private Sub UpDownEmissaoAte_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDownEmissaoAte_UpClick

    lErro = Data_Up_Down_Click(EmissaoAte, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 195455

    Exit Sub

Erro_UpDownEmissaoAte_UpClick:

    Select Case gErr

        Case 195455
            EmissaoAte.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 195456)

    End Select

    Exit Sub

End Sub

