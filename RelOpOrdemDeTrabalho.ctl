VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpOrdensDeTrabalhoOcx 
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   LockControls    =   -1  'True
   ScaleHeight     =   3255
   ScaleWidth      =   8130
   Begin VB.Frame FrameCT 
      Caption         =   "Centros de Trabalho"
      Height          =   1395
      Left            =   90
      TabIndex        =   17
      Top             =   840
      Width           =   7935
      Begin MSMask.MaskEdBox CTInicial 
         Height          =   315
         Left            =   525
         TabIndex        =   1
         Top             =   360
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox CTFinal 
         Height          =   315
         Left            =   525
         TabIndex        =   2
         Top             =   840
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         AllowPrompt     =   -1  'True
         MaxLength       =   20
         PromptChar      =   " "
      End
      Begin VB.Label DescCTFinal 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2085
         TabIndex        =   21
         Top             =   840
         Width           =   5640
      End
      Begin VB.Label DescCTInicial 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2085
         TabIndex        =   20
         Top             =   360
         Width           =   5640
      End
      Begin VB.Label LabelCTDe 
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
         Height          =   255
         Left            =   165
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   19
         Top             =   390
         Width           =   360
      End
      Begin VB.Label LabelCTAte 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   135
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   18
         Top             =   885
         Width           =   435
      End
   End
   Begin VB.Frame FrameData 
      Caption         =   "Data"
      Height          =   810
      Left            =   90
      TabIndex        =   14
      Top             =   2295
      Width           =   7935
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   300
         Left            =   525
         TabIndex        =   3
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataInicial 
         Height          =   300
         Left            =   1680
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   300
         Left            =   3705
         TabIndex        =   5
         Top             =   285
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataFinal 
         Height          =   300
         Left            =   4875
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   285
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label LabelDataDe 
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
         Left            =   180
         TabIndex        =   16
         Top             =   300
         Width           =   315
      End
      Begin VB.Label LabelDataAte 
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
         Left            =   3270
         TabIndex        =   15
         Top             =   300
         Width           =   360
      End
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOpOrdemDeTrabalho.ctx":0000
      Left            =   840
      List            =   "RelOpOrdemDeTrabalho.ctx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   255
      Width           =   2916
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
      Left            =   4080
      Picture         =   "RelOpOrdemDeTrabalho.ctx":0004
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5865
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOpOrdemDeTrabalho.ctx":0106
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "RelOpOrdemDeTrabalho.ctx":0284
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOpOrdemDeTrabalho.ctx":07B6
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOpOrdemDeTrabalho.ctx":0940
         Style           =   1  'Graphical
         TabIndex        =   8
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   135
      TabIndex        =   13
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "RelOpOrdensDeTrabalhoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoCTInic As AdmEvento
Attribute objEventoCTInic.VB_VarHelpID = -1
Private WithEvents objEventoCTFim As AdmEvento
Attribute objEventoCTFim.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio
Dim giFocoInicial As Integer

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    Set objEventoCTInic = Nothing
    Set objEventoCTFim = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 137851
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError 137852
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 137851
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
            
        Case 137852
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170465)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, lCT_I As Long, lCT_F As Long) As Long
'monta a expressão de seleção
'recebe os produtos inicial e final no formato do BD

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    sExpressao = ""

    If lCT_I <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = "CT >= " & Forprint_ConvLong(lCT_I)

    End If

    If lCT_F <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "CT <= " & Forprint_ConvLong(lCT_F)

    End If

    If Len(Trim(DataInicial.ClipText)) <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data >= " & Forprint_ConvData(StrParaDate(DataInicial.Text))

    End If
    
    If Len(Trim(DataFinal.ClipText)) <> 0 Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Data <= " & Forprint_ConvData(StrParaDate(DataFinal.Text))

    End If

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 170466)

    End Select

    Exit Function

End Function

Private Function Formata_E_Critica_Parametros(lCTIni As Long, lCTFim As Long) As Long

Dim objCTInicial As ClassCentrodeTrabalho
Dim objCTFinal As ClassCentrodeTrabalho
Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    Set objCTInicial = New ClassCentrodeTrabalho
    
    objCTInicial.sNomeReduzido = Trim(CTInicial.Text)
    
    'Lê CT Inicial pelo NomeReduzido
    lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCTInicial)
    If lErro <> SUCESSO And lErro <> 134941 Then gError 137853
            
    lCTIni = objCTInicial.lCodigo
            
    Set objCTFinal = New ClassCentrodeTrabalho
    
    objCTFinal.sNomeReduzido = Trim(CTFinal.Text)
    
    'Lê CT Final pelo NomeReduzido
    lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCTFinal)
    If lErro <> SUCESSO And lErro <> 134941 Then gError 137854
            
    lCTFim = objCTFinal.lCodigo
    
    'Valida Centros de Trabalho
    If Len(Trim(CTInicial.Text)) <> 0 And Len(Trim(CTFinal.Text)) <> 0 Then
    
        'codigo do CT inicial não pode ser maior que o final
        If objCTInicial.lCodigo > objCTFinal.lCodigo Then gError 137855
        
    End If
    
    'Valida Datas - data inicial não pode ser maior que a final
    If Len(Trim(DataInicial.ClipText)) <> 0 And Len(Trim(DataFinal.ClipText)) <> 0 Then
        
        If StrParaDate(DataInicial.Text) > StrParaDate(DataFinal.Text) Then gError 137856
    
    End If

    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr

        Case 137853, 137854
        
        Case 137855
            Call Rotina_Erro(vbOKOnly, "ERRO_CT_INICIAL_MAIOR", gErr)
    
        Case 137856
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_INICIAL_MAIOR", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170467)

    End Select

    Exit Function

End Function

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche o arquivo C com os dados fornecidos pelo usuário

Dim lErro As Long
Dim lCTIni As Long
Dim lCTFim As Long

On Error GoTo Erro_PreencherRelOp
           
    lErro = Formata_E_Critica_Parametros(lCTIni, lCTFim)
    If lErro <> SUCESSO Then gError 137857

    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 137858

    lErro = objRelOpcoes.IncluirParametro("NCTINIC", CStr(lCTIni))
    If lErro <> AD_BOOL_TRUE Then gError 137859

    lErro = objRelOpcoes.IncluirParametro("NCTFIM", CStr(lCTFim))
    If lErro <> AD_BOOL_TRUE Then gError 137860

    lErro = objRelOpcoes.IncluirParametro("TCTINIC", CTInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 138190

    lErro = objRelOpcoes.IncluirParametro("TCTFIM", CTFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 138191
    
    If Len(Trim(DataInicial.ClipText)) <> 0 Then
        lErro = objRelOpcoes.IncluirParametro("DDATAINI", DataInicial.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAINI", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 137861
    
    If Len(Trim(DataFinal.ClipText)) <> "" Then
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", DataFinal.Text)
    Else
        lErro = objRelOpcoes.IncluirParametro("DDATAFIM", CStr(DATA_NULA))
    End If
    If lErro <> AD_BOOL_TRUE Then gError 137862

    lErro = Monta_Expressao_Selecao(objRelOpcoes, lCTIni, lCTFim)
    If lErro <> SUCESSO Then gError 137863

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 137857 To 137863, 138190, 138191
            'erro tratado nas rotinas chamadas
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170468)

    End Select

    Exit Function

End Function

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'lê os parâmetros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    Limpar_Tela

    lErro = objRelOpcoes.Carregar
    If lErro Then gError 137863

    'pega CT Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("TCTINIC", sParam)
    If lErro Then gError 137864

    If Len(Trim(sParam)) > 0 Then
 
        CTInicial.Text = sParam
        Call CTInicial_Validate(bSGECancelDummy)
        
    End If

    'pega CT Final e exibe
    lErro = objRelOpcoes.ObterParametro("TCTFIM", sParam)
    If lErro Then gError 137865

    If Len(Trim(sParam)) > 0 Then
 
        CTFinal.Text = sParam
        Call CTFinal_Validate(bSGECancelDummy)
        
    End If

    'pega a Data Inicial e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAINI", sParam)
    If lErro <> SUCESSO Then gError 137866
    Call DateParaMasked(DataInicial, StrParaDate(sParam))
    
    'pega a Data Final e exibe
    lErro = objRelOpcoes.ObterParametro("DDATAFIM", sParam)
    If lErro <> SUCESSO Then gError 137867
    Call DateParaMasked(DataFinal, StrParaDate(sParam))

    PreencherParametrosNaTela = SUCESSO

    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case 137863 To 137867
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170469)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 137868

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError 137869

        'retira nome das opções do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as opções da tela
        Limpar_Tela

    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 137868
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case 137869
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170470)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 137870
    
    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 137870
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170471)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da opção de relatório não pode ser vazia
    If ComboOpcoes.Text = "" Then gError 137871

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro Then gError 137872

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError 137873

    If iResultado = GRAVACAO Then ComboOpcoes.AddItem gobjRelOpcoes.sNome

    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 137871
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case 137872, 137873
            'erro tratado nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170472)

    End Select

    Exit Sub

End Sub

Sub Limpar_Tela()

    Call Limpa_Tela(Me)
    DescCTInicial.Caption = ""
    DescCTFinal.Caption = ""

    ComboOpcoes.SetFocus

End Sub

Private Sub BotaoLimpar_Click()

    ComboOpcoes.Text = ""
    Limpar_Tela

End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_RelOpProdutos_Form_Load
    
    giFocoInicial = 1
    
    Set objEventoCTInic = New AdmEvento
    Set objEventoCTFim = New AdmEvento

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_RelOpProdutos_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 170473)

    End Select
   
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_PRODUTOS
    Set Form_Load_Ocx = Me
    Caption = "Relação de Ordens de Trabalho"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpOrdensDeTrabalho"
    
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

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(Trim(DataFinal.ClipText)) <> 0 Then

        lErro = Data_Critica(DataFinal.Text)
        If lErro <> SUCESSO Then gError 137874

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 137874

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170474)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(Trim(DataInicial.ClipText)) <> 0 Then

        lErro = Data_Critica(DataInicial.Text)
        If lErro <> SUCESSO Then gError 137875

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 137875

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170475)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
    
        If Me.ActiveControl Is CTInicial Then
            Call LabelCTDe_Click
        ElseIf Me.ActiveControl Is CTFinal Then
            Call LabelCTAte_Click
        End If
                
    End If

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Sub Unload(objme As Object)
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

Private Sub CTFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_CTFinal_Validate

    DescCTFinal.Caption = ""

    'Verifica se CTFinal não está preenchido
    If Len(Trim(CTFinal.Text)) <> 0 Then
    
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        'Procura pela empresa toda
        objCentrodeTrabalho.iFilialEmpresa = giFilialEmpresa
        
        'Verifica sua existencia
        lErro = CF("TP_CentrodeTrabalho_Le", CTFinal, objCentrodeTrabalho)
        If lErro <> SUCESSO Then gError 137876
                
        DescCTFinal.Caption = objCentrodeTrabalho.sDescricao
           
    End If
    
    Exit Sub

Erro_CTFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 137876
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170476)

    End Select

    Exit Sub

End Sub

Private Sub CTInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_CTInicial_Validate

    DescCTInicial.Caption = ""

    'Verifica se CTInicial não está preenchido
    If Len(Trim(CTInicial.Text)) <> 0 Then

        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        'Procura pela empresa toda
        objCentrodeTrabalho.iFilialEmpresa = giFilialEmpresa
        
        'Verifica sua existencia
        lErro = CF("TP_CentrodeTrabalho_Le", CTInicial, objCentrodeTrabalho)
        If lErro <> SUCESSO Then gError 137877
                
        DescCTInicial.Caption = objCentrodeTrabalho.sDescricao
       
    End If
    
    Exit Sub

Erro_CTInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 137877
            'erro tratado na rotina chamada
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170477)

    End Select

    Exit Sub

End Sub

Private Sub LabelCTAte_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCTAte

    'Verifica se o CTFinal foi preenchido
    If Len(Trim(CTFinal.Text)) <> 0 Then
            
        objCentrodeTrabalho.sNomeReduzido = Trim(CTFinal.Text)
        
    End If

    Call Chama_Tela("CentrodeTrabalhoLista", colSelecao, objCentrodeTrabalho, objEventoCTFim)

    Exit Sub

Erro_LabelCTAte:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170478)

    End Select

    Exit Sub

End Sub

Private Sub LabelCTDe_Click()

Dim lErro As Long
Dim objCentrodeTrabalho As New ClassCentrodeTrabalho
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCTDe

    'Verifica se o CTInicial foi preenchido
    If Len(Trim(CTInicial.Text)) <> 0 Then
    
        objCentrodeTrabalho.sNomeReduzido = Trim(CTInicial.Text)
        
    End If

    Call Chama_Tela("CentrodeTrabalhoLista", colSelecao, objCentrodeTrabalho, objEventoCTInic)

    Exit Sub

Erro_LabelCTDe:

    Select Case gErr
    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170479)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCTFim_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_objEventoCTFim_evSelecao

    Set objCentrodeTrabalho = obj1

    CTFinal.Text = objCentrodeTrabalho.sNomeReduzido
        
    Call CTFinal_Validate(bSGECancelDummy)
        
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCTFim_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170480)

    End Select

    Exit Sub

End Sub

Private Sub objEventoCTInic_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho

On Error GoTo Erro_objEventoCTInic_evSelecao

    Set objCentrodeTrabalho = obj1

    CTInicial.Text = objCentrodeTrabalho.sNomeReduzido
        
    Call CTInicial_Validate(bSGECancelDummy)
        
    'Fecha comando de setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    Me.Show

    Exit Sub

Erro_objEventoCTInic_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170481)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataFinal_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFinal_DownClick

    DataFinal.SetFocus

    If Len(DataFinal.ClipText) > 0 Then

        sData = DataFinal.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 137878

        DataFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFinal_DownClick:

    Select Case gErr

        Case 137878

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170482)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataFinal_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFinal_UpClick

    DataFinal.SetFocus

    If Len(Trim(DataFinal.ClipText)) > 0 Then

        sData = DataFinal.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 137879

        DataFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFinal_UpClick:

    Select Case gErr

        Case 137879

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170483)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicial_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicial_DownClick

    DataInicial.SetFocus

    If Len(DataInicial.ClipText) > 0 Then

        sData = DataInicial.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 137880

        DataInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicial_DownClick:

    Select Case gErr

        Case 137880

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170484)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataInicial_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataInicial_UpClick

    DataInicial.SetFocus

    If Len(Trim(DataInicial.ClipText)) > 0 Then

        sData = DataInicial.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 137881

        DataInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicial_UpClick:

    Select Case gErr

        Case 137881

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 170485)

    End Select

    Exit Sub

End Sub

