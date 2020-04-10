VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl EscrituracaoFechamentoOcx 
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   ScaleHeight     =   3840
   ScaleWidth      =   4530
   Begin VB.CommandButton BotaoOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1095
      Picture         =   "EscrituracaoFechamento.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3180
      Width           =   975
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancela"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2462
      Picture         =   "EscrituracaoFechamento.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3180
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Caption         =   "Periodo Atual"
      Height          =   1185
      Left            =   240
      TabIndex        =   3
      Top             =   1305
      Width           =   4050
      Begin VB.Label DataImpressao 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1305
         TabIndex        =   10
         Top             =   705
         Width           =   1020
      End
      Begin VB.Label DataFim 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2715
         TabIndex        =   9
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label DataInicio 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   810
         TabIndex        =   8
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Início:"
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
         Left            =   150
         TabIndex        =   6
         Top             =   353
         Width           =   570
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Fim:"
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
         Left            =   2265
         TabIndex        =   5
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Impresso em:"
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
         Left            =   150
         TabIndex        =   4
         Top             =   765
         Width           =   1125
      End
   End
   Begin VB.ComboBox Livro 
      Height          =   315
      Left            =   885
      TabIndex        =   2
      Top             =   825
      Width           =   3420
   End
   Begin VB.ComboBox Tributo 
      Height          =   315
      Left            =   885
      TabIndex        =   0
      Top             =   345
      Width           =   3420
   End
   Begin MSMask.MaskEdBox ProxLivro 
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Top             =   2685
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Folha 
      Height          =   285
      Left            =   3690
      TabIndex        =   15
      Top             =   2670
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   503
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "A partir da Folha:"
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
      Left            =   2220
      TabIndex        =   16
      Top             =   2715
      Width           =   1485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Próximo Livro:"
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
      Left            =   330
      TabIndex        =   14
      Top             =   2715
      Width           =   1215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Livro:"
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
      Left            =   345
      TabIndex        =   7
      Top             =   885
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tributo:"
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
      TabIndex        =   1
      Top             =   405
      Width           =   675
   End
End
Attribute VB_Name = "EscrituracaoFechamentoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'se usuario confirmar entao:
    'criar registro na tabela livroperiodo e atualizar referencias dos registros nas tabelas de reges e outras utilizadas para a "geracao/impressao definitiva".
        'a partir deste momento os registros das tabelas vinculados ao livroperiodo nao podem mais ser alterados.
    'atualizar periodo atual de cada livro do tributo em livrofilial
    'vou precisar de SELECT case p/dar tratamento especifico p/cada tabela ???
    'se for fechto de apuracao, os registros de reges nao podem ser alterados de forma a alterarem o resultado da apuracao.

'Property Variables:
Dim m_Caption As String
Event Unload()

Function Trata_Parametros(Optional objLivrosFilial As ClassLivrosFilial) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159473)

    End Select

    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    'Carrega Tributos
    lErro = Carrega_Tributos()
    If lErro <> SUCESSO Then gError 69596

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 69596

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159474)

    End Select

    Exit Sub

End Sub

Function Carrega_Tributos() As Long

Dim lErro As Long
Dim colTributos As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Carrega_Tributos

    'Lê Tributos da tabela Tributos que possuem Livro Fiscal
    lErro = CF("Tributos_Le",colTributos)
    If lErro <> SUCESSO Then gError 69597

    'Preenche a combo de Tributos
    For iIndice = 1 To colTributos.Count
        Tributo.AddItem colTributos(iIndice).sDescricao
        Tributo.ItemData(Tributo.NewIndex) = colTributos(iIndice).iCodigo
    Next

    Carrega_Tributos = SUCESSO

    Exit Function

Erro_Carrega_Tributos:

    Carrega_Tributos = gErr

    Select Case gErr

        Case 69597

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159475)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoCancela_Click()

    Unload Me

End Sub

Private Sub ProxLivro_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_ProxLivro_Validate

    'Se o campo foi preenchido
    If Len(Trim(ProxLivro.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(ProxLivro.Text)
        If lErro <> SUCESSO Then gError 69598

    End If

    Exit Sub

Erro_ProxLivro_Validate:

    Cancel = True

    Select Case gErr

        Case 69598

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159476)

    End Select

    Exit Sub

End Sub

Private Sub Folha_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Folha_Validate

    'Se o campo foi preenchido
    If Len(Trim(Folha.ClipText)) > 0 Then

        'Critica o valor
        lErro = Valor_Positivo_Critica(Folha.Text)
        If lErro <> SUCESSO Then gError 69599

    End If

    Exit Sub

Erro_Folha_Validate:

    Cancel = True

    Select Case gErr

        Case 69599

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159477)

    End Select

    Exit Sub

End Sub

Private Sub ProxLivro_GotFocus()

    Call MaskEdBox_TrataGotFocus(ProxLivro)

End Sub

Private Sub Folha_GotFocus()

    Call MaskEdBox_TrataGotFocus(Folha)

End Sub

Private Sub Tributo_Click()

Dim lErro As Long
Dim colLivrosFilial As New Collection
Dim iIndice As Integer

On Error GoTo Erro_Tributo_Click

    'Se nenhum Tributo foi selecionado, sai da rotina
    If Tributo.ListIndex = -1 Then Exit Sub

    'Limpa combo de Livros
    Livro.Clear

    'Limpa dados do Livro
    Call Limpa_DadosLivro

    'Lê os Livros Filiais Abertos
    lErro = CF("LivrosFilial_Le_TodosAbertos",Tributo.ItemData(Tributo.ListIndex), colLivrosFilial)
    If lErro <> SUCESSO Then gError 69600

    'Carrega a combo de Livro com Todos os Livros do Tributo selecionado
    For iIndice = 1 To colLivrosFilial.Count
        Livro.AddItem colLivrosFilial(iIndice).sDescricao
        Livro.ItemData(Livro.NewIndex) = colLivrosFilial(iIndice).iCodLivro
    Next

    Exit Sub

Erro_Tributo_Click:

    Select Case gErr

        Case 69600

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159478)

    End Select

    Exit Sub

End Sub

Private Sub Livro_Click()

Dim lErro As Long
Dim objLivrosFilial As New ClassLivrosFilial
Dim objLivroFiscal As New ClassLivrosFiscais
Dim iIndice As Integer
Dim dtData As Date
Dim dtDataFinal As Date

On Error GoTo Erro_Livro_Click

    'Se a Nenhum Livro foi selecionado, sai da rotina
    If Livro.ListIndex = -1 Then Exit Sub

    objLivrosFilial.iCodLivro = CInt(Livro.ItemData(Livro.ListIndex))
    objLivrosFilial.iFilialEmpresa = giFilialEmpresa

    'Verifica se o Livro em questão está cadastrado na Tabela de LivrosFIlial
    lErro = CF("LivrosFilial_Le",objLivrosFilial)
    If lErro <> SUCESSO And lErro <> 67992 Then gError 69601

    If lErro = 67992 Then gError 69602

    'Se encontrou o Livro em questão
    If lErro = SUCESSO Then

        'Preenche os campos da Tela com os dados de LivrosFilial
        Call Traz_LivrosFilial_Tela(objLivrosFilial)

    End If

    Exit Sub

Erro_Livro_Click:

    Select Case gErr

        Case 69601

        Case 69602
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVROFISCAL_NAO_CADASTRADO", gErr, objLivroFiscal.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159479)

    End Select

    Exit Sub

End Sub

Sub Traz_LivrosFilial_Tela(objLivrosFilial As ClassLivrosFilial)
'Traz os dados do LivroFilial para a tela

Dim objLivroFiscal As New ClassLivrosFiscais
Dim iIndice As Integer

    DataInicio.Caption = Format(objLivrosFilial.dtDataInicial, "dd/mm/yyyy")
    DataFim.Caption = Format(objLivrosFilial.dtDataFinal, "dd/mm/yyyy")
    
    'Se a data de impressão estiver preenchida
    If objLivrosFilial.dtImpressoEm <> DATA_NULA Then
        DataImpressao.Caption = Format(objLivrosFilial.dtImpressoEm, "dd/mm/yyyy")
    Else
        DataImpressao.Caption = ""
    End If

End Sub

Private Sub BotaoOk_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoOk_Click

    'Grava um Tributo
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 69603

    'Limpa a tela
    Call Limpa_Tela_EscrituracaoFechamento

    Exit Sub

Erro_BotaoOk_Click:

    Select Case gErr

        Case 69603

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159480)

    End Select

Exit Sub

End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objLivroFilial As New ClassLivrosFilial

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Verifica se Tributo foi preenchido
    If Len(Trim(Tributo.Text)) = 0 Then gError 69604

    'Verifica se o Livro foi preenchido
    If Len(Trim(Livro.Text)) = 0 Then gError 69605

    'Verifica se o número do livro foi preenchido
    If Len(Trim(ProxLivro.ClipText)) = 0 Then gError 69606

    'Verifica se o número da Folha foi preenchida
    If Len(Trim(Folha.ClipText)) = 0 Then gError 69607

    'Recolhe os dados da tela
    Call Move_Tela_Memoria(objLivroFilial)

    'Faz o Fechamento do Livro Selecionado
    lErro = CF("Rotina_Fechamento_Livro",objLivroFilial)
    If lErro <> SUCESSO Then gError 69608

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    GL_objMDIForm.MousePointer = vbDefault
 
    Gravar_Registro = gErr

    Select Case gErr

        Case 69604
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRIBUTO_NAO_PREENCHIDO", gErr)

        Case 69605
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LIVRO_NAO_PREENCHIDO", gErr)

        Case 69606
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NUMEROLIVRO_NAO_PREENCHIDO", gErr)

        Case 69607
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FOLHA_NAO_PREENCHIDA", gErr)

        Case 69608

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 159481)

    End Select

    Exit Function

End Function

Sub Move_Tela_Memoria(objLivroFilial As ClassLivrosFilial)
'Move os dados que seram inseridos na nova Tabela de Livro Filial

    objLivroFilial.dtDataInicial = CDate(DataInicio.Caption)
    objLivroFilial.dtDataFinal = CDate(DataFim.Caption)
    objLivroFilial.iNumeroProxFolha = Folha.Text
    objLivroFilial.iNumeroProxLivro = ProxLivro.Text
    objLivroFilial.iFilialEmpresa = giFilialEmpresa
    objLivroFilial.iCodLivro = Livro.ItemData(Livro.ListIndex)

End Sub

Private Sub Limpa_Tela_EscrituracaoFechamento()
'Limpa os dados da Tela

    Livro.ListIndex = -1
    Call Limpa_DadosLivro

End Sub

Sub Limpa_DadosLivro()
'Limpa os dados do Livro

    DataInicio.Caption = ""
    DataFim.Caption = ""
    DataImpressao.Caption = ""
    ProxLivro.Text = ""
    Folha.Text = ""

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Fechamento dos Livros"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "EscrituracaoFechamento"

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

Private Sub DataImpressao_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataImpressao, Source, X, Y)
End Sub

Private Sub DataImpressao_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataImpressao, Button, Shift, X, Y)
End Sub

Private Sub DataFim_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataFim, Source, X, Y)
End Sub

Private Sub DataFim_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataFim, Button, Shift, X, Y)
End Sub

Private Sub DataInicio_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(DataInicio, Source, X, Y)
End Sub

Private Sub DataInicio_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(DataInicio, Button, Shift, X, Y)
End Sub

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label10, Source, X, Y)
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label10, Button, Shift, X, Y)
End Sub

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label9, Source, X, Y)
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label9, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label3, Source, X, Y)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label3, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label7, Source, X, Y)
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label7, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

