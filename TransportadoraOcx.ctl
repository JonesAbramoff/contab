VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TransportadoraOcx 
   ClientHeight    =   6165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9255
   KeyPreview      =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   9255
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5010
      Index           =   1
      Left            =   195
      TabIndex        =   16
      Top             =   900
      Width           =   8850
      Begin VB.Frame FrameInscricoes 
         Caption         =   "Inscrições"
         Height          =   1515
         Left            =   120
         TabIndex        =   28
         Top             =   3495
         Width           =   6015
         Begin VB.CheckBox IENaoContrib 
            Caption         =   "Não Contribuinte do ICMS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3390
            TabIndex        =   38
            Top             =   1065
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.CheckBox IEIsento 
            Caption         =   "Isento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3390
            TabIndex        =   37
            Top             =   720
            Value           =   1  'Checked
            Width           =   975
         End
         Begin MSMask.MaskEdBox InscricaoEstadual 
            Height          =   315
            Left            =   1920
            TabIndex        =   8
            Top             =   630
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox CGC 
            Height          =   315
            Left            =   1920
            TabIndex        =   7
            Top             =   210
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   14
            Mask            =   "##############"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox InscricaoMunicipal 
            Height          =   315
            Left            =   1920
            TabIndex        =   9
            Top             =   1050
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   18
            PromptChar      =   " "
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Municipal:"
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
            Left            =   120
            TabIndex        =   31
            Top             =   1110
            Width           =   1725
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Estadual:"
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
            Left            =   195
            TabIndex        =   30
            Top             =   690
            Width           =   1650
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "CNPJ/CPF:"
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
            Left            =   885
            TabIndex        =   29
            Top             =   255
            Width           =   975
         End
      End
      Begin VB.Frame FramePrincipal 
         Caption         =   "Principal"
         Height          =   3300
         Left            =   120
         TabIndex        =   23
         Top             =   135
         Width           =   6015
         Begin VB.TextBox Observacao 
            Height          =   840
            Left            =   1785
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   34
            Top             =   2370
            Width           =   3885
         End
         Begin VB.TextBox Guia 
            Height          =   300
            Left            =   4380
            MaxLength       =   10
            TabIndex        =   6
            Top             =   1935
            Width           =   1290
         End
         Begin MSMask.MaskEdBox PesoMinimo 
            Height          =   330
            Left            =   1785
            TabIndex        =   5
            Top             =   1935
            Width           =   525
            _ExtentX        =   926
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
         Begin VB.ComboBox ViaTransporte 
            Height          =   315
            ItemData        =   "TransportadoraOcx.ctx":0000
            Left            =   1785
            List            =   "TransportadoraOcx.ctx":001C
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1500
            Width           =   2040
         End
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2340
            Picture         =   "TransportadoraOcx.ctx":0082
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Numeração Automática"
            Top             =   255
            Width           =   300
         End
         Begin MSMask.MaskEdBox Nome 
            Height          =   315
            Left            =   1785
            TabIndex        =   2
            Top             =   660
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox NomeReduzido 
            Height          =   315
            Left            =   1785
            TabIndex        =   3
            Top             =   1080
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Codigo 
            Height          =   315
            Left            =   1785
            TabIndex        =   0
            Top             =   240
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   4
            Mask            =   "9999"
            PromptChar      =   " "
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Observação:"
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
            Left            =   615
            TabIndex        =   35
            Top             =   2430
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Guia:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   3840
            TabIndex        =   33
            Top             =   1995
            Width           =   555
         End
         Begin VB.Label Label6 
            Caption         =   "Peso Mínimo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   525
            TabIndex        =   32
            Top             =   2010
            Width           =   1185
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Via de Transporte:"
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
            TabIndex        =   27
            Top             =   1545
            Width           =   1590
         End
         Begin VB.Label LabelCodigo 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   1050
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   26
            Top             =   300
            Width           =   660
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
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
            Left            =   1155
            TabIndex        =   25
            Top             =   705
            Width           =   555
         End
         Begin VB.Label LabelNomeReduzido 
            AutoSize        =   -1  'True
            Caption         =   "Nome Reduzido:"
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
            Left            =   300
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   24
            Top             =   1125
            Width           =   1410
         End
      End
      Begin VB.ListBox TransportadoraList 
         Height          =   4545
         ItemData        =   "TransportadoraOcx.ctx":016C
         Left            =   6360
         List            =   "TransportadoraOcx.ctx":016E
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   285
         Width           =   2370
      End
      Begin VB.Label Label13 
         Caption         =   "Transportadoras"
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
         Left            =   6360
         TabIndex        =   20
         Top             =   15
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5025
      Index           =   2
      Left            =   195
      TabIndex        =   11
      Top             =   900
      Visible         =   0   'False
      Width           =   8850
      Begin TelasCpr.TabEndereco TabEnd 
         Height          =   3975
         Index           =   0
         Left            =   15
         TabIndex        =   36
         Top             =   1200
         Width           =   8325
         _ExtentX        =   14684
         _ExtentY        =   7011
      End
      Begin VB.Frame SSFrame9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   105
         TabIndex        =   18
         Top             =   210
         Width           =   8625
         Begin VB.Label Label56 
            Caption         =   "Transportadora:"
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
            Left            =   210
            TabIndex        =   21
            Top             =   195
            Width           =   1365
         End
         Begin VB.Label TransportadoraLabel 
            Height          =   210
            Left            =   1620
            TabIndex        =   22
            Top             =   210
            Width           =   6915
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6945
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   135
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TransportadoraOcx.ctx":0170
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TransportadoraOcx.ctx":02CA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "TransportadoraOcx.ctx":0454
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "TransportadoraOcx.ctx":0986
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip Opcao 
      Height          =   5505
      Left            =   135
      TabIndex        =   19
      Top             =   510
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   9710
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Identificação"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Endereço"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "TransportadoraOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoCidade As AdmEvento
Attribute objEventoCidade.VB_VarHelpID = -1
Private WithEvents objEventoTransportadora As AdmEvento
Attribute objEventoTransportadora.VB_VarHelpID = -1

Dim sIEAnt As String

'DECLARACAO DE VARIAVEIS GLOBAIS
Dim iFrameAtual As Integer
Public iAlterado As Integer
Public gobjTabEnd As New ClassTabEndereco

'Constantes públicas dos tabs
Private Const TAB_Identificacao = 1
Private Const TAB_Endereco = 2

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Gera código automático da próxima Transportadora
    lErro = CF("Transportadora_Automatico", iCodigo)
    If lErro <> SUCESSO Then Error 57558

    'Exibe código na Tela
    Codigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case Err

        Case 57558 'Tratado na Rotina Chamada
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175567)
    
    End Select

    Exit Sub

End Sub

Private Sub CGC_Change()

    iAlterado = REGISTRO_ALTERADO
    
End Sub

Private Sub CGC_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(CGC, iAlterado)

End Sub

Private Sub CGC_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_CGC_Validate

    If Len(Trim(CGC.Text)) = 0 Then Exit Sub
    
    'Pelo Tamanho verifica se é CPF ou CGC
    Select Case Len(Trim(CGC.Text))
        
        Case STRING_CPF 'CPF
            
            lErro = Cpf_Critica(CGC.Text)
            If lErro <> SUCESSO Then Error 33532
            
            CGC.Format = "000\.000\.000-00; ; ; "
            CGC.Text = CGC.Text
        
        Case STRING_CGC  'CGC

            lErro = Cgc_Critica(CGC.Text)
            If lErro <> SUCESSO Then Error 33533
            
            CGC.Format = "00\.000\.000\/0000-00; ; ; "
            CGC.Text = CGC.Text
            
        Case Else

            Error 33534

    End Select

    Exit Sub

Erro_CGC_Validate:

    Cancel = True


    Select Case Err

        Case 33532, 33533

        Case 33534
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TAMANHO_CGC_CPF", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175568)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Codigo_GotFocus()

    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se o Código foi preenchido o Codigo
    If Len(Trim(Codigo.Text)) = 0 Then Exit Sub

    'Critica se é do tipo inteiro positivo
    lErro = Inteiro_Critica(Codigo.Text)
    If lErro <> SUCESSO Then Error 22045

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True


    Select Case Err

        Case 22045
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175569)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim objTransportadora As New ClassTransportadora
Dim colCodigoDescricao As New AdmColCodigoNome
Dim objCodigoDescricao As New AdmCodigoNome
Dim colCodigo As New Collection
Dim vCodigo As Variant
Dim lErro As Long
Dim iIndice As Integer
Dim objTela As Object

On Error GoTo Erro_Form_Load
    
    Set objEventoTransportadora = New AdmEvento
    
    iFrameAtual = 1

    'Lê Códigos e NomesReduzidos da tabela Transportadoras e devolve na coleção
    lErro = CF("Cod_Nomes_Le", "Transportadoras", "Codigo", "NomeReduzido", STRING_TRANSPORTADORA_NOME_REDUZIDO, colCodigoDescricao)
    If lErro <> SUCESSO Then Error 22037

    'Preenche a ListBox TransportadoraList com os objetos da coleção
    For Each objCodigoDescricao In colCodigoDescricao
        TransportadoraList.AddItem objCodigoDescricao.sNome
        TransportadoraList.ItemData(TransportadoraList.NewIndex) = objCodigoDescricao.iCodigo
    Next

    'Lê cada código da tabela Estados
    Set objTela = Me
    lErro = gobjTabEnd.Inicializa(objTela, TabEnd(0))
    If lErro <> SUCESSO Then Error 22038

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case 22037, 22038, 22039

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175570)

    End Select
    
    iAlterado = 0
    
    Exit Sub

End Sub

Function Trata_Parametros(Optional objTransportadora As ClassTransportadora) As Long

Dim lErro As Long
Dim sListBoxItem As String

On Error GoTo Erro_Trata_Parametros

    'Se há um Transportadora selecionada
    If Not (objTransportadora Is Nothing) Then

        'Verifica se a Transportadora existe, lendo no BD a partir do  codigo
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then Error 22061

        'Se a Transportadora existe
        If lErro = SUCESSO Then
        
            lErro = Traz_Transportadora_Tela(objTransportadora)
            If lErro <> SUCESSO Then Error 58592
            
        'Se a Transportadora não existe
        Else

            'Mantém o Código da Transportadora na tela
            Codigo.Text = CStr(objTransportadora.iCodigo)

        End If

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case 22061, 58592 'Tratados nas rotinas chamadas

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175571)

    End Select
    
    iAlterado = 0

    Exit Function

End Function

Function Traz_Transportadora_Tela(objTransportadora As ClassTransportadora) As Long

'Alteracao Daniel em 21/01/2002 - inclusão do campo Inscrição Municipal

Dim lErro As Long
Dim sListBoxItem As String
Dim iCodigo As Integer
Dim objEndereco As New ClassEndereco
Dim iIndice As Integer
Dim colEnderecos As New Collection

On Error GoTo Erro_Traz_Transportadora_Tela
    
    'Carrega lEndereco em objTransportadora
    objEndereco.lCodigo = objTransportadora.lEndereco

    'Lê o endereço a partir do Código
    lErro = CF("Endereco_Le", objEndereco)
    If lErro <> SUCESSO And lErro <> 12309 Then Error 22063

    If lErro = 12309 Then Error 22118
    
    colEnderecos.Add objEndereco

    'Exibe os dados de objTransportadora na tela
    If objTransportadora.iCodigo = 0 Then
        Codigo.Text = ""
    Else
        Codigo.Text = CStr(objTransportadora.iCodigo)
    End If
    Nome.Text = objTransportadora.sNome
    NomeReduzido.Text = objTransportadora.sNomeReduzido
    Call NomeReduzido_Validate(bSGECancelDummy)
        
    ViaTransporte.ListIndex = -1
    
    'traz Via Transporte a tela
    For iIndice = 0 To (ViaTransporte.ListCount - 1)
        If ViaTransporte.ItemData(iIndice) = objTransportadora.iViaTransporte Then
            ViaTransporte.ListIndex = iIndice
            Exit For
        End If
    Next
    
    'Preenche o CGC
    If objTransportadora.sCgc <> "" Then
        CGC.Text = objTransportadora.sCgc
        Call CGC_Validate(bSGECancelDummy)
    Else
        CGC.Text = ""
    End If
    
    'Preenche a Incricao Estadual e Municipal
    InscricaoEstadual.Text = objTransportadora.sInscricaoEstadual
    Call Trata_IE
    If objTransportadora.iIEIsento = MARCADO Then
        IEIsento.Value = vbChecked
    Else
        IEIsento.Value = vbUnchecked
    End If
    If objTransportadora.iIENaoContrib = MARCADO Then
        IENaoContrib.Value = vbChecked
    Else
        IENaoContrib.Value = vbUnchecked
    End If
    InscricaoMunicipal.Text = objTransportadora.sInscricaoMunicipal
    
    PesoMinimo.PromptInclude = False
    If objTransportadora.dPesoMinimo <> 0 Then
        PesoMinimo.Text = Format(objTransportadora.dPesoMinimo, PesoMinimo.Format)
    Else
        PesoMinimo.Text = ""
    End If
    PesoMinimo.PromptInclude = True
    
    Guia.Text = objTransportadora.sGuia
    Observacao.Text = objTransportadora.sObservacao
    
    lErro = gobjTabEnd.Traz_Endereco_Tela(colEnderecos)
    If lErro <> SUCESSO Then Error 22063

    iAlterado = 0

    Traz_Transportadora_Tela = SUCESSO

    Exit Function

Erro_Traz_Transportadora_Tela:

    Traz_Transportadora_Tela = Err

    Select Case Err

        Case 22063 'Tratado na Rotina chamada

        Case 22118
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ENDERECO_NAO_CADASTRADO", Err)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 175572)

    End Select

    Exit Function

End Function

Private Sub InscricaoEstadual_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub InscricaoMunicipal_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Nome_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NomeReduzido_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NomeReduzido_Validate
    
    'Se está preenchido, testa se começa por letra
    If Len(Trim(NomeReduzido.Text)) > 0 Then

        If Not IniciaLetra(NomeReduzido.Text) Then Error 57821

    End If
    
    TransportadoraLabel.Caption = Trim(NomeReduzido.Text)
    
    Exit Sub

Erro_NomeReduzido_Validate:

    Cancel = True

    
    Select Case Err
    
        Case 57821
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_COMECA_LETRA", Err, NomeReduzido.Text)
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175574)
    
    End Select
    
    Exit Sub

End Sub

Private Sub Opcao_Click()

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If Opcao.SelectedItem.Index <> iFrameAtual Then

        If TabStrip_PodeTrocarTab(iFrameAtual, Opcao, Me) <> SUCESSO Then Exit Sub
        
        'Torna Frame correspondente ao Tab selecionado visivel
        Frame1(Opcao.SelectedItem.Index).Visible = True
        'Torna Frame atual visivel
        Frame1(iFrameAtual).Visible = False
        'Armazena novo valor de iFrameAtual
        iFrameAtual = Opcao.SelectedItem.Index
        
        Select Case iFrameAtual
        
            Case TAB_Identificacao
                Parent.HelpContextID = IDH_TRANSPORTADORA_ID
                
            Case TAB_Endereco
                Parent.HelpContextID = IDH_TRANSPORTADORA_ENDERECO
                        
        End Select
        
    End If

End Sub

Private Sub TransportadoraList_DblClick()

Dim lErro As Long
Dim sListBoxItem As String
Dim objTransportadora As New ClassTransportadora

On Error GoTo Erro_TransportadoraList_DblClick

    'Guarda o valor do código da Transportadora selecionada na ListBox TransportadoraList
    objTransportadora.iCodigo = TransportadoraList.ItemData(TransportadoraList.ListIndex)

    'Lê a Transportadora no BD
    lErro = CF("Transportadora_Le", objTransportadora)
    If lErro <> SUCESSO And lErro <> 19250 Then Error 22112

    'Se Transportadora não está cadastrada, erro
    If lErro = 19250 Then Error 22113

    'Exibe os dados da Transportadora
    lErro = Traz_Transportadora_Tela(objTransportadora)
    If lErro <> SUCESSO Then Error 22114

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    Exit Sub

Erro_TransportadoraList_DblClick:

    Select Case Err

        Case 22112, 22114 'Tratado nas Rotinas chamadas

        Case 22113
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objTransportadora.iCodigo)
            TransportadoraList.RemoveItem (TransportadoraList.ListIndex)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175576)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoGravar_Click

    'Grava a Transportadora
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 22046

    'Limpa a Tela
    Call Limpa_Tela_Transportadora

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case Err

        Case 22046

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175577)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTransportadora As New ClassTransportadora
Dim objEndereco As New ClassEndereco

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se foi preenchido o Código
    If Len(Trim(Codigo.Text)) = 0 Then Error 22048

    'Verifica se foi preenchido o Nome
    If Len(Trim(Nome.Text)) = 0 Then Error 22049

    'Verifica se foi preenchido o Nome Reduzido
    If Len(Trim(NomeReduzido.Text)) = 0 Then Error 22050
    
    'Verifica se foi prenchido a transportadora
    If Len(Trim(ViaTransporte.List(ViaTransporte.ListIndex))) = 0 Then Error 52284

    'Preenche os objetos com os dados da tela
    lErro = Move_Tela_Memoria(objTransportadora, objEndereco)
    If lErro <> SUCESSO Then Error 22051

    lErro = Trata_Alteracao(objTransportadora, objTransportadora.iCodigo)
    If lErro <> SUCESSO Then Error 22054

    'Grava a Transportadora no BD
    lErro = CF("Transportadora_Grava", objTransportadora, objEndereco)
    If lErro <> SUCESSO Then Error 22052

    'Atualiza ListBox de Transportadora
    Call TransportadoraList_Remove(objTransportadora)
    Call TransportadoraList_Adiciona(objTransportadora)

    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 22048
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", Err)

        Case 22049
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_NAO_PREENCHIDO", Err)

        Case 22050
            lErro = Rotina_Erro(vbOKOnly, "ERRO_NOME_REDUZIDO_NAO_PREENCHIDO", Err)

        Case 22051, 22052, 22054
        
        Case 52284
            lErro = Rotina_Erro(vbOKOnly, "ERRO_VIA_TRANSPORTE_NAO_PREENCHIDO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175578)

    End Select

    Exit Function

End Function

Private Sub TransportadoraList_Remove(objTransportadora As ClassTransportadora)
'Percorre a ListBox Transportadoralist para remover a Transportadora caso ela exista

Dim iIndice As Integer

    For iIndice = 0 To TransportadoraList.ListCount - 1
    
        If TransportadoraList.ItemData(iIndice) = objTransportadora.iCodigo Then
    
            TransportadoraList.RemoveItem iIndice
            Exit For
    
        End If
    
    Next

End Sub

Private Sub TransportadoraList_Adiciona(objTransportadora As ClassTransportadora)
'Inclui Transportadora na List

    TransportadoraList.AddItem objTransportadora.sNomeReduzido
    TransportadoraList.ItemData(TransportadoraList.NewIndex) = objTransportadora.iCodigo

End Sub

Private Function Move_Tela_Memoria(objTransportadora As ClassTransportadora, objEndereco As ClassEndereco, Optional ByVal bMovEnd As Boolean = True) As Long
'Lê os dados que estão na tela Transportadora e coloca em objTransportadora

'Alteracao Daniel em 21/01/2002 - Inclusão do campo Inscrição Municipal

Dim lErro As Long
Dim iPais As Integer
Dim colEnderecos As New Collection

    'IDENTIFICACAO :
    If Len(Trim(Codigo.Text)) > 0 Then objTransportadora.iCodigo = CInt(Codigo.Text)
    objTransportadora.sNome = Trim(Nome.Text)
    objTransportadora.sNomeReduzido = Trim(NomeReduzido.Text)
    objTransportadora.sCgc = CGC.Text
    objTransportadora.sInscricaoEstadual = InscricaoEstadual.Text
    objTransportadora.sInscricaoMunicipal = InscricaoMunicipal.Text
    If Len(Trim(ViaTransporte.List(ViaTransporte.ListIndex))) > 0 Then objTransportadora.iViaTransporte = ViaTransporte.ItemData(ViaTransporte.ListIndex)
    objTransportadora.dPesoMinimo = StrParaDbl(PesoMinimo.Text)
    objTransportadora.sGuia = Trim(Guia.Text)
    objTransportadora.sObservacao = Trim(Observacao.Text)
    
    If IEIsento.Value = vbChecked Then
        objTransportadora.iIEIsento = MARCADO
    Else
        objTransportadora.iIEIsento = DESMARCADO
    End If
    If IENaoContrib.Value = vbChecked Then
        objTransportadora.iIENaoContrib = MARCADO
    Else
        objTransportadora.iIENaoContrib = DESMARCADO
    End If
    
    'ENDERECO
    If bMovEnd Then
        lErro = gobjTabEnd.Move_Endereco_Memoria(colEnderecos)
        If lErro <> SUCESSO Then gError 22048
        Set objEndereco = colEnderecos.Item(1)
    End If

    Move_Tela_Memoria = SUCESSO
    
    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 22048

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175578)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTransportadora As New ClassTransportadora
Dim colCodNomeFiliais As New AdmColCodigoNome
Dim iCodigo As Integer
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Verifica se o codigo foi preenchido
    If Len(Trim(Codigo.Text)) = 0 Then Error 22060

    objTransportadora.iCodigo = CInt(Codigo.Text)

    'Lê os dados da Transportadora a ser excluida
    lErro = CF("Transportadora_Le", objTransportadora)
    If lErro <> SUCESSO And lErro <> 19250 Then Error 22058

    'Verifica se Transportadora está cadastrada
    If lErro <> SUCESSO Then Error 22059

    'Envia aviso perguntando se realmente deseja excluir Transportadora
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_TRANSPORTADORA", objTransportadora.iCodigo, colCodNomeFiliais.Count - 1)

    If vbMsgRes = vbYes Then

        'Exclui Transportadora
        lErro = CF("Transportadora_Exclui", objTransportadora)
        If lErro <> SUCESSO Then Error 22078

        'Exclui da ListBox
        Call TransportadoraList_Remove(objTransportadora)

        'Limpa a Tela
        Call Limpa_Tela_Transportadora

    End If

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err

        Case 22060
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CODTRANSPORTADORA_NAO_PREENCHIDO", Err)

        Case 22058, 22078

        Case 22059   'Transportadora com codigo %i nao esta cadastrada
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", Err, objTransportadora.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175579)

    End Select

    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoLimpar_Click

    'Testa se deseja salvar mudanças
    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 80461

    'Limpa a Tela
    Call Limpa_Tela_Transportadora

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 80461

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175580)

    End Select

    Exit Sub
    
End Sub

Sub Limpa_Tela_Transportadora()

Dim iIndice As Integer, lErro As Long, iCodigo As Integer

On Error GoTo Erro_Limpa_Tela_Transportadora

    'Fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Fechar(Me.Name)

    'Limpa TextBox e MaskedEditBox
    Call Limpa_Tela(Me)
    
    Codigo.Text = ""
    
    'Limpa os textos das Combos
    Call gobjTabEnd.Limpa_Tela
    
    'Limpa o Label Transportadora
    TransportadoraLabel.Caption = ""
    
    IEIsento.Value = vbChecked
    IENaoContrib.Value = vbChecked
    
    'limpa via transporte
    ViaTransporte.ListIndex = -1
    
    iAlterado = 0

    Exit Sub
    
Erro_Limpa_Tela_Transportadora:

    Select Case Err

        Case Else

            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175581)

    End Select

    Exit Sub
    
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

    Set objEventoTransportadora = Nothing
    
    Call gobjTabEnd.Limpa_Tela
    Set gobjTabEnd = Nothing
    
    'Libera a referencia da tela e fecha o comando das setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)
 
End Sub

Private Sub TransportadoraLabel_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

'Alteracao Daniel em 21/01/2002 - Inclusão do cmapo Inscrição Municipal

Dim lErro As Long
Dim objTransportadora As New ClassTransportadora
Dim objEndereco As New ClassEndereco
On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "Transportadoras"

    'Le os dados da Tela Transportadora
    lErro = Move_Tela_Memoria(objTransportadora, objEndereco, False)
    If lErro <> SUCESSO Then gError 22117

    'No lEndereco armazena  0
    objTransportadora.lEndereco = 0

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objTransportadora.iCodigo, 0, "Codigo"
    colCampoValor.Add "Nome", objTransportadora.sNome, STRING_TRANSPORTADORA_NOME, "Nome"
    colCampoValor.Add "NomeReduzido", objTransportadora.sNomeReduzido, STRING_TRANSPORTADORA_NOME_REDUZIDO, "NomeReduzido"
    colCampoValor.Add "CGC", objTransportadora.sCgc, STRING_CGC, "CGC"
    colCampoValor.Add "InscricaoEstadual", objTransportadora.sInscricaoEstadual, STRING_INSCR_EST, "InscricaoEstadual"
    colCampoValor.Add "InscricaoMunicipal", objTransportadora.sInscricaoMunicipal, STRING_INSCR_MUN, "InscricaoMunicipal"
    colCampoValor.Add "Endereco", objTransportadora.lEndereco, 0, "Endereco"
    colCampoValor.Add "ViaTransporte", objTransportadora.iViaTransporte, 0, "ViaTransporte"
    colCampoValor.Add "PesoMinimo", objTransportadora.dPesoMinimo, 0, "PesoMinimo"
    colCampoValor.Add "Guia", objTransportadora.sGuia, STRING_TRANSPORTADORA_GUIA, "Guia"
    colCampoValor.Add "Observacao", objTransportadora.sObservacao, STRING_TRANSPORTADORA_OBS, "Observacao"

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 22117

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175582)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

'Alteracao Daniel em 21/01/2002 - Inclusão do campo Inscrição Municipal

Dim lErro As Long
Dim objTransportadora As New ClassTransportadora

On Error GoTo Erro_Tela_Preenche

    objTransportadora.iCodigo = colCampoValor.Item("Codigo").vValor

    If objTransportadora.iCodigo > 0 Then

        'Carrega objTransportadora com os dados passados em colCampoValor
        objTransportadora.sNome = colCampoValor.Item("Nome").vValor
        objTransportadora.sNomeReduzido = colCampoValor.Item("NomeReduzido").vValor
        objTransportadora.sCgc = colCampoValor.Item("CGC").vValor
        objTransportadora.sInscricaoEstadual = colCampoValor.Item("InscricaoEstadual").vValor
        objTransportadora.sInscricaoMunicipal = colCampoValor.Item("InscricaoMunicipal").vValor
        objTransportadora.lEndereco = colCampoValor.Item("Endereco").vValor
        objTransportadora.iViaTransporte = colCampoValor.Item("ViaTransporte").vValor
        objTransportadora.dPesoMinimo = colCampoValor.Item("PesoMinimo").vValor
        objTransportadora.sGuia = colCampoValor.Item("Guia").vValor
        objTransportadora.sObservacao = colCampoValor.Item("Observacao").vValor

        'Lê a Transportadora no BD
        lErro = CF("Transportadora_Le", objTransportadora)
        If lErro <> SUCESSO And lErro <> 19250 Then Error 22116

        'Traz dados da Transportadora para a Tela
        lErro = Traz_Transportadora_Tela(objTransportadora)
        If lErro <> SUCESSO Then Error 22116

    End If

    Exit Sub

Erro_Tela_Preenche:

    Select Case Err

        Case 22116 'Tratado na Rotina Chamada

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 175583)

    End Select

    Exit Sub

End Sub

Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_TRANSPORTADORA_ID
    Set Form_Load_Ocx = Me
    Caption = "Transportadora"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "Transportadora"
    
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

'***** fim do trecho a ser copiado ******

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_PROXIMO_NUMERO Then
        Call BotaoProxNum_Click
    End If

    If KeyCode = KEYCODE_BROWSER Then
        If Me.ActiveControl Is Codigo Then
            Call LabelCodigo_Click
        ElseIf Me.ActiveControl Is NomeReduzido Then
            Call LabelNomeReduzido_Click
        End If
    End If
    
End Sub

Private Sub LabelCodigo_Click()

Dim objTransportadora As New ClassTransportadora
Dim colSelecao As New Collection
    
    'Se a Transportadora estiver preenchida passa o código para o objTransportadora
    If Len(Trim(Codigo.Text)) > 0 Then objTransportadora.iCodigo = StrParaInt(Codigo.Text)
    
    'Chama a tela que lista as transportadoras
    Call Chama_Tela("TransportadoraLista", colSelecao, objTransportadora, objEventoTransportadora)

End Sub

Private Sub LabelNomeReduzido_Click()

Dim objTransportadora As New ClassTransportadora
Dim colSelecao As New Collection
    
    'Se a Transportadora estiver preenchida passa o código para o objTransportadora
    If Len(Trim(NomeReduzido.Text)) > 0 Then objTransportadora.sNomeReduzido = NomeReduzido.Text
    
    'Chama a tela que lista as transportadoras
    Call Chama_Tela("TransportadoraLista", colSelecao, objTransportadora, objEventoTransportadora)

End Sub

Private Sub objEventoTransportadora_evSelecao(obj1 As Object)

Dim objTransportadora As ClassTransportadora
Dim lErro As Long

On Error GoTo Erro_objEventoTransportadora_evSelecao

    Set objTransportadora = obj1

    'Lê a Transportadora no BD
    lErro = CF("Transportadora_Le", objTransportadora)
    If lErro <> SUCESSO And lErro <> 19250 Then gError 196960

    'Se Transportadora não está cadastrada, erro
    If lErro = 19250 Then gError 196961

    'Exibe os dados da Transportadora
    lErro = Traz_Transportadora_Tela(objTransportadora)
    If lErro <> SUCESSO Then gError 196962

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    Me.Show

    Exit Sub

Erro_objEventoTransportadora_evSelecao:

    Select Case gErr

        Case 196960, 196962

        Case 196961
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TRANSPORTADORA_NAO_CADASTRADA", gErr, objTransportadora.iCodigo)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 175584)

    End Select

    Exit Sub

End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label13, Source, X, Y)
End Sub

Private Sub Label13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label13, Button, Shift, X, Y)
End Sub

Private Sub LabelNomeReduzido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelNomeReduzido, Source, X, Y)
End Sub

Private Sub LabelNomeReduzido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNomeReduzido, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub LabelCodigo_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelCodigo, Source, X, Y)
End Sub

Private Sub LabelCodigo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCodigo, Button, Shift, X, Y)
End Sub

Private Sub Label35_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label35, Source, X, Y)
End Sub

Private Sub Label35_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label35, Button, Shift, X, Y)
End Sub

Private Sub Label34_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label34, Source, X, Y)
End Sub

Private Sub Label34_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label34, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label56_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label56, Source, X, Y)
End Sub

Private Sub Label56_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label56, Button, Shift, X, Y)
End Sub

Private Sub TransportadoraLabel_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TransportadoraLabel, Source, X, Y)
End Sub

Private Sub TransportadoraLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TransportadoraLabel, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub

Private Sub Opcao_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, Opcao)
End Sub

Private Sub IEIsento_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub IENaoContrib_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Trata_IE()
    If Len(Trim(InscricaoEstadual.Text)) > 0 Then
        IEIsento.Value = vbUnchecked
        IEIsento.Enabled = False
        If InscricaoEstadual.Text <> sIEAnt Then
            IENaoContrib.Value = vbUnchecked
        End If
    Else
        If InscricaoEstadual.Text <> sIEAnt Then
            IEIsento.Value = vbChecked
            IENaoContrib.Value = vbChecked
        End If
        IEIsento.Enabled = True
    End If
    sIEAnt = InscricaoEstadual.Text
End Sub

Private Sub InscricaoEstadual_Validate(Cancel As Boolean)
    Call Trata_IE
End Sub
