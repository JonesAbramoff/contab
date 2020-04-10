VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TRVAcordos 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9300
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleMode       =   0  'User
   ScaleWidth      =   9510
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5010
      Index           =   2
      Left            =   210
      TabIndex        =   20
      Top             =   810
      Visible         =   0   'False
      Width           =   8850
      Begin VB.CommandButton BotaoProdutos 
         Caption         =   "Produtos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   15
         TabIndex        =   12
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Frame Frame5 
         Caption         =   "Percentual de comissão por Produto\Destino"
         Height          =   4320
         Left            =   0
         TabIndex        =   32
         Top             =   120
         Width           =   8790
         Begin VB.ComboBox Destino 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2220
            TabIndex        =   36
            Top             =   1980
            Width           =   2040
         End
         Begin VB.TextBox DescricaoProduto 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   4215
            MaxLength       =   250
            TabIndex        =   34
            Top             =   1590
            Width           =   2745
         End
         Begin MSMask.MaskEdBox Produto 
            Height          =   315
            Left            =   780
            TabIndex        =   35
            Top             =   1575
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            AllowPrompt     =   -1  'True
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox PercComiss 
            Height          =   315
            Left            =   7005
            TabIndex        =   33
            Top             =   1590
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   0
            PromptInclude   =   0   'False
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0%"
            PromptChar      =   " "
         End
         Begin MSFlexGridLib.MSFlexGrid GridComissao 
            Height          =   1200
            Left            =   75
            TabIndex        =   11
            Top             =   300
            Width           =   8610
            _ExtentX        =   15187
            _ExtentY        =   2117
            _Version        =   393216
            Cols            =   8
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483640
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            FocusRect       =   2
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5040
      Index           =   1
      Left            =   165
      TabIndex        =   19
      Top             =   825
      Width           =   8910
      Begin VB.Frame Frame3 
         Caption         =   "Detalhamento"
         Height          =   2700
         Left            =   120
         TabIndex        =   29
         Top             =   2280
         Width           =   8640
         Begin VB.TextBox Descricao 
            Height          =   1050
            Left            =   1260
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   300
            Width           =   7185
         End
         Begin VB.TextBox Observacao 
            Height          =   1050
            Left            =   1275
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   1560
            Width           =   7185
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   31
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Index           =   8
            Left            =   135
            TabIndex        =   30
            Top             =   1560
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Identificação"
         Height          =   1305
         Left            =   90
         TabIndex        =   24
         Top             =   90
         Width           =   8670
         Begin VB.ComboBox Filial 
            Height          =   315
            Left            =   5520
            TabIndex        =   4
            Top             =   810
            Width           =   2085
         End
         Begin VB.CommandButton BotaoProxNum 
            Height          =   285
            Left            =   2115
            Picture         =   "TRVAcordos.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Numeração Automática"
            Top             =   330
            Width           =   300
         End
         Begin MSMask.MaskEdBox Numero 
            Height          =   315
            Left            =   1230
            TabIndex        =   0
            Top             =   315
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   6
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Cliente 
            Height          =   300
            Left            =   1245
            TabIndex        =   3
            Top             =   795
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox Contrato 
            Height          =   300
            Left            =   5520
            TabIndex        =   2
            Top             =   315
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   529
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   20
            PromptChar      =   " "
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Número do Contrato:"
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
            Index           =   1
            Left            =   3690
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   28
            Top             =   345
            Width           =   1770
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Filial:"
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
            Index           =   13
            Left            =   4995
            TabIndex        =   27
            Top             =   870
            Width           =   480
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
            Left            =   510
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   26
            Top             =   840
            Width           =   660
         End
         Begin VB.Label LabelNumero 
            Alignment       =   1  'Right Justify
            Caption         =   "Número:"
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
            Height          =   315
            Left            =   360
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   25
            Top             =   345
            Width           =   810
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Validade"
         Height          =   705
         Left            =   105
         TabIndex        =   21
         Top             =   1425
         Width           =   8655
         Begin MSComCtl2.UpDown UpDownDe 
            Height          =   300
            Left            =   2370
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   255
            Width           =   225
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox DataValidadeDe 
            Height          =   315
            Left            =   1245
            TabIndex        =   5
            Top             =   240
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DataValidadeAte 
            Height          =   300
            Left            =   5505
            TabIndex        =   7
            Top             =   240
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   8
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/##"
            PromptChar      =   " "
         End
         Begin MSComCtl2.UpDown UpDownAte 
            Height          =   300
            Left            =   6675
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label1 
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
            Index           =   4
            Left            =   5100
            TabIndex        =   23
            Top             =   285
            Width           =   360
         End
         Begin VB.Label Label1 
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
            Index           =   3
            Left            =   855
            TabIndex        =   22
            Top             =   285
            Width           =   315
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   6975
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   120
         Picture         =   "TRVAcordos.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "TRVAcordos.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1110
         Picture         =   "TRVAcordos.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1620
         Picture         =   "TRVAcordos.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5460
      Left            =   105
      TabIndex        =   18
      Top             =   450
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   9631
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Inicial"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comissão"
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
Attribute VB_Name = "TRVAcordos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

'Variáveis Globais
Dim iFrameAtual As Integer
Dim iAlterado As Integer

Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1
Private WithEvents objEventoCliente As AdmEvento
Attribute objEventoCliente.VB_VarHelpID = -1
Private WithEvents objEventoProduto As AdmEvento
Attribute objEventoProduto.VB_VarHelpID = -1

Dim objGridComissao As AdmGrid

Dim iGrid_Produto_Col As Integer
Dim iGrid_DescricaoProduto_Col As Integer
Dim iGrid_PercComiss_Col As Integer
Dim iGrid_Destino_Col As Integer

Public Sub BotaoProxNum_Click()

Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click
    
    lErro = CF("Config_ObterAutomatico", "FATConfig", "NUM_PROX_TRVACORDO", "TRVAcordos", "Numero", lCodigo)
    If lErro <> SUCESSO Then gError 197034
    
    Numero.PromptInclude = False
    Numero.Text = CStr(lCodigo)
    Numero.PromptInclude = True

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr
        
        Case 197034

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197035)
    
    End Select

    Exit Sub
    
End Sub

Private Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Numero_Validate

    'Verifica se Numero está preenchido
    If Len(Trim(Numero.Text)) <> 0 Then

       'Critica o Numero
       lErro = Long_Critica(Numero.Text)
       If lErro <> SUCESSO Then gError 197011

    End If

    Exit Sub

Erro_Numero_Validate:

    Cancel = True

    Select Case gErr

        Case 197011

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197012)

    End Select

    Exit Sub

End Sub

Private Sub Numero_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Numero, iAlterado)
    
End Sub

Private Sub Numero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LabelNumero_Click()

Dim lErro As Long
Dim objTRVAcordo As New ClassTRVAcordos
Dim colSelecao As New Collection

On Error GoTo Erro_LabelNumero_Click

    'Verifica se o Numero foi preenchido
    If Len(Trim(Numero.Text)) <> 0 Then

        objTRVAcordo.lNumero = Numero.Text

    End If

    Call Chama_Tela("TRVAcordosLista", colSelecao, objTRVAcordo, objEventoNumero)

    Exit Sub

Erro_LabelNumero_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197010)

    End Select

    Exit Sub

End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objTRVAcordo As ClassTRVAcordos

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objTRVAcordo = obj1

    'Mostra os dados do TRVAporte na tela
    lErro = Traz_TRVAcordo_Tela(objTRVAcordo)
    If lErro <> SUCESSO Then gError 197008

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case 197008

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197009)

    End Select

    Exit Sub

End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Acordo"
    Call Form_Load

End Function

Public Function Name() As String
    Name = "TRVAcordos"
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property
'**** fim do trecho a ser copiado *****

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Activate()

    'Carrega os índices da tela
    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub

Sub Form_Unload(Cancel As Integer)

Dim lErro As Long

On Error GoTo Erro_Form_Unload

    Set objEventoNumero = Nothing
    Set objEventoCliente = Nothing
    Set objEventoProduto = Nothing
    
    Set objGridComissao = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197067)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    iAlterado = 0

    Set objEventoNumero = New AdmEvento
    Set objEventoCliente = New AdmEvento
    Set objEventoProduto = New AdmEvento
    
    Set objGridComissao = New AdmGrid
    
    'Carrega a combo Destino
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_DESTINO_VIAGEM, Destino)
    If lErro <> SUCESSO Then gError 197036

    lErro = Inicializa_Grid_Comissao(objGridComissao)
    If lErro <> SUCESSO Then gError 197037
           
    iFrameAtual = 1

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 197036, 197037, 197068, 197158

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197038)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Private Sub Contrato_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cliente_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Cliente_GotFocus()
    Call MaskEdBox_TrataGotFocus(Cliente, iAlterado)
End Sub

Private Sub Cliente_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome

On Error GoTo Erro_Cliente_Validate

    'Verifica se o Cliente está preenchido
    If Len(Trim(Cliente.Text)) > 0 Then

        'Busca o Cliente no BD
        lErro = TP_Cliente_Le(Cliente, objcliente, iCodFilial)
        If lErro <> SUCESSO Then gError 197013
                   
        lErro = CF("FiliaisClientes_Le_Cliente", objcliente, colCodigoNome)
        If lErro <> SUCESSO Then gError 197014

        'Preenche ComboBox de Filiais
        Call CF("Filial_Preenche", Filial, colCodigoNome)
        
        If iCodFilial = 0 Then iCodFilial = FILIAL_MATRIZ

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, iCodFilial)

    'Se não estiver preenchido
    ElseIf Len(Trim(Cliente.Text)) = 0 Then

        'Limpa a Combo de Filiais
        Filial.Clear

    End If

    Exit Sub

Erro_Cliente_Validate:

    Cancel = True

    Select Case gErr

        Case 197013, 197014

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197015)

    End Select

    Exit Sub

End Sub

Private Sub Filial_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Filial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
Dim objFilialCliente As New ClassFilialCliente
Dim sCliente As String
Dim vbMsgRes As VbMsgBoxResult
Dim objcliente As New ClassCliente

On Error GoTo Erro_Filial_Validate

    'Verifica se a filial foi preenchida ou alterada
    If Len(Trim(Filial.Text)) = 0 Then Exit Sub

    'Verifica se é uma filial selecionada
    If Filial.Text = Filial.List(Filial.ListIndex) Then Exit Sub

    'Tenta selecionar na combo
    lErro = Combo_Seleciona(Filial, iCodigo)
    If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 197016

    'Se não encontrou o CÓDIGO
    If lErro = 6730 Then

        'Verifica se o cliente foi digitado
        If Len(Trim(Cliente.Text)) = 0 Then gError 197017

        sCliente = Cliente.Text
        objFilialCliente.iCodFilial = iCodigo

        'Pesquisa se existe Filial com o código extraído
        lErro = CF("FilialCliente_Le_NomeRed_CodFilial", sCliente, objFilialCliente)
        If lErro <> SUCESSO And lErro <> 17660 Then gError 197018

        If lErro = 17660 Then

            'Lê o Cliente
            objcliente.sNomeReduzido = sCliente
            lErro = CF("Cliente_Le_NomeReduzido", objcliente)
            If lErro <> SUCESSO And lErro <> 12348 Then gError 197019

            'Se encontrou o Cliente
            If lErro = SUCESSO Then
                
                objFilialCliente.lCodCliente = objcliente.lCodigo

                gError 197020
            
            End If
            
        End If
        
        If iCodigo <> 0 Then
        
            'Coloca na tela a Filial lida
            Filial.Text = iCodigo & SEPARADOR & objFilialCliente.sNome
        
        Else
            
            objcliente.lCodigo = 0
            objFilialCliente.iCodFilial = 0
            
        End If
        
    'Não encontrou a STRING
    ElseIf lErro = 6731 Then
        
        'trecho incluido por Leo em 17/04/02
        objcliente.sNomeReduzido = Cliente.Text
        
        'Lê o Cliente
        lErro = CF("Cliente_Le_NomeReduzido", objcliente)
        If lErro <> SUCESSO And lErro <> 12348 Then gError 197021
        
        If lErro = SUCESSO Then gError 197022
        
    End If

    Exit Sub

Erro_Filial_Validate:

    Cancel = True

    Select Case gErr

        Case 197016, 197018, 197019, 197021

        Case 197017
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)
        
        Case 197020
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_FILIALCLIENTE", iCodigo, Cliente.Text)

            If vbMsgRes = vbYes Then
                Call Chama_Tela("FiliaisClientes", objFilialCliente)
            End If

        Case 197022
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIALCLIENTE_NAO_ENCONTRADA", gErr, Filial.Text)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197023)

    End Select

    Exit Sub

End Sub

Public Sub LabelCliente_Click()

Dim objcliente As New ClassCliente
Dim colSelecao As New Collection

    'Prenche o Nome Reduzido do Cliente com o Cliente da Tela
    objcliente.sNomeReduzido = Cliente.Text

    Call Chama_Tela("ClientesLista", colSelecao, objcliente, objEventoCliente)

End Sub

Private Sub objEventoCliente_evSelecao(obj1 As Object)

Dim objcliente As ClassCliente
Dim bCancel As Boolean

    Set objcliente = obj1

    'Preenche campo Cliente
    Cliente.Text = objcliente.sNomeReduzido

    'Executa o Validate
    Call Cliente_Validate(bCancel)

    Me.Show

    Exit Sub

End Sub

Private Sub DataValidadeDe_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataValidadeDe_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataValidadeDe, iAlterado)
    
End Sub

Private Sub DataValidadeDe_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataValidadeDe_Validate

    If Len(Trim(DataValidadeDe.ClipText)) <> 0 Then

        lErro = Data_Critica(DataValidadeDe.Text)
        If lErro <> SUCESSO Then gError 190716

    End If
    
    Exit Sub

Erro_DataValidadeDe_Validate:

    Cancel = True

    Select Case gErr

        Case 190716

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 190717)

    End Select

    Exit Sub

End Sub

Private Sub DataValidadeAte_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub DataValidadeAte_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataValidadeAte, iAlterado)
    
End Sub

Private Sub DataValidadeAte_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataValidadeAte_Validate

    If Len(Trim(DataValidadeAte.ClipText)) <> 0 Then

        lErro = Data_Critica(DataValidadeAte.Text)
        If lErro <> SUCESSO Then gError 197024

    End If
    
    Exit Sub

Erro_DataValidadeAte_Validate:

    Cancel = True

    Select Case gErr

        Case 197024

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197025)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDe_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDe_DownClick

    DataValidadeDe.SetFocus

    If Len(DataValidadeDe.ClipText) > 0 Then

        sData = DataValidadeDe.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 197026

        DataValidadeDe.Text = sData

    End If

    Exit Sub

Erro_UpDownDe_DownClick:

    Select Case gErr

        Case 197026

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197027)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDe_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDe_UpClick

    DataValidadeDe.SetFocus

    If Len(Trim(DataValidadeDe.ClipText)) > 0 Then

        sData = DataValidadeDe.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 197028

        DataValidadeDe.Text = sData

    End If

    Exit Sub

Erro_UpDownDe_UpClick:

    Select Case gErr

        Case 197028

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197029)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAte_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownAte_DownClick

    DataValidadeAte.SetFocus

    If Len(DataValidadeAte.ClipText) > 0 Then

        sData = DataValidadeAte.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 197030

        DataValidadeAte.Text = sData

    End If

    Exit Sub

Erro_UpDownAte_DownClick:

    Select Case gErr

        Case 197030

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197031)

    End Select

    Exit Sub

End Sub

Private Sub UpDownAte_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownAte_UpClick

    DataValidadeAte.SetFocus

    If Len(Trim(DataValidadeAte.ClipText)) > 0 Then

        sData = DataValidadeAte.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError 197032

        DataValidadeAte.Text = sData

    End If

    Exit Sub

Erro_UpDownAte_UpClick:

    Select Case gErr

        Case 197032

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197033)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Observacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub


Private Function Inicializa_Grid_Comissao(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Produto")
    objGridInt.colColuna.Add ("Descrição")
    objGridInt.colColuna.Add ("Destino")
    objGridInt.colColuna.Add ("Comissão")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (Produto.Name)
    objGridInt.colCampo.Add (DescricaoProduto.Name)
    objGridInt.colCampo.Add (Destino.Name)
    objGridInt.colCampo.Add (PercComiss.Name)

    'Colunas do Grid
    iGrid_Produto_Col = 1
    iGrid_DescricaoProduto_Col = 2
    iGrid_Destino_Col = 3
    iGrid_PercComiss_Col = 4

    'Grid do GridInterno
    objGridInt.objGrid = GridComissao

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 901

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 10

    'Largura da primeira coluna
    GridComissao.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_AUTOMATICA

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Comissao = SUCESSO

    Exit Function

End Function

Function Trata_Parametros(Optional objTRVAcordo As ClassTRVAcordos) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objTRVAcordo Is Nothing) Then

        lErro = Traz_TRVAcordo_Tela(objTRVAcordo)
        If lErro <> SUCESSO Then gError 197044

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 197044

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197045)

    End Select

    iAlterado = 0

    Exit Function

End Function

Public Sub GridComissao_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridComissao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComissao, iAlterado)
    End If
    
End Sub

Public Sub GridComissao_EnterCell()
    Call Grid_Entrada_Celula(objGridComissao, iAlterado)
End Sub

Public Sub GridComissao_GotFocus()
    Call Grid_Recebe_Foco(objGridComissao)
End Sub

Public Sub GridComissao_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridComissao)
    
End Sub

Public Sub GridComissao_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridComissao, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridComissao, iAlterado)
    End If
    
End Sub

Public Sub GridComissao_LeaveCell()
    Call Saida_Celula(objGridComissao)
End Sub

Public Sub GridComissao_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridComissao)
End Sub

Public Sub GridComissao_RowColChange()
    Call Grid_RowColChange(objGridComissao)
End Sub

Public Sub GridComissao_Scroll()
    Call Grid_Scroll(objGridComissao)
End Sub

Public Sub Produto_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Produto_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridComissao)
End Sub

Public Sub Produto_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissao)
End Sub

Public Sub Produto_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissao.objControle = Produto
    lErro = Grid_Campo_Libera_Foco(objGridComissao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub Destino_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Destino_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridComissao)
End Sub

Public Sub Destino_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissao)
End Sub

Public Sub Destino_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissao.objControle = Destino
    lErro = Grid_Campo_Libera_Foco(objGridComissao)
    If lErro <> SUCESSO Then Cancel = True
    
End Sub

Public Sub PercComiss_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub PercComiss_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridComissao)
End Sub

Public Sub PercComiss_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridComissao)
End Sub

Public Sub PercComiss_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridComissao.objControle = PercComiss
    
    lErro = Grid_Campo_Libera_Foco(objGridComissao)
    If lErro <> SUCESSO Then Cancel = True
    
    
End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'aquii está devolvendo erro em vez de sucesso
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        'Verifica qual o Grid em questão
        Select Case objGridInt.objGrid.Name
    
            'Se for o GridComissao
            
            Case GridComissao.Name
    
                lErro = Saida_Celula_GridComissao(objGridInt)
                If lErro <> SUCESSO Then gError 197046
    
    
        End Select
    
        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 197172
    
    End If
    
    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 197046, 197047, 197172

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197048)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_GridComissao(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_GridComissao

    'Verifica qual a coluna atual do Grid
    Select Case objGridInt.objGrid.Col

        'Se for a de Produto
        Case iGrid_Produto_Col
            lErro = Saida_Celula_Produto(objGridInt)
            If lErro <> SUCESSO Then gError 197049

        Case iGrid_Destino_Col
            lErro = Saida_Celula_Destino(objGridInt)
            If lErro <> SUCESSO Then gError 197050

        Case iGrid_PercComiss_Col
            lErro = Saida_Celula_PercComiss(objGridInt, PercComiss)
            If lErro <> SUCESSO Then gError 197051

    End Select

    Saida_Celula_GridComissao = SUCESSO

    Exit Function

Erro_Saida_Celula_GridComissao:

    Saida_Celula_GridComissao = gErr

    Select Case gErr

        Case 197049 To 197051

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197052)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Produto(objGridInt As AdmGrid) As Long
'Faz a crítica da célula Produto que está deixando de ser a corrente

Dim lErro As Long
Dim objProduto As New ClassProduto
Dim iIndice As Integer
Dim iProdutoPreenchido As Integer
Dim sProdutoEnxuto As String
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_Saida_Celula_Produto

    Set objGridInt.objControle = Produto

    If Len(Trim(Produto.ClipText)) > 0 Then

        'Critica o Produto
        lErro = CF("Produto_Critica_Filial2", Produto.Text, objProduto, iProdutoPreenchido)
        If lErro <> SUCESSO And lErro <> 51381 And lErro <> 86295 Then gError 197053
        
        'Se o produto é gerencial ==> erro
        If lErro = 86295 Then gError 197054
               
        'Se o produto não foi encontrado ==> Pergunta se deseja criar
        If lErro = 51381 Then gError 197055

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 197056
    
        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True
        
'
'        'Verifica se já está em outra linha do Grid
'        For iIndice = 1 To objGridInt.iLinhasExistentes
'            If iIndice <> GridComissao.Row Then
'                If GridComissao.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text Then gError 197077
'            End If
'        Next

        'Verifica se já está em outra linha do Grid
        For iIndice = 1 To objGridInt.iLinhasExistentes
            If iIndice <> GridComissao.Row Then
                If GridComissao.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text And _
                GridComissao.TextMatrix(GridComissao.Row, iGrid_Destino_Col) = GridComissao.TextMatrix(iIndice, iGrid_Destino_Col) And _
                Len(Trim(GridComissao.TextMatrix(GridComissao.Row, iGrid_Destino_Col))) > 0 Then gError 197057
            End If
        Next
    
        If GridComissao.Row - GridComissao.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
        GridComissao.TextMatrix(GridComissao.Row, iGrid_DescricaoProduto_Col) = objProduto.sDescricao

    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 197058

    Saida_Celula_Produto = SUCESSO

    Exit Function

Erro_Saida_Celula_Produto:

    Saida_Celula_Produto = gErr

    Select Case gErr

        Case 197053, 197058
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 197054
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_GERENCIAL", gErr, objProduto.sCodigo)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 197055
            vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CRIAR_PRODUTO", Produto.Text)
            If vbMsgRes = vbYes Then
            
                Call Grid_Trata_Erro_Saida_Celula_Chama_Tela(objGridInt)
                
                Call Chama_Tela("Produto", objProduto)
            Else
                Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            End If

        Case 197056
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
                   
        Case 197057
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_JA_PREENCHIDO_LINHA_GRID", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197059)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Destino(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iCodigo As Integer
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula_Destino

    Set objGridInt.objControle = Destino

    'Verifica se o Destino preenchido
    If Len(Trim(Destino.Text)) > 0 Then

        'Verifica se ele foi selecionado
        If Destino.Text <> Destino.List(Destino.ListIndex) Then

            'Seleciona o Tipo de Cobrança
            lErro = Combo_Seleciona(Destino, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 197060

            If lErro = 6730 Then gError 197061
            If lErro = 6731 Then gError 197062

        End If

        'Verifica se já está em outra linha do Grid
        For iIndice = 1 To objGridInt.iLinhasExistentes
            If iIndice <> GridComissao.Row Then
                If GridComissao.TextMatrix(iIndice, iGrid_Destino_Col) = Destino.Text And _
                GridComissao.TextMatrix(GridComissao.Row, iGrid_Produto_Col) = GridComissao.TextMatrix(iIndice, iGrid_Produto_Col) And _
                Len(Trim(GridComissao.TextMatrix(GridComissao.Row, iGrid_Produto_Col))) > 0 Then gError 197065
            End If
        Next

        'Acrescenta uma linha no Grid se for o caso
        If GridComissao.Row - GridComissao.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If

    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 197063

    Saida_Celula_Destino = SUCESSO

    Exit Function

Erro_Saida_Celula_Destino:

    Saida_Celula_Destino = gErr

    Select Case gErr

        Case 197060, 197063
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 197061, 197062
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DESTINO_NAO_CADASTRADO", gErr, Destino.Text)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case 197065
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODDEST_JA_PREENCHIDO_LINHA_GRID", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197064)

    End Select

    Exit Function

End Function

Function Saida_Celula_PercComiss(objGridInt As AdmGrid, ByVal objControle As Object) As Long
'Faz a crítica da célula Percentual Comissao que está deixando de ser a corrente

Dim lErro As Long
Dim dPercentDesc As Double

On Error GoTo Erro_Saida_Celula_PercComiss

    Set objGridInt.objControle = objControle

    If Len(objControle.Text) > 0 Then
        
        'Critica a porcentagem
        lErro = Porcentagem_Critica(objControle.Text)
        If lErro <> SUCESSO Then gError 197066

        dPercentDesc = CDbl(objControle.Text)
        
        objControle.Text = Format(dPercentDesc, "Fixed")
        
        If objGridInt.objGrid.Row - objGridInt.objGrid.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If
    
    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 197067

    Saida_Celula_PercComiss = SUCESSO

    Exit Function

Erro_Saida_Celula_PercComiss:

    Saida_Celula_PercComiss = gErr

    Select Case gErr

        Case 197066, 197067
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197068)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 197088

    'Limpa Tela
    Call Limpa_Tela_TRVAcordos

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 197088

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197089)

    End Select

    Exit Sub

End Sub

Sub BotaoFechar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoFechar_Click

    Unload Me

    Exit Sub

Erro_BotaoFechar_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197090)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 197091

    Call Limpa_Tela_TRVAcordos

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 197091

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197092)

    End Select

    Exit Sub

End Sub

Function Move_Tela_Memoria(objTRVAcordo As ClassTRVAcordos) As Long

Dim lErro As Long
Dim objcliente As New ClassCliente
Dim iLinha As Integer
Dim objTRVAcordoDif As ClassTRVAcordoTarifaDif
Dim objTRVAcordoComiss As ClassTRVAcordoComissao
Dim sProduto As String
Dim sProduto1 As String
Dim iPreenchido As Integer

On Error GoTo Erro_Move_Tela_Memoria

    objcliente.sNomeReduzido = Cliente.Text

    'Lê o Cliente através do Nome Reduzido
    lErro = CF("Cliente_Le_NomeReduzido", objcliente)
    If lErro <> SUCESSO And lErro <> 12348 Then gError 197093

    objTRVAcordo.lNumero = StrParaLong(Numero.Text)
    objTRVAcordo.sContrato = Contrato.Text
    objTRVAcordo.lCliente = objcliente.lCodigo
    objTRVAcordo.iFilialCliente = Codigo_Extrai(Filial.Text)
    objTRVAcordo.dtValidadeDe = StrParaDate(DataValidadeDe.Text)
    objTRVAcordo.dtValidadeAte = StrParaDate(DataValidadeAte.Text)
    objTRVAcordo.sDescricao = Descricao.Text
    objTRVAcordo.sObservacao = Observacao.Text
    
    For iLinha = 1 To objGridComissao.iLinhasExistentes
    
        Set objTRVAcordoComiss = New ClassTRVAcordoComissao
    
        sProduto1 = GridComissao.TextMatrix(iLinha, iGrid_Produto_Col)
        
        'Formata o produto
        lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
        If lErro <> SUCESSO Then gError 197094

        objTRVAcordoComiss.sProduto = sProduto
        
        objTRVAcordoComiss.iDestino = Codigo_Extrai(GridComissao.TextMatrix(iLinha, iGrid_Destino_Col))
        objTRVAcordoComiss.dPercComissao = PercentParaDbl(GridComissao.TextMatrix(iLinha, iGrid_PercComiss_Col))
        
        objTRVAcordo.colTRVAcordoComiss.Add objTRVAcordoComiss
    
    Next
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 197093 To 197095
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197096)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objTRVAcordo As New ClassTRVAcordos
Dim iLinha As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Numero.Text)) = 0 Then gError 197097
    If Len(Trim(Cliente.Text)) = 0 Then gError 197098
    If Codigo_Extrai(Filial.Text) = 0 Then gError 197099
    
    If StrParaDate(DataValidadeDe.Text) = DATA_NULA Then gError 197100
    If StrParaDate(DataValidadeAte.Text) = DATA_NULA Then gError 197101
    
    If StrParaDate(DataValidadeDe.Text) > StrParaDate(DataValidadeAte.Text) Then gError 197102
    
    For iLinha = 1 To objGridComissao.iLinhasExistentes
    
        If Len(Trim(GridComissao.TextMatrix(iLinha, iGrid_Produto_Col))) = 0 Then gError 197103
        'If PercentParaDbl(GridComissao.TextMatrix(iLinha, iGrid_PercComiss_Col)) = 0 Then gError 197105
        If Len(Trim(GridComissao.TextMatrix(iLinha, iGrid_Destino_Col))) = 0 Then gError 197104
    
    Next

    'Preenche o objTRVAporte
    lErro = Move_Tela_Memoria(objTRVAcordo)
    If lErro <> SUCESSO Then gError 197109

    lErro = Trata_Alteracao(objTRVAcordo, objTRVAcordo.lNumero)
    If lErro <> SUCESSO Then gError 197110

    'Grava c/a TRVAcordo no Banco de Dados
    lErro = CF("TRVAcordo_Grava", objTRVAcordo)
    If lErro <> SUCESSO Then gError 197111
    
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 197097
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_PREENCHIDO", gErr)

        Case 197098
            Call Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_NAO_PREENCHIDO", gErr)

        Case 197099
            Call Rotina_Erro(vbOKOnly, "ERRO_FILIAL_NAO_PREENCHIDA", gErr)
            
        Case 197100
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_NAO_PREENCHIDA", gErr)

        Case 197101
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAFINAL_NAO_PREENCHIDA", gErr)

        Case 197102
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)
        
        Case 197103
            Call Rotina_Erro(vbOKOnly, "ERRO_PRODUTO_NAO_PREENCHIDO_GRID_COMISS", gErr, iLinha)
        
        Case 197104
            Call Rotina_Erro(vbOKOnly, "ERRO_DESTINO_NAO_PREENCHIDO_GRID_COMISS", gErr, iLinha)
        
        Case 197105
            Call Rotina_Erro(vbOKOnly, "ERRO_COMIS_NAO_PREENCHIDO_GRID_COMISS", gErr, iLinha)

        Case 197109 To 197111
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197112)

    End Select

    Exit Function

End Function

Function Limpa_Tela_TRVAcordos() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_TRVAcordos

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Call Grid_Limpa(objGridComissao)
    
    Filial.Clear
    
    iAlterado = 0

    Limpa_Tela_TRVAcordos = SUCESSO

    Exit Function

Erro_Limpa_Tela_TRVAcordos:

    Limpa_Tela_TRVAcordos = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197113)

    End Select

    Exit Function

End Function

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objTRVAcordo As New ClassTRVAcordos
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Numero.Text)) = 0 Then gError 197136

    objTRVAcordo.lNumero = StrParaLong(Numero.Text)

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_TRVACORDOS", objTRVAcordo.lNumero)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("TRVAcordo_Exclui", objTRVAcordo)
        If lErro <> SUCESSO Then gError 197137

        'Limpa Tela
        Call Limpa_Tela_TRVAcordos

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 197136
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_NAO_PREENCHIDO", gErr)

        Case 197137

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197138)

    End Select

    Exit Sub

End Sub

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "TRVAcordos"

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Numero", StrParaLong(Numero.Text), 0, "Numero"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197156)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objTRVAcordo As New ClassTRVAcordos

On Error GoTo Erro_Tela_Preenche

    objTRVAcordo.lNumero = colCampoValor.Item("Numero").vValor

    If objTRVAcordo.lNumero <> 0 Then
    
        lErro = Traz_TRVAcordo_Tela(objTRVAcordo)
        If lErro <> SUCESSO Then gError 197157
        
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 197157

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197158)

    End Select

    Exit Function

End Function

Function Traz_TRVAcordo_Tela(objTRVAcordo As ClassTRVAcordos) As Long

Dim lErro As Long
Dim objTRVAcordoComiss As ClassTRVAcordoComissao
Dim objTRVAcordoDif As ClassTRVAcordoTarifaDif
Dim iLinha As Integer
Dim iIndice As Integer
Dim sProdutoEnxuto As String
Dim objProduto As New ClassProduto
Dim iCodigo As Integer

On Error GoTo Erro_Traz_TRVAcordo_Tela

    Call Limpa_Tela_TRVAcordos
    
    If objTRVAcordo.lNumero <> 0 Then
        Numero.PromptInclude = False
        Numero.Text = objTRVAcordo.lNumero
        Numero.PromptInclude = True
    End If

    'Lê o TRVAporte que está sendo Passado
    lErro = CF("TRVAcordo_Le", objTRVAcordo)
    If lErro <> SUCESSO And lErro <> 197148 Then gError 197159
    
    If lErro = SUCESSO Then
        
        Contrato.Text = objTRVAcordo.sContrato
        Cliente.Text = CStr(objTRVAcordo.lCliente)
        Call Cliente_Validate(bSGECancelDummy)

        'Seleciona filial na Combo Filial
        Call CF("Filial_Seleciona", Filial, objTRVAcordo.iFilialCliente)

        If objTRVAcordo.dtValidadeDe <> DATA_NULA Then
            DataValidadeDe.PromptInclude = False
            DataValidadeDe.Text = Format(objTRVAcordo.dtValidadeDe, "dd/mm/yy")
            DataValidadeDe.PromptInclude = True
        End If

        If objTRVAcordo.dtValidadeAte <> DATA_NULA Then
            DataValidadeAte.PromptInclude = False
            DataValidadeAte.Text = Format(objTRVAcordo.dtValidadeAte, "dd/mm/yy")
            DataValidadeAte.PromptInclude = True
        End If

        Descricao.Text = objTRVAcordo.sDescricao
        Observacao.Text = objTRVAcordo.sObservacao
        
        iIndice = 0
        
        For Each objTRVAcordoComiss In objTRVAcordo.colTRVAcordoComiss
        
            iIndice = iIndice + 1
                    
            lErro = Mascara_RetornaProdutoEnxuto(objTRVAcordoComiss.sProduto, sProdutoEnxuto)
            If lErro <> SUCESSO Then gError 197160

            Produto.PromptInclude = False
            Produto.Text = sProdutoEnxuto
            Produto.PromptInclude = True
    
            objProduto.sCodigo = objTRVAcordoComiss.sProduto
            
            'Lê o Produto
            lErro = CF("Produto_Le", objProduto)
            If lErro <> SUCESSO And lErro <> 28030 Then gError 197161
            
            GridComissao.TextMatrix(iIndice, iGrid_Produto_Col) = Produto.Text
            GridComissao.TextMatrix(iIndice, iGrid_DescricaoProduto_Col) = objProduto.sDescricao
            
            Destino.Text = objTRVAcordoComiss.iDestino

            'Seleciona o Tipo de Cobrança
            lErro = Combo_Seleciona(Destino, iCodigo)
            If lErro <> SUCESSO And lErro <> 6730 And lErro <> 6731 Then gError 197162

            GridComissao.TextMatrix(iIndice, iGrid_Destino_Col) = Destino.Text
            GridComissao.TextMatrix(iIndice, iGrid_PercComiss_Col) = Format(objTRVAcordoComiss.dPercComissao, "Percent")
            
        Next

        objGridComissao.iLinhasExistentes = objTRVAcordo.colTRVAcordoComiss.Count

    End If

    iAlterado = 0

    Traz_TRVAcordo_Tela = SUCESSO

    Exit Function

Erro_Traz_TRVAcordo_Tela:

    Traz_TRVAcordo_Tela = gErr

    Select Case gErr

        Case 197159, 197161, 197162, 197164, 195165

        Case 197160, 197163
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197166)

    End Select

    Exit Function

End Function

Private Sub TabStrip1_BeforeClick(Cancel As Integer)
    Call TabStrip_TrataBeforeClick(Cancel, TabStrip1)
End Sub

Private Sub TabStrip1_Click()

Dim lErro As Long
Dim iLinha As Integer
Dim iFrameAnterior

On Error GoTo Erro_TabStrip1_Click

    'Se frame selecionado não for o atual esconde o frame atual, mostra o novo.
    If TabStrip1.SelectedItem.Index = iFrameAtual Then Exit Sub

    If TabStrip_PodeTrocarTab(iFrameAtual, TabStrip1, Me) <> SUCESSO Then Exit Sub

    'Torna Frame correspondente ao Tab selecionado visivel
    Frame1(TabStrip1.SelectedItem.Index).Visible = True
    'Torna Frame atual invisivel
    Frame1(iFrameAtual).Visible = False
    'Armazena novo valor de iFrameAtual
    iFrameAtual = TabStrip1.SelectedItem.Index

    Exit Sub

Erro_TabStrip1_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 190634)

    End Select

    Exit Sub

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Cliente Then Call LabelCliente_Click
        If Me.ActiveControl Is Numero Then Call LabelNumero_Click
        If Me.ActiveControl Is Produto Then Call BotaoProdutos_Click
    
    End If
    
End Sub

Private Sub Label1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(Label1(Index), Source, X, Y)
End Sub
Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1(Index), Button, Shift, X, Y)
End Sub
Private Sub LabelNumero_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelNumero, Source, X, Y)
End Sub
Private Sub LabelNumero_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelNumero, Button, Shift, X, Y)
End Sub
Private Sub LabelCliente_DragDrop(Source As Control, X As Single, Y As Single)
    Call Controle_DragDrop(LabelCliente, Source, X, Y)
End Sub
Private Sub LabelCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelCliente, Button, Shift, X, Y)
End Sub

Sub BotaoProdutos_Click()

Dim lErro As Long
Dim sProduto As String
Dim iPreenchido As Integer
Dim objProduto As New ClassProduto
Dim colSelecao As Collection
Dim sProduto1 As String

On Error GoTo Erro_BotaoProdutos_Click

    If Me.ActiveControl Is Produto Then
    
        sProduto1 = Produto.Text
        
    Else
    
        'Verifica se tem alguma linha selecionada no Grid
        If GridComissao.Row = 0 Then gError 197160

        sProduto1 = GridComissao.TextMatrix(GridComissao.Row, iGrid_Produto_Col)
        
    End If
    
    lErro = CF("Produto_Formata", sProduto1, sProduto, iPreenchido)
    If lErro <> SUCESSO Then gError 197161
    
    If iPreenchido <> PRODUTO_PREENCHIDO Then sProduto = ""

    objProduto.sCodigo = sProduto

    'Chama a Tela ProdutoVendaLista
    Call Chama_Tela("ProdutoVendaLista", colSelecao, objProduto, objEventoProduto)

    Exit Sub
    
Erro_BotaoProdutos_Click:

    Select Case gErr
    
        Case 197160
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
        
        Case 197161
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197162)
    
    End Select
    
    Exit Sub

End Sub

Private Sub objEventoProduto_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objProduto As ClassProduto
Dim sProdutoEnxuto As String
Dim iIndice As Integer

On Error GoTo Erro_objEventoProduto_evSelecao

    'verifica se tem alguma linha do Grid selecionada
    If GridComissao.Row > 0 Then

        Set objProduto = obj1

        lErro = Mascara_RetornaProdutoEnxuto(objProduto.sCodigo, sProdutoEnxuto)
        If lErro <> SUCESSO Then gError 197163

        Produto.PromptInclude = False
        Produto.Text = sProdutoEnxuto
        Produto.PromptInclude = True
        
        'Lê o Produto
        lErro = CF("Produto_Le", objProduto)
        If lErro <> SUCESSO And lErro <> 28030 Then gError 197164
        
        If Not (Me.ActiveControl Is Produto) Then
    
            'Preenche o Grid
            GridComissao.TextMatrix(GridComissao.Row, iGrid_Produto_Col) = Produto.Text
            GridComissao.TextMatrix(GridComissao.Row, iGrid_DescricaoProduto_Col) = objProduto.sDescricao
    
            If GridComissao.Row - GridComissao.FixedRows = objGridComissao.iLinhasExistentes Then
        
                objGridComissao.iLinhasExistentes = objGridComissao.iLinhasExistentes + 1
        
            End If
    
    
        End If
        

    End If

    Me.Show

    Exit Sub

Erro_objEventoProduto_evSelecao:

    Select Case gErr

        Case 197163
            Call Rotina_Erro(vbOKOnly, "ERRO_MASCARA_RETORNAPRODUTOENXUTO", gErr, objProduto.sCodigo)
        
        Case 197164
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197165)

    End Select

    Exit Sub

End Sub
