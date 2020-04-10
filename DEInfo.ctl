VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl DEInfoOcx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.CommandButton BotaoNF 
      Caption         =   "Notas Fiscais Vinculadas"
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
      Left            =   5820
      TabIndex        =   12
      ToolTipText     =   "Lista de Fórmulas Utilizadas na Formação de Preço"
      Top             =   30
      Width           =   1380
   End
   Begin VB.Frame Frame2 
      Caption         =   "Embarque"
      Height          =   2100
      Left            =   3990
      TabIndex        =   29
      Top             =   3660
      Width           =   5400
      Begin VB.ComboBox TipoConhEmbarque 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "DEInfo.ctx":0000
         Left            =   975
         List            =   "DEInfo.ctx":0054
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   780
         Width           =   2535
      End
      Begin VB.ComboBox UFEmbarque 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4590
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   765
         Width           =   750
      End
      Begin VB.TextBox LocalEmbarque 
         Height          =   795
         Left            =   975
         MaxLength       =   250
         TabIndex        =   11
         Top             =   1215
         Width           =   4365
      End
      Begin MSMask.MaskEdBox NumConhEmbarque 
         Height          =   315
         Left            =   975
         TabIndex        =   7
         Top             =   360
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   50
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataConhEmbarque 
         Height          =   315
         Left            =   3975
         TabIndex        =   8
         Top             =   330
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataConhEmbarque 
         Height          =   300
         Left            =   5070
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   330
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label LabelTipoConhEmbarque 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo:"
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
         Left            =   450
         TabIndex        =   38
         Top             =   795
         Width           =   495
      End
      Begin VB.Label LabelDataConhEmbarque 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3480
         TabIndex        =   37
         Top             =   375
         Width           =   480
      End
      Begin VB.Label LabelLocalEmbarque 
         Alignment       =   1  'Right Justify
         Caption         =   "Local:"
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
         Left            =   105
         TabIndex        =   35
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label LabelUFEmbarque 
         Alignment       =   1  'Right Justify
         Caption         =   "UF:"
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
         Left            =   3960
         TabIndex        =   34
         Top             =   825
         Width           =   525
      End
      Begin VB.Label LabelNumConhEmbarque 
         Alignment       =   1  'Right Justify
         Caption         =   "Núm.Cnh.:"
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
         Left            =   75
         TabIndex        =   33
         Top             =   390
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação"
      Height          =   2985
      Left            =   45
      TabIndex        =   22
      Top             =   615
      Width           =   9345
      Begin VB.ComboBox CodPais 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2430
         Width           =   3270
      End
      Begin VB.ComboBox Natureza 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "DEInfo.ctx":014A
         Left            =   5805
         List            =   "DEInfo.ctx":0154
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1995
         Width           =   3435
      End
      Begin VB.ComboBox TipoDoc 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "DEInfo.ctx":0184
         Left            =   1245
         List            =   "DEInfo.ctx":0191
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1980
         Width           =   3255
      End
      Begin VB.TextBox Descricao 
         Height          =   1095
         Left            =   1245
         MaxLength       =   250
         TabIndex        =   3
         Top             =   735
         Width           =   7995
      End
      Begin MSMask.MaskEdBox Numero 
         Height          =   315
         Left            =   1245
         TabIndex        =   0
         Top             =   285
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   14
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Data 
         Height          =   315
         Left            =   4410
         TabIndex        =   1
         Top             =   270
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownData 
         Height          =   300
         Left            =   5580
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataAverbacao 
         Height          =   315
         Left            =   7845
         TabIndex        =   2
         Top             =   240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataAverbacao 
         Height          =   300
         Left            =   9000
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label LabelCodPais 
         Alignment       =   1  'Right Justify
         Caption         =   "País Destino:"
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
         Left            =   15
         TabIndex        =   32
         Top             =   2460
         Width           =   1185
      End
      Begin VB.Label LabelDataAverbacao 
         Alignment       =   1  'Right Justify
         Caption         =   "Averbação:"
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
         Left            =   6540
         TabIndex        =   31
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label LabelNatureza 
         Alignment       =   1  'Right Justify
         Caption         =   "Natureza:"
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
         Left            =   4560
         TabIndex        =   27
         Top             =   2010
         Width           =   1200
      End
      Begin VB.Label LabelTipoDoc 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo:"
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
         Left            =   105
         TabIndex        =   26
         Top             =   2010
         Width           =   1095
      End
      Begin VB.Label LabelDescricao 
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
         Left            =   240
         TabIndex        =   25
         Top             =   765
         Width           =   945
      End
      Begin VB.Label LabelData 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   3690
         TabIndex        =   24
         Top             =   300
         Width           =   690
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
         Left            =   90
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   23
         Top             =   315
         Width           =   1080
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Registros da Exportação"
      Height          =   2100
      Left            =   45
      TabIndex        =   18
      Top             =   3660
      Width           =   3885
      Begin MSMask.MaskEdBox NumRegistro 
         Height          =   225
         Left            =   1500
         TabIndex        =   19
         Top             =   690
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         MaxLength       =   12
         Mask            =   "############"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox DataRegistro 
         Height          =   225
         Left            =   315
         TabIndex        =   20
         Top             =   690
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridItens 
         Height          =   465
         Left            =   30
         TabIndex        =   21
         Top             =   210
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   820
         _Version        =   393216
         Rows            =   10
         Cols            =   4
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         FocusRect       =   2
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   7320
      ScaleHeight     =   450
      ScaleWidth      =   2025
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   30
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "DEInfo.ctx":01FE
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "DEInfo.ctx":0358
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "DEInfo.ctx":04E2
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "DEInfo.ctx":0A14
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
End
Attribute VB_Name = "DEInfoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoNumero As AdmEvento
Attribute objEventoNumero.VB_VarHelpID = -1

Private objGridItens As AdmGrid
Private iGrid_DataRegistro_Col As Integer
Private iGrid_NumRegistro_Col As Integer

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Declaração de Exportação"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "DEInfo"

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
    Set objGridItens = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216001)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long
Dim colCodigo As New Collection
Dim vCodigo As Variant
Dim objCodigoDescricao As AdmCodigoNome
Dim colCodigoDescricao As New AdmColCodigoNome

On Error GoTo Erro_Form_Load

    Set objEventoNumero = New AdmEvento
    
    Set objGridItens = New AdmGrid
    
    'Lê cada codigo da tabela Estados
    lErro = CF("Codigos_Le", "Estados", "Sigla", TIPO_STR, colCodigo, STRING_ESTADOS_SIGLA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Lê cada codigo e descricao da tabela Paises
    lErro = CF("Cod_Nomes_Le", "PaisesSISCOMEX", "Codigo", "Nome", STRING_PAISES_NOME, colCodigoDescricao)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'Preenche as ComboBox Estados com os objetos da colecao colCodigo
    UFEmbarque.AddItem ""
    For Each vCodigo In colCodigo
        UFEmbarque.AddItem vCodigo
    Next
    
    'Preenche cada ComboBox País com os objetos da colecao colCodigoDescricao
    For Each objCodigoDescricao In colCodigoDescricao
        CodPais.AddItem CStr(objCodigoDescricao.iCodigo) & SEPARADOR & objCodigoDescricao.sNome
        CodPais.ItemData(CodPais.NewIndex) = objCodigoDescricao.iCodigo
    Next
    
    TipoDoc.ListIndex = 0
    Natureza.ListIndex = 0
    CodPais.ListIndex = -1
    TipoConhEmbarque.ListIndex = -1
    UFEmbarque.ListIndex = -1


    lErro = Inicializa_GridItens(objGridItens)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216002)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objDEInfo As ClassDEInfo) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objDEInfo Is Nothing) Then

        lErro = Traz_DEInfo_Tela(objDEInfo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216003)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objDEInfo As ClassDEInfo) As Long

Dim lErro As Long, iIndice As Integer
Dim objRE As ClassDERegistro

On Error GoTo Erro_Move_Tela_Memoria

    objDEInfo.sNumero = Trim(numero.Text)
    objDEInfo.dtData = StrParaDate(Data.Text)
    objDEInfo.iFilialEmpresa = giFilialEmpresa
    objDEInfo.sDescricao = Descricao.Text
    objDEInfo.iTipoDoc = Codigo_Extrai(TipoDoc.Text)
    objDEInfo.iNatureza = Codigo_Extrai(Natureza.Text)
    objDEInfo.sNumConhEmbarque = Trim(NumConhEmbarque.Text)
    objDEInfo.sUFEmbarque = UFEmbarque.Text
    objDEInfo.sLocalEmbarque = Trim(LocalEmbarque.Text)
    objDEInfo.dtDataConhEmbarque = StrParaDate(DataConhEmbarque.Text)
    objDEInfo.iTipoConhEmbarque = Codigo_Extrai(TipoConhEmbarque.Text)
    objDEInfo.iCodPais = Codigo_Extrai(CodPais.Text)
    objDEInfo.dtDataAverbacao = StrParaDate(DataAverbacao.Text)

    For iIndice = 1 To objGridItens.iLinhasExistentes

        Set objRE = New ClassDERegistro

        objRE.sNumRegistro = Trim(GridItens.TextMatrix(iIndice, iGrid_NumRegistro_Col))
        objRE.dtDataRegistro = StrParaDate(GridItens.TextMatrix(iIndice, iGrid_DataRegistro_Col))

        If Len(Trim(objRE.sNumRegistro)) > 0 Then
        
            If objRE.dtDataRegistro = DATA_NULA Then gError 216004
        
            objDEInfo.colRE.Add objRE
        End If
        
    Next
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 216004
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_EMISSAO_NAO_PREENCHIDA_GRID", gErr, iIndice)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216005)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objDEInfo As New ClassDEInfo

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "DEInfo"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objDEInfo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Numero", objDEInfo.sNumero, STRING_MAXIMO, "Numero"

    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216006)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objDEInfo As New ClassDEInfo

On Error GoTo Erro_Tela_Preenche

    objDEInfo.sNumero = colCampoValor.Item("Numero").vValor

    If Len(Trim(objDEInfo.sNumero)) > 0 Then

        lErro = Traz_DEInfo_Tela(objDEInfo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216007)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objDEInfo As New ClassDEInfo

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(numero.Text)) = 0 Then gError 216008
    '#####################

    'Preenche o objDEInfo
    lErro = Move_Tela_Memoria(objDEInfo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = Trata_Alteracao(objDEInfo, objDEInfo.sNumero)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Grava o/a DEInfo no Banco de Dados
    lErro = CF("DEInfo_Grava", objDEInfo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 216008
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DEINFO_NAO_PREENCHIDO", gErr)
            numero.SetFocus

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216009)

    End Select

    Exit Function

End Function

Function Limpa_Tela_DEInfo() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_DEInfo

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)
    
    Call Grid_Limpa(objGridItens)
    
    TipoDoc.ListIndex = 0
    Natureza.ListIndex = 0
    CodPais.ListIndex = -1
    TipoConhEmbarque.ListIndex = -1
    UFEmbarque.ListIndex = -1

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Limpa_Tela_DEInfo = SUCESSO

    Exit Function

Erro_Limpa_Tela_DEInfo:

    Limpa_Tela_DEInfo = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216010)

    End Select

    Exit Function

End Function

Function Traz_DEInfo_Tela(objDEInfo As ClassDEInfo) As Long

Dim lErro As Long, iIndice As Integer
Dim objRE As ClassDERegistro

On Error GoTo Erro_Traz_DEInfo_Tela

    Call Limpa_Tela_DEInfo
    
    numero.PromptInclude = False
    numero.Text = objDEInfo.sNumero
    numero.PromptInclude = True

    'Lê o DEInfo que está sendo Passado
    lErro = CF("DEInfo_Le", objDEInfo)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM

    If lErro = SUCESSO Then

        If objDEInfo.dtData <> DATA_NULA Then
            Data.PromptInclude = False
            Data.Text = Format(objDEInfo.dtData, "dd/mm/yy")
            Data.PromptInclude = True
        End If

        Descricao.Text = objDEInfo.sDescricao
        
        Call Combo_Seleciona_ItemData(TipoDoc, objDEInfo.iTipoDoc)

        Call Combo_Seleciona_ItemData(Natureza, objDEInfo.iNatureza)

        NumConhEmbarque.Text = objDEInfo.sNumConhEmbarque
        
        If Len(Trim(objDEInfo.sUFEmbarque)) > 0 Then
            Call CF("SCombo_Seleciona2", UFEmbarque, objDEInfo.sUFEmbarque)
        End If
        
        LocalEmbarque.Text = objDEInfo.sLocalEmbarque

        If objDEInfo.dtDataConhEmbarque <> DATA_NULA Then
            DataConhEmbarque.PromptInclude = False
            DataConhEmbarque.Text = Format(objDEInfo.dtDataConhEmbarque, "dd/mm/yy")
            DataConhEmbarque.PromptInclude = True
        End If
        
        Call Combo_Seleciona_ItemData(TipoConhEmbarque, objDEInfo.iTipoConhEmbarque)

        Call Combo_Seleciona_ItemData(CodPais, objDEInfo.iCodPais)

        If objDEInfo.dtDataAverbacao <> DATA_NULA Then
            DataAverbacao.PromptInclude = False
            DataAverbacao.Text = Format(objDEInfo.dtDataAverbacao, "dd/mm/yy")
            DataAverbacao.PromptInclude = True
        End If
        
        iIndice = 0
        For Each objRE In objDEInfo.colRE
            iIndice = iIndice + 1
            GridItens.TextMatrix(iIndice, iGrid_NumRegistro_Col) = objRE.sNumRegistro
            GridItens.TextMatrix(iIndice, iGrid_DataRegistro_Col) = Format(objRE.dtDataRegistro, "dd/mm/yyyy")
        Next
        objGridItens.iLinhasExistentes = iIndice


    End If

    iAlterado = 0

    Traz_DEInfo_Tela = SUCESSO

    Exit Function

Erro_Traz_DEInfo_Tela:

    Traz_DEInfo_Tela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216011)

    End Select
    
    Resume Next

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    'Limpa Tela
    Call Limpa_Tela_DEInfo

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216012)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216013)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call Limpa_Tela_DEInfo

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216014)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objDEInfo As New ClassDEInfo
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(numero.Text)) = 0 Then gError 216015
    '#####################

    objDEInfo.sNumero = numero.Text

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_DEINFO", objDEInfo.sNumero)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("DEInfo_Exclui", objDEInfo)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'Limpa Tela
        Call Limpa_Tela_DEInfo

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 216015
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DEINFO_NAO_PREENCHIDO", gErr)
            numero.SetFocus

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216016)

    End Select

    Exit Sub

End Sub

Private Sub Numero_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Numero_Validate

    'Verifica se Numero está preenchida
    If Len(Trim(numero.Text)) <> 0 Then

       '#######################################
       'CRITICA Numero
       '#######################################

    End If

    Exit Sub

Erro_Numero_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216017)

    End Select

    Exit Sub

End Sub

Private Sub Numero_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Numero_GotFocus()
    Call MaskEdBox_TrataGotFocus(numero, iAlterado)
End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216018)

    End Select

    Exit Sub

End Sub

Private Sub UpDownData_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_UpClick

    Data.SetFocus

    If Len(Trim(Data.ClipText)) > 0 Then

        sData = Data.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216019)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)
    
End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) <> 0 Then

        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216020)

    End Select

    Exit Sub

End Sub

Private Sub Data_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Descricao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Descricao_Validate

    'Verifica se Descricao está preenchida
    If Len(Trim(Descricao.Text)) <> 0 Then

       '#######################################
       'CRITICA Descricao
       '#######################################

    End If

    Exit Sub

Erro_Descricao_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216021)

    End Select

    Exit Sub

End Sub

Private Sub Descricao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TipoDoc_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub Natureza_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub NumConhEmbarque_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_NumConhEmbarque_Validate

    'Verifica se NumConhEmbarque está preenchida
    If Len(Trim(NumConhEmbarque.Text)) <> 0 Then

       '#######################################
       'CRITICA NumConhEmbarque
       '#######################################

    End If

    Exit Sub

Erro_NumConhEmbarque_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216022)

    End Select

    Exit Sub

End Sub

Private Sub NumConhEmbarque_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UFEmbarque_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub LocalEmbarque_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_LocalEmbarque_Validate

    'Verifica se LocalEmbarque está preenchida
    If Len(Trim(LocalEmbarque.Text)) <> 0 Then

       '#######################################
       'CRITICA LocalEmbarque
       '#######################################

    End If

    Exit Sub

Erro_LocalEmbarque_Validate:

    Cancel = True

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216023)

    End Select

    Exit Sub

End Sub

Private Sub LocalEmbarque_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataConhEmbarque_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataConhEmbarque_DownClick

    DataConhEmbarque.SetFocus

    If Len(DataConhEmbarque.ClipText) > 0 Then

        sData = DataConhEmbarque.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataConhEmbarque.Text = sData

    End If

    Exit Sub

Erro_UpDownDataConhEmbarque_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216024)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataConhEmbarque_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataConhEmbarque_UpClick

    DataConhEmbarque.SetFocus

    If Len(Trim(DataConhEmbarque.ClipText)) > 0 Then

        sData = DataConhEmbarque.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataConhEmbarque.Text = sData

    End If

    Exit Sub

Erro_UpDownDataConhEmbarque_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216025)

    End Select

    Exit Sub

End Sub

Private Sub DataConhEmbarque_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataConhEmbarque, iAlterado)
    
End Sub

Private Sub DataConhEmbarque_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataConhEmbarque_Validate

    If Len(Trim(DataConhEmbarque.ClipText)) <> 0 Then

        lErro = Data_Critica(DataConhEmbarque.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_DataConhEmbarque_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216026)

    End Select

    Exit Sub

End Sub

Private Sub DataConhEmbarque_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub TipoConhEmbarque_Validate(Cancel As Boolean)

'Dim lErro As Long
'
'On Error GoTo Erro_TipoConhEmbarque_Validate
'
'    'Verifica se TipoConhEmbarque está preenchida
'    If Len(Trim(TipoConhEmbarque.Text)) <> 0 Then
'
'       'Critica a TipoConhEmbarque
'       lErro = Inteiro_Critica(TipoConhEmbarque.Text)
'       If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
'
'    End If
'
'    Exit Sub
'
'Erro_TipoConhEmbarque_Validate:
'
'    Cancel = True
'
'    Select Case gErr
'
'        Case ERRO_SEM_MENSAGEM
'
'        Case Else
'            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216027)
'
'    End Select
'
'    Exit Sub

End Sub

Private Sub TipoConhEmbarque_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub CodPais_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub UpDownDataAverbacao_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataAverbacao_DownClick

    DataAverbacao.SetFocus

    If Len(DataAverbacao.ClipText) > 0 Then

        sData = DataAverbacao.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataAverbacao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataAverbacao_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216028)

    End Select

    Exit Sub

End Sub

Private Sub UpDownDataAverbacao_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataAverbacao_UpClick

    DataAverbacao.SetFocus

    If Len(Trim(DataAverbacao.ClipText)) > 0 Then

        sData = DataAverbacao.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataAverbacao.Text = sData

    End If

    Exit Sub

Erro_UpDownDataAverbacao_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216029)

    End Select

    Exit Sub

End Sub

Private Sub DataAverbacao_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataAverbacao, iAlterado)
    
End Sub

Private Sub DataAverbacao_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataAverbacao_Validate

    If Len(Trim(DataAverbacao.ClipText)) <> 0 Then

        lErro = Data_Critica(DataAverbacao.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_DataAverbacao_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216030)

    End Select

    Exit Sub

End Sub

Private Sub DataAverbacao_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub objEventoNumero_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objDEInfo As ClassDEInfo

On Error GoTo Erro_objEventoNumero_evSelecao

    Set objDEInfo = obj1

    'Mostra os dados do DEInfo na tela
    lErro = Traz_DEInfo_Tela(objDEInfo)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Me.Show

    Exit Sub

Erro_objEventoNumero_evSelecao:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216031)

    End Select

    Exit Sub

End Sub

Private Sub LabelNumero_Click()

Dim lErro As Long
Dim objDEInfo As New ClassDEInfo
Dim colSelecao As New Collection

On Error GoTo Erro_LabelNumero_Click

    'Verifica se o Numero foi preenchido
    If Len(Trim(numero.Text)) <> 0 Then

        objDEInfo.sNumero = numero.Text

    End If

    Call Chama_Tela("DEInfoLista", colSelecao, objDEInfo, objEventoNumero)

    Exit Sub

Erro_LabelNumero_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216032)

    End Select

    Exit Sub

End Sub

Public Sub GridItens_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridItens, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Public Sub GridItens_EnterCell()
    Call Grid_Entrada_Celula(objGridItens, iAlterado)
End Sub

Public Sub GridItens_GotFocus()
    Call Grid_Recebe_Foco(objGridItens)
End Sub

Public Sub GridItens_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridItens, iExecutaEntradaCelula)

   If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridItens, iAlterado)
    End If

End Sub

Public Sub GridItens_LeaveCell()
    Call Saida_Celula(objGridItens)
End Sub

Public Sub GridItens_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridItens)
End Sub

Public Sub GridItens_RowColChange()
    Call Grid_RowColChange(objGridItens)
End Sub

Public Sub GridItens_Scroll()
    Call Grid_Scroll(objGridItens)
End Sub

Public Sub GridItens_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_Trata_Tecla1(KeyCode, objGridItens)
End Sub

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iCaminho As Integer)

Dim lErro As Long
Dim sNumReg As String

On Error GoTo Erro_Rotina_Grid_Enable

    sNumReg = Trim(GridItens.TextMatrix(iLinha, iGrid_NumRegistro_Col))

    Select Case objControl.Name

        Case NumRegistro.Name
            NumRegistro.Enabled = True

        Case DataRegistro.Name
            If sNumReg = "" Then
                objControl.Enabled = False
            Else
                objControl.Enabled = True
            End If
                                 
    End Select

    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216033)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_GridItens(objGridInt As AdmGrid) As Long
'Inicializa o Grid de Alocação

Dim iIndice As Integer

    Set objGridItens.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Data")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (NumRegistro.Name)
    objGridInt.colCampo.Add (DataRegistro.Name)

    'Colunas da Grid
    iGrid_NumRegistro_Col = 1
    iGrid_DataRegistro_Col = 2
 
    'Grid do GridInterno
    objGridInt.objGrid = GridItens

    'Largura da primeira coluna
    GridItens.ColWidth(0) = 200

    'Habilita a execução da Rotina_Grid_Enable
    objGridInt.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGridInt.objGrid.Rows = 100

    objGridInt.iLinhasVisiveis = 6
       
    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridItens)

    Inicializa_GridItens = SUCESSO

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'Faz a crítica da célula do grid que está deixando de ser a corrente

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
                
        If objGridInt.objGrid Is GridItens Then
        
            Select Case GridItens.Col
    
                Case iGrid_NumRegistro_Col
    
                    lErro = Saida_Celula_Padrao(objGridInt, NumRegistro)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                
                Case iGrid_DataRegistro_Col
    
                    lErro = Saida_Celula_Data(objGridInt, DataRegistro, True)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
             End Select
                
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 216034
        
        iAlterado = REGISTRO_ALTERADO

    End If
       
    Saida_Celula = SUCESSO
    
    Exit Function

Erro_Saida_Celula:
    
    Saida_Celula = gErr
    
    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case 216034
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
        
        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 216035)

    End Select

    Exit Function

End Function

Public Sub DataRegistro_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataRegistro_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub DataRegistro_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub DataRegistro_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = DataRegistro
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Public Sub NumRegistro_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub NumRegistro_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridItens)
End Sub

Public Sub NumRegistro_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridItens)
End Sub

Public Sub NumRegistro_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridItens.objControle = NumRegistro
    lErro = Grid_Campo_Libera_Foco(objGridItens)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub BotaoNF_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objDE As New ClassDEInfo

On Error GoTo Erro_BotaoNF_Click

    If Len(Trim(numero.Text)) = 0 Then gError 216043

    objDE.sNumero = numero.Text

    lErro = CF("DEInfo_Le", objDE)
    If lErro <> SUCESSO And lErro <> ERRO_LEITURA_SEM_DADOS Then gError ERRO_SEM_MENSAGEM
    
    If lErro = ERRO_LEITURA_SEM_DADOS Then gError 216044

    colSelecao.Add objDE.lNumIntDoc
    colSelecao.Add objDE.lNumIntDoc
    
    'O filtro é para pegar só os itens de NF ligados a DE sendo que
    'a associação pode estar a nível de item ou a nível de NF
    Call Chama_Tela("ItensNFiscalTodosSaida_Lista", colSelecao, Nothing, Nothing, "(NumIntDoc IN (SELECT NumIntDocItem FROM InfoAdicDocItem WHERE TipoDoc = 9 AND NumIntDE = ?) OR NumIntDoc IN (SELECT NumIntDocItem FROM InfoAdicDocItem AS X, ItensNFiscal AS I, InfoAdicExportacao AS Y WHERE X.TipoDoc = 9 AND X.NumIntDE = 0 AND I.NumIntDoc = X.NumIntDocItem AND Y.TipoDoc = 0 AND Y.NumIntDoc = I.NumIntNF AND Y.NumIntDE = ?))")
    
    Exit Sub

Erro_BotaoNF_Click:

    Select Case gErr
        
        Case 216043
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMERO_DEINFO_NAO_PREENCHIDO", gErr)
            numero.SetFocus
        
        Case 216044
            Call Rotina_Erro(vbOKOnly, "ERRO_DEINFO_NAO_CADASTRADO", gErr, objDE.sNumero)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 216045)

    End Select

    Exit Sub
    
End Sub
