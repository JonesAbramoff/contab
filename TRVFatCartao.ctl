VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl TRVFatCartao 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   KeyPreview      =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9510
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   8235
      ScaleHeight     =   495
      ScaleWidth      =   1140
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   30
      Width           =   1200
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "TRVFatCartao.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   630
         Picture         =   "TRVFatCartao.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Autorizações"
      Height          =   3465
      Left            =   30
      TabIndex        =   9
      Top             =   555
      Width           =   9450
      Begin VB.CommandButton BotaoMarcarTodos 
         Caption         =   "Marcar Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   60
         Picture         =   "TRVFatCartao.ctx":02D8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2835
         Width           =   1770
      End
      Begin VB.CommandButton BotaoDesmarcarTodos 
         Caption         =   "Desmarcar Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1995
         Picture         =   "TRVFatCartao.ctx":12F2
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2835
         Width           =   1770
      End
      Begin MSMask.MaskEdBox AutoVou 
         Height          =   255
         Left            =   165
         TabIndex        =   27
         Top             =   1125
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
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
      Begin VB.TextBox AutoNumParc 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   7920
         TabIndex        =   26
         Top             =   825
         Width           =   420
      End
      Begin VB.TextBox AutoValorL 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   7065
         TabIndex        =   25
         Top             =   855
         Width           =   915
      End
      Begin VB.TextBox AutoTarifa 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   6240
         TabIndex        =   24
         Top             =   825
         Width           =   735
      End
      Begin VB.TextBox AutoValorB 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   5400
         TabIndex        =   23
         Top             =   840
         Width           =   915
      End
      Begin VB.TextBox AutoNumero 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   3540
         TabIndex        =   21
         Top             =   825
         Width           =   720
      End
      Begin VB.TextBox AutoCliente 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   1905
         TabIndex        =   20
         Top             =   825
         Width           =   1620
      End
      Begin VB.TextBox AutoBandeira 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   900
         TabIndex        =   19
         Top             =   810
         Width           =   495
      End
      Begin VB.CheckBox AutoSelecionado 
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
         Left            =   390
         TabIndex        =   18
         Top             =   810
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid GridAuto 
         Height          =   405
         Left            =   30
         TabIndex        =   0
         Top             =   210
         Width           =   9390
         _ExtentX        =   16563
         _ExtentY        =   714
         _Version        =   393216
         Rows            =   15
         Cols            =   8
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
      End
      Begin MSMask.MaskEdBox AutoData 
         Height          =   255
         Left            =   4380
         TabIndex        =   22
         Top             =   825
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Vouchers"
      Height          =   1830
      Left            =   30
      TabIndex        =   8
      Top             =   4080
      Width           =   9450
      Begin VB.CommandButton BotaoVoucher 
         Caption         =   "Voucher ..."
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
         Left            =   60
         TabIndex        =   4
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox VouCliente 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   7095
         TabIndex        =   17
         Top             =   960
         Width           =   2325
      End
      Begin VB.TextBox VouPax 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   5355
         TabIndex        =   16
         Top             =   960
         Width           =   1515
      End
      Begin MSMask.MaskEdBox VouData 
         Height          =   255
         Left            =   2250
         TabIndex        =   10
         Top             =   960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         Enabled         =   0   'False
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox VouNumero 
         Height          =   255
         Left            =   1485
         TabIndex        =   11
         Top             =   960
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
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
      Begin MSMask.MaskEdBox VouSerie 
         Height          =   255
         Left            =   465
         TabIndex        =   12
         Top             =   960
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
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
      Begin MSMask.MaskEdBox VouTipo 
         Height          =   255
         Left            =   915
         TabIndex        =   13
         Top             =   960
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
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
      Begin MSFlexGridLib.MSFlexGrid GridVou 
         Height          =   1215
         Left            =   30
         TabIndex        =   3
         Top             =   225
         Width           =   9390
         _ExtentX        =   16563
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   15
         Cols            =   8
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         FocusRect       =   2
      End
      Begin MSMask.MaskEdBox VouValorComi 
         Height          =   255
         Left            =   4365
         TabIndex        =   14
         Top             =   960
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
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
      Begin MSMask.MaskEdBox VouValor 
         Height          =   255
         Left            =   3285
         TabIndex        =   15
         Top             =   960
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   0
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   15
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
   End
End
Attribute VB_Name = "TRVFatCartao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gcolAuto As Collection
Dim iLinhaAnt As Integer

Dim objGridAuto As AdmGrid
Dim iGrid_AutoSelecionado_Col As Integer
Dim iGrid_AutoBandeira_Col As Integer
Dim iGrid_AutoCliente_Col As Integer
Dim iGrid_AutoNumero_Col As Integer
Dim iGrid_AutoVou_Col As Integer
Dim iGrid_AutoData_Col As Integer
Dim iGrid_AutoValorB_Col As Integer
Dim iGrid_AutoTarifa_Col As Integer
Dim iGrid_AutoValorL_Col As Integer
Dim iGrid_AutoNumParc_Col As Integer

Dim objGridVou As AdmGrid
Dim iGrid_VouTipo_Col As Integer
Dim iGrid_VouSerie_Col As Integer
Dim iGrid_VouNumero_Col As Integer
Dim iGrid_VouData_Col As Integer
Dim iGrid_VouValor_Col As Integer
Dim iGrid_VouValorComi_Col As Integer
Dim iGrid_VouPax_Col As Integer
Dim iGrid_VouCliente_Col As Integer

'Variáveis globais
Dim iAlterado As Integer

'*** CARREGAMENTO DA TELA - INÍCIO ***
Public Sub Form_Load()

Dim lErro As Long
Dim colAuto As New Collection

On Error GoTo Erro_Form_Load
    
    iAlterado = 0
    
    Set objGridAuto = New AdmGrid
    Set objGridVou = New AdmGrid
    
    Call Inicializa_Grid_Auto(objGridAuto)
    Call Inicializa_Grid_Vou(objGridVou)
    
    DoEvents
    
    lErro = CF("TRVFatCartao_Le", colAuto)
    If lErro <> SUCESSO Then gError 200653
    
    lErro = Traz_Dados_Tela(colAuto)
    If lErro <> SUCESSO Then gError 200654
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    Select Case gErr
    
        Case 200653, 200654

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200655)

    End Select
    
    Exit Sub
    
End Sub

Public Function Trata_Parametros() As Long
'A tela não espera recebimento de parâmetros, portanto, essa função sempre retorna sucesso
    Trata_Parametros = SUCESSO
End Function
'*** CARREGAMENTO DA TELA - FIM ***

'*** FECHAMENTO DA TELA - INÍCIO ***
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set objGridAuto = Nothing
    Set objGridVou = Nothing

End Sub
'*** FECHAMENTO DA TELA - FIM ***

Private Sub BotaoLimpar_Click()
'Dispara a limpeza da tela

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    'limpa a tela
    Call Limpa_Tela_Fatura

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192674)

    End Select

    Exit Sub

End Sub

Private Sub BotaoDesmarcarTodos_Click()
    Call Grid_Marca_Desmarca(objGridAuto, iGrid_AutoSelecionado_Col, DESMARCADO)
End Sub

Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Private Sub Limpa_Tela_Fatura()
'Limpa a tela com exceção do campo 'Modelo'

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_Fatura

    'Limpa os controles básicos da tela
    Call Limpa_Tela(Me)
    Call Grid_Limpa(objGridAuto)
    Call Grid_Limpa(objGridVou)
    Call Ordenacao_Limpa(objGridAuto)
    
    Set gcolAuto = New Collection
    iLinhaAnt = 0
    
    iAlterado = 0

    Exit Sub

Erro_Limpa_Tela_Fatura:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200656)

    End Select
    
    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Faturamento - Cartão"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "TRVFatCartao"

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

Private Sub Motivo_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub BotaoMarcarTodos_Click()
    Call Grid_Marca_Desmarca(objGridAuto, iGrid_AutoSelecionado_Col, MARCADO)
End Sub

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

Public Property Let MousePointer(ByVal iTipo As Integer)
    Parent.MousePointer = iTipo
End Property

Public Property Get MousePointer() As Integer
    MousePointer = Parent.MousePointer
End Property
'**** fim do trecho a ser copiado *****

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 200657

    'Limpa Tela
'    Call Limpa_Tela_Fatura

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 200657

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200658)

    End Select

    Exit Sub

End Sub

Function Gravar_Registro() As Long

Dim lErro As Long
Dim colAuto As New Collection
Dim objVouAuto As ClassTRVVouAutoCC
Dim iLinha As Integer

On Error GoTo Erro_Gravar_Registro
 
    GL_objMDIForm.MousePointer = vbHourglass

    iLinha = 0
    For Each objVouAuto In gcolAuto
        iLinha = iLinha + 1
        If StrParaInt(GridAuto.TextMatrix(iLinha, iGrid_AutoSelecionado_Col)) = MARCADO Then
            colAuto.Add objVouAuto
        End If
    Next
    
    If colAuto.Count = 0 Then gError 200664
    
    lErro = CF("TRVFatCartao_Grava", colAuto)
    If lErro <> SUCESSO And lErro <> 200689 Then gError 200665
    
    If lErro = 200689 Then Call Rotina_Aviso(vbOKOnly, "AVISO_FAT_GERADA_SUCESSO_SEM_HTML")
    
    Call Limpa_Tela_Fatura
    
    Set colAuto = New Collection

    lErro = CF("TRVFatCartao_Le", colAuto)
    If lErro <> SUCESSO Then gError 200666
    
    lErro = Traz_Dados_Tela(colAuto)
    If lErro <> SUCESSO Then gError 200667

    GL_objMDIForm.MousePointer = vbDefault
       
    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr
    
        Case 200664
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)
    
        Case 200665 To 200667

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200659)

    End Select

    Exit Function

End Function

Private Function Inicializa_Grid_Vou(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("T")
    objGridInt.colColuna.Add ("S")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Valor")
    objGridInt.colColuna.Add ("CMCC")
    objGridInt.colColuna.Add ("Passageiro")
    objGridInt.colColuna.Add ("Cliente")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (VouTipo.Name)
    objGridInt.colCampo.Add (VouSerie.Name)
    objGridInt.colCampo.Add (VouNumero.Name)
    objGridInt.colCampo.Add (VouData.Name)
    objGridInt.colCampo.Add (VouValor.Name)
    objGridInt.colCampo.Add (VouValorComi.Name)
    objGridInt.colCampo.Add (VouPax.Name)
    objGridInt.colCampo.Add (VouCliente.Name)

    'Colunas do GridRepr
    iGrid_VouTipo_Col = 1
    iGrid_VouSerie_Col = 2
    iGrid_VouNumero_Col = 3
    iGrid_VouData_Col = 4
    iGrid_VouValor_Col = 5
    iGrid_VouValorComi_Col = 6
    iGrid_VouPax_Col = 7
    iGrid_VouCliente_Col = 8

    'Grid do GridInterno
    objGridInt.objGrid = GridVou

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 20
    
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 3

    'Largura da primeira coluna
    GridVou.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Vou = SUCESSO

    Exit Function

End Function

Private Function Inicializa_Grid_Auto(objGridInt As AdmGrid) As Long
'Inicializa o Grid

    'Form do Grid
    Set objGridInt.objForm = Me

    'Títulos das colunas
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add (" ")
    objGridInt.colColuna.Add ("Bnd")
    objGridInt.colColuna.Add ("Cliente")
    objGridInt.colColuna.Add ("Número")
    objGridInt.colColuna.Add ("Vouchers")
    objGridInt.colColuna.Add ("Data")
    objGridInt.colColuna.Add ("Bruto")
    objGridInt.colColuna.Add ("Tarifa")
    objGridInt.colColuna.Add ("Líquido")
    objGridInt.colColuna.Add ("Parc.")

    'Controles que participam do Grid
    objGridInt.colCampo.Add (AutoSelecionado.Name)
    objGridInt.colCampo.Add (AutoBandeira.Name)
    objGridInt.colCampo.Add (AutoCliente.Name)
    objGridInt.colCampo.Add (AutoNumero.Name)
    objGridInt.colCampo.Add (AutoVou.Name)
    objGridInt.colCampo.Add (AutoData.Name)
    objGridInt.colCampo.Add (AutoValorB.Name)
    objGridInt.colCampo.Add (AutoTarifa.Name)
    objGridInt.colCampo.Add (AutoValorL.Name)
    objGridInt.colCampo.Add (AutoNumParc.Name)

    'Colunas do GridRepr
    iGrid_AutoSelecionado_Col = 1
    iGrid_AutoBandeira_Col = 2
    iGrid_AutoCliente_Col = 3
    iGrid_AutoNumero_Col = 4
    iGrid_AutoVou_Col = 5
    iGrid_AutoData_Col = 6
    iGrid_AutoValorB_Col = 7
    iGrid_AutoTarifa_Col = 8
    iGrid_AutoValorL_Col = 9
    iGrid_AutoNumParc_Col = 10

    'Grid do GridInterno
    objGridInt.objGrid = GridAuto

    'Todas as linhas do grid
    objGridInt.objGrid.Rows = 100

    'Linhas visíveis do grid
    objGridInt.iLinhasVisiveis = 7
    
    objGridInt.iProibidoIncluir = GRID_PROIBIDO_INCLUIR
    objGridInt.iProibidoExcluir = GRID_PROIBIDO_EXCLUIR

    'Largura da primeira coluna
    GridAuto.ColWidth(0) = 400

    'Largura automática para as outras colunas
    objGridInt.iGridLargAuto = GRID_LARGURA_MANUAL

    'Chama função que inicializa o Grid
    Call Grid_Inicializa(objGridInt)

    Inicializa_Grid_Auto = SUCESSO

    Exit Function

End Function


Public Sub GridAuto_Click()

Dim iExecutaEntradaCelula As Integer
Dim colcolColecao As New Collection

    Call Grid_Click(objGridAuto, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAuto, iAlterado)
    End If
    
    colcolColecao.Add gcolAuto
    
    Call Ordenacao_ClickGrid(objGridAuto, , colcolColecao)

    
End Sub

Public Sub GridAuto_EnterCell()
    Call Grid_Entrada_Celula(objGridAuto, iAlterado)
End Sub

Public Sub GridAuto_GotFocus()
    Call Grid_Recebe_Foco(objGridAuto)
End Sub

Public Sub GridAuto_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridAuto)
    
End Sub

Public Sub GridAuto_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridAuto, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridAuto, iAlterado)
    End If
    
End Sub

Public Sub GridAuto_LeaveCell()
    Call Saida_Celula(objGridAuto)
End Sub

Public Sub GridAuto_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridAuto)
End Sub

Public Sub GridAuto_RowColChange()
    Call Grid_RowColChange(objGridAuto)
    Call Traz_Vou_Tela(GridAuto.Row)
End Sub

Public Sub GridAuto_Scroll()
    Call Grid_Scroll(objGridAuto)
End Sub

Public Sub GridVou_Click()

Dim iExecutaEntradaCelula As Integer

    Call Grid_Click(objGridVou, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridVou, iAlterado)
    End If
    
End Sub

Public Sub GridVou_EnterCell()
    Call Grid_Entrada_Celula(objGridVou, iAlterado)
End Sub

Public Sub GridVou_GotFocus()
    Call Grid_Recebe_Foco(objGridVou)
End Sub

Public Sub GridVou_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call Grid_Trata_Tecla1(KeyCode, objGridVou)
    
End Sub

Public Sub GridVou_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridVou, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridVou, iAlterado)
    End If
    
End Sub

Public Sub GridVou_LeaveCell()
    Call Saida_Celula(objGridVou)
End Sub

Public Sub GridVou_Validate(Cancel As Boolean)
    Call Grid_Libera_Foco(objGridVou)
End Sub

Public Sub GridVou_RowColChange()
    Call Grid_RowColChange(objGridVou)
End Sub

Public Sub GridVou_Scroll()
    Call Grid_Scroll(objGridVou)
End Sub

Function Traz_Dados_Tela(ByVal colAuto As Collection) As Long

Dim lErro As Long
Dim objAuto As ClassTRVVouAutoCC
Dim iLinha As Integer
Dim sNumVou As String
Dim objVou As ClassTRVFATInfoVoucher

On Error GoTo Erro_Traz_Dados_Tela

    If colAuto.Count >= objGridAuto.objGrid.Rows Then
        Call Refaz_Grid(objGridAuto, colAuto.Count)
    End If
    
    Set gcolAuto = colAuto

    Call Grid_Limpa(objGridAuto)
    iLinha = 0
    For Each objAuto In gcolAuto
    
        iLinha = iLinha + 1
        
        sNumVou = ""
        For Each objVou In objAuto.colVou
            If Len(Trim(sNumVou)) = 0 Then
                sNumVou = CStr(objVou.lNumVou)
            Else
                sNumVou = sNumVou & ";" & CStr(objVou.lNumVou)
            End If
        Next
        
        GridAuto.TextMatrix(iLinha, iGrid_AutoSelecionado_Col) = DESMARCADO
        GridAuto.TextMatrix(iLinha, iGrid_AutoBandeira_Col) = objAuto.sBandeira
        GridAuto.TextMatrix(iLinha, iGrid_AutoCliente_Col) = objAuto.lClienteFat & SEPARADOR & objAuto.sClienteFat
        GridAuto.TextMatrix(iLinha, iGrid_AutoVou_Col) = sNumVou
        'GridAuto.TextMatrix(iLinha, iGrid_AutoCC_Col) = right(objAuto.sNumCCred, 4)
        GridAuto.TextMatrix(iLinha, iGrid_AutoNumero_Col) = CStr(objAuto.sNumAuto)
        GridAuto.TextMatrix(iLinha, iGrid_AutoData_Col) = Format(objAuto.dtDataAutoCC, "dd/mm/yyyy")
        GridAuto.TextMatrix(iLinha, iGrid_AutoValorB_Col) = Format(objAuto.dValorB, "STANDARD")
        GridAuto.TextMatrix(iLinha, iGrid_AutoTarifa_Col) = Format(objAuto.dTarifa, "PERCENT")
        GridAuto.TextMatrix(iLinha, iGrid_AutoValorL_Col) = Format(objAuto.dValorL, "STANDARD")
        GridAuto.TextMatrix(iLinha, iGrid_AutoNumParc_Col) = CStr(objAuto.iQuantParc)
       
    Next
    
    objGridAuto.iLinhasExistentes = iLinha
    
    Call Grid_Refresh_Checkbox(objGridAuto)

    Traz_Dados_Tela = SUCESSO

    Exit Function

Erro_Traz_Dados_Tela:

    Traz_Dados_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200660)

    End Select

    Exit Function

End Function

Function Traz_Vou_Tela(ByVal iAuto As Integer) As Long

Dim lErro As Long
Dim objAuto As ClassTRVVouAutoCC
Dim objVou As ClassTRVFATInfoVoucher
Dim iLinha As Integer

On Error GoTo Erro_Traz_Vou_Tela

    If Not (gcolAuto Is Nothing) Then

        If iAuto <> 0 And iAuto <= gcolAuto.Count And iAuto <> iLinhaAnt Then
  
            Call Grid_Limpa(objGridVou)
            
            Set objAuto = gcolAuto.Item(iAuto)
            
            If objAuto.colVou.Count >= objGridVou.objGrid.Rows Then
                Call Refaz_Grid(objGridVou, objAuto.colVou.Count)
            End If
            
            iLinha = 0
            For Each objVou In objAuto.colVou
            
                iLinha = iLinha + 1
                
                GridVou.TextMatrix(iLinha, iGrid_VouTipo_Col) = objVou.sTipoVou
                GridVou.TextMatrix(iLinha, iGrid_VouSerie_Col) = objVou.sSerie
                GridVou.TextMatrix(iLinha, iGrid_VouNumero_Col) = CStr(objVou.lNumVou)
                GridVou.TextMatrix(iLinha, iGrid_VouData_Col) = Format(objVou.dtDataEmissao, "dd/mm/yyyy")
                GridVou.TextMatrix(iLinha, iGrid_VouValor_Col) = Format(objVou.dValor, "STANDARD")
                GridVou.TextMatrix(iLinha, iGrid_VouValorComi_Col) = Format(objVou.dValorComissao, "STANDARD")
                GridVou.TextMatrix(iLinha, iGrid_VouPax_Col) = objVou.sPassageiroNome & " " & objVou.sPassageiroSobreNome
                GridVou.TextMatrix(iLinha, iGrid_VouCliente_Col) = objVou.lCliente & SEPARADOR & objVou.sNomeCliVou
    
            Next
            
            objGridVou.iLinhasExistentes = iLinha
        
        End If
    
        iLinhaAnt = iAuto
    
    End If

    Traz_Vou_Tela = SUCESSO

    Exit Function

Erro_Traz_Vou_Tela:

    Traz_Vou_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200661)

    End Select

    Exit Function

End Function

Public Function Saida_Celula(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula

    'aquii está devolvendo erro em vez de sucesso
    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro <> SUCESSO Then gError 200662
    
    End If
    
    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 200662

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 200663)

    End Select

    Exit Function

End Function

Public Sub AutoSelecionado_GotFocus()
    Call Grid_Campo_Recebe_Foco(objGridAuto)
End Sub

Public Sub AutoSelecionado_KeyPress(KeyAscii As Integer)
    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridAuto)
End Sub

Public Sub AutoSelecionado_Validate(Cancel As Boolean)
Dim lErro As Long
    Set objGridAuto.objControle = AutoSelecionado
    lErro = Grid_Campo_Libera_Foco(objGridAuto)
    If lErro <> SUCESSO Then Cancel = True
End Sub

Private Sub BotaoVoucher_Click()

Dim objVoucher As New ClassTRVVouchers

On Error GoTo Erro_BotaoVoucher_Click

    If GridVou.Row = 0 Then gError 192875

    objVoucher.lNumVou = StrParaLong(GridVou.TextMatrix(GridVou.Row, iGrid_VouNumero_Col))
    objVoucher.sSerie = GridVou.TextMatrix(GridVou.Row, iGrid_VouSerie_Col)
    objVoucher.sTipoDoc = TRV_TIPODOC_VOU_TEXTO
    objVoucher.sTipVou = GridVou.TextMatrix(GridVou.Row, iGrid_VouTipo_Col)

    Call Chama_Tela("TRVVoucher", objVoucher)

    Exit Sub

Erro_BotaoVoucher_Click:

    Select Case gErr
    
        Case 192875
            Call Rotina_Erro(vbOKOnly, "ERRO_LINHA_GRID_NAO_SELECIONADA", gErr)

        Case Else
             Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 192876)

    End Select

    Exit Sub
    
End Sub

Sub Refaz_Grid(ByVal objGridInt As AdmGrid, ByVal iNumLinhas As Integer)
    objGridInt.objGrid.Rows = iNumLinhas + 1
    'Chama rotina de Inicialização do Grid
    Call Grid_Inicializa(objGridInt)
End Sub
