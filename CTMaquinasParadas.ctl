VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl CTMaquinasParadas 
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8340
   KeyPreview      =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   8340
   Begin VB.TextBox Observacao 
      Height          =   315
      Left            =   2250
      MaxLength       =   255
      TabIndex        =   25
      Top             =   3825
      Width           =   5895
   End
   Begin VB.ComboBox Tipo 
      Height          =   315
      Left            =   2250
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   2490
      Width           =   2910
   End
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   3135
      Picture         =   "CTMaquinasParadas.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Numeração Automática"
      Top             =   1575
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   510
      Left            =   6015
      ScaleHeight     =   450
      ScaleWidth      =   2025
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "CTMaquinasParadas.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "CTMaquinasParadas.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "CTMaquinasParadas.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   45
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "CTMaquinasParadas.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   45
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Data 
      Height          =   315
      Left            =   2250
      TabIndex        =   7
      Top             =   2040
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownData 
      Height          =   300
      Left            =   3615
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2040
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSMask.MaskEdBox Horas 
      Height          =   315
      Left            =   2250
      TabIndex        =   11
      Top             =   2940
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   8
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox QtdMaquinas 
      Height          =   315
      Left            =   2250
      TabIndex        =   13
      Top             =   3390
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   2
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   2250
      TabIndex        =   21
      Top             =   1575
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   6
      Mask            =   "######"
      PromptChar      =   " "
   End
   Begin VB.Label HorasDia 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   7020
      TabIndex        =   28
      Top             =   2925
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Horas Disponível no dia:"
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
      Left            =   4635
      TabIndex        =   27
      Top             =   2955
      Width           =   2325
   End
   Begin VB.Label LabelObservacao 
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
      Left            =   1125
      TabIndex        =   26
      Top             =   3855
      Width           =   1140
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Qtde Disponível no dia:"
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
      Left            =   4710
      TabIndex        =   24
      Top             =   3405
      Width           =   2250
   End
   Begin VB.Label QtdDia 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   7020
      TabIndex        =   23
      Top             =   3390
      Width           =   885
   End
   Begin VB.Label DescricaoCTPadrao 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4305
      TabIndex        =   19
      Top             =   660
      Width           =   3825
   End
   Begin VB.Label CTLabel 
      Caption         =   "Centro de Trabalho:"
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
      Left            =   465
      TabIndex        =   18
      Top             =   690
      Width           =   1770
   End
   Begin VB.Label DescMaquina 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4305
      TabIndex        =   17
      Top             =   1110
      Width           =   3825
   End
   Begin VB.Label LabelMaquina 
      Alignment       =   1  'Right Justify
      Caption         =   "Máquina:"
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
      Left            =   1335
      TabIndex        =   16
      Top             =   1140
      Width           =   900
   End
   Begin VB.Label Maquina 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2250
      TabIndex        =   15
      Top             =   1110
      Width           =   2025
   End
   Begin VB.Label CodigoCT 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2250
      TabIndex        =   14
      Top             =   660
      Width           =   2025
   End
   Begin VB.Label LabelCodigo 
      Alignment       =   1  'Right Justify
      Caption         =   "Codigo:"
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
      Left            =   735
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   6
      Top             =   1605
      Width           =   1500
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
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   735
      TabIndex        =   9
      Top             =   2055
      Width           =   1500
   End
   Begin VB.Label LabelTipo 
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
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   750
      TabIndex        =   10
      Top             =   2535
      Width           =   1500
   End
   Begin VB.Label LabelHoras 
      Alignment       =   1  'Right Justify
      Caption         =   "Horas Paradas:"
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
      Left            =   735
      TabIndex        =   12
      Top             =   2955
      Width           =   1500
   End
   Begin VB.Label LabelQtdMaquinas 
      Alignment       =   1  'Right Justify
      Caption         =   "Qtde Máquinas Paradas:"
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
      Left            =   60
      TabIndex        =   5
      Top             =   3420
      Width           =   2175
   End
End
Attribute VB_Name = "CTMaquinasParadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoCodigo As AdmEvento
Attribute objEventoCodigo.VB_VarHelpID = -1

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Paradas não programadas das Máquinas"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "CTMaquinasParadas"

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

    Set objEventoCodigo = Nothing
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156084)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoCodigo = New AdmEvento
    
    Horas.Format = FORMATO_ESTOQUE
    
    Call Carrega_Tipo

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156085)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objCTMaquinasParadas As ClassCTMaquinasParadas) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objCTMaquinasParadas Is Nothing) Then

        lErro = Traz_CTMaquinasParadas_Tela(objCTMaquinasParadas)
        If lErro <> SUCESSO Then gError 140724

        Data.Text = Format(gdtDataAtual, "dd/mm/yy")
        Call Data_Validate(bSGECancelDummy)
    
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 140724

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156086)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objCTMaquinasParadas As ClassCTMaquinasParadas) As Long

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objMaquinas As ClassMaquinas

On Error GoTo Erro_Move_Tela_Memoria

    objCTMaquinasParadas.lCodigo = StrParaLong(Codigo.Text)
    objCTMaquinasParadas.iFilialEmpresa = giFilialEmpresa
    If Len(Trim(Data.ClipText)) <> 0 Then objCTMaquinasParadas.dtData = Format(Data.Text, Data.Format)
    objCTMaquinasParadas.iTipo = Tipo.ItemData(Tipo.ListIndex)
    objCTMaquinasParadas.dHoras = StrParaDbl(Horas.Text)
    objCTMaquinasParadas.iQtdMaquinas = StrParaInt(QtdMaquinas.Text)
    
    lErro = Move_CTMaquina_Memoria(objCentrodeTrabalho, objMaquinas)
    If lErro <> SUCESSO Then gError 140978
    
    objCTMaquinasParadas.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
    objCTMaquinasParadas.lNumIntDocMaq = objMaquinas.lNumIntDoc
 
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 140978

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156087)

    End Select

    Exit Function

End Function

Function Move_CTMaquina_Memoria(objCentrodeTrabalho As ClassCentrodeTrabalho, objMaquinas As ClassMaquinas) As Long

Dim lErro As Long

On Error GoTo Erro_Move_CTMaquina_Memoria
    
    Set objCentrodeTrabalho = New ClassCentrodeTrabalho
    Set objMaquinas = New ClassMaquinas
    
    If Len(Trim(CodigoCT.Caption)) <> 0 Then
    
        objCentrodeTrabalho.sNomeReduzido = CodigoCT.Caption
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 140976
    
    End If
    
    If Len(Maquina.Caption) > 0 Then
                
        objMaquinas.sNomeReduzido = Maquina.Caption
        
        'Le a Máquina no BD a partir do NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 140977
                
    End If

    Move_CTMaquina_Memoria = SUCESSO

    Exit Function

Erro_Move_CTMaquina_Memoria:

    Move_CTMaquina_Memoria = gErr

    Select Case gErr
    
        Case 140976 To 140977

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156088)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objCTMaquinasParadas As New ClassCTMaquinasParadas

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "CTMaquinasParadas"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objCTMaquinasParadas)
    If lErro <> SUCESSO Then gError 140725

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objCTMaquinasParadas.lCodigo, 0, "Codigo"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, giFilialEmpresa

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 140725

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156089)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objCTMaquinasParadas As New ClassCTMaquinasParadas

On Error GoTo Erro_Tela_Preenche

    objCTMaquinasParadas.lCodigo = colCampoValor.Item("Codigo").vValor

    objCTMaquinasParadas.iFilialEmpresa = giFilialEmpresa

    If objCTMaquinasParadas.lCodigo <> 0 And objCTMaquinasParadas.iFilialEmpresa <> 0 Then
        lErro = Traz_CTMaquinasParadas_Tela(objCTMaquinasParadas)
        If lErro <> SUCESSO Then gError 140726
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 140726

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156090)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim vbMsgBox As VbMsgBoxResult
Dim objCTMaquinasParadas As New ClassCTMaquinasParadas

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 140727
    If StrParaDate(Data.Text) = DATA_NULA Then gError 140972
    If Tipo.ListIndex = -1 Then gError 140973
    If StrParaDbl(Horas.Text) = 0 Then gError 140974
    If StrParaInt(QtdMaquinas.Text) = 0 Then gError 140975
    
    If StrParaDbl(Horas.Text) - StrParaDbl(HorasDia.Caption) > QTDE_ESTOQUE_DELTA Then
        vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_HORAS_DISP_MENOR_HORAS_CAD", Horas.Text, HorasDia.Caption)
        If vbMsgBox = vbNo Then gError 140981
    End If
    
    If StrParaInt(QtdMaquinas.Text) > StrParaInt(QtdDia.Caption) Then
        vbMsgBox = Rotina_Aviso(vbYesNo, "AVISO_QTD_DISP_MENOR_QTD_CAD", QtdMaquinas.Text, QtdDia.Caption)
        If vbMsgBox = vbNo Then gError 140982
    End If
    '#####################

    'Preenche o objCTMaquinasParadas
    lErro = Move_Tela_Memoria(objCTMaquinasParadas)
    If lErro <> SUCESSO Then gError 140728

    lErro = Trata_Alteracao(objCTMaquinasParadas, objCTMaquinasParadas.lCodigo, objCTMaquinasParadas.iFilialEmpresa)
    If lErro <> SUCESSO Then gError 140729

    'Grava o/a CTMaquinasParadas no Banco de Dados
    lErro = CF("CTMaquinasParadas_Grava", objCTMaquinasParadas)
    If lErro <> SUCESSO Then gError 140730

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 140727
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CTMAQUINASPARADAS_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 140728, 140729, 140730

        Case 140972
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
            Data.SetFocus

        Case 140973
            Call Rotina_Erro(vbOKOnly, "ERRO_TIPO_DOCUMENTO_NAO_PREENCHIDO", gErr)
            Tipo.SetFocus

        Case 140974
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
            Horas.SetFocus

        Case 140975
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDO1", gErr)
            QtdMaquinas.SetFocus
            
        Case 140981
            Horas.SetFocus
            
        Case 140982
            QtdMaquinas.SetFocus
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156091)

    End Select

    Exit Function

End Function

Function Limpa_Tela_CTMaquinasParadas() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_CTMaquinasParadas

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    Data.Text = Format(gdtDataAtual, "dd/mm/yy")
    Call Data_Validate(bSGECancelDummy)
    
    QtdDia.Caption = ""

    iAlterado = 0

    Limpa_Tela_CTMaquinasParadas = SUCESSO

    Exit Function

Erro_Limpa_Tela_CTMaquinasParadas:

    Limpa_Tela_CTMaquinasParadas = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156092)

    End Select

    Exit Function

End Function

Function Traz_CTMaquinasParadas_Tela(objCTMaquinasParadas As ClassCTMaquinasParadas) As Long

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objMaquinas As ClassMaquinas

On Error GoTo Erro_Traz_CTMaquinasParadas_Tela

    If objCTMaquinasParadas.lNumIntDocCT <> 0 Then
        
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        objCentrodeTrabalho.lNumIntDoc = objCTMaquinasParadas.lNumIntDocCT
        
        lErro = CF("CentroDeTrabalho_Le_NumIntDoc", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134590 Then gError 140969
        
        CodigoCT.Caption = objCentrodeTrabalho.sNomeReduzido
        DescricaoCTPadrao.Caption = objCentrodeTrabalho.sDescricao
    
    End If

    If objCTMaquinasParadas.lNumIntDocMaq <> 0 Then

        Set objMaquinas = New ClassMaquinas
        
        objMaquinas.lNumIntDoc = objCTMaquinasParadas.lNumIntDocMaq
        
        lErro = CF("Maquinas_Le_NumIntDoc", objMaquinas)
        If lErro <> SUCESSO And lErro <> 106353 Then gError 140970
        
        Maquina.Caption = objMaquinas.sNomeReduzido
        DescMaquina.Caption = objMaquinas.sDescricao
        
    End If
    
    'Lê o CTMaquinasParadas que está sendo Passado
    lErro = CF("CTMaquinasParadas_Le", objCTMaquinasParadas)
    If lErro <> SUCESSO And lErro <> 140704 Then gError 140731

    If lErro = SUCESSO Then

        If objCTMaquinasParadas.lCodigo <> 0 Then Codigo.Text = CStr(objCTMaquinasParadas.lCodigo)

        If objCTMaquinasParadas.iTipo <> 0 Then Call Combo_Seleciona_ItemData(Tipo, objCTMaquinasParadas.iTipo)
        If objCTMaquinasParadas.dHoras <> 0 Then Horas.Text = Formata_Estoque(objCTMaquinasParadas.dHoras)
        If objCTMaquinasParadas.iQtdMaquinas <> 0 Then QtdMaquinas.Text = CStr(objCTMaquinasParadas.iQtdMaquinas)

        If objCTMaquinasParadas.dtData <> DATA_NULA Then
            Data.PromptInclude = False
            Data.Text = Format(objCTMaquinasParadas.dtData, "dd/mm/yy")
            Data.PromptInclude = True
            Call Data_Validate(bSGECancelDummy)
        End If

    End If

    iAlterado = 0

    Traz_CTMaquinasParadas_Tela = SUCESSO

    Exit Function

Erro_Traz_CTMaquinasParadas_Tela:

    Traz_CTMaquinasParadas_Tela = gErr

    Select Case gErr

        Case 140731, 140969, 140970

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156093)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 140732

    'Limpa Tela
    Call Limpa_Tela_CTMaquinasParadas

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 140732

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156094)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156095)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 140733

    Call Limpa_Tela_CTMaquinasParadas

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 140733

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156096)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objCTMaquinasParadas As New ClassCTMaquinasParadas
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    '#####################
    'CRITICA DADOS DA TELA
    If Len(Trim(Codigo.Text)) = 0 Then gError 140734
    '#####################

    objCTMaquinasParadas.lCodigo = StrParaLong(Codigo.Text)
    objCTMaquinasParadas.iFilialEmpresa = giFilialEmpresa

    'Pergunta ao usuário se confirma a exclusão
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_CTMAQUINASPARADAS", objCTMaquinasParadas.lCodigo)

    If vbMsgRes = vbYes Then

        'Exclui a requisição de consumo
        lErro = CF("CTMaquinasParadas_Exclui", objCTMaquinasParadas)
        If lErro <> SUCESSO Then gError 140735

        'Limpa Tela
        Call Limpa_Tela_CTMaquinasParadas

    End If

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 140734
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_CTMAQUINASPARADAS_NAO_PREENCHIDO", gErr)
            Codigo.SetFocus

        Case 140735

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156097)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Codigo_Validate

    'Verifica se Codigo está preenchida
    If Len(Trim(Codigo.Text)) <> 0 Then

       'Critica a Codigo
       lErro = Long_Critica(Codigo.Text)
       If lErro <> SUCESSO Then gError 140736

    End If

    Exit Sub

Erro_Codigo_Validate:

    Cancel = True

    Select Case gErr

        Case 140736

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156098)

    End Select

    Exit Sub

End Sub

Private Sub Codigo_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Codigo, iAlterado)
    
End Sub

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownData_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownData_DownClick

    Data.SetFocus

    If Len(Data.ClipText) > 0 Then

        sData = Data.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 140737

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_DownClick:

    Select Case gErr

        Case 140737

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156099)

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
        If lErro <> SUCESSO Then gError 140738

        Data.Text = sData

    End If

    Exit Sub

Erro_UpDownData_UpClick:

    Select Case gErr

        Case 140738

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156100)

    End Select

    Exit Sub

End Sub

Private Sub Data_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Data, iAlterado)
    
End Sub

Private Sub Data_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCTMaquinas As New ClassCTMaquinas
Dim dHoras As Double
Dim iQtd As Integer
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objMaquinas As ClassMaquinas

On Error GoTo Erro_Data_Validate

    If Len(Trim(Data.ClipText)) <> 0 Then

        lErro = Data_Critica(Data.Text)
        If lErro <> SUCESSO Then gError 140739

        lErro = Move_CTMaquina_Memoria(objCentrodeTrabalho, objMaquinas)
        If lErro <> SUCESSO Then gError 140979
        
        objCTMaquinas.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
        objCTMaquinas.lNumIntDocMaq = objMaquinas.lNumIntDoc
        
        lErro = CF("Verifica_Maquinas_Disponiveis_Dia", objCTMaquinas, StrParaDate(Data.Text), giFilialEmpresa, iQtd, dHoras)
        If lErro <> SUCESSO Then gError 140980
        
        QtdDia.Caption = CStr(iQtd)
        
        HorasDia.Caption = Formata_Estoque(dHoras)
    
    End If

    Exit Sub

Erro_Data_Validate:

    Cancel = True

    Select Case gErr

        Case 140739, 140979, 140980

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156101)

    End Select

    Exit Sub

End Sub

Private Sub Data_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Tipo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Horas_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Horas_Validate

    'Verifica se Horas está preenchida
    If Len(Trim(Horas.Text)) <> 0 Then

       'Critica a Horas
       lErro = Valor_Positivo_Critica(Horas.Text)
       If lErro <> SUCESSO Then gError 140741

    End If

    Exit Sub

Erro_Horas_Validate:

    Cancel = True

    Select Case gErr

        Case 140741

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156102)

    End Select

    Exit Sub

End Sub

Private Sub Horas_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Horas, iAlterado)
    
End Sub

Private Sub Horas_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub QtdMaquinas_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_QtdMaquinas_Validate

    'Verifica se QtdMaquinas está preenchida
    If Len(Trim(QtdMaquinas.Text)) <> 0 Then

       'Critica a QtdMaquinas
       lErro = Inteiro_Critica(QtdMaquinas.Text)
       If lErro <> SUCESSO Then gError 140742

    End If

    Exit Sub

Erro_QtdMaquinas_Validate:

    Cancel = True

    Select Case gErr

        Case 140742

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156103)

    End Select

    Exit Sub

End Sub

Private Sub QtdMaquinas_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(QtdMaquinas, iAlterado)
    
End Sub

Private Sub QtdMaquinas_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoCodigo_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCTMaquinasParadas As ClassCTMaquinasParadas

On Error GoTo Erro_objEventoCodigo_evSelecao

    Set objCTMaquinasParadas = obj1

    'Mostra os dados do CTMaquinasParadas na tela
    lErro = Traz_CTMaquinasParadas_Tela(objCTMaquinasParadas)
    If lErro <> SUCESSO Then gError 140743

    Me.Show

    Exit Sub

Erro_objEventoCodigo_evSelecao:

    Select Case gErr

        Case 140743


        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156104)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim lErro As Long
Dim objCTMaquinasParadas As New ClassCTMaquinasParadas
Dim colSelecao As New Collection

On Error GoTo Erro_LabelCodigo_Click

    'Verifica se o Codigo foi preenchido
    If Len(Trim(Codigo.Text)) <> 0 Then

        objCTMaquinasParadas.lCodigo = Codigo.Text

    End If

    Call Chama_Tela("CTMaquinasParadasLista", colSelecao, objCTMaquinasParadas, objEventoCodigo)

    Exit Sub

Erro_LabelCodigo_Click:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156105)

    End Select

    Exit Sub

End Sub

Private Function Carrega_Tipo() As Long
'Carrega a combo de Tipo

Dim lErro As Long

On Error GoTo Erro_Carrega_Tipo

    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_TIPOPARADA, Tipo)
    If lErro <> SUCESSO Then gError 140968

    Carrega_Tipo = SUCESSO

    Exit Function

Erro_Carrega_Tipo:

    Carrega_Tipo = gErr

    Select Case gErr
    
        Case 140968

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156106)

    End Select

    Exit Function

End Function

Private Sub Tipo_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Private Sub BotaoProxNum_Click()
Dim lErro As Long
Dim lCodigo As Long

On Error GoTo Erro_BotaoProxNum_Click

    'Mostra número do proximo numero disponível para um Centro de Trabalho
    lErro = CF("CTMaquinasParadas_Automatico", lCodigo)
    If lErro <> SUCESSO Then gError 134339
    
    Codigo.Text = CStr(lCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 134339
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156107)
    
    End Select

    Exit Sub
    
End Sub
