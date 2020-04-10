VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl CTMaqProgTurno 
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   8055
   Begin VB.Frame Frame3 
      Caption         =   "Turnos"
      Height          =   1575
      Left            =   4485
      TabIndex        =   19
      Top             =   1665
      Width           =   3450
      Begin MSMask.MaskEdBox Turno 
         Height          =   315
         Left            =   990
         TabIndex        =   23
         Top             =   885
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   2
         Format          =   "##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Horas 
         Height          =   315
         Left            =   1710
         TabIndex        =   24
         Top             =   885
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   8
         Format          =   "#,##0.0#"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid GridTurnos 
         Height          =   1215
         Left            =   390
         TabIndex        =   20
         Top             =   240
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   2143
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Período"
      Height          =   1575
      Left            =   495
      TabIndex        =   12
      Top             =   1665
      Width           =   3780
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   315
         Left            =   1320
         TabIndex        =   13
         Top             =   330
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataInicial 
         Height          =   300
         Left            =   2625
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   330
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   315
         Left            =   1320
         TabIndex        =   15
         Top             =   855
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSComCtl2.UpDown UpDownDataFinal 
         Height          =   300
         Left            =   2625
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   855
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label LabelData 
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
         Height          =   315
         Left            =   930
         TabIndex        =   18
         Top             =   390
         Width           =   375
      End
      Begin VB.Label Label1 
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
         Height          =   315
         Left            =   885
         TabIndex        =   17
         Top             =   885
         Width           =   375
      End
   End
   Begin VB.TextBox Observacao 
      Height          =   315
      Left            =   1830
      MaxLength       =   255
      TabIndex        =   6
      Top             =   3380
      Width           =   6105
   End
   Begin VB.CommandButton BotaoProgTurno 
      Caption         =   "Programações dos Turnos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   210
      TabIndex        =   5
      ToolTipText     =   "Abre o Browse para as Programações dosTurnos cadastradas para este CT/Máquina"
      Top             =   3945
      Width           =   1875
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5835
      ScaleHeight     =   495
      ScaleWidth      =   2025
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   60
         Picture         =   "CTMaqProgTurno.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   570
         Picture         =   "CTMaqProgTurno.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1065
         Picture         =   "CTMaqProgTurno.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1545
         Picture         =   "CTMaqProgTurno.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
   End
   Begin VB.Label Maquina 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1830
      TabIndex        =   22
      Top             =   1260
      Width           =   2025
   End
   Begin VB.Label CodigoCT 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1830
      TabIndex        =   21
      Top             =   810
      Width           =   2025
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
      Left            =   660
      TabIndex        =   11
      Top             =   3410
      Width           =   1140
   End
   Begin VB.Label DescricaoCTPadrao 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3885
      TabIndex        =   10
      Top             =   810
      Width           =   4035
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
      Left            =   60
      TabIndex        =   9
      Top             =   840
      Width           =   1830
   End
   Begin VB.Label DescMaquina 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3885
      TabIndex        =   8
      Top             =   1260
      Width           =   4035
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
      Left            =   840
      TabIndex        =   7
      Top             =   1290
      Width           =   900
   End
End
Attribute VB_Name = "CTMaqProgTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoData As AdmEvento
Attribute objEventoData.VB_VarHelpID = -1

'Grid Turnos
Dim objGridTurnos As AdmGrid
Dim iGrid_Turno_Col As Integer
Dim iGrid_Horas_Col As Integer

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Programação dos Turnos da Máquina"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "CTMaquinaProgTurno"

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

Private Sub GridTurnos_KeyDown(KeyCode As Integer, Shift As Integer)

    Call Grid_Trata_Tecla1(KeyCode, objGridTurnos)
        
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty(True, UserControl.Enabled, True)
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

    Set objEventoData = Nothing
    
    Set objGridTurnos = Nothing
    
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156036)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoData = New AdmEvento
    
    'Grid Turnos
    Set objGridTurnos = New AdmGrid
    
    'tela em questão
    Set objGridTurnos.objForm = Me
        
    lErro = Inicializa_GridTurnos(objGridTurnos)
    If lErro <> SUCESSO Then gError 137386
        
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr
    
        Case 137386

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156037)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objCTMaqProgTurno As ClassCTMaqProgTurno) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objCTMaqProgTurno Is Nothing) Then

        lErro = Traz_CTMaquinaProgTurno_Tela(objCTMaqProgTurno)
        If lErro <> SUCESSO Then gError 137387

    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 137387

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156038)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objCTMaqProgTurno As ClassCTMaqProgTurno) As Long

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objMaquinas As ClassMaquinas
Dim iIndice As Integer
Dim objCTMaqProgTurnoItem As ClassCTMaqProgTurnoItens

On Error GoTo Erro_Move_Tela_Memoria

    If Len(Trim(CodigoCT.Caption)) <> 0 Then
            
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
    
        objCentrodeTrabalho.sNomeReduzido = CodigoCT.Caption
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 137388
        
        objCTMaqProgTurno.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
    
    End If
    
    If Len(Maquina.Caption) > 0 Then
        
        Set objMaquinas = New ClassMaquinas
        
        objMaquinas.sNomeReduzido = Maquina.Caption
        
        'Le a Máquina no BD a partir do NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 137389
        
        objCTMaqProgTurno.lNumIntDocMaq = objMaquinas.lNumIntDoc
        
    End If
    
    If Len(Trim(DataInicial.ClipText)) <> 0 Then objCTMaqProgTurno.dtData = StrParaDate(DataInicial.Text)
    If Len(Trim(DataInicial.ClipText)) <> 0 Then objCTMaqProgTurno.dtDataDe = StrParaDate(DataInicial.Text)
    If Len(Trim(DataFinal.ClipText)) <> 0 Then objCTMaqProgTurno.dtDataAte = StrParaDate(DataFinal.Text)
    objCTMaqProgTurno.sObservacao = Observacao.Text
    
    'Move o GridTurnos para memória
    For iIndice = 1 To objGridTurnos.iLinhasExistentes
    
        Set objCTMaqProgTurnoItem = New ClassCTMaqProgTurnoItens
        
        objCTMaqProgTurnoItem.iTurno = StrParaInt(GridTurnos.TextMatrix(iIndice, iGrid_Turno_Col))
        objCTMaqProgTurnoItem.dHoras = StrParaDbl(GridTurnos.TextMatrix(iIndice, iGrid_Horas_Col))
        
        objCTMaqProgTurno.colTurnos.Add objCTMaqProgTurnoItem
    
    Next

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 137388, 137389
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156039)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objCTMaqProgTurno As New ClassCTMaqProgTurno

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "CTMaquinaProgTurno"

    'Lê os dados da Tela PedidoVenda
    lErro = Move_Tela_Memoria(objCTMaqProgTurno)
    If lErro <> SUCESSO Then gError 137390

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDocCT", objCTMaqProgTurno.lNumIntDocCT, 0, "NumIntDocCt"
    colCampoValor.Add "NumIntDocMaq", objCTMaqProgTurno.lNumIntDocMaq, 0, "NumIntDocMaq"
    colCampoValor.Add "Data", objCTMaqProgTurno.dtData, 0, "Data"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 137390

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156040)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objCTMaqProgTurno As New ClassCTMaqProgTurno
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objMaquinas As ClassMaquinas

On Error GoTo Erro_Tela_Preenche

    If Len(Trim(CodigoCT.Caption)) <> 0 Then
            
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
    
        objCentrodeTrabalho.sNomeReduzido = CodigoCT.Caption
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 137704
        
        objCTMaqProgTurno.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
    
    End If
    
    If Len(Maquina.Caption) > 0 Then
        
        Set objMaquinas = New ClassMaquinas
        
        objMaquinas.sNomeReduzido = Maquina.Caption
        
        'Le a Máquina no BD a partir do NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 137705
        
        objCTMaqProgTurno.lNumIntDocMaq = objMaquinas.lNumIntDoc
        
    End If

    objCTMaqProgTurno.dtData = colCampoValor.Item("Data").vValor

    If objCTMaqProgTurno.lNumIntDocCT <> 0 And objCTMaqProgTurno.lNumIntDocMaq <> 0 And objCTMaqProgTurno.dtData <> DATA_NULA Then
        lErro = Traz_CTMaquinaProgTurno_Tela(objCTMaqProgTurno)
        If lErro <> SUCESSO Then gError 137391
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 137391, 137704, 137705

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156041)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objCTMaqProgTurno As New ClassCTMaqProgTurno
Dim iIndice As Integer

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Se Data Inicial está vazio
    If Len(Trim(DataInicial.ClipText)) = 0 Then gError 137392
    
    'Se Data Final está vazio
    If Len(Trim(DataFinal.ClipText)) = 0 Then gError 137535
    
    'Verifica se existem Turnos cadastrados
    If objGridTurnos.iLinhasExistentes = 0 Then gError 137393
    
    'Para cada CTMaqProgTurnoItens
    For iIndice = 1 To objGridTurnos.iLinhasExistentes
        
        'Verifica se as Horas foram informadas
        If Len(Trim(GridTurnos.TextMatrix(iIndice, iGrid_Horas_Col))) = 0 Then gError 137394

    Next

    'Preenche o objCTMaqProgTurno
    lErro = Move_Tela_Memoria(objCTMaqProgTurno)
    If lErro <> SUCESSO Then gError 137395

    lErro = Trata_Alteracao(objCTMaqProgTurno, objCTMaqProgTurno.dtDataDe, objCTMaqProgTurno.dtDataAte, objCTMaqProgTurno.lNumIntDocCT, objCTMaqProgTurno.lNumIntDocMaq)
    If lErro <> SUCESSO Then gError 137687
    
    'Grava CTMaquinaProgTurno no Banco de Dados - conforme periodo informado
    lErro = CF("CTMaquinaProgTurno_Grava_Periodo", objCTMaqProgTurno)
    If lErro <> SUCESSO Then gError 137397
    
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 137392, 137535
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
            
        Case 137393
            Call Rotina_Erro(vbOKOnly, "ERRO_GRIDTURNOS_NAO_PREENCHIDO", gErr)
        
        Case 137394
            Call Rotina_Erro(vbOKOnly, "ERRO_HORASGRIDTURNOS_NAO_PREENCHIDA", gErr, iIndice)

        Case 137395, 137397, 137687
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156042)

    End Select

    Exit Function

End Function

Function Limpa_Tela_CTMaquinaProgTurno() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_CTMaquinaProgTurno
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)
    
    'Limpa o grid
    Call Grid_Limpa(objGridTurnos)

    iAlterado = 0

    Limpa_Tela_CTMaquinaProgTurno = SUCESSO

    Exit Function

Erro_Limpa_Tela_CTMaquinaProgTurno:

    Limpa_Tela_CTMaquinaProgTurno = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156043)

    End Select

    Exit Function

End Function

Function Traz_CTMaquinaProgTurno_Tela(objCTMaqProgTurno As ClassCTMaqProgTurno) As Long

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objMaquinas As ClassMaquinas
Dim iLinha As Integer
Dim objCTMaqProgTurnoItens As New ClassCTMaqProgTurnoItens

On Error GoTo Erro_Traz_CTMaquinaProgTurno_Tela

    If objCTMaqProgTurno.lNumIntDocCT <> 0 Then
        
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        objCentrodeTrabalho.lNumIntDoc = objCTMaqProgTurno.lNumIntDocCT
        
        lErro = CF("CentroDeTrabalho_Le_NumIntDoc", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134590 Then gError 137398
        
        CodigoCT.Caption = objCentrodeTrabalho.sNomeReduzido
        DescricaoCTPadrao.Caption = objCentrodeTrabalho.sDescricao
    
    End If

    If objCTMaqProgTurno.lNumIntDocMaq <> 0 Then

        Set objMaquinas = New ClassMaquinas
        
        objMaquinas.lNumIntDoc = objCTMaqProgTurno.lNumIntDocMaq
        
        lErro = CF("Maquinas_Le_NumIntDoc", objMaquinas)
        If lErro <> SUCESSO And lErro <> 106353 Then gError 137399
        
        Maquina.Caption = objMaquinas.sNomeReduzido
        DescMaquina.Caption = objMaquinas.sDescricao
        
    End If

    'Lê o CTMaquinaProgTurno que está sendo Passado
    lErro = CF("CTMaquinaProgTurno_Le", objCTMaqProgTurno)
    If lErro <> SUCESSO And lErro <> 136704 Then gError 137400

    If lErro = SUCESSO Then

        If objCTMaqProgTurno.dtData <> 0 Then
        
            DataInicial.PromptInclude = False
            DataInicial.Text = Format(objCTMaqProgTurno.dtData, "dd/mm/yy")
            DataInicial.PromptInclude = True
        
            DataFinal.PromptInclude = False
            DataFinal.Text = Format(objCTMaqProgTurno.dtData, "dd/mm/yy")
            DataFinal.PromptInclude = True
        
        End If

        Observacao.Text = objCTMaqProgTurno.sObservacao
        
        'Limpa o Grid antes de colocar algo nele
        Call Grid_Limpa(objGridTurnos)
        
        iLinha = 1
        
        'Exibe os dados da coleção de Turnos na tela
        For Each objCTMaqProgTurnoItens In objCTMaqProgTurno.colTurnos
            
            'Insere no Grid Turnos
            GridTurnos.TextMatrix(iLinha, iGrid_Turno_Col) = objCTMaqProgTurnoItens.iTurno
            GridTurnos.TextMatrix(iLinha, iGrid_Horas_Col) = Formata_Estoque(objCTMaqProgTurnoItens.dHoras)
        
            iLinha = iLinha + 1
        
        Next
        
        objGridTurnos.iLinhasExistentes = objCTMaqProgTurno.colTurnos.Count
        
    End If
    
    iAlterado = 0

    Traz_CTMaquinaProgTurno_Tela = SUCESSO

    Exit Function

Erro_Traz_CTMaquinaProgTurno_Tela:

    Traz_CTMaquinaProgTurno_Tela = gErr

    Select Case gErr

        Case 137398 To 137400
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156044)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 137401

    'Limpa Tela
    Call Limpa_Tela_CTMaquinaProgTurno

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 137401

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156045)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156046)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 137402

    Call Limpa_Tela_CTMaquinaProgTurno

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 137402

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156047)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objCTMaqProgTurno As New ClassCTMaqProgTurno
Dim vbMsgRes As VbMsgBoxResult
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objMaquinas As ClassMaquinas

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    If Len(Trim(DataInicial.ClipText)) = 0 Then gError 137403

    If Len(Trim(CodigoCT.Caption)) <> 0 Then
            
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
    
        objCentrodeTrabalho.sNomeReduzido = CodigoCT.Caption
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 137404
        
        objCTMaqProgTurno.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
    
    End If
    
    If Len(Maquina.Caption) > 0 Then
        
        Set objMaquinas = New ClassMaquinas
        
        objMaquinas.sNomeReduzido = Maquina.Caption
        
        'Le a Máquina no BD a partir do NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 137405
        
        objCTMaqProgTurno.lNumIntDocMaq = objMaquinas.lNumIntDoc
        
    End If
    
    objCTMaqProgTurno.dtData = StrParaDate(DataInicial.Text)
    
    'Se a data final não estiver preenchida
    If Len(Trim(DataFinal.ClipText)) = 0 Then
        
        'Pergunta ao usuário se confirma a exclusão
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_CTMAQUINAPROGTURNO", objCTMaqProgTurno.dtData)
    
        If vbMsgRes = vbNo Then
            GL_objMDIForm.MousePointer = vbDefault
            Exit Sub
        End If
    
        'Exclui a requisição de consumo
        lErro = CF("CTMaquinaProgTurno_Exclui", objCTMaqProgTurno)
        If lErro <> SUCESSO Then gError 137406
    
    Else
    
        objCTMaqProgTurno.dtDataDe = StrParaDate(DataInicial.Text)
        objCTMaqProgTurno.dtDataAte = StrParaDate(DataFinal.Text)
    
        'Pergunta ao usuário se confirma a exclusão
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PERIODO_CTMAQPROGTURNO", objCTMaqProgTurno.dtDataDe, objCTMaqProgTurno.dtDataAte)
    
        If vbMsgRes = vbNo Then
            GL_objMDIForm.MousePointer = vbDefault
            Exit Sub
        End If
    
        'Exclui a requisição de consumo
        lErro = CF("CTMaquinaProgTurno_Exclui_Periodo", objCTMaqProgTurno)
        If lErro <> SUCESSO Then gError 137407
    
    End If
    
    'Limpa Tela
    Call Limpa_Tela_CTMaquinaProgTurno

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 137403
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
            DataInicial.SetFocus

        Case 137404 To 137407
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156048)

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
        If lErro <> SUCESSO Then gError 137408

        DataInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicial_DownClick:

    Select Case gErr

        Case 137408

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156049)

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
        If lErro <> SUCESSO Then gError 137409

        DataInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicial_UpClick:

    Select Case gErr

        Case 137409

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156050)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataInicial, iAlterado)
    
End Sub

Private Sub DataInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim lIntervalo As Long

On Error GoTo Erro_DataInicial_Validate

    If Len(Trim(DataInicial.ClipText)) <> 0 Then

        lErro = Data_Critica(DataInicial.Text)
        If lErro <> SUCESSO Then gError 137410

        'Se a data final também está preenchida
        If Len(Trim(DataFinal.ClipText)) <> 0 Then
        
            'Verifica qual é o intervalo entre as datas
            lIntervalo = DateDiff("d", StrParaDate(DataInicial.Text), StrParaDate(DataFinal.Text))
            
            'Se o intervalo for negativo -> Erro
            If lIntervalo < 0 Then gError 137411
        
        End If

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 137410

        Case 137411
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156051)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub UpDownDataFinal_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFinal_DownClick

    DataFinal.SetFocus

    If Len(DataFinal.ClipText) > 0 Then

        sData = DataFinal.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 137412

        DataFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFinal_DownClick:

    Select Case gErr

        Case 137412

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156052)

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
        If lErro <> SUCESSO Then gError 137413

        DataFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFinal_UpClick:

    Select Case gErr

        Case 137413

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156053)

    End Select

    Exit Sub

End Sub

Private Sub DataFinal_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(DataFinal, iAlterado)
    
End Sub

Private Sub DataFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim lIntervalo As Long

On Error GoTo Erro_DataFinal_Validate

    If Len(Trim(DataFinal.ClipText)) <> 0 Then

        lErro = Data_Critica(DataFinal.Text)
        If lErro <> SUCESSO Then gError 137414

        'Se a data inicial também está preenchida
        If Len(Trim(DataInicial.ClipText)) <> 0 Then
        
            'Verifica qual é o intervalo entre as datas
            lIntervalo = DateDiff("d", StrParaDate(DataInicial.Text), StrParaDate(DataFinal.Text))
            
            'Se o intervalo for negativo -> Erro
            If lIntervalo < 0 Then gError 137415
            
        End If

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 137414
        
        Case 137415
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156054)

    End Select

    Exit Sub

End Sub

Private Sub DataFinal_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoData_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCTMaqProgTurno As ClassCTMaqProgTurno

On Error GoTo Erro_objEventoData_evSelecao

    Set objCTMaqProgTurno = obj1

    'Mostra os dados do CTMaquinaProgTurno na tela
    lErro = Traz_CTMaquinaProgTurno_Tela(objCTMaqProgTurno)
    If lErro <> SUCESSO Then gError 137416

    Me.Show

    Exit Sub

Erro_objEventoData_evSelecao:

    Select Case gErr

        Case 137416

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156055)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProgTurno_Click()

Dim lErro As Long
Dim objCTMaqProgTurno As New ClassCTMaqProgTurno
Dim colSelecao As New Collection
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objMaquinas As ClassMaquinas
Dim sFiltro As String

On Error GoTo Erro_BotaoProgTurno_Click

    'Verifica se o Data foi preenchido
    If Len(Trim(DataInicial.ClipText)) <> 0 Then

        objCTMaqProgTurno.dtData = StrParaDate(DataInicial.Text)

    End If

    If Len(Trim(CodigoCT.Caption)) <> 0 Then
            
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
    
        objCentrodeTrabalho.sNomeReduzido = CodigoCT.Caption
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 137417
        
        objCTMaqProgTurno.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
    
    End If
    
    If Len(Maquina.Caption) > 0 Then
        
        Set objMaquinas = New ClassMaquinas
        
        objMaquinas.sNomeReduzido = Maquina.Caption
        
        'Le a Máquina no BD a partir do NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 137418
        
        objCTMaqProgTurno.lNumIntDocMaq = objMaquinas.lNumIntDoc
        
    End If

    sFiltro = "NumIntDocCT = ? And NumIntDocMaq = ?"
    colSelecao.Add objCTMaqProgTurno.lNumIntDocCT
    colSelecao.Add objCTMaqProgTurno.lNumIntDocMaq

    Call Chama_Tela("CTMaquinaProgTurnoLista", colSelecao, objCTMaqProgTurno, objEventoData, sFiltro)

    Exit Sub

Erro_BotaoProgTurno_Click:

    Select Case gErr
    
        Case 137417, 137418
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156056)

    End Select

    Exit Sub

End Sub

Private Function Inicializa_GridTurnos(objGrid As AdmGrid) As Long

Dim iIndice As Integer

    'tela em questão
    Set objGrid.objForm = Me

    'titulos do grid
    objGrid.colColuna.Add ("")
    objGrid.colColuna.Add ("Turno")
    objGrid.colColuna.Add ("Horas")

    'Controles que participam do Grid
    objGrid.colCampo.Add (Turno.Name)
    objGrid.colCampo.Add (Horas.Name)
    
    'Colunas do Grid
    iGrid_Turno_Col = 1
    iGrid_Horas_Col = 2
    
    objGrid.objGrid = GridTurnos

    'Todas as linhas do grid
    objGrid.objGrid.Rows = NUM_MAXIMO_ITENS + 1

    objGrid.iExecutaRotinaEnable = GRID_EXECUTAR_ROTINA_ENABLE

    objGrid.iLinhasVisiveis = 2

    'Largura da primeira coluna
    GridTurnos.ColWidth(0) = 600

    objGrid.iGridLargAuto = GRID_LARGURA_MANUAL

    Call Grid_Inicializa(objGrid)

    Inicializa_GridTurnos = SUCESSO

End Function

Private Sub GridTurnos_Click()

Dim iExecutaEntradaCelula As Integer

        Call Grid_Click(objGridTurnos, iExecutaEntradaCelula)

        If iExecutaEntradaCelula = 1 Then
            Call Grid_Entrada_Celula(objGridTurnos, iAlterado)
        End If

End Sub

Private Sub GridTurnos_GotFocus()
    
    Call Grid_Recebe_Foco(objGridTurnos)

End Sub

Private Sub GridTurnos_EnterCell()

    Call Grid_Entrada_Celula(objGridTurnos, iAlterado)

End Sub

Private Sub GridTurnos_LeaveCell()
    
    Call Saida_Celula(objGridTurnos)

End Sub

Private Sub GridTurnos_KeyPress(KeyAscii As Integer)

Dim iExecutaEntradaCelula As Integer

    Call Grid_Trata_Tecla(KeyAscii, objGridTurnos, iExecutaEntradaCelula)

    If iExecutaEntradaCelula = 1 Then
        Call Grid_Entrada_Celula(objGridTurnos, iAlterado)
    End If

End Sub

Private Sub GridTurnos_RowColChange()

    Call Grid_RowColChange(objGridTurnos)

End Sub

Private Sub GridTurnos_Scroll()

    Call Grid_Scroll(objGridTurnos)

End Sub

Public Function Saida_Celula(objGridInt As AdmGrid) As Long
'faz a critica da celula do grid que está deixando de ser a corrente

Dim lErro As Long
Dim iIndice As Integer

On Error GoTo Erro_Saida_Celula

    lErro = Grid_Inicializa_Saida_Celula(objGridInt)

    If lErro = SUCESSO Then
    
        'Turnos
        If objGridInt.objGrid.Name = GridTurnos.Name Then
            
            'Verifica qual a coluna do Grid em questão
            Select Case objGridInt.objGrid.Col
                
                Case iGrid_Turno_Col
                
                    lErro = Saida_Celula_Turno(objGridInt)
                    If lErro <> SUCESSO Then gError 137419
                
                Case iGrid_Horas_Col
                
                    lErro = Saida_Celula_Horas(objGridInt)
                    If lErro <> SUCESSO Then gError 137420
                    
            End Select
                
        End If

        lErro = Grid_Finaliza_Saida_Celula(objGridInt)
        If lErro Then gError 137421

    End If

    Saida_Celula = SUCESSO

    Exit Function

Erro_Saida_Celula:

    Saida_Celula = gErr

    Select Case gErr

        Case 137419, 137420
            'erros tratatos nas rotinas chamadas
        
        Case 137421
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156057)

    End Select

    Exit Function

End Function

Private Sub Horas_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Horas_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridTurnos)

End Sub

Private Sub Horas_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTurnos)

End Sub

Private Sub Horas_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridTurnos.objControle = Horas
    lErro = Grid_Campo_Libera_Foco(objGridTurnos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Sub Turno_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Turno_GotFocus()

    Call Grid_Campo_Recebe_Foco(objGridTurnos)

End Sub

Private Sub Turno_KeyPress(KeyAscii As Integer)

    Call Grid_Trata_Tecla_Campo(KeyAscii, objGridTurnos)

End Sub

Private Sub Turno_Validate(Cancel As Boolean)

Dim lErro As Long

    Set objGridTurnos.objControle = Turno
    lErro = Grid_Campo_Libera_Foco(objGridTurnos)
    If lErro <> SUCESSO Then Cancel = True

End Sub

Private Function Saida_Celula_Turno(objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iLinha As Integer
Dim sTurno As String

On Error GoTo Erro_Saida_Celula_Turno

    Set objGridInt.objControle = Turno

    'Se o campo foi preenchido
    If Len(Turno.Text) > 0 Then

        'Critica o Turno
        lErro = Inteiro_Critica(Turno.Text)
        If lErro <> SUCESSO Then gError 137422
        
        'Verifica se o turno está repetido no grid
        For iLinha = 1 To objGridTurnos.iLinhasExistentes
            
            If iLinha <> GridTurnos.Row Then
                                                    
                If GridTurnos.TextMatrix(iLinha, iGrid_Turno_Col) = Turno.Text Then
                    
                    sTurno = Turno.Text
                    Turno.Text = ""
                    gError 137893
                    
                End If
                    
            End If
                           
        Next
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridTurnos.Row - GridTurnos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 137423

    Saida_Celula_Turno = SUCESSO

    Exit Function

Erro_Saida_Celula_Turno:

    Saida_Celula_Turno = gErr

    Select Case gErr

        Case 137422, 137423
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)
            
        Case 137893
            Call Rotina_Erro(vbOKOnly, "ERRO_TURNO_REPETIDO", gErr, sTurno, iLinha)
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 156058)

    End Select

    Exit Function

End Function

Private Function Saida_Celula_Horas(objGridInt As AdmGrid) As Long

Dim lErro As Long

On Error GoTo Erro_Saida_Celula_Horas

    Set objGridInt.objControle = Horas

    'Se o campo foi preenchido
    If Len(Horas.Text) > 0 Then

        'Critica a Quantidade de Horas do Turno
        lErro = Horas_Turno_Critica(Horas.Text, GridTurnos.Row, objGridInt)
        If lErro <> SUCESSO Then gError 137424
        
        Horas.Text = Formata_Estoque(Horas.Text)
        
        'verifica se precisa preencher o grid com uma nova linha
        If GridTurnos.Row - GridTurnos.FixedRows = objGridInt.iLinhasExistentes Then
            objGridInt.iLinhasExistentes = objGridInt.iLinhasExistentes + 1
        End If

    End If
    
    lErro = Grid_Abandona_Celula(objGridInt)
    If lErro <> SUCESSO Then gError 137425

    Saida_Celula_Horas = SUCESSO

    Exit Function

Erro_Saida_Celula_Horas:

    Saida_Celula_Horas = gErr

    Select Case gErr

        Case 137424, 137425
            Call Grid_Trata_Erro_Saida_Celula(objGridInt)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 156059)

    End Select

    Exit Function

End Function

Function Horas_Turno_Critica(sDispHorasTurno As String, iGridLinha As Integer, objGridInt As AdmGrid) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim dQtdeTotalHoras As Double

On Error GoTo Erro_Horas_Turno_Critica

    'Critica a Quantidade de Horas
    lErro = Valor_Positivo_Critica(sDispHorasTurno)
    If lErro <> SUCESSO Then gError 137426
    
    'Efetua a Somatória da Quantidade Total de Horas
    For iIndice = 1 To objGridInt.iLinhasExistentes
    
        'se é a linha que estou alterando ...
        If iIndice = iGridLinha Then
        
            'despreza a hora do grid e acumula a que esta sendo alterada
            dQtdeTotalHoras = dQtdeTotalHoras + StrParaDbl(sDispHorasTurno)
        
        Else
            
            'acumula as horas
            dQtdeTotalHoras = dQtdeTotalHoras + StrParaDbl(GridTurnos.TextMatrix(iIndice, iGrid_Horas_Col))
        
        End If
    
    Next
    
    'Verifica se a Somatória das Horas é maior que 24 horas
    If dQtdeTotalHoras > HORAS_DO_DIA Then gError 137427
    
    Horas_Turno_Critica = SUCESSO
    
    Exit Function
    
Erro_Horas_Turno_Critica:

    Horas_Turno_Critica = gErr
    
    Select Case gErr
    
        Case 137426
        
        Case 137427
            Call Rotina_Erro(vbOKOnly, "ERRO_QTDEHORASGRID_EXCEDE_DIA", gErr, dQtdeTotalHoras, "Horas")
   
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 156060)
    
    End Select
    
    Exit Function

End Function

Public Sub Rotina_Grid_Enable(iLinha As Integer, objControl As Object, iLocalChamada As Integer)

Dim lErro As Long
Dim iTurno As String

On Error GoTo Erro_Rotina_Grid_Enable

    'Guardo o valor do Turno
    iTurno = StrParaInt(GridTurnos.TextMatrix(GridTurnos.Row, iGrid_Turno_Col))

    Select Case objControl.Name
    
        Case Is = "Turno"
            
            If iTurno > 0 Then
                objControl.Enabled = False
    
            Else
                objControl.Enabled = True
            
            End If
    
        Case Is = "Horas"
            
            If iTurno > 0 Then
                objControl.Enabled = True
    
            Else
                objControl.Enabled = False
            
            End If
    
    End Select
        
    Exit Sub

Erro_Rotina_Grid_Enable:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, 156061)

    End Select

    Exit Sub

End Sub


