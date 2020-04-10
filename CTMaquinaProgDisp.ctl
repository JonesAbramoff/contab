VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl CTMaquinaProgDisp 
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8145
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4320
   ScaleWidth      =   8145
   Begin VB.Frame Frame1 
      Caption         =   "Período"
      Height          =   765
      Left            =   150
      TabIndex        =   14
      Top             =   1740
      Width           =   7875
      Begin MSMask.MaskEdBox DataInicial 
         Height          =   315
         Left            =   1755
         TabIndex        =   15
         Top             =   255
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
         Left            =   3060
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataFinal 
         Height          =   315
         Left            =   4830
         TabIndex        =   18
         Top             =   255
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
         Left            =   6150
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
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
         Left            =   4440
         TabIndex        =   20
         Top             =   285
         Width           =   375
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
         Left            =   1350
         TabIndex        =   17
         Top             =   285
         Width           =   375
      End
   End
   Begin VB.CommandButton BotaoProgDisponibilidade 
      Caption         =   "Programações da Disponibilidade"
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
      Left            =   150
      TabIndex        =   13
      ToolTipText     =   "Abre o Browse para as Programações da Disponibilidade cadastradas para este CT/Máquina"
      Top             =   3690
      Width           =   1875
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   5925
      ScaleHeight     =   495
      ScaleWidth      =   2025
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2085
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   75
         Picture         =   "CTMaquinaProgDisp.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   585
         Picture         =   "CTMaquinaProgDisp.ctx":015A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1080
         Picture         =   "CTMaquinaProgDisp.ctx":02E4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1560
         Picture         =   "CTMaquinaProgDisp.ctx":0816
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin MSMask.MaskEdBox Quantidade 
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Top             =   2700
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   2
      Format          =   "##"
      PromptChar      =   " "
   End
   Begin VB.TextBox Observacao 
      Height          =   315
      Left            =   1920
      MaxLength       =   255
      TabIndex        =   7
      Top             =   3165
      Width           =   6105
   End
   Begin VB.Label CodigoCT 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1920
      TabIndex        =   22
      Top             =   870
      Width           =   2025
   End
   Begin VB.Label Maquina 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1920
      TabIndex        =   21
      Top             =   1320
      Width           =   2025
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
      Left            =   930
      TabIndex        =   12
      Top             =   1350
      Width           =   900
   End
   Begin VB.Label DescMaquina 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3975
      TabIndex        =   11
      Top             =   1320
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
      Left            =   150
      TabIndex        =   10
      Top             =   900
      Width           =   1830
   End
   Begin VB.Label DescricaoCTPadrao 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3975
      TabIndex        =   9
      Top             =   870
      Width           =   4035
   End
   Begin VB.Label LabelQuantidade 
      Caption         =   "Quantidade:"
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
      Left            =   825
      TabIndex        =   6
      Top             =   2730
      Width           =   1020
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
      Left            =   750
      TabIndex        =   8
      Top             =   3195
      Width           =   1140
   End
End
Attribute VB_Name = "CTMaquinaProgDisp"
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

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Programação da Disponibilidade das Máquinas"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "CTMaquinaProgDisp"

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

Private Sub UpDownDataFinal_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataFinal_DownClick

    DataFinal.SetFocus

    If Len(DataFinal.ClipText) > 0 Then

        sData = DataFinal.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError 137354

        DataFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFinal_DownClick:

    Select Case gErr

        Case 137354

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156062)

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
        If lErro <> SUCESSO Then gError 137355

        DataFinal.Text = sData

    End If

    Exit Sub

Erro_UpDownDataFinal_UpClick:

    Select Case gErr

        Case 137355

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156063)

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
        If lErro <> SUCESSO Then gError 137356
        
        'Se a data inicial também está preenchida
        If Len(Trim(DataInicial.ClipText)) <> 0 Then
        
            'Verifica qual é o intervalo entre as datas
            lIntervalo = DateDiff("d", StrParaDate(DataInicial.Text), StrParaDate(DataFinal.Text))
            
            'Se o intervalo for negativo -> Erro
            If lIntervalo < 0 Then gError 137357
        
        End If

    End If

    Exit Sub

Erro_DataFinal_Validate:

    Cancel = True

    Select Case gErr

        Case 137356
            'erro tratado na rotina chamada
        
        Case 137357
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156064)

    End Select

    Exit Sub

End Sub

Private Sub DataFinal_Change()

    iAlterado = REGISTRO_ALTERADO

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
    Call ComandoSeta_Liberar(Me.Name)

    Exit Sub

Erro_Form_Unload:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156065)

    End Select

    Exit Sub

End Sub

Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set objEventoData = New AdmEvento
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156066)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objCTMaqProgDisp As ClassCTMaqProgDisp) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (objCTMaqProgDisp Is Nothing) Then

        lErro = Traz_CTMaquinaProgDisponibilidade_Tela(objCTMaqProgDisp)
        If lErro <> SUCESSO Then gError 137358
        
    End If

    iAlterado = 0

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 137358
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156067)

    End Select

    iAlterado = 0

    Exit Function

End Function

Function Move_Tela_Memoria(objCTMaqProgDisp As ClassCTMaqProgDisp) As Long

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objMaquinas As ClassMaquinas

On Error GoTo Erro_Move_Tela_Memoria

    If Len(Trim(CodigoCT.Caption)) <> 0 Then
            
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
    
        objCentrodeTrabalho.sNomeReduzido = CodigoCT.Caption
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 137359
        
        objCTMaqProgDisp.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
    
    End If
    
    If Len(Maquina.Caption) > 0 Then
        
        Set objMaquinas = New ClassMaquinas
        
        objMaquinas.sNomeReduzido = Maquina.Caption
        
        'Le a Máquina no BD a partir do NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 137360
        
        objCTMaqProgDisp.lNumIntDocMaq = objMaquinas.lNumIntDoc
        
    End If
    
    If Len(Trim(DataInicial.ClipText)) <> 0 Then objCTMaqProgDisp.dtData = StrParaDate(DataInicial.Text)
    If Len(Trim(DataInicial.ClipText)) <> 0 Then objCTMaqProgDisp.dtDataDe = StrParaDate(DataInicial.Text)
    If Len(Trim(DataFinal.ClipText)) <> 0 Then objCTMaqProgDisp.dtDataAte = StrParaDate(DataFinal.Text)
    objCTMaqProgDisp.iQuantidade = StrParaInt(Quantidade.Text)
    objCTMaqProgDisp.sObservacao = Observacao.Text

    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr
    
        Case 137359, 137360
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156068)

    End Select

    Exit Function

End Function

Function Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro) As Long

Dim lErro As Long
Dim objCTMaqProgDisp As New ClassCTMaqProgDisp

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "CTMaquinaProgDisponibilidade"

    'Lê os dados da Tela
    lErro = Move_Tela_Memoria(objCTMaqProgDisp)
    If lErro <> SUCESSO Then gError 137361

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "NumIntDocCT", objCTMaqProgDisp.lNumIntDocCT, 0, "NumIntDocCt"
    colCampoValor.Add "NumIntDocMaq", objCTMaqProgDisp.lNumIntDocMaq, 0, "NumIntDocMaq"
    colCampoValor.Add "Data", objCTMaqProgDisp.dtData, 0, "Data"

    Tela_Extrai = SUCESSO

    Exit Function

Erro_Tela_Extrai:

    Tela_Extrai = gErr

    Select Case gErr

        Case 137361

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156069)

    End Select

    Exit Function

End Function

Function Tela_Preenche(colCampoValor As AdmColCampoValor) As Long

Dim lErro As Long
Dim objCTMaqProgDisp As New ClassCTMaqProgDisp
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objMaquinas As ClassMaquinas

On Error GoTo Erro_Tela_Preenche

    If Len(Trim(CodigoCT.Caption)) <> 0 Then
            
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
    
        objCentrodeTrabalho.sNomeReduzido = CodigoCT.Caption
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 137706
        
        objCTMaqProgDisp.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
    
    End If
    
    If Len(Maquina.Caption) > 0 Then
        
        Set objMaquinas = New ClassMaquinas
        
        objMaquinas.sNomeReduzido = Maquina.Caption
        
        'Le a Máquina no BD a partir do NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 137707
        
        objCTMaqProgDisp.lNumIntDocMaq = objMaquinas.lNumIntDoc
        
    End If

    objCTMaqProgDisp.dtData = colCampoValor.Item("Data").vValor

    If objCTMaqProgDisp.lNumIntDocCT <> 0 And objCTMaqProgDisp.lNumIntDocMaq <> 0 And objCTMaqProgDisp.dtData <> DATA_NULA Then
        lErro = Traz_CTMaquinaProgDisponibilidade_Tela(objCTMaqProgDisp)
        If lErro <> SUCESSO Then gError 137362
    End If

    Tela_Preenche = SUCESSO

    Exit Function

Erro_Tela_Preenche:

    Tela_Preenche = gErr

    Select Case gErr

        Case 137362, 137706, 137707

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156070)

    End Select

    Exit Function

End Function

Function Gravar_Registro() As Long

Dim lErro As Long
Dim objCTMaqProgDisp As New ClassCTMaqProgDisp

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    'Se Data Inicial está vazio
    If Len(Trim(DataInicial.ClipText)) = 0 Then gError 137363
    
    'Se Data Final está vazio
    If Len(Trim(DataFinal.ClipText)) = 0 Then gError 137536
    
    'Se Quantidade está vazio
    If Len(Trim(Quantidade.Text)) = 0 Then gError 137364

    'Preenche o objCTMaqProgDisp
    lErro = Move_Tela_Memoria(objCTMaqProgDisp)
    If lErro <> SUCESSO Then gError 137365
    
    lErro = Trata_Alteracao(objCTMaqProgDisp, objCTMaqProgDisp.dtDataDe, objCTMaqProgDisp.dtDataAte, objCTMaqProgDisp.lNumIntDocCT, objCTMaqProgDisp.lNumIntDocMaq)
    If lErro <> SUCESSO Then gError 137686
    
    'Grava CTMaquinaProgDisponibilidade no Banco de Dados - conforme periodo informado
    lErro = CF("CTMaquinaProgDisponibilidade_Grava_Periodo", objCTMaqProgDisp)
    If lErro <> SUCESSO Then gError 137367
        
    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 137363, 137536
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
            
        Case 137364
            Call Rotina_Erro(vbOKOnly, "ERRO_QUANTIDADE_NAO_PREENCHIDO1", gErr)

        Case 137365, 137367, 137686

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156071)

    End Select

    Exit Function

End Function

Function Limpa_Tela_CTMaquinaProgDisponibilidade() As Long

Dim lErro As Long

On Error GoTo Erro_Limpa_Tela_CTMaquinaProgDisponibilidade
        
    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    'Função genérica que limpa campos da tela
    Call Limpa_Tela(Me)

    iAlterado = 0

    Limpa_Tela_CTMaquinaProgDisponibilidade = SUCESSO

    Exit Function

Erro_Limpa_Tela_CTMaquinaProgDisponibilidade:

    Limpa_Tela_CTMaquinaProgDisponibilidade = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156072)

    End Select

    Exit Function

End Function

Function Traz_CTMaquinaProgDisponibilidade_Tela(objCTMaqProgDisp As ClassCTMaqProgDisp) As Long

Dim lErro As Long
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objMaquinas As ClassMaquinas

On Error GoTo Erro_Traz_CTMaquinaProgDisponibilidade_Tela

    If objCTMaqProgDisp.lNumIntDocCT <> 0 Then
        
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
        
        objCentrodeTrabalho.lNumIntDoc = objCTMaqProgDisp.lNumIntDocCT
        
        lErro = CF("CentroDeTrabalho_Le_NumIntDoc", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134590 Then gError 137368
        
        CodigoCT.Caption = objCentrodeTrabalho.sNomeReduzido
        DescricaoCTPadrao.Caption = objCentrodeTrabalho.sDescricao
    
    End If

    If objCTMaqProgDisp.lNumIntDocMaq <> 0 Then

        Set objMaquinas = New ClassMaquinas
        
        objMaquinas.lNumIntDoc = objCTMaqProgDisp.lNumIntDocMaq
        
        lErro = CF("Maquinas_Le_NumIntDoc", objMaquinas)
        If lErro <> SUCESSO And lErro <> 106353 Then gError 137369
        
        Maquina.Caption = objMaquinas.sNomeReduzido
        DescMaquina.Caption = objMaquinas.sDescricao
        
    End If

    'Lê o CTMaquinaProgDisponibilidade que está sendo Passado
    lErro = CF("CTMaquinaProgDisponibilidade_Le", objCTMaqProgDisp)
    If lErro <> SUCESSO And lErro <> 136564 Then gError 137370
    
    If lErro = SUCESSO Then

        If objCTMaqProgDisp.dtData <> 0 Then
        
            DataInicial.PromptInclude = False
            DataInicial.Text = Format(objCTMaqProgDisp.dtData, "dd/mm/yy")
            DataInicial.PromptInclude = True
        
            DataFinal.PromptInclude = False
            DataFinal.Text = Format(objCTMaqProgDisp.dtData, "dd/mm/yy")
            DataFinal.PromptInclude = True
        
        End If

        If objCTMaqProgDisp.iQuantidade <> 0 Then Quantidade.Text = CStr(objCTMaqProgDisp.iQuantidade)
        Observacao.Text = objCTMaqProgDisp.sObservacao

    End If
    
    iAlterado = 0

    Traz_CTMaquinaProgDisponibilidade_Tela = SUCESSO

    Exit Function

Erro_Traz_CTMaquinaProgDisponibilidade_Tela:

    Traz_CTMaquinaProgDisponibilidade_Tela = gErr

    Select Case gErr
    
        Case 137368 To 137370
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156073)

    End Select

    Exit Function

End Function

Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro
    If lErro <> SUCESSO Then gError 137371

    'Limpa Tela
    Call Limpa_Tela_CTMaquinaProgDisponibilidade

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 137371

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156074)

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
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156075)

    End Select

    Exit Sub

End Sub

Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 137372

    Call Limpa_Tela_CTMaquinaProgDisponibilidade

    Exit Sub

Erro_BotaoLimpar_Click:

    Select Case gErr

        Case 137372
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156076)

    End Select

    Exit Sub

End Sub

Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objCTMaqProgDisp As New ClassCTMaqProgDisp
Dim vbMsgRes As VbMsgBoxResult
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objMaquinas As ClassMaquinas

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass

    'Se a data inicial não estiver preenchida
    If Len(Trim(DataInicial.ClipText)) = 0 Then gError 137373

    If Len(Trim(CodigoCT.Caption)) <> 0 Then
            
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
    
        objCentrodeTrabalho.sNomeReduzido = CodigoCT.Caption
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 137374
        
        objCTMaqProgDisp.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
    
    End If
    
    If Len(Maquina.Caption) > 0 Then
        
        Set objMaquinas = New ClassMaquinas
        
        objMaquinas.sNomeReduzido = Maquina.Caption
        
        'Le a Máquina no BD a partir do NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 137375
        
        objCTMaqProgDisp.lNumIntDocMaq = objMaquinas.lNumIntDoc
        
    End If

    objCTMaqProgDisp.dtData = StrParaDate(DataInicial.Text)
    
    'Se a data final não estiver preenchida
    If Len(Trim(DataFinal.ClipText)) = 0 Then
    
        'Pergunta ao usuário se confirma a exclusão
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_CTMAQPROGDISP", objCTMaqProgDisp.dtData)
    
        If vbMsgRes = vbNo Then
            GL_objMDIForm.MousePointer = vbDefault
            Exit Sub
        End If
    
        'Exclui CTMaquinaProgDisponibilidade - uma única vez
        lErro = CF("CTMaquinaProgDisponibilidade_Exclui", objCTMaqProgDisp)
        If lErro <> SUCESSO Then gError 137376
        
    Else
    
        objCTMaqProgDisp.dtDataDe = StrParaDate(DataInicial.Text)
        objCTMaqProgDisp.dtDataAte = StrParaDate(DataFinal.Text)
        
        'Pergunta ao usuário se confirma a exclusão
        vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_CONFIRMA_EXCLUSAO_PERIODO_CTMAQPROGDISP", objCTMaqProgDisp.dtDataDe, objCTMaqProgDisp.dtDataAte)
    
        If vbMsgRes = vbNo Then
            GL_objMDIForm.MousePointer = vbDefault
            Exit Sub
        End If
    
        'Exclui CTMaquinaProgDisponibilidade - uma única vez
        lErro = CF("CTMaquinaProgDisponibilidade_Exclui_Periodo", objCTMaqProgDisp)
        If lErro <> SUCESSO Then gError 137377
    
    End If

    'Limpa Tela
    Call Limpa_Tela_CTMaquinaProgDisponibilidade

    GL_objMDIForm.MousePointer = vbDefault

    Exit Sub

Erro_BotaoExcluir_Click:

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 137373
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
            DataInicial.SetFocus

        Case 137374 To 137377
            'erros tratados nas rotinas chamadas
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156077)

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
        If lErro <> SUCESSO Then gError 137378

        DataInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicial_DownClick:

    Select Case gErr

        Case 137378

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156078)

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
        If lErro <> SUCESSO Then gError 137379

        DataInicial.Text = sData

    End If

    Exit Sub

Erro_UpDownDataInicial_UpClick:

    Select Case gErr

        Case 137379

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156079)

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
        If lErro <> SUCESSO Then gError 137380
        
        'Se a data final também está preenchida
        If Len(Trim(DataFinal.ClipText)) <> 0 Then
        
            'Verifica qual é o intervalo entre as datas
            lIntervalo = DateDiff("d", StrParaDate(DataInicial.Text), StrParaDate(DataFinal.Text))
            
            'Se o intervalo for negativo -> Erro
            If lIntervalo < 0 Then gError 137381
        
        End If

    End If

    Exit Sub

Erro_DataInicial_Validate:

    Cancel = True

    Select Case gErr

        Case 137380
            'erro tratado na rotina chamada

        Case 137381
            Call Rotina_Erro(vbOKOnly, "ERRO_DATAINICIAL_MAIOR_DATAFINAL", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156080)

    End Select

    Exit Sub

End Sub

Private Sub DataInicial_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Quantidade_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Quantidade_Validate

    'Verifica se Quantidade está preenchida
    If Len(Trim(Quantidade.Text)) <> 0 Then

       'Critica a Quantidade
       lErro = Inteiro_Critica(Quantidade.Text)
       If lErro <> SUCESSO Then gError 137382

    End If

    Exit Sub

Erro_Quantidade_Validate:

    Cancel = True

    Select Case gErr

        Case 137382

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156081)

    End Select

    Exit Sub

End Sub

Private Sub Quantidade_GotFocus()
    
    Call MaskEdBox_TrataGotFocus(Quantidade, iAlterado)
    
End Sub

Private Sub Quantidade_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub Observacao_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub objEventoData_evSelecao(obj1 As Object)

Dim lErro As Long
Dim objCTMaqProgDisp As ClassCTMaqProgDisp

On Error GoTo Erro_objEventoData_evSelecao

    Set objCTMaqProgDisp = obj1

    'Mostra os dados do CTMaquinaProgDisp na tela
    lErro = Traz_CTMaquinaProgDisponibilidade_Tela(objCTMaqProgDisp)
    If lErro <> SUCESSO Then gError 137383

    Me.Show

    Exit Sub

Erro_objEventoData_evSelecao:

    Select Case gErr

        Case 137383
            'erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156082)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProgDisponibilidade_Click()

Dim lErro As Long
Dim objCTMaqProgDisp As New ClassCTMaqProgDisp
Dim colSelecao As New Collection
Dim objCentrodeTrabalho As ClassCentrodeTrabalho
Dim objMaquinas As ClassMaquinas
Dim sFiltro As String

On Error GoTo Erro_BotaoProgDisponibilidade_Click

    'Verifica se o Data foi preenchido
    If Len(Trim(DataInicial.ClipText)) <> 0 Then

        objCTMaqProgDisp.dtData = StrParaDate(DataInicial.Text)

    End If
    
    If Len(Trim(CodigoCT.Caption)) <> 0 Then
            
        Set objCentrodeTrabalho = New ClassCentrodeTrabalho
    
        objCentrodeTrabalho.sNomeReduzido = CodigoCT.Caption
        
        'Lê o CentrodeTrabalho que está sendo Passado
        lErro = CF("CentrodeTrabalho_Le_NomeReduzido", objCentrodeTrabalho)
        If lErro <> SUCESSO And lErro <> 134941 Then gError 137384
        
        objCTMaqProgDisp.lNumIntDocCT = objCentrodeTrabalho.lNumIntDoc
    
    End If
    
    If Len(Maquina.Caption) > 0 Then
        
        Set objMaquinas = New ClassMaquinas
        
        objMaquinas.sNomeReduzido = Maquina.Caption
        
        'Le a Máquina no BD a partir do NomeReduzido
        lErro = CF("Maquinas_Le_NomeReduzido", objMaquinas)
        If lErro <> SUCESSO And lErro <> 103100 Then gError 137385
        
        objCTMaqProgDisp.lNumIntDocMaq = objMaquinas.lNumIntDoc
        
    End If

    sFiltro = "NumIntDocCT = ? And NumIntDocMaq = ?"
    colSelecao.Add objCTMaqProgDisp.lNumIntDocCT
    colSelecao.Add objCTMaqProgDisp.lNumIntDocMaq

    Call Chama_Tela("CTMaquinaProgDispLista", colSelecao, objCTMaqProgDisp, objEventoData, sFiltro)

    Exit Sub

Erro_BotaoProgDisponibilidade_Click:

    Select Case gErr
    
        Case 137384, 137385
            'erros tratados nas rotinas chamadas

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 156083)

    End Select

    Exit Sub

End Sub

