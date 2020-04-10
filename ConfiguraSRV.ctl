VERSION 5.00
Begin VB.UserControl ConfiguraSRV 
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6450
   ScaleHeight     =   2010
   ScaleWidth      =   6450
   Begin VB.ListBox ListaConfigura 
      Height          =   1635
      ItemData        =   "ConfiguraSRV.ctx":0000
      Left            =   135
      List            =   "ConfiguraSRV.ctx":0019
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   120
      Width           =   4320
   End
   Begin VB.PictureBox Picture5 
      Height          =   555
      Left            =   4695
      ScaleHeight     =   495
      ScaleWidth      =   1110
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   105
      Width           =   1170
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   600
         Picture         =   "ConfiguraSRV.ctx":0162
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fechar"
         Top             =   75
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ConfiguraSRV.ctx":02E0
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Gravar"
         Top             =   75
         Width           =   420
      End
   End
End
Attribute VB_Name = "ConfiguraSRV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Dim iAlterado As Integer

Const SRVCONFIG_AGLUTINA_LANCAM_POR_DIA = 0
Const SRVCONFIG_GERA_LOTE_AUTOMATICO = 1
Const SRVCONFIG_VALIDA_GARANTIA = 2
Const SRVCONFIG_VALIDA_MANUTENCAO = 3
Const SRVCONFIG_GARANTIA_AUTOMATICA_SOLICITACAO = 4
Const SRVCONFIG_CONTRATO_AUTOMATICO_SOLICITACAO = 5
Const SRVCONFIG_VERIFICA_LOTE = 6

Dim m_Caption As String


Private Sub BotaoFechar_Click()
    Unload Me
End Sub

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183410)

    End Select

    Exit Function

End Function

Public Sub Form_Load()
           
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Checa Aglutina lançamentos por dia
    If gobjSRV.iAglutinaLancamDia = AGLUTINA_LANCAM_POR_DIA Then
        ListaConfigura.Selected(SRVCONFIG_AGLUTINA_LANCAM_POR_DIA) = True
    Else
        ListaConfigura.Selected(SRVCONFIG_AGLUTINA_LANCAM_POR_DIA) = False
    End If
    
    'Checa Gera Lote Automatico
    If gobjSRV.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO Then
        ListaConfigura.Selected(SRVCONFIG_GERA_LOTE_AUTOMATICO) = True
    Else
        ListaConfigura.Selected(SRVCONFIG_GERA_LOTE_AUTOMATICO) = False
    End If
            
    'Checa Validacao de Garantia
    If gobjSRV.iValidaGarantia = VALIDA_GARANTIA Then
        ListaConfigura.Selected(SRVCONFIG_VALIDA_GARANTIA) = True
    Else
        ListaConfigura.Selected(SRVCONFIG_VALIDA_GARANTIA) = False
    End If
            
    'Checa Validacao de Manutencao
    If gobjSRV.iValidaGarantia = VALIDA_MANUTENCAO Then
        ListaConfigura.Selected(SRVCONFIG_VALIDA_MANUTENCAO) = True
    Else
        ListaConfigura.Selected(SRVCONFIG_VALIDA_MANUTENCAO) = False
    End If
            
    If gobjSRV.iValidaGarantia = VERIFICA_LOTE Then
        ListaConfigura.Selected(SRVCONFIG_VERIFICA_LOTE) = True
    Else
        ListaConfigura.Selected(SRVCONFIG_VERIFICA_LOTE) = False
    End If
            
    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183411)

    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 183412
    
    iAlterado = 0
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case gErr

        Case 183412

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183413)

    End Select

    Exit Sub
    
End Sub

Public Function Gravar_Registro() As Long

Dim lErro As Long
    
On Error GoTo Erro_Gravar_Registro
    
    'Move os Dados para Memoria
    If ListaConfigura.Selected(SRVCONFIG_AGLUTINA_LANCAM_POR_DIA) = True Then
        gobjSRV.iAglutinaLancamDia = AGLUTINA_LANCAM_POR_DIA
    Else
        gobjSRV.iAglutinaLancamDia = NAO_AGLUTINA_LANCAM_POR_DIA
    End If
    
    If ListaConfigura.Selected(SRVCONFIG_GERA_LOTE_AUTOMATICO) = True Then
        gobjSRV.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO
    Else
        gobjSRV.iGeraLoteAutomatico = NAO_GERA_LOTE_AUTOMATICO
    End If
    
    If ListaConfigura.Selected(SRVCONFIG_VALIDA_GARANTIA) = True Then
        gobjSRV.iValidaGarantia = VALIDA_GARANTIA
    Else
        gobjSRV.iValidaGarantia = NAO_VALIDA_GARANTIA
    End If
    
    If ListaConfigura.Selected(SRVCONFIG_VALIDA_MANUTENCAO) = True Then
        gobjSRV.iValidaManutencao = VALIDA_MANUTENCAO
    Else
        gobjSRV.iValidaManutencao = NAO_VALIDA_MANUTENCAO
    End If
    
    
    'Grava na Tabela  "CPConfig" as Configurações
    lErro = gobjSRV.Gravar()
    If lErro <> SUCESSO Then gError 183414
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = gErr

    Select Case gErr
    
        Case 183414
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 183415)
            
    End Select

    Exit Function
    
End Function

Private Sub ListaConfigura_ItemCheck(Item As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)
    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)
End Sub

Public Function Form_Load_Ocx() As Object

    Set Form_Load_Ocx = Me
    Caption = "Configuração do Módulo de Serviços"
    Call Form_Load

End Function

Public Sub Form_Unload(Cancel As Integer)

End Sub

Public Function Name() As String
    Name = "ConfiguraSRV"
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

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Parent.Caption = New_Caption
    m_Caption = New_Caption
End Property



