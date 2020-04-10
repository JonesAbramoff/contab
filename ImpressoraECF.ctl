VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl ImpressoraECF 
   ClientHeight    =   2970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   LockControls    =   -1  'True
   ScaleHeight     =   2970
   ScaleWidth      =   6210
   Begin VB.CommandButton BotaoProxNum 
      Height          =   285
      Left            =   1860
      Picture         =   "ImpressoraECF.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Numeração Automática"
      Top             =   255
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3945
      ScaleHeight     =   495
      ScaleWidth      =   2085
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   90
      Width           =   2145
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "ImpressoraECF.ctx":00EA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "ImpressoraECF.ctx":0244
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1100
         Picture         =   "ImpressoraECF.ctx":03CE
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "ImpressoraECF.ctx":0900
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impressora"
      Height          =   1545
      Left            =   90
      TabIndex        =   5
      Top             =   1290
      Width           =   5970
      Begin MSMask.MaskEdBox CodImp 
         Height          =   315
         Left            =   1215
         TabIndex        =   7
         Top             =   345
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         Mask            =   "####"
         PromptChar      =   " "
      End
      Begin VB.Label LabelCodImp 
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
         Left            =   465
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   6
         Top             =   405
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Modelo:"
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
         Left            =   3150
         TabIndex        =   10
         Top             =   1020
         Width           =   690
      End
      Begin VB.Label LabelModelo 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3930
         TabIndex        =   11
         Top             =   975
         Width           =   1785
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fabricante:"
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
         Left            =   150
         TabIndex        =   8
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label LabelFabricante 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1215
         TabIndex        =   9
         Top             =   975
         Width           =   1785
      End
   End
   Begin VB.TextBox NumSerie 
      Height          =   315
      Left            =   1245
      MaxLength       =   50
      TabIndex        =   4
      Top             =   840
      Width           =   2595
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   315
      Left            =   1245
      TabIndex        =   1
      Top             =   240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   " "
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Num. Serie:"
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
      Left            =   150
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   3
      Top             =   885
      Width           =   1005
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
      Left            =   495
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   0
      Top             =   285
      Width           =   660
   End
End
Attribute VB_Name = "ImpressoraECF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iAlterado As Integer

Private WithEvents objEventoImp As AdmEvento
Attribute objEventoImp.VB_VarHelpID = -1
Private WithEvents objEventoModelo As AdmEvento
Attribute objEventoModelo.VB_VarHelpID = -1

Option Explicit

Public Function Gravar_Registro() As Long

Dim lErro As Long
Dim objImpressoraECF As New ClassImpressoraECF

On Error GoTo Erro_Gravar_Registro

    GL_objMDIForm.MousePointer = vbHourglass

    If Len(Trim(Codigo.Text)) = 0 Then gError 103466
    If Len(Trim(NumSerie.Text)) = 0 Then gError 103467
    If Len(Trim(CodImp.Text)) = 0 Then gError 103468

    lErro = Move_Tela_Memoria(objImpressoraECF)
    If lErro <> AD_SQL_SUCESSO Then gError 103469

    lErro = Trata_Alteracao(objImpressoraECF, objImpressoraECF.iCodigo)
    If lErro <> SUCESSO Then gError 103470

    lErro = CF("ImpressoraECF_Grava", objImpressoraECF)
    If lErro <> SUCESSO Then gError 103471

    GL_objMDIForm.MousePointer = vbDefault

    Gravar_Registro = SUCESSO

    Exit Function

Erro_Gravar_Registro:

    Gravar_Registro = gErr

    GL_objMDIForm.MousePointer = vbDefault

    Select Case gErr

        Case 103466
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)

        Case 103467
            Call Rotina_Erro(vbOKOnly, "ERRO_NUMSERIE_NAO_PREENCHIDA", gErr)

        Case 103468
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIMPRESSORA_NAO_PREENCHIDA", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161861)

    End Select

    Exit Function

End Function

Function Move_Tela_Memoria(objImp As ClassImpressoraECF) As Long
'Guarda no objECF os dados informados na tela

On Error GoTo Erro_Move_Tela_Memoria

    'Guarda Codigo, FilialEmpresa
    objImp.iCodigo = StrParaInt(Codigo.Text)
    objImp.iFilialEmpresa = giFilialEmpresa
    objImp.iCodModelo = StrParaInt(CodImp.Text)
    objImp.sNumSerie = NumSerie.Text
    
    Move_Tela_Memoria = SUCESSO

    Exit Function

Erro_Move_Tela_Memoria:

    Move_Tela_Memoria = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161862)

    End Select

    Exit Function

End Function

Sub Limpa_Tela_ImpressoraECF()
    
    Call Limpa_Tela(Me)
    
    LabelFabricante.Caption = ""
    LabelModelo.Caption = ""

    'Fecha o comando das setas se estiver aberto
    Call ComandoSeta_Fechar(Me.Name)

    iAlterado = 0

End Sub

Function Traz_ImpressoraECF_Tela(objImpressoraECF As ClassImpressoraECF) As Long

Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Traz_ImpressoraECF_Tela

    lErro = CF("ImpressoraECF_Le", objImpressoraECF)
    If lErro <> SUCESSO And lErro <> 103447 Then gError 103457

    If lErro <> SUCESSO Then gError 103458

    Call Limpa_Tela_ImpressoraECF

    'Traz os dados para tela
    Codigo.Text = objImpressoraECF.iCodigo
    NumSerie.Text = objImpressoraECF.sNumSerie

    CodImp.Text = objImpressoraECF.iCodModelo
    Call CodImp_Validate(bSGECancelDummy)

    iAlterado = 0

    Exit Function

    Traz_ImpressoraECF_Tela = SUCESSO

Erro_Traz_ImpressoraECF_Tela:

    Traz_ImpressoraECF_Tela = gErr

    Select Case gErr

        Case 103457

        Case 103458
            'Tratado na Rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161863)

    End Select

    Exit Function

End Function

Public Sub Tela_Extrai(sTabela As String, colCampoValor As AdmColCampoValor, colSelecao As AdmColFiltro)
'Extrai os campos da tela que correspondem aos campos no BD

Dim lErro As Long
Dim objImpressoraECF As New ClassImpressoraECF

On Error GoTo Erro_Tela_Extrai

    'Informa tabela associada à Tela
    sTabela = "ImpressoraECF"

    'Le os dados da Tela ImpressoraECF
    lErro = Move_Tela_Memoria(objImpressoraECF)
    If lErro <> SUCESSO Then gError 103450

    'Preenche a coleção colCampoValor, com nome do campo,
    'valor atual (com a tipagem do BD), tamanho do campo
    'no BD no caso de STRING e Key igual ao nome do campo
    colCampoValor.Add "Codigo", objImpressoraECF.iCodigo, 0, "Codigo"
    colCampoValor.Add "CodModelo", objImpressoraECF.iCodModelo, 0, "CodModelo"
    colCampoValor.Add "NumSerie", objImpressoraECF.sNumSerie, STRING_IMPRESSORAECF_NUMSERIE, "NumSerie"

    'Filtros para o Sistema de Setas
    colSelecao.Add "FilialEmpresa", OP_IGUAL, objImpressoraECF.iFilialEmpresa

    Exit Sub

Erro_Tela_Extrai:

    Select Case gErr

        Case 103450
        'Erro tratado na rotina chamadora
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161864)

    End Select

    Exit Sub

End Sub

Public Sub Tela_Preenche(colCampoValor As AdmColCampoValor)
'Preenche os campos da tela com os correspondentes do BD

Dim lErro As Long
Dim objImpressoraECF As New ClassImpressoraECF

On Error GoTo Erro_Tela_Preenche

    objImpressoraECF.iCodigo = colCampoValor.Item("Codigo").vValor
    objImpressoraECF.iCodModelo = colCampoValor.Item("CodModelo").vValor
    objImpressoraECF.sNumSerie = colCampoValor.Item("NumSerie").vValor

    If objImpressoraECF.iCodigo <> 0 Then
        
        objImpressoraECF.iFilialEmpresa = giFilialEmpresa
        
        'Traz dados do ImpressoraECF para a Tela
        lErro = Traz_ImpressoraECF_Tela(objImpressoraECF)
        If lErro <> SUCESSO And lErro <> 103458 Then gError 103451

    End If

    iAlterado = 0

    Exit Sub

Erro_Tela_Preenche:

    Select Case gErr

        Case 103451

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161865)

    End Select

    Exit Sub

End Sub
Public Sub Form_Activate()

    Call TelaIndice_Preenche(Me)

End Sub

Public Sub Form_Deactivate()

    gi_ST_SetaIgnoraClick = 1

End Sub


Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoImp = New AdmEvento
    Set objEventoModelo = New AdmEvento
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161866)

    End Select

    iAlterado = 0

    Exit Sub

End Sub

Function Trata_Parametros(Optional objImpressoraECF As ClassImpressoraECF) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    'Se houver impressora ECF passada como parâmetro
    If Not (objImpressoraECF Is Nothing) Then

        If objImpressoraECF.iCodigo <> 0 Then

            objImpressoraECF.iFilialEmpresa = giFilialEmpresa

            'Lê a Impressora no BD a partir do código
            lErro = CF("ImpressoraECF_Le", objImpressoraECF)
            If lErro <> SUCESSO And lErro <> 103447 Then gError 103448

            If lErro = SUCESSO Then

                'Exibe os dados da Impressora
                lErro = Traz_ImpressoraECF_Tela(objImpressoraECF)
                If lErro <> SUCESSO And lErro <> 103458 Then gError 103449

            Else
                Codigo.Text = objImpressoraECF.iCodigo

            End If

        End If

    End If

    iAlterado = 0

    Exit Function

    Trata_Parametros = SUCESSO

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 103448, 103449

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161867)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim lErro As Long
Dim objImp As New ClassImpressoraECF
Dim vbMsgRes As VbMsgBoxResult

On Error GoTo Erro_BotaoExcluir_Click

    GL_objMDIForm.MousePointer = vbHourglass
    
    'Se o código não foi preenchido = > erro
    If Len(Trim(Codigo.ClipText)) = 0 Then gError 112780

    'Guarda em objECF o código que será passado como parâmtero para ECF_Le
    objImp.iCodigo = StrParaInt(Codigo.Text)
    objImp.iFilialEmpresa = giFilialEmpresa
    
    'Lê no BD os dados do ECF que será excluído
    lErro = CF("ImpressoraECF_Le", objImp)
    If lErro <> SUCESSO And lErro <> 103447 Then gError 112781

    'Se o ECF não estiver cadastrado => erro
    If lErro = 103447 Then gError 112782
    
    'Envia aviso perguntando se realmente deseja excluir ECF
    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUIR_IMPRESSORA", objImp.iCodigo)

    If vbMsgRes = vbYes Then
    'Se sim
    
        'Chama a Função Move_Tela_Memória
        lErro = Move_Tela_Memoria(objImp)
        If lErro <> SUCESSO Then gError 112783
        
        'Chama a função que irá excluir o ECF
        lErro = CF("ImpressoraECF_Exclui", objImp)
        If lErro <> SUCESSO Then gError 112784

        'Limpa a Tela
        Call Limpa_Tela_ImpressoraECF
                
    End If

    iAlterado = 0
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr
        
        Case 112780
            Call Rotina_Erro(vbOKOnly, "ERRO_CODIGO_NAO_PREENCHIDO", gErr)
        
        Case 112781, 112783, 112784
            'Erro Tratado dentro da função chamadora

        Case 112782
            Call Rotina_Erro(vbOKOnly, "ERRO_IMPRESSORA_NAO_CADASTRADA", gErr, objImp.iCodigo)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161868)

    End Select

    GL_objMDIForm.MousePointer = vbDefault
    
    Exit Sub

End Sub

Private Sub BotaoFechar_Click()

    Unload Me
    
End Sub

Private Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click

    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then gError 103465

    Call Limpa_Tela_ImpressoraECF

    iAlterado = 0

    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 103465
            'Erro tratado na rotina chamadora

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161869)

    End Select

    Exit Sub

End Sub

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_Botaolimpar_Click

    lErro = Teste_Salva(Me, iAlterado)
    If lErro <> SUCESSO Then gError 103453

    Call Limpa_Tela_ImpressoraECF

    Exit Sub

Erro_Botaolimpar_Click:

    Select Case gErr

        Case 103453

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161870)

    End Select

End Sub

Private Sub BotaoProxNum_Click()

Dim lErro As Long
Dim iCodigo As Integer

On Error GoTo Erro_BotaoProxNum_Click

    'Chama a função que gera o sequencial do Código Automático para uma nova Impressora
    lErro = CF("Config_Obter_Inteiro_Automatico", "LojaConfig", "NUM_PROXIMO_IMPRESSORAECF", "ImpressoraECF", "Codigo", iCodigo)

    If lErro <> SUCESSO Then gError 103452

    'Exibe o novo código na tela
    Codigo.Text = CStr(iCodigo)

    Exit Sub

Erro_BotaoProxNum_Click:

    Select Case gErr

        Case 103452
            'Erro tratado na rotina chamada

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161871)

    End Select

    Exit Sub

End Sub

Private Sub LabelCodigo_Click()

Dim objImp As New ClassImpressoraECF
Dim colSelecao As Collection
    
    If Len(Trim(Codigo.Text)) <> 0 Then
        objImp.iCodigo = StrParaInt(Codigo.Text)
    End If
    
    Call Chama_Tela("ImpressoraECFLista", colSelecao, objImp, objEventoImp)

    Exit Sub

End Sub

Private Sub objEventoImp_evSelecao(obj1 As Object)

Dim objImp As ClassImpressoraECF
Dim lErro As Long
Dim iIndex As Integer

On Error GoTo Erro_objEventoImp_evSelecao

    Set objImp = obj1
    
    lErro = Traz_ImpressoraECF_Tela(objImp)
    If lErro <> SUCESSO Then gError 118003
    
    Me.Show
        
    Exit Sub

Erro_objEventoImp_evSelecao:
    
    Select Case gErr
                
        Case 118003
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161872)

    End Select
    
    Exit Sub

End Sub

Private Sub LabelCodImp_Click()

Dim objModelo As New ClassModeloECF
Dim colSelecao As Collection
    
    If Len(Trim(CodImp.Text)) <> 0 Then
        objModelo.iCodigo = StrParaInt(CodImp.Text)
    End If
    
    Call Chama_Tela("ModeloECFLista", colSelecao, objModelo, objEventoModelo)

    Exit Sub

End Sub

Private Sub objEventoModelo_evSelecao(obj1 As Object)

Dim objModelo As ClassModeloECF
Dim lErro As Long
Dim iIndex As Integer

On Error GoTo Erro_objEventoModelo_evSelecao

    Set objModelo = obj1
    
    CodImp.Text = objModelo.iCodigo
    Call CodImp_Validate(False)
    
    Me.Show
        
    Exit Sub

Erro_objEventoModelo_evSelecao:
    
    Select Case gErr
                
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 161873)

    End Select
    
    Exit Sub

End Sub

Public Sub form_unload(Cancel As Integer)

Dim lErro As Long
    
    Set objEventoImp = Nothing
    
    'Fecha o comando de setas se estiver aberto
    lErro = ComandoSeta_Liberar(Me.Name)

End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    '??? Parent.HelpContextID = IDH_
    Set Form_Load_Ocx = Me
    Caption = "Impressoras ECF"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ImpressoraECF"

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

Private Sub Codigo_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub NumSerie_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodImp_Change()

    iAlterado = REGISTRO_ALTERADO

End Sub

Private Sub CodImp_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objModeloECF As New ClassModeloECF

On Error GoTo Erro_CodImp_Validate

    'se o código estiver preenchido
    If Len(Trim(CodImp.Text)) = 0 Then Exit Sub

    'preenche os atributos para buscar o ModeloECF
    objModeloECF.iCodigo = StrParaInt(CodImp.Text)
    
    'busca na tabela ModeloECF
    lErro = CF("ModeloECF_Le", objModeloECF)
    If lErro <> SUCESSO And lErro <> 103459 Then gError 103460

    If lErro <> SUCESSO Then

        LabelFabricante.Caption = ""
        LabelModelo.Caption = ""

        gError 103461

    Else

        LabelFabricante.Caption = objModeloECF.sFabricante
        LabelModelo.Caption = objModeloECF.sNome

    End If

    Exit Sub

Erro_CodImp_Validate:

    Cancel = True

    Select Case gErr

        Case 103460

        Case 103461
            Call Rotina_Erro(vbOKOnly, "ERRO_MODELOECF_NAO_CADASTRADO", gErr)

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 161874)

    End Select

    Exit Sub

End Sub

