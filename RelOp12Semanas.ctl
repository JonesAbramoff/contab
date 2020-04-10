VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOp12SemanasOcx 
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6045
   LockControls    =   -1  'True
   ScaleHeight     =   3495
   ScaleWidth      =   6045
   Begin VB.Frame FrameVendedor 
      Caption         =   "Vendedor"
      Height          =   600
      Left            =   60
      TabIndex        =   21
      Top             =   2745
      Width           =   5910
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   300
         Left            =   1035
         TabIndex        =   6
         Top             =   195
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelVendedor 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
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
         Left            =   105
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   22
         Top             =   225
         Width           =   885
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Somente acima ou igual a"
      Height          =   690
      Left            =   3330
      TabIndex        =   19
      Top             =   1995
      Width           =   2640
      Begin VB.ComboBox Status 
         Height          =   315
         Left            =   735
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   255
         Width           =   1845
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
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
         Left            =   75
         TabIndex        =   20
         Top             =   285
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Agrupar or�amentos menores a partir de"
      Height          =   690
      Left            =   45
      TabIndex        =   17
      Top             =   1995
      Width           =   3210
      Begin MSMask.MaskEdBox Valor 
         Height          =   300
         Left            =   1050
         TabIndex        =   4
         Top             =   285
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         Format          =   "#,##0.00"
         PromptChar      =   " "
      End
      Begin VB.Label LabelClienteAte 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
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
         Left            =   480
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   18
         Top             =   345
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Considerar a partir de"
      Height          =   1140
      Left            =   45
      TabIndex        =   14
      Top             =   780
      Width           =   3210
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   2070
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   675
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSMask.MaskEdBox DataRel 
         Height          =   300
         Left            =   1065
         TabIndex        =   2
         Top             =   690
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   8
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Semana 
         Height          =   300
         Left            =   1065
         TabIndex        =   1
         Top             =   315
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Semana:"
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
         Left            =   255
         TabIndex        =   16
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Data 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   525
         TabIndex        =   15
         Top             =   750
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   3780
      ScaleHeight     =   495
      ScaleWidth      =   2115
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   180
      Width           =   2175
      Begin VB.CommandButton BotaoExcluir 
         Height          =   360
         Left            =   595
         Picture         =   "RelOp12Semanas.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Excluir"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoLimpar 
         Height          =   360
         Left            =   1125
         Picture         =   "RelOp12Semanas.ctx":018A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Limpar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   360
         Left            =   1605
         Picture         =   "RelOp12Semanas.ctx":06BC
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoGravar 
         Height          =   360
         Left            =   90
         Picture         =   "RelOp12Semanas.ctx":083A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Gravar"
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.CommandButton BotaoExecutar 
      Caption         =   "Executar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4425
      Picture         =   "RelOp12Semanas.ctx":0994
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   915
      Width           =   1515
   End
   Begin VB.ComboBox ComboOpcoes 
      Height          =   315
      ItemData        =   "RelOp12Semanas.ctx":0A96
      Left            =   900
      List            =   "RelOp12Semanas.ctx":0A98
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   323
      Width           =   2610
   End
   Begin VB.Label Label1 
      Caption         =   "Op��o:"
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
      Left            =   210
      TabIndex        =   13
      Top             =   353
      Width           =   555
   End
End
Attribute VB_Name = "RelOp12SemanasOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Property Variables:
Dim m_Caption As String
Event Unload()

Dim iSemanaAnt As Integer
Dim dtDataAnt As Date

Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Private Sub Form_Load()

Dim lErro As Long
Dim iSemana As Integer
Dim dtDataIniRel As Date
Dim colVend As New Collection
Dim objVend As ClassVendedor

On Error GoTo Erro_Form_Load

    Set objEventoVendedor = New AdmEvento

    Call CF("Rel12Semanas_Semana", Date, iSemana, dtDataIniRel)
    
    'mostra na tela a data de dia atual
    DataRel.PromptInclude = False
    DataRel.Text = Format(dtDataIniRel, "dd/mm/yy")
    DataRel.PromptInclude = True
    
    Semana.PromptInclude = False
    Semana.Text = CStr(iSemana)
    Semana.PromptInclude = True
    
    Call Carrega_Status(Status)
    
    lErro = CF("VendedorAtivo_Le_Todos", colVend)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    For Each objVend In colVend
        If objVend.sCodUsuario = gsUsuario Then
            FrameVendedor.Enabled = False
            Vendedor.Text = CStr(objVend.iCodigo)
            Call Vendedor_Validate(bSGECancelDummy)
            Exit For
        End If
    Next
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208011)

    End Select

    Exit Sub

End Sub

Function PreencherParametrosNaTela(objRelOpcoes As AdmRelOpcoes) As Long
'l� os par�metros do arquivo C e exibe na tela

Dim lErro As Long
Dim sParam As String

On Error GoTo Erro_PreencherParametrosNaTela

    lErro = objRelOpcoes.Carregar
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
   
    lErro = objRelOpcoes.ObterParametro("DDATA", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call DateParaMasked(DataRel, CDate(sParam))
    
    lErro = objRelOpcoes.ObterParametro("NSEMANA", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Semana.PromptInclude = False
    Semana.Text = sParam
    Semana.PromptInclude = True
    
    lErro = objRelOpcoes.ObterParametro("NVALOR", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Valor.Text = sParam
    
    lErro = objRelOpcoes.ObterParametro("NVENDEDOR", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If StrParaInt(sParam) > 0 Then
        Vendedor.Text = sParam
        Call Vendedor_Validate(bSGECancelDummy)
    End If
    
    lErro = objRelOpcoes.ObterParametro("NSTATUS", sParam)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call Combo_Seleciona_ItemData(Status, StrParaInt(sParam))
    
    PreencherParametrosNaTela = SUCESSO
    
    Exit Function

Erro_PreencherParametrosNaTela:

    PreencherParametrosNaTela = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208012)

    End Select

    Exit Function

End Function

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 208013
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    'Preenche com as Opcoes
    lErro = RelOpcoes_ComboOpcoes_Preenche(objRelatorio, ComboOpcoes, objRelOpcoes, Me)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr
        
        Case 208013
            Call Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208014)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()

    Unload Me

End Sub

Private Function Formata_E_Critica_Parametros() As Long
'Formata os produtos retornando em sProd_I e sProd_F
'Verifica se os par�metros iniciais s�o maiores que os finais

Dim lErro As Long

On Error GoTo Erro_Formata_E_Critica_Parametros

    If Len(DataRel.ClipText) = 0 Then gError 208015
    
    Formata_E_Critica_Parametros = SUCESSO

    Exit Function

Erro_Formata_E_Critica_Parametros:

    Formata_E_Critica_Parametros = gErr

    Select Case gErr
            
        Case 208015
            Call Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208016)

    End Select

    Exit Function

End Function

Private Sub BotaoLimpar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoLimpar_Click
  
    lErro = Limpa_Relatorio(Me)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    ComboOpcoes.Text = ""
   
    ComboOpcoes.SetFocus
    
    Status.ListIndex = -1
    
    Exit Sub
    
Erro_BotaoLimpar_Click:
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208017)

    End Select

    Exit Sub
   
End Sub

Private Sub ComboOpcoes_Click()

    Call RelOpcoes_ComboOpcoes_Click(gobjRelOpcoes, ComboOpcoes, Me)
    
End Sub

Private Sub ComboOpcoes_Validate(Cancel As Boolean)

    Call RelOpcoes_ComboOpcoes_Validate(ComboOpcoes, Cancel)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoVendedor = Nothing
    
End Sub

Private Sub DataRel_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataRel_Validate

    If Len(DataRel.ClipText) > 0 Then
        
        lErro = Data_Critica(DataRel.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    If dtDataAnt <> StrParaDate(DataRel.Text) Then
        dtDataAnt = StrParaDate(DataRel.Text)
        Call Ajusta_Semana
    End If
    
    Exit Sub

Erro_DataRel_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208018)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes, Optional ByVal bExecutando As Boolean = False) As Long
'preenche o arquivo C com os dados fornecidos pelo usu�rio

Dim lErro As Long
Dim lNumIntRel As Long
Dim lStatus As Long

On Error GoTo Erro_PreencherRelOp
      
    lErro = Formata_E_Critica_Parametros()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
             
    lErro = objRelOpcoes.IncluirParametro("NVALOR", Format(Valor.Text, "STANDARD"))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("NSEMANA", Semana.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("DDATA", DataRel.Text)
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    lErro = objRelOpcoes.IncluirParametro("NVENDEDOR", CStr(Codigo_Extrai(Vendedor.Text)))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If Status.ListIndex <> -1 Then
        lStatus = Status.ItemData(Status.ListIndex)
    Else
        lStatus = 0
    End If
    
    lErro = objRelOpcoes.IncluirParametro("NSTATUS", CStr(lStatus))
    If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    If bExecutando Then
    
        lErro = CF("Rel12Semanas_Prepara", lNumIntRel, giFilialEmpresa, StrParaDate(DataRel.Text), StrParaDbl(Valor.Text), lStatus, Codigo_Extrai(Vendedor.Text))
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        lErro = objRelOpcoes.IncluirParametro("NNUMINTREL", CStr(lNumIntRel))
        If lErro <> AD_BOOL_TRUE Then gError ERRO_SEM_MENSAGEM
    
    End If
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208019)

    End Select

    Exit Function

End Function

Private Sub BotaoExcluir_Click()

Dim vbMsgRes As VbMsgBoxResult
Dim lErro As Long

On Error GoTo Erro_BotaoExcluir_Click

    'verifica se nao existe elemento selecionado na ComboBox
    If ComboOpcoes.ListIndex = -1 Then gError 208020

    vbMsgRes = Rotina_Aviso(vbYesNo, "AVISO_EXCLUSAO_RELOPRAZAO")

    If vbMsgRes = vbYes Then

        lErro = CF("RelOpcoes_Exclui", gobjRelOpcoes)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        'retira nome das op��es do ComboBox
        ComboOpcoes.RemoveItem ComboOpcoes.ListIndex

        'limpa as op��es da tela
         lErro = Limpa_Relatorio(Me)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
        ComboOpcoes.Text = ""
    
    End If

    Exit Sub

Erro_BotaoExcluir_Click:

    Select Case gErr

        Case 208020
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_NAO_SELEC", gErr)
            ComboOpcoes.SetFocus

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208021)

    End Select

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes, True)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208022)

    End Select

    Exit Sub

End Sub

Private Sub BotaoGravar_Click()
'Grava a op��o de relat�rio com os par�metros da tela

Dim lErro As Long
Dim iResultado As Integer

On Error GoTo Erro_BotaoGravar_Click

    'nome da op��o de relat�rio n�o pode ser vazia
    If ComboOpcoes.Text = "" Then gError 208023

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    gobjRelOpcoes.sNome = ComboOpcoes.Text

    lErro = CF("RelOpcoes_Grava", gobjRelOpcoes, iResultado)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    lErro = RelOpcoes_Testa_Combo(ComboOpcoes, gobjRelOpcoes.sNome)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    Call BotaoLimpar_Click
    
    Exit Sub

Erro_BotaoGravar_Click:

    Select Case gErr

        Case 208023
            Call Rotina_Erro(vbOKOnly, "ERRO_NOME_RELOP_VAZIO", gErr)
            ComboOpcoes.SetFocus

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208024)

    End Select

    Exit Sub

End Sub

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes) As Long
'monta a express�o de sele��o de relat�rio

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208025)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_ANALISE_ESTOQUE
    Set Form_Load_Ocx = Me
    Caption = "12 Semanas"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOp12Semanas"
    
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

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = KEYCODE_BROWSER Then

    End If

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Sub Unload(objme As Object)
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

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick
   
    lErro = Data_Up_Down_Click(DataRel, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If dtDataAnt <> StrParaDate(DataRel.Text) Then
        dtDataAnt = StrParaDate(DataRel.Text)
        Call Ajusta_Semana
    End If
    
    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataRel.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208026)

    End Select

    Exit Sub

End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(DataRel, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    If dtDataAnt <> StrParaDate(DataRel.Text) Then
        dtDataAnt = StrParaDate(DataRel.Text)
        Call Ajusta_Semana
    End If
    
    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM
            DataRel.SetFocus

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 208027)

    End Select

    Exit Sub

End Sub

Private Sub Valor_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_Valor_Validate

    'Veifica se Valor est� preenchida
    If Len(Trim(Valor.Text)) <> 0 Then

       'Critica a Valor
       lErro = Valor_Positivo_Critica(Valor.Text)
       If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If

    Exit Sub

Erro_Valor_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208028)

    End Select

    Exit Sub

End Sub

Private Sub Valor_GotFocus()
    Dim iAlterado As Integer
    Call MaskEdBox_TrataGotFocus(Valor, iAlterado)
    
End Sub

Private Sub Semana_GotFocus()
    Dim iAlterado As Integer
    Call MaskEdBox_TrataGotFocus(Semana, iAlterado)
End Sub

Private Sub Semana_Validate(Cancel As Boolean)
    If iSemanaAnt <> StrParaInt(Semana.Text) Then
        iSemanaAnt = StrParaInt(Semana.Text)
        Call Ajusta_Data
    End If
End Sub

Private Sub Ajusta_Semana()

Dim iSemana As Integer
Dim dtDataIniRel As Date

    If StrParaDate(DataRel.Text) <> DATA_NULA Then
    
        Call CF("Rel12Semanas_Semana", StrParaDate(DataRel.Text), iSemana, dtDataIniRel)
    
        Semana.PromptInclude = False
        Semana.Text = CStr(iSemana)
        Semana.PromptInclude = True
        
        iSemanaAnt = iSemana
        
    End If

End Sub

Private Sub Ajusta_Data()

Dim iSemanaAux As Integer
Dim iSemana As Integer
Dim dtDataIniRel As Date
Dim dtData As Date

    iSemanaAux = StrParaInt(Semana.Text)
    
    If iSemanaAux <> 0 Then

        Call CF("Rel12Semanas_Semana", StrParaDate(DataRel.Text), iSemana, dtDataIniRel)
        
        dtData = DateAdd("d", (iSemanaAux - iSemana) * 7, dtDataIniRel)
        
        If Year(dtDataIniRel) > Year(dtData) Then
            dtData = StrParaDate("01/01/" & CStr(Year(dtDataIniRel)))
        End If
    
        Call DateParaMasked(DataRel, dtData)
    
    End If
    
End Sub

Private Function Carrega_Status(ByVal objComboBox As ComboBox) As Long
'Carrega a combo de Tipo

Dim lErro As Long

On Error GoTo Erro_Carrega_Status

    'carregar tipos de desconto
    lErro = CF("Carrega_CamposGenericos", CAMPOSGENERICOS_STATUSOV, objComboBox)
    If lErro <> SUCESSO Then gError 141371

    objComboBox.AddItem ""
    objComboBox.ItemData(objComboBox.NewIndex) = 0
    
    objComboBox.ListIndex = -1
    
    Carrega_Status = SUCESSO

    Exit Function

Erro_Carrega_Status:

    Carrega_Status = gErr

    Select Case gErr
    
        Case 141371

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157851)

    End Select

    Exit Function

End Function

Private Sub LabelVendedor_Click()

Dim objVendedor As New ClassVendedor
Dim colSelecao As Collection
   
    If Len(Trim(Vendedor.Text)) > 0 Then
        'Preenche com o Vendedor da tela
        objVendedor.iCodigo = Codigo_Extrai(Vendedor.Text)
    End If
    
    'Chama Tela VendedorLista
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1
    
    Vendedor.Text = CStr(objVendedor.iCodigo)
    Call Vendedor_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_Vendedor_Validate

    If Len(Trim(Vendedor.Text)) > 0 Then
   
        'Tenta ler o vendedor (NomeReduzido ou C�digo)
        lErro = TP_Vendedor_Le2(Vendedor, objVendedor, 0)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    End If
    
    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 167058)

    End Select

End Sub
