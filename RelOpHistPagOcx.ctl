VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl RelOpHistPagOcx 
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5610
   LockControls    =   -1  'True
   ScaleHeight     =   960
   ScaleWidth      =   5610
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
      Left            =   4335
      Picture         =   "RelOpHistPagOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   105
      Width           =   1155
   End
   Begin MSMask.MaskEdBox FornecedorDesde 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   540
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   8
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox Fornecedor 
      Height          =   300
      Left            =   1215
      TabIndex        =   0
      Top             =   105
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   20
      PromptChar      =   "_"
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   540
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Label LabelFornecedor 
      Caption         =   "Fornecedor:"
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
      Height          =   255
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   5
      Top             =   135
      Width           =   1050
   End
   Begin VB.Label Label5 
      Caption         =   " Desde:"
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
      Height          =   255
      Left            =   465
      TabIndex        =   4
      Top             =   570
      Width           =   780
   End
End
Attribute VB_Name = "RelOpHistPagOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'William - 27/04/2001
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjRelatorio As AdmRelatorio
Dim gobjRelOpcoes As AdmRelOpcoes


Private WithEvents objEventoFornecedor As AdmEvento
Attribute objEventoFornecedor.VB_VarHelpID = -1


Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados fornecidos pelo usuário

Dim lErro As Long, dtDataAberto As Date

On Error GoTo Erro_PreencherRelOp

    lErro = ParcelaPag_Le_MenorData(dtDataAberto, LCodigo_Extrai(Fornecedor.Text))
    If lErro <> SUCESSO Then gError 87577

    If dtDataAberto <> DATA_NULA And StrParaDate(FornecedorDesde.Text) > dtDataAberto Then gError 87594
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 87578

    'Pegar parametros da tela
    lErro = objRelOpcoes.IncluirParametro("NFORNECEDOR", LCodigo_Extrai(Fornecedor.Text))
    If lErro <> AD_BOOL_TRUE Then gError 87579
    
    lErro = objRelOpcoes.IncluirParametro("TFORNECEDOR", Fornecedor.Text)
    If lErro <> AD_BOOL_TRUE Then gError 87580
        
    lErro = objRelOpcoes.IncluirParametro("DDATA", StrParaDate(FornecedorDesde.Text))
    If lErro <> AD_BOOL_TRUE Then gError 87581
            
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr

        Case 87594
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TITULOSPAG_ABERTO", gErr, dtDataAberto)
            
        Case 87578, 87579, 87580, 87581, 87577

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169372)

    End Select

    Exit Function

End Function

Public Sub Form_Unload(Cancel As Integer)

    Set objEventoFornecedor = Nothing
    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
End Sub

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 87582
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 87582
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 169373)

    End Select

    Exit Function

End Function

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    If Len(Trim(Fornecedor.Text)) = 0 Then gError 87583
    If Len(Trim(FornecedorDesde.ClipText)) = 0 Then gError 87584
    
    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 87585

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr
            
        Case 87583
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECEDOR_NAO_PREENCHIDO", gErr)
            Fornecedor.SetFocus
        
        Case 87584
            lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_NAO_PREENCHIDA", gErr)
            FornecedorDesde.SetFocus
                
        Case 87585
                
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169374)

    End Select
                
    Exit Sub

End Sub

Private Sub Fornecedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objFornecedor As New ClassFornecedor
Dim iCodFilial As Integer
Dim colCodigoNome As New AdmColCodigoNome
Dim iCria As Integer

On Error GoTo Erro_Fornecedor_Validate

        If Len(Trim(Fornecedor.Text)) > 0 Then

            iCria = 0 'Não deseja criar Fornecedor caso não exista
            lErro = TP_Fornecedor_Le2(Fornecedor, objFornecedor, iCria)
            If lErro <> SUCESSO Then gError 87586

        End If


    Exit Sub

Erro_Fornecedor_Validate:

    Cancel = True


    Select Case gErr

        Case 87586
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169375)

    End Select

    Exit Sub

End Sub

Private Sub FornecedorDesde_Validate(Cancel As Boolean)

Dim lErro As Long
Dim iCodigo As Integer
'Dim objFilialFornecedor As New ClassFilialFornecedor
Dim sFornecedor As String

On Error GoTo Erro_FornecedorDesde_Validate

    'Verifica se a data foi preenchida
    If Len(Trim(FornecedorDesde.ClipText)) = 0 Then Exit Sub
    
    'Verifica se é uma data válida
    lErro = Data_Critica(FornecedorDesde.Text)
    If lErro <> SUCESSO Then gError 87587
    
    'Verifica se a data informada é maoir que a data atual
    If StrParaDate(FornecedorDesde.Text) > gdtDataAtual Then gError 87588

    Exit Sub

Erro_FornecedorDesde_Validate:

    Cancel = True


    Select Case gErr

        Case 87587

        Case 87588
             lErro = Rotina_Erro(vbOKOnly, "ERRO_DATA_INFORMADA_MENOR_DATA_HOJE", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169376)

    End Select

    Exit Sub

End Sub

Public Sub Form_Load()

    Set objEventoFornecedor = New AdmEvento
    lErro_Chama_Tela = SUCESSO

End Sub

Private Sub LabelFornecedor_Click()

Dim objFornecedor As New ClassFornecedor
Dim colSelecao As Collection

    If Len(Trim(Fornecedor.Text)) > 0 Then
        'Preenche com o Fornecedor da tela
        objFornecedor.lCodigo = LCodigo_Extrai(Fornecedor.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("FornecedorLista", colSelecao, objFornecedor, objEventoFornecedor)

End Sub

Private Sub objEventoFornecedor_evSelecao(obj1 As Object)

Dim objFornecedor As ClassFornecedor

    Set objFornecedor = obj1
    
    'Preenche campo Cliente
    Fornecedor.Text = CStr(objFornecedor.lCodigo)
    Call Fornecedor_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub
    
End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_POSCLI
    Set Form_Load_Ocx = Me
    Caption = "Histórico de Pagamentos"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpHistPag"
    
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

Private Sub UpDown1_DownClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_DownClick

    lErro = Data_Up_Down_Click(FornecedorDesde, DIMINUI_DATA)
    If lErro <> SUCESSO Then gError 84460

    Exit Sub

Erro_UpDown1_DownClick:

    Select Case gErr

        Case 84460
             FornecedorDesde.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169377)

    End Select

    Exit Sub


End Sub

Private Sub UpDown1_UpClick()

Dim lErro As Long

On Error GoTo Erro_UpDown1_UpClick

    lErro = Data_Up_Down_Click(FornecedorDesde, AUMENTA_DATA)
    If lErro <> SUCESSO Then gError 84459

    Exit Sub

Erro_UpDown1_UpClick:

    Select Case gErr

        Case 84459
            FornecedorDesde.SetFocus

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169378)

    End Select

    Exit Sub

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
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If Me.ActiveControl Is Fornecedor Then
            Call LabelFornecedor_Click
        End If
    
    End If

End Sub


Private Sub LabelFornecedor_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelFornecedor, Source, X, Y)
End Sub

Private Sub LabelFornecedor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelFornecedor, Button, Shift, X, Y)
End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label5, Source, X, Y)
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label5, Button, Shift, X, Y)
End Sub


Public Function ParcelaPag_Le_MenorData(dtData As Date, lCodFornecedor As Long) As Long
'obtem a data de vencimento mais antiga de uma parcela a pagar em aberto de um fornecedor

Dim lErro As Long
Dim lComando As Long

On Error GoTo Erro_ParcelaPag_Le_MenorData

    'Abre Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then gError 87573
    
    'Le Último titulo mais antigo em aberto
    lErro = Comando_Executar(lComando, "SELECT MIN(ParcelasPag.DataVencimentoReal) FROM ParcelasPag, TitulosPag WHERE TitulosPag.NumIntDoc = ParcelasPag.NumIntTitulo AND ParcelasPag.Status = ? AND TitulosPag.Fornecedor = ? " _
    , dtData, STATUS_ABERTO, lCodFornecedor)
    If lErro <> AD_SQL_SUCESSO Then gError 87574

    'Busca o primeiro titulo
    lErro = Comando_BuscarPrimeiro(lComando)
    If lErro <> AD_SQL_SUCESSO And lErro <> AD_SQL_SEM_DADOS Then gError 87575
    
    'Fecha Comando
    Call Comando_Fechar(lComando)
    
    ParcelaPag_Le_MenorData = SUCESSO

    Exit Function

Erro_ParcelaPag_Le_MenorData:

    ParcelaPag_Le_MenorData = gErr
    
    Select Case gErr
    
        Case 87573
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", gErr)
                
        Case 87574, 87575
            lErro = Rotina_Erro(vbOKOnly, "ERRO_LEITURA_TITULOS_PAGAR", gErr)
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 169379)
    
    End Select
    
    'Fecha Comando --> saída por erro
    Call Comando_Fechar(lComando)
    
    Exit Function

End Function
