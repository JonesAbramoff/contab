VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.UserControl RelOpTitAtrasoOcx 
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5940
   ScaleHeight     =   2550
   ScaleWidth      =   5940
   Begin VB.Frame Frame7 
      Caption         =   "Vendedor"
      Height          =   720
      Left            =   255
      TabIndex        =   8
      Top             =   1035
      Width           =   5355
      Begin MSMask.MaskEdBox Vendedor 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   20
         PromptChar      =   "_"
      End
      Begin VB.Label VendedorLabel 
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
         Left            =   210
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   9
         Top             =   300
         Width           =   885
      End
   End
   Begin VB.CommandButton BotaoFechar 
      Height          =   600
      Left            =   3165
      Picture         =   "RelOpTitAtrasoOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Fechar"
      Top             =   1815
      Width           =   1575
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
      Left            =   885
      Picture         =   "RelOpTitAtrasoOcx.ctx":017E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1815
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clientes"
      Height          =   825
      Left            =   255
      TabIndex        =   5
      Top             =   180
      Width           =   5355
      Begin MSMask.MaskEdBox ClienteInicial 
         Height          =   300
         Left            =   615
         TabIndex        =   0
         Top             =   315
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox ClienteFinal 
         Height          =   300
         Left            =   3210
         TabIndex        =   1
         Top             =   300
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   529
         _Version        =   393216
         PromptInclude   =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label LabelClienteAte 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2775
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   7
         Top             =   360
         Width           =   360
      End
      Begin VB.Label LabelClienteDe 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   225
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   6
         Top             =   345
         Width           =   315
      End
   End
End
Attribute VB_Name = "RelOpTitAtrasoOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Private WithEvents objEventoClienteInic As AdmEvento
Attribute objEventoClienteInic.VB_VarHelpID = -1
Private WithEvents objEventoClienteFim As AdmEvento
Attribute objEventoClienteFim.VB_VarHelpID = -1
Private WithEvents objEventoVendedor As AdmEvento
Attribute objEventoVendedor.VB_VarHelpID = -1

Dim gobjRelOpcoes As AdmRelOpcoes
Dim gobjRelatorio As AdmRelatorio

Function Trata_Parametros(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    If Not (gobjRelatorio Is Nothing) Then gError 90608
    
    Set gobjRelatorio = objRelatorio
    Set gobjRelOpcoes = objRelOpcoes

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case 90608
            lErro = Rotina_Erro(vbOKOnly, "ERRO_RELATORIO_EXECUTANDO", gErr)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173376)

    End Select

    Exit Function

End Function

Private Sub BotaoFechar_Click()
    
    Unload Me
    
End Sub

Private Sub ClienteFinal_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteFinal_Validate

    If Len(Trim(ClienteFinal.Text)) > 0 Then

        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_cliente_Le2(ClienteFinal, objCliente, 0)
        If lErro <> SUCESSO Then gError 90609

    End If
    
    Exit Sub

Erro_ClienteFinal_Validate:

    Cancel = True


    Select Case gErr

        Case 90609
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173377)

    End Select

End Sub

Private Sub ClienteInicial_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objCliente As New ClassCliente

On Error GoTo Erro_ClienteInicial_Validate

    If Len(Trim(ClienteInicial.Text)) > 0 Then
   
        'Tenta ler o Cliente (NomeReduzido ou Código)
        lErro = TP_cliente_Le2(ClienteInicial, objCliente, 0)
        If lErro <> SUCESSO Then gError 90610

    End If
        
    Exit Sub

Erro_ClienteInicial_Validate:

    Cancel = True


    Select Case gErr

        Case 90610
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 173378)

    End Select

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    Set objEventoClienteInic = New AdmEvento
    Set objEventoClienteFim = New AdmEvento
    Set objEventoVendedor = New AdmEvento
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

   lErro_Chama_Tela = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173379)

    End Select

    Exit Sub

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjRelatorio = Nothing
    Set gobjRelOpcoes = Nothing
    
    Set objEventoClienteInic = Nothing
    Set objEventoClienteFim = Nothing
    Set objEventoVendedor = Nothing
    
End Sub

Private Sub LabelClienteAte_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente

On Error GoTo Erro_LabelClienteAte_Click
    
    If Len(Trim(ClienteFinal.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteFinal.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteFim)

   Exit Sub

Erro_LabelClienteAte_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173380)

    End Select

    Exit Sub

End Sub

Private Sub LabelClienteDe_Click()

Dim lErro As Long
Dim colSelecao As New Collection
Dim objCliente As New ClassCliente

On Error GoTo Erro_LabelClienteDe_Click
    
    If Len(Trim(ClienteInicial.Text)) > 0 Then
        'Preenche com o cliente da tela
        objCliente.lCodigo = LCodigo_Extrai(ClienteInicial.Text)
    End If
    
    'Chama Tela ClientesLista
    Call Chama_Tela("ClientesLista", colSelecao, objCliente, objEventoClienteInic)

   Exit Sub

Erro_LabelClienteDe_Click:

    Select Case gErr

         Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173381)

    End Select

    Exit Sub
    
End Sub

Private Sub objEventoClienteFim_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    ClienteFinal.Text = CStr(objCliente.lCodigo)
    Call ClienteFinal_Validate(bSGECancelDummy)
    
    Me.Show

    Exit Sub

End Sub

Private Sub objEventoClienteInic_evSelecao(obj1 As Object)

Dim objCliente As ClassCliente

    Set objCliente = obj1
    
    ClienteInicial.Text = CStr(objCliente.lCodigo)
    Call ClienteInicial_Validate(bSGECancelDummy)

    Me.Show

    Exit Sub

End Sub

Private Sub BotaoExecutar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoExecutar_Click

    lErro = PreencherRelOp(gobjRelOpcoes)
    If lErro <> SUCESSO Then gError 90611

    Call gobjRelatorio.Executar_Prossegue2(Me)

    Exit Sub

Erro_BotaoExecutar_Click:

    Select Case gErr

        Case 90611

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173382)

    End Select

    Exit Sub

End Sub

Function PreencherRelOp(objRelOpcoes As AdmRelOpcoes) As Long
'preenche objRelOpcoes com os dados da tela

Dim lErro As Long
Dim sClienteIni As String
Dim sClienteFim As String
Dim iVendedor As Integer
Dim objVendedor As New ClassVendedor

On Error GoTo Erro_PreencherRelOp

    If Len(Trim(ClienteInicial.Text)) <> 0 And Len(Trim(ClienteFinal.Text)) <> 0 Then
        If (Codigo_Extrai(ClienteInicial.Text)) > (Codigo_Extrai(ClienteFinal.Text)) Then gError 90612
    End If
       
    sClienteIni = SCodigo_Extrai(ClienteInicial.Text)
    sClienteFim = SCodigo_Extrai(ClienteFinal.Text)
    
    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor.Text

    'Verifica se vendedor existe
    If objVendedor.sNomeReduzido <> "" Then
    
        lErro = CF("Vendedor_Le_NomeReduzido", objVendedor)
        If lErro <> SUCESSO And lErro <> 25008 Then gError ERRO_SEM_MENSAGEM

        iVendedor = objVendedor.iCodigo

    End If
    
    lErro = objRelOpcoes.Limpar
    If lErro <> AD_BOOL_TRUE Then gError 90613
         
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEINIC", sClienteIni)
    If lErro <> AD_BOOL_TRUE Then gError 90614
         
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEINIC", ClienteInicial.Text)
    If lErro <> AD_BOOL_TRUE Then gError 90615
    
    lErro = objRelOpcoes.IncluirParametro("NCLIENTEFIM", sClienteFim)
    If lErro <> AD_BOOL_TRUE Then gError 90616
    
    lErro = objRelOpcoes.IncluirParametro("TCLIENTEFIM", ClienteFinal.Text)
    If lErro <> AD_BOOL_TRUE Then gError 90617
    
    lErro = objRelOpcoes.IncluirParametro("NVENDEDOR", CStr(iVendedor))
    If lErro <> AD_BOOL_TRUE Then gError 90617
    
    lErro = objRelOpcoes.IncluirParametro("DDATA_HOJE", CStr(gdtDataAtual))
    If lErro <> AD_BOOL_TRUE Then gError 90617
    
    lErro = Monta_Expressao_Selecao(objRelOpcoes, sClienteIni, sClienteFim)
    If lErro <> SUCESSO Then gError 90618
    
    PreencherRelOp = SUCESSO

    Exit Function

Erro_PreencherRelOp:

    PreencherRelOp = gErr

    Select Case gErr
    
        Case 90612
            lErro = Rotina_Erro(vbOKOnly, "ERRO_CLIENTE_INICIAL_MAIOR", gErr)

        Case 90613 To 90618
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173383)

    End Select

    Exit Function

End Function

Function Monta_Expressao_Selecao(objRelOpcoes As AdmRelOpcoes, sClienteIni As String, sClienteFim As String) As Long
'monta a expressão de seleção de relatório

Dim sExpressao As String
Dim lErro As Long

On Error GoTo Erro_Monta_Expressao_Selecao

   If sClienteIni <> "" Then sExpressao = "Cliente >= " & Forprint_ConvLong(CLng(sClienteIni))

   If sClienteFim <> "" Then

        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "Cliente <= " & Forprint_ConvLong(CLng(sClienteFim))

    End If
        
    If giFilialEmpresa <> EMPRESA_TODA And giFilialEmpresa <> gobjCR.iFilialCentralizadora Then
        If sExpressao <> "" Then sExpressao = sExpressao & " E "
        sExpressao = sExpressao & "FilialEmpresa = " & Forprint_ConvInt(CInt(giFilialEmpresa))
    End If

    If sExpressao <> "" Then

        objRelOpcoes.sSelecao = sExpressao

    End If

    Monta_Expressao_Selecao = SUCESSO

    Exit Function

Erro_Monta_Expressao_Selecao:

    Monta_Expressao_Selecao = gErr

    Select Case gErr

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 173384)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_RELOP_CAD_CLI
    Set Form_Load_Ocx = Me
    Caption = "Títulos em Atraso"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "RelOpTitAtraso"
    
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
        
        If Me.ActiveControl Is ClienteInicial Then
            Call LabelClienteDe_Click
        ElseIf Me.ActiveControl Is ClienteFinal Then
            Call LabelClienteAte_Click
        End If
    
    End If

End Sub

Private Sub LabelClienteDe_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteDe, Source, X, Y)
End Sub

Private Sub LabelClienteDe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteDe, Button, Shift, X, Y)
End Sub

Private Sub LabelClienteAte_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelClienteAte, Source, X, Y)
End Sub

Private Sub LabelClienteAte_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelClienteAte, Button, Shift, X, Y)
End Sub


Public Sub Vendedor_Validate(Cancel As Boolean)

Dim lErro As Long
Dim objVendedor As New ClassVendedor
Dim dPercComissao As Double

On Error GoTo Erro_Vendedor_Validate

    'Se Vendedor foi alterado,
    If Len(Trim(Vendedor.Text)) <> 0 Then

        'Tenta ler o Vendedor (NomeReduzido ou Código)
        lErro = TP_Vendedor_Le(Vendedor, objVendedor)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        

    End If

    Exit Sub

Erro_Vendedor_Validate:

    Cancel = True
    
    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 209207)
    
    End Select

End Sub

Public Sub VendedorLabel_Click()

'BROWSE VENDEDOR :

Dim objVendedor As New ClassVendedor
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(Vendedor.Text)) > 0 Then objVendedor.sNomeReduzido = Vendedor.Text
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("VendedorLista", colSelecao, objVendedor, objEventoVendedor)

End Sub

Private Sub objEventoVendedor_evSelecao(obj1 As Object)

Dim objVendedor As ClassVendedor

    Set objVendedor = obj1

    'Preenche campo Vendedor
    Vendedor.Text = objVendedor.sNomeReduzido

    Me.Show

    Vendedor.SetFocus 'Inserido por Wagner
    
    Exit Sub

End Sub

