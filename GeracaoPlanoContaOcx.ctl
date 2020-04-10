VERSION 5.00
Begin VB.UserControl GeracaoPlanoContaOcx 
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
   LockControls    =   -1  'True
   ScaleHeight     =   1965
   ScaleWidth      =   5310
   Begin VB.CommandButton Gera 
      Caption         =   "GeraPlanoCont"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1140
      Width           =   1215
   End
   Begin VB.CommandButton GeraPedido 
      Caption         =   "GeraPedido"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   1140
      Width           =   1215
   End
   Begin VB.CommandButton GeraPedido1 
      Caption         =   "GeraPedido1"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   1155
      Width           =   1215
   End
   Begin VB.Label LabelPedido 
      Height          =   315
      Left            =   2430
      Top             =   240
      Width           =   1170
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Height          =   225
      Left            =   1680
      Top             =   315
      Width           =   600
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Pedido:"
   End
End
Attribute VB_Name = "GeracaoPlanoContaOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Const CONTA_SINTETICA = 1
Const CONTA_ANALITICA = 3
Const CONSTANTE = 1
Const NIVEL1 = 1
Const NIVEL2 = 2
Const NIVEL3 = 3

Private Sub Gera_Click()

Dim lErro As Long

On Error GoTo Erro_Gera_Click

    Call Gera_PlanoConta

    Exit Sub

Erro_Gera_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161421)

    End Select

    Exit Sub

End Sub

Public Function Gera_PlanoConta() As Long

Dim lErro As Long
Dim lComando As Long
Dim lTransacao As Long
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim iIndice3 As Integer
Dim sConta1 As String
Dim sConta2 As String
Dim sConta3 As String

On Error GoTo Erro_Gera_PlanoConta

    'Abertura Comando
    lComando = Comando_Abrir()
    If lComando = 0 Then Error 28514

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 28515

    For iIndice1 = 1 To 9

        sConta1 = Trim(Str(iIndice1) & "00000000")

        'Insere PlanoConta nível 1 no BD
        lErro = Comando_Executar(lComando, "INSERT INTO PlanoConta (Conta, NivelConta, TipoConta, CP, CR, TES, FAT, EST) VALUES (?,?,?,?,?,?,?,?)", sConta1, 1, CONTA_SINTETICA, CONSTANTE, CONSTANTE, CONSTANTE, CONSTANTE, CONSTANTE)
        If lErro <> AD_SQL_SUCESSO Then Error 28516

        For iIndice2 = 1 To 100

            sConta2 = Trim(Str(iIndice1) & Format(iIndice2, "000") & "00000")

            'Insere PlanoConta nível 2 no BD
            lErro = Comando_Executar(lComando, "INSERT INTO PlanoConta (Conta, NivelConta, TipoConta, CP, CR, TES, FAT, EST) VALUES (?,?,?,?,?,?,?,?)", sConta2, 2, CONTA_SINTETICA, CONSTANTE, CONSTANTE, CONSTANTE, CONSTANTE, CONSTANTE)
            If lErro <> AD_SQL_SUCESSO Then Error 28517

            For iIndice3 = 1 To 10

                sConta3 = Trim(Str(iIndice1) & Format(iIndice2, "000") & Format(iIndice3, "00") & "000")

                'Insere PlanoConta nível 3 no BD
                lErro = Comando_Executar(lComando, "INSERT INTO PlanoConta (Conta, NivelConta, TipoConta, CP, CR, TES, FAT, EST) VALUES (?,?,?,?,?,?,?,?)", sConta3, 3, CONTA_ANALITICA, CONSTANTE, CONSTANTE, CONSTANTE, CONSTANTE, CONSTANTE)
                If lErro <> AD_SQL_SUCESSO Then Error 28518


            Next

        Next

    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 28519

    Call Comando_Fechar(lComando)

    Gera_PlanoConta = SUCESSO

    Exit Function

Erro_Gera_PlanoConta:

    Gera_PlanoConta = Err

    Select Case Err

        Case 28514
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 28515
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 28516, 28517, 28518, 28519

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161422)

     End Select

    Call Transacao_Rollback
    Call Comando_Fechar(lComando)

     Exit Function

End Function

Function Trata_Parametros() As Long

    Trata_Parametros = SUCESSO

End Function

Private Sub GeraPedido_Click()

Dim lErro As Long

On Error GoTo Erro_GeraPedido_Click

    Call Gera_Pedido

    Exit Sub

Erro_GeraPedido_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161423)

    End Select

    Exit Sub

End Sub

Public Function Gera_Pedido() As Long

Dim lErro As Long
Dim lComando1 As Long, lComando2 As Long
Dim lTransacao As Long
Dim iFilialEmpresa As Integer
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim dQuantidade As Double
Dim lPedido As Long
Dim lCliente As Long
Dim iItem As Integer
Dim sProduto As String
Dim dPrecoTotal As Double
Dim iNumero As Integer
Dim iTransportadora As Integer

On Error GoTo Erro_Gera_Pedido

    'Abertura Comando
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 24631

    'Abertura Comando
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 24631

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 24632

    iFilialEmpresa = 1
    iTransportadora = 51
    
    For iIndice1 = 1 To 10000

        lPedido = iIndice1
        lCliente = Int((100 * Rnd) + 1)

        'Insere PedidoVendaBaixado no BD
        lErro = Comando_Executar(lComando1, "INSERT INTO PedidosDeVendaBaixados (FilialEmpresa,Codigo,Cliente,CodTransportadora) VALUES (?,?,?,?)", iFilialEmpresa, lPedido, lCliente, iTransportadora)
        If lErro <> AD_SQL_SUCESSO Then Error 24633

        For iIndice2 = 1 To 5

            iItem = iIndice2
            iNumero = Int((100 * Rnd) + 1)
            sProduto = Str(iNumero)
            dQuantidade = iNumero
            dPrecoTotal = iNumero

            'Insere ItemPedidoVendaBaixado no BD
            lErro = Comando_Executar(lComando2, "INSERT INTO ItensPedidoDeVendaBaixados (FilialEmpresa,CodPedido,Item,Produto,Quantidade,PrecoTotal) VALUES (?,?,?,?,?,?)", iFilialEmpresa, lPedido, iItem, sProduto, dQuantidade, dPrecoTotal)
            If lErro <> AD_SQL_SUCESSO Then Error 24634

        Next

        LabelPedido.Caption = lPedido

    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 24635

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Gera_Pedido = SUCESSO

    Exit Function

Erro_Gera_Pedido:

    Gera_Pedido = Err

    Select Case Err

        Case 24631
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 24632
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 24633, 24634, 24635

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161424)

     End Select

    Call Transacao_Rollback
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

     Exit Function

End Function

Private Sub GeraPedido1_Click()

Dim lErro As Long

On Error GoTo Erro_GeraPedido1_Click

    Call Gera_Pedido1

    Exit Sub

Erro_GeraPedido1_Click:

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 161425)

    End Select

    Exit Sub

End Sub

Public Function Gera_Pedido1() As Long

Dim lErro As Long
Dim lComando1 As Long, lComando2 As Long
Dim lTransacao As Long
Dim iFilialEmpresa As Integer
Dim iIndice1 As Integer
Dim iIndice2 As Integer
Dim dQuantidade As Double
Dim lPedido As Long
Dim lCliente As Long
Dim iItem As Integer
Dim sProduto As String
Dim dPrecoTotal As Double
Dim iNumero As Integer
Dim iTransportadora As Integer

On Error GoTo Erro_Gera_Pedido1

    'Abertura Comando
    lComando1 = Comando_Abrir()
    If lComando1 = 0 Then Error 24642

    'Abertura Comando
    lComando2 = Comando_Abrir()
    If lComando2 = 0 Then Error 24643

    'Inicia a Transacao
    lTransacao = Transacao_Abrir()
    If lTransacao = 0 Then Error 24644

    iFilialEmpresa = 1
    iTransportadora = 51
        
    For iIndice1 = 10001 To 11000

        lPedido = iIndice1
        lCliente = Int((100 * Rnd) + 1)

        'Insere PedidoVendaBaixado no BD
        lErro = Comando_Executar(lComando1, "INSERT INTO PedidosDeVenda1 (FilialEmpresa,Codigo,Cliente,CodTransportadora) VALUES (?,?,?,?)", iFilialEmpresa, lPedido, lCliente, iTransportadora)
        If lErro <> AD_SQL_SUCESSO Then Error 24641

        For iIndice2 = 1 To 5

            iItem = iIndice2
            iNumero = Int((100 * Rnd) + 1)
            sProduto = Str(iNumero)
            dQuantidade = iNumero
            dPrecoTotal = iNumero

            'Insere ItemPedidoVendaBaixado no BD
            lErro = Comando_Executar(lComando2, "INSERT INTO ItensPedidoDeVenda1 (FilialEmpresa,CodPedido,Item,Produto,Quantidade,PrecoTotal) VALUES (?,?,?,?,?,?)", iFilialEmpresa, lPedido, iItem, sProduto, dQuantidade, dPrecoTotal)
            If lErro <> AD_SQL_SUCESSO Then Error 24645

        Next

        LabelPedido.Caption = lPedido

    Next

    'Confirma a transação
    lErro = Transacao_Commit()
    If lErro <> AD_SQL_SUCESSO Then Error 24646

    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

    Gera_Pedido1 = SUCESSO

    Exit Function

Erro_Gera_Pedido1:

    Gera_Pedido1 = Err

    Select Case Err

        Case 24642, 24643
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_COMANDO", Err)

        Case 24644
            lErro = Rotina_Erro(vbOKOnly, "ERRO_ABERTURA_TRANSACAO", Err)

        Case 24641, 24645, 24646

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 161426)

     End Select

    Call Transacao_Rollback
    
    Call Comando_Fechar(lComando1)
    Call Comando_Fechar(lComando2)

     Exit Function

End Function

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_GERACAO_PLANOCONTA
    Set Form_Load_Ocx = Me
    Caption = "Geração Plano Conta"
    'Call Form_Load
    
End Function

Public Function Name() As String

    Name = "GeracaoPlanoConta"
    
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

'***** fim do trecho a ser copiado ******



Private Sub LabelPedido_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(LabelPedido, Source, X, Y)
End Sub

Private Sub LabelPedido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(LabelPedido, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

