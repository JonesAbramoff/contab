VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl TransfCaixaLista 
   ClientHeight    =   4275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5820
   DefaultCancel   =   -1  'True
   ScaleHeight     =   4275
   ScaleWidth      =   5820
   Begin VB.CommandButton BotaoFechar 
      Cancel          =   -1  'True
      Caption         =   "Fechar"
      Height          =   780
      Left            =   2925
      Picture         =   "TransfCaixaLista.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3315
      Width           =   1830
   End
   Begin VB.CommandButton BotaoSelecionar 
      Caption         =   "Selecionar"
      Default         =   -1  'True
      Height          =   780
      Left            =   990
      Picture         =   "TransfCaixaLista.ctx":0272
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3315
      Width           =   1860
   End
   Begin MSFlexGridLib.MSFlexGrid GridTransf 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   5318
      _Version        =   393216
      Rows            =   11
      Cols            =   5
      FixedCols       =   0
      ForeColorSel    =   16777215
      AllowBigSelection=   0   'False
      Enabled         =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "TransfCaixaLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim gobjTransfCaixa As ClassTransfCaixa
Public iAlterado As Integer
Dim gdQuant As Double

'Constantes Relacionadas as Colunas do Grid

Dim iGrid_Codigo_Col As Integer
Dim iGrid_Data_Col As Integer
Dim iGrid_Valor_Col As Integer
Dim iGrid_TipoOrigem_Col As Integer
Dim iGrid_TipoDestino_Col As Integer

'******************************************************************
Private Sub BotaoSelecionar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoSelecionar_Click

    If GridTransf.Row = 0 Or GridTransf.Row > gdQuant Then Exit Sub

    Set gobjTransfCaixa.objMovCaixaDe = New ClassMovimentoCaixa
    
    Set gobjTransfCaixa.objMovCaixaPara = New ClassMovimentoCaixa
   
    gobjTransfCaixa.objMovCaixaDe.lTransferencia = StrParaLong(GridTransf.TextMatrix(GridTransf.Row, iGrid_Codigo_Col))
    gobjTransfCaixa.objMovCaixaPara.lTransferencia = StrParaLong(GridTransf.TextMatrix(GridTransf.Row, iGrid_Codigo_Col))
    gobjTransfCaixa.objMovCaixaDe.dtDataMovimento = StrParaDate(GridTransf.TextMatrix(GridTransf.Row, iGrid_Data_Col))
    gobjTransfCaixa.objMovCaixaPara.dtDataMovimento = StrParaDate(GridTransf.TextMatrix(GridTransf.Row, iGrid_Data_Col))
    gobjTransfCaixa.objMovCaixaDe.dValor = StrParaDbl(GridTransf.TextMatrix(GridTransf.Row, iGrid_Valor_Col))
    gobjTransfCaixa.objMovCaixaPara.dValor = StrParaDbl(GridTransf.TextMatrix(GridTransf.Row, iGrid_Valor_Col))
    gobjTransfCaixa.objMovCaixaDe.iTipo = Converte_Tipo_String_Saida(GridTransf.TextMatrix(GridTransf.Row, iGrid_TipoOrigem_Col))
    gobjTransfCaixa.objMovCaixaPara.iTipo = Converte_Tipo_String_Entrada(GridTransf.TextMatrix(GridTransf.Row, iGrid_TipoDestino_Col))

    giRetornoTela = vbOK

    Unload Me

    Exit Sub

Erro_BotaoSelecionar_Click:

    Select Case Err

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, Err, Error, 175378)

    End Select

    Exit Sub

End Sub

Private Sub Gridcheque_DblClick()
    
    Call BotaoSelecionar_Click
    
End Sub

Private Sub BotaoFechar_Click()
    
    giRetornoTela = vbCancel
    
    Unload Me

End Sub

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    Set gobjTransfCaixa = New ClassTransfCaixa
    
    Set gobjTransfCaixa.objMovCaixaDe = New ClassMovimentoCaixa
    
    Set gobjTransfCaixa.objMovCaixaPara = New ClassMovimentoCaixa
    
    iGrid_Codigo_Col = 0
    iGrid_Data_Col = 1
    iGrid_Valor_Col = 2
    iGrid_TipoOrigem_Col = 3
    iGrid_TipoDestino_Col = 4
    
    GridTransf.TextMatrix(0, iGrid_Codigo_Col) = "Código"
    GridTransf.TextMatrix(0, iGrid_Data_Col) = "Data"
    GridTransf.TextMatrix(0, iGrid_Valor_Col) = "Valor"
    GridTransf.TextMatrix(0, iGrid_TipoOrigem_Col) = "Origem"
    GridTransf.TextMatrix(0, iGrid_TipoDestino_Col) = "Destino"
     
    lErro = Preenche_Grid_Transf()
    If lErro <> SUCESSO Then gError 101958

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = gErr

    Select Case gErr

        Case 101958

        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175379)

    End Select

    Exit Sub

End Sub

Private Function Preenche_Grid_Transf() As Long

Dim objMC_Entrada As ClassMovimentoCaixa
Dim objMC_Saida As ClassMovimentoCaixa
Dim iIndice As Integer
Dim colMovTransf_Entrada As New Collection
Dim colMovTransf_Saida As New Collection
Dim iIndiceGrid As Integer
Dim iTipoEntrada As Integer

On Error GoTo Erro_Preenche_Grid_Transf

    'lê todos os movtos de caixa que são do tipo transf caixa, separando os de entrada dos de saída
    Call Monta_Colecao_MovTransf(colMovTransf_Entrada, colMovTransf_Saida)

    If colMovTransf_Saida.Count > 11 Then
        GridTransf.Rows = colMovTransf_Saida.Count + 1
    Else
        GridTransf.Rows = 11
    End If
    
    'varre a coleção de movto de transf
    For Each objMC_Saida In colMovTransf_Saida
        
        'incrementa a posição de preenchimento do grid
        iIndiceGrid = iIndiceGrid + 1
        
        'zera o indexador de elementos da col de movtos de entrada
        iIndice = 0
        
        'tenho o movimento de saída (o movimento "de"), preciso o de entrada ("para")
        For Each objMC_Entrada In colMovTransf_Entrada
        
            iIndice = iIndice + 1
            
            'quando encontrar, sai do loop, removendo o elemento da coleção de movtos de entrada
            'isso otimiza o algoritmo, uma vez que um elemento encontrado não será mais necessário,
            'ele pode ser removido para não "atrapalhar" na próxima rodada de busca
            If objMC_Saida.lTransferencia = objMC_Entrada.lTransferencia Then
                
                'guarda o tipo de entrada (o "para")
                iTipoEntrada = objMC_Entrada.iTipo
                
                'remove o elemento da coleção
                colMovTransf_Entrada.Remove (iIndice)
                
                'sai do loop
                Exit For
            
            End If
        
        Next
        
        'preenche a linha do grid
        GridTransf.TextMatrix(iIndiceGrid, iGrid_Codigo_Col) = objMC_Saida.lTransferencia
        GridTransf.TextMatrix(iIndiceGrid, iGrid_Data_Col) = Format(objMC_Saida.dtDataMovimento, "dd/mm/yyyy")
        GridTransf.TextMatrix(iIndiceGrid, iGrid_Valor_Col) = Format(objMC_Saida.dValor, "STANDARD")
        GridTransf.TextMatrix(iIndiceGrid, iGrid_TipoOrigem_Col) = Converte_Tipo_String(objMC_Saida.iTipo)
        GridTransf.TextMatrix(iIndiceGrid, iGrid_TipoDestino_Col) = Converte_Tipo_String(objMC_Entrada.iTipo)
                
    Next
    
    gdQuant = iIndiceGrid
    
    Preenche_Grid_Transf = SUCESSO

    Exit Function

Erro_Preenche_Grid_Transf:

    Preenche_Grid_Transf = gErr

    Select Case gErr
    
        Case Else
            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175380)

    End Select

    Exit Function

End Function

Private Sub Monta_Colecao_MovTransf(ByVal colMovTransf_Entrada As Collection, ByVal colMovTransf_Saida As Collection)

Dim objMC As ClassMovimentoCaixa

    'varro a coleção global de movimentos de caixa
    For Each objMC In gcolMovimentosCaixa
        
        'se for um movimento de transferência
        If objMC.lTransferencia > 0 Then
            
            'verifico se é de saída ou entrada, adicionando na coleção correspondente
            If CF_ECF("Movimento_Eh_De_Saida", objMC) Then
                colMovTransf_Saida.Add objMC
            Else
                colMovTransf_Entrada.Add objMC
            End If
        
        End If
    
    Next

End Sub

Private Function RetornaRemove_Movto_Transf(lCodTransf As Long, colMovTransf As Collection) As ClassMovimentoCaixa

Dim objMC As ClassMovimentoCaixa
Dim iIndice As Long

   For Each objMC In colMovTransf
   
      iIndice = iIndice + 1
   
      If objMC.lTransferencia = lCodTransf Then
      
         Set RetornaRemove_Movto_Transf = objMC
   
         colMovTransf.Remove (iIndice)
      
      End If
   
   Next

End Function

Public Sub Form_Unload(Cancel As Integer)

'    Set gobjTransfCaixa = Nothing

End Sub

Private Function Converte_Tipo_String(iTipo As Integer) As String
'tendo em mente q movimentos de exclusao nao sao considerados

   If iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_DINHEIRO Or _
      iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_DINHEIRO Then

      Converte_Tipo_String = "DINHEIRO"
   
   ElseIf iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_CHEQUE Or _
          iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_CHEQUE Then
   
      Converte_Tipo_String = "CHEQUE"
      
   ElseIf iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_CARTAO_CREDITO Or _
          iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_CARTAO_CREDITO Then
          
      Converte_Tipo_String = "CARTÃO DE CRÉDITO"
   
   ElseIf iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_CARTAO_DEBITO Or _
          iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_CARTAO_DEBITO Then
          
      Converte_Tipo_String = "CARTÃO DE DÉBITO"
      
   ElseIf iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_VALETICKET Or _
          iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_VALETICKET Then
          
      Converte_Tipo_String = "VALE/TICKET"
   
   ElseIf iTipo = MOVIMENTOCAIXA_SAIDA_TRANSF_OUTROS Or _
          iTipo = MOVIMENTOCAIXA_ENTRADA_TRANSF_OUTROS Then
          
      Converte_Tipo_String = "OUTROS"
   
   Else
   
      Converte_Tipo_String = "DESCONHECIDO"
   
   End If
   
End Function

Private Function Converte_Tipo_String_Entrada(sString As String) As Integer

   If sString = "DINHEIRO" Then
      
      Converte_Tipo_String_Entrada = MOVIMENTOCAIXA_ENTRADA_TRANSF_DINHEIRO
   
   ElseIf sString = "CHEQUE" Then
   
      Converte_Tipo_String_Entrada = MOVIMENTOCAIXA_ENTRADA_TRANSF_CHEQUE
   
   ElseIf sString = "CARTÃO DE CRÉDITO" Then
   
      Converte_Tipo_String_Entrada = MOVIMENTOCAIXA_ENTRADA_TRANSF_CARTAO_CREDITO
      
   ElseIf sString = "CARTÃO DE DÉBITO" Then
   
      Converte_Tipo_String_Entrada = MOVIMENTOCAIXA_ENTRADA_TRANSF_CARTAO_DEBITO
   
   ElseIf sString = "VALE/TICKET" Then
   
      Converte_Tipo_String_Entrada = MOVIMENTOCAIXA_ENTRADA_TRANSF_VALETICKET
      
   ElseIf sString = "OUTROS" Then
   
      Converte_Tipo_String_Entrada = MOVIMENTOCAIXA_ENTRADA_TRANSF_OUTROS
      
   Else
   
      Converte_Tipo_String_Entrada = -1

   End If

End Function

Private Function Converte_Tipo_String_Saida(sString As String) As Integer

   If sString = "DINHEIRO" Then
      
      Converte_Tipo_String_Saida = MOVIMENTOCAIXA_SAIDA_TRANSF_DINHEIRO
   
   ElseIf sString = "CHEQUE" Then
   
      Converte_Tipo_String_Saida = MOVIMENTOCAIXA_SAIDA_TRANSF_CHEQUE
   
   ElseIf sString = "CARTÃO DE CRÉDITO" Then
   
      Converte_Tipo_String_Saida = MOVIMENTOCAIXA_SAIDA_TRANSF_CARTAO_CREDITO
      
   ElseIf sString = "CARTÃO DE DÉBITO" Then
   
      Converte_Tipo_String_Saida = MOVIMENTOCAIXA_SAIDA_TRANSF_CARTAO_DEBITO
   
   ElseIf sString = "VALE/TICKET" Then
   
      Converte_Tipo_String_Saida = MOVIMENTOCAIXA_SAIDA_TRANSF_VALETICKET
      
   ElseIf sString = "OUTROS" Then
   
      Converte_Tipo_String_Saida = MOVIMENTOCAIXA_SAIDA_TRANSF_OUTROS
      
   Else
   
      Converte_Tipo_String_Saida = -1

   End If

End Function

Function Trata_Parametros(objTransfCaixa As ClassTransfCaixa) As Long

On Error GoTo Erro_Trata_Parametros

    Set gobjTransfCaixa = objTransfCaixa

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = gErr

    Select Case gErr

        Case Else

            Call Rotina_ErroECF(vbOKOnly, ERRO_FORNECIDO_PELO_VB_1, gErr, Error, 175381)

    End Select

    Exit Function

End Function

'**** inicio do trecho a ser copiado *****

Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BROWSE
    Set Form_Load_Ocx = Me
    Caption = "Lista de Transferências"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "ChequeLista"

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

Private Sub GridVendedor_Click()

End Sub

Private Sub GridTransf_DblClick()

   Call BotaoSelecionar_Click

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

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

End Sub

'**** fim do trecho a ser copiado *****

