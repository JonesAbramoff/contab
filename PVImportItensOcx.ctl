VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl PVImportItensOcx 
   ClientHeight    =   2355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7425
   ScaleHeight     =   2355
   ScaleWidth      =   7425
   Begin VB.CommandButton BotaoOK 
      Caption         =   "OK"
      Height          =   540
      Left            =   2850
      Picture         =   "PVImportItensOcx.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1725
      Width           =   855
   End
   Begin VB.CommandButton BotaoCancela 
      Caption         =   "Cancelar"
      Height          =   540
      Left            =   3780
      Picture         =   "PVImportItensOcx.ctx":015A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1725
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Identificação do Arquivo"
      Height          =   1470
      Left            =   105
      TabIndex        =   0
      Top             =   195
      Width           =   7200
      Begin VB.Frame Frame2 
         Caption         =   "Colunas"
         Height          =   645
         Left            =   75
         TabIndex        =   6
         Top             =   720
         Width           =   7050
         Begin VB.ComboBox colPreco 
            Height          =   315
            ItemData        =   "PVImportItensOcx.ctx":025C
            Left            =   5790
            List            =   "PVImportItensOcx.ctx":02BF
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   195
            Width           =   1185
         End
         Begin VB.ComboBox colQtd 
            Height          =   315
            ItemData        =   "PVImportItensOcx.ctx":0311
            Left            =   3345
            List            =   "PVImportItensOcx.ctx":0374
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   225
            Width           =   1185
         End
         Begin VB.ComboBox colProduto 
            Height          =   315
            ItemData        =   "PVImportItensOcx.ctx":03C6
            Left            =   915
            List            =   "PVImportItensOcx.ctx":03DF
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   225
            Width           =   1185
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Preço Unit.:"
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
            Height          =   315
            Left            =   4635
            TabIndex        =   12
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
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
            Left            =   2190
            TabIndex        =   10
            Top             =   270
            Width           =   1125
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Produto:"
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
            Left            =   75
            TabIndex        =   8
            Top             =   270
            Width           =   795
         End
      End
      Begin VB.TextBox NomeArquivo 
         Height          =   315
         Left            =   1740
         TabIndex        =   2
         Top             =   315
         Width           =   4995
      End
      Begin VB.CommandButton BotaoProcurar 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6750
         TabIndex        =   1
         Top             =   270
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "Nome do Arquivo:"
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
         Left            =   150
         TabIndex        =   5
         Top             =   345
         Width           =   1560
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   1590
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "PVImportItensOcx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim gobjPV As ClassPedidoDeVenda

'Property Variables:
Dim m_Caption As String
Event Unload()

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

'    Parent.HelpContextID = IDH_PROCESSA_ARQRETCOBRANCA
    Set Form_Load_Ocx = Me
    Caption = "Importar Itens do Pedido de Venda"
    Call Form_Load

End Function

Public Function Name() As String

    Name = "PVImportItens"

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

Private Sub BotaoCancela_Click()
    Unload Me
End Sub

Private Sub BotaoOK_Click()

Dim lErro As Long
Dim sNomeArq As String
Dim sNomeDir As String
Dim iPos As Integer
Dim iPosAnt As Integer

On Error GoTo Erro_BotaoOK_Click

    If Len(Trim(NomeArquivo.Text)) = 0 Then gError 192919
    
    'NomeArquivo.Text
    iPos = 1
    Do While iPos <> 0
        iPosAnt = iPos
        iPos = InStr(iPosAnt + 1, NomeArquivo.Text, "\")
    Loop
    
    sNomeDir = left(NomeArquivo.Text, iPosAnt)
    sNomeArq = Mid(NomeArquivo.Text, iPosAnt + 1)
    
    lErro = XLS_Importa(sNomeDir, sNomeArq)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    Unload Me

    Exit Sub

Erro_BotaoOK_Click:

    Select Case gErr
        
        Case 192919
            Call Rotina_Erro(vbOKOnly, "ERRO_ARQUIVO_NAO_PREENCHIDO", gErr)
            
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 192921)

    End Select

    Exit Sub

End Sub

Private Sub BotaoProcurar_Click()

    On Error GoTo Erro_BotaoProcurar_Click

    ' Set CancelError is True
    CommonDialog1.CancelError = True
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNNoChangeDir
    ' Set filters
    CommonDialog1.Filter = "xls Files (*.xls)|*.xls|xlsx Files" & _
    "(*.xlsx)|*.xlsx"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen

    ' Display name of selected file

    NomeArquivo.Text = CommonDialog1.FileName
    Exit Sub

Erro_BotaoProcurar_Click:
    'User pressed the Cancel button
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

Private Sub Unload(objme As Object)

    Set gobjPV = Nothing

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

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load

    ColProduto.ListIndex = 0
    colQtd.ListIndex = 5
    colPreco.ListIndex = -1
    
    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 192922)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objPV As ClassPedidoDeVenda) As Long

    Set gobjPV = objPV

    Trata_Parametros = SUCESSO

End Function

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

Public Function XLS_Importa(ByVal sDir As String, ByVal sArq As String) As Long

Dim lErro As Long
Dim objPV As New ClassPedidoDeVenda
Dim objItem As ClassItemPedido
Dim objExcelApp As New ClassExcelApp
Dim sValColProd As String, sValColQtde As String, sValColPreco As String
Dim sProduto As String, sProdutoBD As String, iProdPreenchido As Integer
Dim bComMask As Boolean, iLinha As Integer
Dim iCountSemProd As Integer, iIndice As Integer

On Error GoTo Erro_XLS_Importa

    'Abre o excel
    lErro = objExcelApp.Abrir()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

    lErro = objExcelApp.Abrir_Planilha(sDir & sArq)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    For iLinha = 1 To 10000
    
        sValColProd = objExcelApp.Obtem_Valor_Celula(iLinha, ColProduto.ItemData(ColProduto.ListIndex))
        sValColQtde = objExcelApp.Obtem_Valor_Celula(iLinha, colQtd.ItemData(colQtd.ListIndex))
        If colPreco.ListIndex <> -1 Then
            sValColPreco = objExcelApp.Obtem_Valor_Celula(iLinha, colPreco.ItemData(colPreco.ListIndex))
        End If
        
        If Len(Trim(sValColProd)) > 0 Then
            iCountSemProd = 0
        Else
            iCountSemProd = iCountSemProd + 1
        End If
        If iCountSemProd > 50 Or UCase(sValColProd) = "VALOR TOTAL" Or UCase(sValColProd) = "FIM" Then
            Exit For
        End If
        
        If Len(Trim(sValColQtde)) > 0 And Len(Trim(sValColProd)) > 0 Then
        
            If IsNumeric(sValColQtde) Then
            
                If StrParaDbl(sValColQtde) > QTDE_ESTOQUE_DELTA Then
            
                    sProduto = sValColProd
                
                    lErro = Produto_Ajusta_Formato(sProduto, bComMask)
                    If lErro = SUCESSO Then
                    
                        If bComMask Then
                            lErro = CF("Produto_Formata", sProduto, sProdutoBD, iProdPreenchido)
                            If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                        Else
                            sProdutoBD = sProduto
                        End If
                        
                        Set objItem = New ClassItemPedido
                        
                        objItem.sProduto = sProdutoBD
                        objItem.dQuantidade = StrParaDbl(sValColQtde)
                        
                        If IsNumeric(sValColPreco) Then objItem.dPrecoUnitario = StrParaDbl(sValColPreco)
                        
                        objPV.colItensPedido.AddObj objItem
                        
                    End If
                End If
            End If
        End If
    Next
    
    For iIndice = gobjPV.colItensPedido.Count To 1 Step -1
        gobjPV.colItensPedido.Remove iIndice
    Next
    For Each objItem In objPV.colItensPedido
        gobjPV.colItensPedido.AddObj objItem
    Next
    
    lErro = objExcelApp.Fechar()
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    XLS_Importa = SUCESSO

    Exit Function

Erro_XLS_Importa:

    XLS_Importa = gErr

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 157922)

    End Select
    
    Call objExcelApp.Fechar

    Exit Function

End Function

Function Produto_Ajusta_Formato(sProduto As String, Optional bComMask As Boolean) As Long

Dim lErro As Long
Dim sProdutoNovo As String
Dim iSeg As Integer, iNumSeg As Integer
Dim objSegmento As New ClassSegmento, sSeg As String
Dim colSegmento As New Collection
Dim iPos As Integer, iTamFalta As Integer
Dim sProdSeg As String
Dim objProd As ClassProduto, iTeste As Integer, bAchouProd As Boolean
Dim objProdAux As ClassProduto, sProdutoBD As String, iProdPreenchido As Integer

On Error GoTo Erro_Produto_Ajusta_Formato

    objSegmento.sCodigo = "produto"

    'preenche toda colecao(colSegmento) em relacao ao formato corrente
    lErro = CF("Segmento_Le_Codigo", objSegmento, colSegmento)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    sProdSeg = ""
    For Each objSegmento In colSegmento
        If objSegmento.iTipo = SEGMENTO_NUMERICO Then
            Select Case objSegmento.iPreenchimento
                Case ZEROS_ESPACOS
                    sProdSeg = sProdSeg & String(objSegmento.iTamanho, "0")
                Case ESPACOS
                    sProdSeg = sProdSeg & String(objSegmento.iTamanho, " ")
            End Select
        Else
            sProdSeg = sProdSeg & String(objSegmento.iTamanho, " ")
        End If
    Next
    
    Set objSegmento = colSegmento.Item(1)
    
    iPos = InStr(1, sProduto, objSegmento.sDelimitador)
        
    'Se tiver ponto e a quantidade de pontos corresponde a quantidade de segmentos - 1
    If iPos <> 0 And (colSegmento.Count - 1) = (Len(sProduto) - Len(Replace(sProduto, ".", ""))) Then
    
        bComMask = True
        
        sSeg = Mid(sProduto, 1, iPos - 1)
        
        iTamFalta = objSegmento.iTamanho - Len(sSeg)
        
        If iTamFalta = 0 Or objSegmento.iPreenchimento = PREENCH_LIMPA_BRANCOS Then
            sProdutoNovo = sProduto
        Else
            If objSegmento.iTipo = SEGMENTO_NUMERICO Then
                Select Case objSegmento.iPreenchimento
                    Case ZEROS_ESPACOS
                        sSeg = String(iTamFalta, "0") & sSeg
                    Case ESPACOS
                        sSeg = String(iTamFalta, " ") & sSeg
                End Select
            Else
                sSeg = sSeg & String(iTamFalta, " ")
            End If
            sProdutoNovo = sSeg & Mid(sProduto, iPos)
        End If
    Else
        bComMask = False
        iTeste = 0
        bAchouProd = False
        Do While Not bAchouProd
        
            Set objProd = New ClassProduto
            Set objProdAux = New ClassProduto
            
            iTeste = iTeste + 1
            iTamFalta = Len(sProdSeg) - Len(sProduto)
            
            Select Case iTeste
                Case 1
                    objProd.sCodigo = sProduto
                Case 2
                    objProd.sCodigo = String(iTamFalta, " ") & sProduto
                Case 3
                    objProd.sCodigo = String(iTamFalta, "0") & sProduto
                Case 4
                    objProd.sCodigo = sProduto & String(iTamFalta, " ")
                Case 5
                    objProd.sCodigo = sProduto & String(iTamFalta, " ")
                Case 6
                    lErro = CF("Produto_Formata", sProduto, sProdutoBD, iProdPreenchido)
                    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
                    
                    objProd.sCodigo = sProdutoBD
                Case 7
                    objProd.sCodigo = sProduto
                    Exit Do
            End Select
            
            lErro = CF("Produto_Le", objProd)
            If lErro <> SUCESSO And lErro <> 28030 Then gError ERRO_SEM_MENSAGEM
            
            If lErro = SUCESSO Then
            
                objProdAux.sNomeReduzido = objProd.sNomeReduzido
            
                lErro = CF("Produto_Le_NomeReduzido", objProdAux)
                If lErro <> SUCESSO And lErro <> 26927 Then gError ERRO_SEM_MENSAGEM
                
                objProd.sCodigo = objProdAux.sCodigo
            
                If lErro = SUCESSO Then bAchouProd = True
            End If
            
        Loop
        
        sProdutoNovo = objProd.sCodigo
    End If
    
    sProduto = sProdutoNovo

    Produto_Ajusta_Formato = SUCESSO

    Exit Function

Erro_Produto_Ajusta_Formato:

    Produto_Ajusta_Formato = gErr

    Select Case gErr
        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 208961)

    End Select

    Exit Function

End Function
