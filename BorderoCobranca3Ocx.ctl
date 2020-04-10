VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl BorderoCobranca3Ocx 
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4395
   ScaleHeight     =   3360
   ScaleWidth      =   4395
   Begin VB.PictureBox Picture7 
      Height          =   555
      Left            =   810
      ScaleHeight     =   495
      ScaleWidth      =   2640
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2700
      Begin VB.CommandButton BotaoVoltar 
         Height          =   345
         Left            =   90
         Picture         =   "BorderoCobranca3Ocx.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   90
         Width           =   885
      End
      Begin VB.CommandButton BotaoFechar 
         Height          =   345
         Left            =   2160
         Picture         =   "BorderoCobranca3Ocx.ctx":075E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fechar"
         Top             =   90
         Width           =   420
      End
      Begin VB.CommandButton BotaoAtualizar 
         Caption         =   "Atualizar"
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
         Left            =   1035
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   90
         Width           =   1050
      End
   End
   Begin VB.CommandButton BotaoIntAtualiza 
      Caption         =   "Interromper Atualização"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   630
      TabIndex        =   0
      Top             =   2100
      Width           =   3105
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   360
      Left            =   615
      TabIndex        =   4
      Top             =   1590
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label TotalTitulos 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2430
      TabIndex        =   5
      Top             =   645
      Width           =   1350
   End
   Begin VB.Label TitulosProcessados 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   2430
      TabIndex        =   6
      Top             =   1095
      Width           =   1350
   End
   Begin VB.Label Label4 
      Caption         =   "Atualização de Arquivos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   705
      TabIndex        =   7
      Top             =   165
      Width           =   3015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Parcelas Processadas:"
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
      Left            =   435
      TabIndex        =   8
      Top             =   1125
      Width           =   1950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total de Parcelas:"
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
      Left            =   810
      TabIndex        =   9
      Top             =   675
      Width           =   1575
   End
End
Attribute VB_Name = "BorderoCobranca3Ocx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Property Variables:
Dim m_Caption As String
Event Unload()

Dim giCancelaBatch As Integer
Dim giExecutando As Integer ' 0: nao está executando, 1: em andamento

Dim gobjBorderoCobrancaEmissao As ClassBorderoCobrancaEmissao

'variaveis auxiliares para criacao da contabilizacao
Private gobjContabAutomatica As ClassContabAutomatica
Private gobjInfoParcRec As ClassInfoParcRec
Private gobjOcorrRemParcRec As ClassOcorrRemParcRec
'para evitar acessos desnecessarios durante o calculo de mnemonicos (contabilizacao)
Private giCartCobrInfoOK As Integer 'indica que as contas abaixo já foram lidas
Private gsContaCartCobr As String, gsContaDesconto As String 'contas contábeis associadas à carteira de cobranca para onde este bordero está transferindo as parcelas a receber

Private Sub BotaoFechar_Click()

    If giExecutando = ESTADO_ANDAMENTO Then
        giCancelaBatch = CANCELA_BATCH
        BotaoFechar.Enabled = False
        Exit Sub
    End If

    'Fecha a tela
    Unload Me

End Sub

Private Sub BotaoAtualizar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoAtualizar_Click

    BotaoAtualizar.Enabled = False

    BotaoIntAtualiza.Enabled = True
    
    If giCancelaBatch <> CANCELA_BATCH Then

        giExecutando = ESTADO_ANDAMENTO
        
        gobjBorderoCobrancaEmissao.objTelaAtualizacao = Me
        lErro = CF("BorderoCobranca_Grava",gobjBorderoCobrancaEmissao)
                
        giExecutando = ESTADO_PARADO

        BotaoIntAtualiza.Enabled = False

        If lErro <> SUCESSO And lErro <> 59190 Then Error 59191

        If lErro = 59190 Then Error 59192 'interrompeu

        'Chama a tela do passo seguinte
        Call Chama_Tela("BorderoCobranca4", gobjBorderoCobrancaEmissao)
                
        'Fecha a tela
        Unload Me

        'Fecha a tela
        Unload Me
    
    End If

    Exit Sub

Erro_BotaoAtualizar_Click:

    Select Case Err

        Case 59192
            lErro = Rotina_Aviso(vbOKOnly, "AVISO_BATCH_CANCELADO")
            Unload Me

        Case 59191
            Unload Me

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143661)

    End Select

    Exit Sub

End Sub

Function Trata_Parametros(objBorderoCobrancaEmissao As ClassBorderoCobrancaEmissao) As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros

    giCancelaBatch = 0
    giExecutando = ESTADO_PARADO

    Set gobjBorderoCobrancaEmissao = objBorderoCobrancaEmissao
        
    'Passa para a tela os dados dos Títulos selecionados
    TotalTitulos.Caption = CStr(objBorderoCobrancaEmissao.iQtdeParcelasSelecionadas)
    TitulosProcessados.Caption = "0"

    ProgressBar1.Min = 0
    ProgressBar1.Max = 100

    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143662)

    End Select

    giCancelaBatch = CANCELA_BATCH

    Exit Function

End Function

Public Function Mostra_Evolucao(iCancela As Integer, iNumProc As Integer) As Long

Dim lErro As Long
Dim iEventos As Integer
Dim iProcessados As Integer
Dim iTotal As Integer

On Error GoTo Erro_Mostra_Evolucao

    iEventos = DoEvents()

    If giCancelaBatch = CANCELA_BATCH Then

        iCancela = CANCELA_BATCH
        giExecutando = ESTADO_PARADO

    Else
        'atualiza dados da tela ( registros atualizados e a barra )

        iProcessados = CInt(TitulosProcessados.Caption)
        iTotal = CInt(TotalTitulos.Caption)

        iProcessados = iProcessados + iNumProc
        TitulosProcessados.Caption = CStr(iProcessados)

        ProgressBar1.Value = CInt((iProcessados / iTotal) * 100)

        giExecutando = ESTADO_ANDAMENTO

    End If

    Mostra_Evolucao = SUCESSO

    Exit Function

Erro_Mostra_Evolucao:

    Mostra_Evolucao = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143663)

    End Select

    giCancelaBatch = CANCELA_BATCH

    Exit Function

End Function

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    If giExecutando = ESTADO_ANDAMENTO Then
        If giCancelaBatch <> CANCELA_BATCH Then giCancelaBatch = CANCELA_BATCH
        Cancel = 1
    End If

End Sub

Public Sub Form_Unload(Cancel As Integer)

    Set gobjContabAutomatica = Nothing
    Set gobjInfoParcRec = Nothing
    Set gobjOcorrRemParcRec = Nothing

    Set gobjBorderoCobrancaEmissao = Nothing
    
End Sub

Private Function CarteiraCobrador_ObtemInfoContab() As Long
'funcao auxiliar a calcula_mnemonico para a obtencao de informacoes
'associadas à carteira de cobranca para onde foram transferidos as parcelas a receber

Dim lErro As Long, objCarteiraCobrador As New ClassCarteiraCobrador
Dim sContaTela As String

On Error GoTo Erro_CarteiraCobrador_ObtemInfoContab

    objCarteiraCobrador.iCobrador = gobjBorderoCobrancaEmissao.iCobrador
    objCarteiraCobrador.iCodCarteiraCobranca = gobjBorderoCobrancaEmissao.iCarteira
    
    lErro = CF("CarteiraCobrador_Le",objCarteiraCobrador)
    If lErro <> SUCESSO Then Error 32228
    
    If objCarteiraCobrador.sContaContabil <> "" Then
    
        lErro = Mascara_RetornaContaTela(objCarteiraCobrador.sContaContabil, sContaTela)
        If lErro <> SUCESSO Then Error 32229
    
        gsContaCartCobr = sContaTela
        
    Else
    
        gsContaCartCobr = ""
        
    End If
    
    If objCarteiraCobrador.sContaDuplDescontadas <> "" Then
    
        lErro = Mascara_RetornaContaTela(objCarteiraCobrador.sContaDuplDescontadas, sContaTela)
        If lErro <> SUCESSO Then Error 32230
    
        gsContaDesconto = sContaTela
        
    Else
    
        gsContaDesconto = ""
        
    End If
    
    giCartCobrInfoOK = 1
    
    CarteiraCobrador_ObtemInfoContab = SUCESSO
     
    Exit Function
    
Erro_CarteiraCobrador_ObtemInfoContab:

    CarteiraCobrador_ObtemInfoContab = Err
     
    Select Case Err
          
        Case 32228, 32229, 32230
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143664)
     
    End Select
     
    Exit Function

End Function

Public Function Calcula_Mnemonico(objMnemonicoValor As ClassMnemonicoValor) As Long

Dim lErro As Long

On Error GoTo Erro_Calcula_Mnemonico

    Select Case objMnemonicoValor.sMnemonico

        Case "Cobranca_Descontada"
        
            objMnemonicoValor.colValor.Add IIf(gobjBorderoCobrancaEmissao.iCarteira = CARTEIRA_DESCONTADA, 1, 0)
        
        Case "CartCobr_Conta"
        
            If giCartCobrInfoOK = 0 Then
            
                lErro = CarteiraCobrador_ObtemInfoContab
                If lErro <> SUCESSO Then Error 56532
                
            End If
            
            objMnemonicoValor.colValor.Add gsContaCartCobr
            
        Case "CartCobr_CtaDesconto"
        
            If giCartCobrInfoOK = 0 Then
            
                lErro = CarteiraCobrador_ObtemInfoContab
                If lErro <> SUCESSO Then Error 56533
                                
            End If
            
            objMnemonicoValor.colValor.Add gsContaDesconto
        
        Case "Numero_Bordero"
        
            objMnemonicoValor.colValor.Add gobjOcorrRemParcRec.lNumBordero

        Case "Valor_Cobrado"

            objMnemonicoValor.colValor.Add gobjOcorrRemParcRec.dValorCobrado

        Case "Cliente_Codigo"
        
            objMnemonicoValor.colValor.Add gobjInfoParcRec.lCliente
        
        Case "FilialCli_Codigo"
        
            objMnemonicoValor.colValor.Add gobjInfoParcRec.iFilialCliente
        
        Case "Titulo_Numero"
        
            objMnemonicoValor.colValor.Add gobjInfoParcRec.lNumTitulo
        
        Case "Titulo_Filial"
        
            objMnemonicoValor.colValor.Add gobjInfoParcRec.iFilialEmpresa
        
        Case "Parcela_Numero"
        
            objMnemonicoValor.colValor.Add gobjInfoParcRec.iNumParcela
        
        Case Else

            Error 56531

    End Select

    Calcula_Mnemonico = SUCESSO

    Exit Function

Erro_Calcula_Mnemonico:

    Calcula_Mnemonico = Err

    Select Case Err

        Case 56532, 56533

        Case 56531
            Calcula_Mnemonico = CONTABIL_MNEMONICO_NAO_ENCONTRADO

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143665)

    End Select

    Exit Function

End Function

Public Function GeraContabilizacao(objContabAutomatica As ClassContabAutomatica, vParams As Variant) As Long
'esta funcao é chamada a cada atualizacao de parcela e é responsavel por gerar a contabilizacao correspondente

Dim lErro As Long, lDoc As Long, iFilialEmpresaLote As Integer

On Error GoTo Erro_GeraContabilizacao

    Set gobjContabAutomatica = objContabAutomatica
    Set gobjInfoParcRec = vParams(0)
    Set gobjOcorrRemParcRec = vParams(1)
    
    'teste de filiais com autonomia contabil
    iFilialEmpresaLote = IIf(giContabCentralizada, giFilialEmpresa, gobjInfoParcRec.iFilialEmpresa)
    
    'obtem numero de doc
    lErro = objContabAutomatica.Obter_Doc(lDoc, iFilialEmpresaLote)
    If lErro <> SUCESSO Then Error 32226

    'grava a contabilizacao
    lErro = objContabAutomatica.Gravar_Registro(Me, "BorderoCobranca", gobjOcorrRemParcRec.lNumIntDoc, gobjInfoParcRec.lCliente, gobjInfoParcRec.iFilialCliente, LANPENDENTE_NAO_APROPR_CRPROD, lDoc, iFilialEmpresaLote)
    If lErro <> SUCESSO Then Error 32227
            
    GeraContabilizacao = SUCESSO
     
    Exit Function
    
Erro_GeraContabilizacao:

    GeraContabilizacao = Err
     
    Select Case Err
          
        Case 32226, 32227
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 143666)
     
    End Select
     
    Exit Function

End Function

Public Sub Form_Load()

Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    lErro_Chama_Tela = SUCESSO
    
    Exit Sub
    
Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err
    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 143667)
    
    End Select
    
    Exit Sub

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_BORDERO_COBRANCA_P3
    Set Form_Load_Ocx = Me
    Caption = "Bordero de Cobranca - Passo 3"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "BorderoCobranca3"
    
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

Private Sub BotaoVoltar_Click()

    'Chama a tela do passo anterior
    Call Chama_Tela("BorderoCobranca2", gobjBorderoCobrancaEmissao)
    
    'Fecha a tela
    Unload Me

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

'***** fim do trecho a ser copiado ******

Private Sub BotaoIntAtualiza_Click()

    If giExecutando = ESTADO_ANDAMENTO Then
        
        giCancelaBatch = CANCELA_BATCH
        Exit Sub
    
    End If
    
    'Fecha a tela
    Unload Me

End Sub


Private Sub TotalTitulos_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TotalTitulos, Source, X, Y)
End Sub

Private Sub TotalTitulos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TotalTitulos, Button, Shift, X, Y)
End Sub

Private Sub TitulosProcessados_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(TitulosProcessados, Source, X, Y)
End Sub

Private Sub TitulosProcessados_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(TitulosProcessados, Button, Shift, X, Y)
End Sub

Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label4, Source, X, Y)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label4, Button, Shift, X, Y)
End Sub

Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label2, Source, X, Y)
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label2, Button, Shift, X, Y)
End Sub

Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
   Call Controle_DragDrop(Label1, Source, X, Y)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Controle_MouseDown(Label1, Button, Shift, X, Y)
End Sub

