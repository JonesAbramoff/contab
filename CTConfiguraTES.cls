VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTConfiguraTES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_objUserControl As Object

'Property Variables:
Dim m_Caption As String
Event Unload()

'Obs.: Não transferir.
Private Const TESCONFIG_AGLUTINA_LANCAM_POR_DIA = 0
Private Const TESCONFIG_GERA_LOTE_AUTOMATICO = 1
Private Const TESCONFIG_FLUXO_INCLUI_PEDIDOS = 2

Dim iAlterado As Integer

'--- inicio dos properties get dos controles da tela

Public Property Get BloqueioCTB() As Object
     Set BloqueioCTB = objUserControl.Controls("BloqueioCTB")
End Property

Public Property Get DataBloqLimite() As Object
     Set DataBloqLimite = objUserControl.Controls("DataBloqLimite")
End Property

Public Property Get UpDownDataBloqLimite() As Object
     Set UpDownDataBloqLimite = objUserControl.Controls("UpDownDataBloqLimite")
End Property

Public Property Get BotaoGravar() As Object
     Set BotaoGravar = objUserControl.Controls("BotaoGravar")
End Property

Public Property Get BotaoFechar() As Object
     Set BotaoFechar = objUserControl.Controls("BotaoFechar")
End Property

Public Property Get ListaConfigura() As Object
     Set ListaConfigura = objUserControl.Controls("ListaConfigura")
End Property

'--- fim dos properties get dos controles da tela

Public Sub BotaoFechar_Click()

    Unload Me

End Sub

Function Trata_Parametros() As Long

Dim lErro As Long

On Error GoTo Erro_Trata_Parametros
    
    Trata_Parametros = SUCESSO

    Exit Function

Erro_Trata_Parametros:

    Trata_Parametros = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155664)

    End Select

    Exit Function

End Function

Public Sub Form_Load()
           
Dim iIndice As Integer
Dim lErro As Long

On Error GoTo Erro_Form_Load
    
    'Checa Aglutina lançamentos por dia
    If gobjTES.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA Then
        ListaConfigura.Selected(TESCONFIG_AGLUTINA_LANCAM_POR_DIA) = True
    Else
        ListaConfigura.Selected(TESCONFIG_AGLUTINA_LANCAM_POR_DIA) = False
    End If
    
'    If giTipoVersao = VERSAO_FULL Then
    
        'Checa Exige Preenchimento Data de Saída
        If gobjTES.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO Then
            ListaConfigura.Selected(TESCONFIG_GERA_LOTE_AUTOMATICO) = True
        Else
            ListaConfigura.Selected(TESCONFIG_GERA_LOTE_AUTOMATICO) = False
        End If
        
'    End If

    If gobjTES.iFluxoIncluiPedidos = FLUXO_INCLUI_PEDIDOS Then
        ListaConfigura.Selected(TESCONFIG_FLUXO_INCLUI_PEDIDOS) = True
    Else
        ListaConfigura.Selected(TESCONFIG_FLUXO_INCLUI_PEDIDOS) = False
    End If
    
    If gobjTES.iMovCCVerificaBloqData = MARCADO Then
        BloqueioCTB.Value = vbChecked
    Else
        BloqueioCTB.Value = vbUnchecked
    End If
    
    If gobjTES.dtMovCCBloqDataLimite <> DATA_NULA Then Call DateParaMasked(DataBloqLimite, gobjTES.dtMovCCBloqDataLimite)

    iAlterado = 0

    lErro_Chama_Tela = SUCESSO

    Exit Sub

Erro_Form_Load:

    lErro_Chama_Tela = Err

    Select Case Err

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155665)

    End Select
    
    iAlterado = 0
    
    Exit Sub
    
End Sub

Public Sub BotaoGravar_Click()

Dim lErro As Long

On Error GoTo Erro_BotaoGravar_Click
    
    lErro = Gravar_Registro()
    If lErro <> SUCESSO Then Error 44685
    
    iAlterado = 0
    
    Unload Me
    
    Exit Sub
    
Erro_BotaoGravar_Click:

    Select Case Err

        Case 44685

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155666)

    End Select

    Exit Sub
    
End Sub
    
Public Function Gravar_Registro() As Long

Dim lErro As Long
    
On Error GoTo Erro_Gravar_Registro
    
    GL_objMDIForm.MousePointer = vbHourglass
    
    If ListaConfigura.Selected(TESCONFIG_AGLUTINA_LANCAM_POR_DIA) = True Then
        gobjTES.iAglutinaLancamPorDia = AGLUTINA_LANCAM_POR_DIA
    Else
        gobjTES.iAglutinaLancamPorDia = NAO_AGLUTINA_LANCAM_POR_DIA
    End If
    
'    If giTipoVersao = VERSAO_FULL Then
    
        If ListaConfigura.Selected(TESCONFIG_GERA_LOTE_AUTOMATICO) = True Then
            gobjTES.iGeraLoteAutomatico = GERA_LOTE_AUTOMATICO
        Else
            gobjTES.iGeraLoteAutomatico = NAO_GERA_LOTE_AUTOMATICO
        End If
    
'    End If
    
    If ListaConfigura.Selected(TESCONFIG_FLUXO_INCLUI_PEDIDOS) = True Then
        gobjTES.iFluxoIncluiPedidos = FLUXO_INCLUI_PEDIDOS
    Else
        gobjTES.iFluxoIncluiPedidos = FLUXO_NAO_INCLUI_PEDIDOS
    End If
    
    If BloqueioCTB.Value = vbChecked Then
        gobjTES.iMovCCVerificaBloqData = MARCADO
    Else
        gobjTES.iMovCCVerificaBloqData = DESMARCADO
    End If
    
    gobjTES.dtMovCCBloqDataLimite = StrParaDate(DataBloqLimite.Text)
    
    'chama gobjFAT.Gravar()
    lErro = gobjTES.Gravar()
    If lErro <> SUCESSO Then Error 34516
    
    GL_objMDIForm.MousePointer = vbDefault
    
    Gravar_Registro = SUCESSO
    
    Exit Function
    
Erro_Gravar_Registro:

    Gravar_Registro = Err

    GL_objMDIForm.MousePointer = vbDefault
    
    Select Case Err
    
        Case 34516
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error, 155667)
    End Select

    Exit Function
    
End Function

Public Sub ListaConfigura_ItemCheck(Item As Integer)
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer, iTelaCorrenteAtiva As Integer)

    Call Tela_QueryUnload(Me, iAlterado, Cancel, UnloadMode, iTelaCorrenteAtiva)

End Sub

Public Sub Form_Unload(Cancel As Integer)

End Sub

'**** inicio do trecho a ser copiado *****
Public Function Form_Load_Ocx() As Object

    Parent.HelpContextID = IDH_CONFIGURA_TES
    Set Form_Load_Ocx = Me
    Caption = "Configuração do Módulo de Tesouraria"
    Call Form_Load
    
End Function

Public Function Name() As String

    Name = "ConfiguraTES"
    
End Function

Public Sub Show()
    Parent.Show
    Parent.SetFocus
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



Public Property Get objUserControl() As Object
    Set objUserControl = m_objUserControl
End Property

Public Property Set objUserControl(ByVal vData As Object)
    Set m_objUserControl = vData
End Property

'Devolve Parent do User Control
Public Property Get Parent() As Object
    Set Parent = objUserControl.Parent
End Property

Public Property Get Controls() As Object
    Set Controls = objUserControl.Controls
End Property

Public Property Get ActiveControl() As Object
    Set ActiveControl = objUserControl.ActiveControl
End Property

Public Property Get Enabled() As Boolean
    Enabled = objUserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    objUserControl.Enabled = New_Enabled
End Property

Public Sub UpDownDataBloqLimite_DownClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataBloqLimite_DownClick

    DataBloqLimite.SetFocus

    If Len(DataBloqLimite.ClipText) > 0 Then

        sData = DataBloqLimite.Text

        lErro = Data_Diminui(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataBloqLimite.Text = sData

    End If

    Exit Sub

Erro_UpDownDataBloqLimite_DownClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155647)

    End Select

    Exit Sub

End Sub

Public Sub UpDownDataBloqLimite_UpClick()

Dim lErro As Long
Dim sData As String

On Error GoTo Erro_UpDownDataBloqLimite_UpClick

    DataBloqLimite.SetFocus

    If Len(Trim(DataBloqLimite.ClipText)) > 0 Then

        sData = DataBloqLimite.Text

        lErro = Data_Aumenta(sData)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM

        DataBloqLimite.Text = sData

    End If

    Exit Sub

Erro_UpDownDataBloqLimite_UpClick:

    Select Case gErr

        Case ERRO_SEM_MENSAGEM

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155648)

    End Select

    Exit Sub

End Sub

Public Sub DataBloqLimite_Validate(Cancel As Boolean)

Dim lErro As Long

On Error GoTo Erro_DataBloqLimite_Validate

    If Len(Trim(DataBloqLimite.ClipText)) > 0 Then
    
        lErro = Data_Critica(DataBloqLimite.Text)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    End If

    Exit Sub

Erro_DataBloqLimite_Validate:

    Cancel = True

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
            'erro tratado na rotina chamada
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 155649)

    End Select

    Exit Sub
    
End Sub

Public Sub BloqueioCTB_Click()
    iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub DataBloqLimite_Change()
    iAlterado = REGISTRO_ALTERADO
End Sub

