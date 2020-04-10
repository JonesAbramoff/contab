VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassIdiomaTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mvarsNomeTela As String
Private mvarcolTabelas As New Collection
Private mvarcolControles As New Collection
Private mvarcolIdiomas As New Collection
Private mvarobjControleAtivo As Object

Public Property Set objControleAtivo(ByVal vData As Object)
    Set mvarobjControleAtivo = vData
End Property

Public Property Get objControleAtivo() As Object
    Set objControleAtivo = mvarobjControleAtivo
End Property

Public Property Set colIdiomas(ByVal vData As Collection)
    Set mvarcolIdiomas = vData
End Property

Public Property Get colIdiomas() As Collection
    Set colIdiomas = mvarcolIdiomas
End Property

Public Property Set colControles(ByVal vData As Collection)
    Set mvarcolControles = vData
End Property

Public Property Get colControles() As Collection
    Set colControles = mvarcolControles
End Property

Public Property Set colTabelas(ByVal vData As Collection)
    Set mvarcolTabelas = vData
End Property

Public Property Get colTabelas() As Collection
    Set colTabelas = mvarcolTabelas
End Property

Public Property Let sNomeTela(ByVal vData As String)
    mvarsNomeTela = vData
End Property

Public Property Get sNomeTela() As String
    sNomeTela = mvarsNomeTela
End Property

Public Function Iniciar(ByVal sNomeTelaAux As String) As Long

Dim lErro As Long
Dim objIdioma As New ClassIdioma
Dim objControle As ClassIdiomaTelaControls
Dim objTab As ClassIdiomaTab
Dim objCampo As ClassIdiomaTabCampo
Dim objTexto As ClassIdiomaTabCampoTexto
Dim bAchou As Boolean

On Error GoTo Erro_Iniciar
    
    mvarsNomeTela = sNomeTelaAux
    
    'LÊ OS CONTROLES
    lErro = CF("IdiomaTelaControles_Le", sNomeTelaAux, mvarcolControles)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
    
    'LÊ OS IDIOMAS
    lErro = CF("Idiomas_Le", mvarcolIdiomas)
    If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
    'Monta a hierarquia de tabelas/campos/idiomas
    For Each objControle In mvarcolControles
        bAchou = False
        For Each objTab In mvarcolTabelas
            If UCase(objTab.sNomeTabela) = UCase(objControle.sNomeTabela) Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            Set objTab = New ClassIdiomaTab
            objTab.sNomeTabela = objControle.sNomeTabela
            mvarcolTabelas.Add objTab
        End If
        bAchou = False
        For Each objCampo In objTab.colCampos
            If UCase(objCampo.sNomeCampo) = UCase(objControle.sNomeCampo) Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            Set objCampo = New ClassIdiomaTabCampo
            objCampo.sNomeCampo = objControle.sNomeCampo
            objTab.colCampos.Add objCampo
        End If
    Next
    
    For Each objTab In mvarcolTabelas
        For Each objCampo In objTab.colCampos
            For Each objIdioma In mvarcolIdiomas
                If objIdioma.iPadrao = DESMARCADO Then
                    Set objTexto = New ClassIdiomaTabCampoTexto
                    objTexto.sNomeTabela = objTab.sNomeTabela
                    objTexto.sNomeCampo = objCampo.sNomeCampo
                    objTexto.iIdioma = objIdioma.iCodigo
                    objTexto.sTexto = ""
                    objCampo.colTextos.Add objTexto
                End If
            Next
        Next
    Next
    
    Iniciar = SUCESSO

    Exit Function

Erro_Iniciar:

    Iniciar = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202853)

    End Select

    Exit Function
    
End Function

Public Function Reiniciar() As Long
Dim objTab As ClassIdiomaTab
    For Each objTab In mvarcolTabelas
        Call objTab.Reiniciar
    Next
End Function

Public Function AbreTela(ByVal objControleAtivo As Object) As Long
Dim objControle As ClassIdiomaTelaControls
Dim objChamaTela As New AdmChamaTela
    For Each objControle In mvarcolControles
        'Se é um controle com essa funcionalidade
        If UCase(objControle.sNomeControle) = UCase(objControleAtivo.Name) Then
            Set mvarobjControleAtivo = objControleAtivo
            Call objChamaTela.Chama_Tela_Modal("IdiomaTextos", Me)
            Exit For
        End If
    Next
End Function

Public Function Associar(ByVal sNomeTabela As String, ByVal iLinhaGrid As Integer, ByVal sChaveDocS As String, ByVal lChaveDocL As Long, iChaveDocI As Integer) As Long
Dim objTab As ClassIdiomaTab
    For Each objTab In mvarcolTabelas
        'Se é um controle com essa funcionalidade
        If UCase(objTab.sNomeTabela) = UCase(sNomeTabela) And objTab.iLinhaGrid = iLinhaGrid Then
            objTab.sChaveDocS = sChaveDocS
            objTab.lChaveDocL = lChaveDocL
            objTab.iChaveDocI = iChaveDocI
            Exit For
        End If
    Next
End Function

Public Function Carrega_Combo_Idioma(ByVal objCombo As Object) As Long

Dim lErro As Long, iIndice As Integer, iIndexPadrao As Integer
Dim objIdioma As New ClassIdioma

On Error GoTo Erro_Carrega_Combo_Idioma
    
    iIndice = 0
    For Each objIdioma In mvarcolIdiomas
        If objIdioma.iPadrao = MARCADO Then iIndexPadrao = iIndice
        objCombo.AddItem objIdioma.iCodigo & SEPARADOR & objIdioma.sDescricao
        objCombo.ItemData(objCombo.NewIndex) = objIdioma.iCodigo
        iIndice = iIndice + 1
    Next
    objCombo.ListIndex = iIndexPadrao
    
    Carrega_Combo_Idioma = SUCESSO

    Exit Function

Erro_Carrega_Combo_Idioma:

    Carrega_Combo_Idioma = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202853)

    End Select

    Exit Function
    
End Function

Public Function Traduzir(ByVal sNomeTabela As String, ByVal sNomeCampo As String, ByVal sValorPadrao As String, ByVal sChaveDocS As String, ByVal lChaveDocL As Long, ByVal iChaveDocI As Integer, ByVal iIdioma As Integer, sRetorno As String) As Long

Dim lErro As Long, sVal As String

On Error GoTo Erro_Traduzir

    sRetorno = sValorPadrao

    If iIdioma <> 1 Then
    
        lErro = CF("IdiomaTextos_Le_Texto", sNomeTabela, sNomeCampo, sChaveDocS, lChaveDocL, iChaveDocI, iIdioma, sVal)
        If lErro <> SUCESSO Then gError ERRO_SEM_MENSAGEM
        
        If Len(Trim(sVal)) > 0 Then sRetorno = sVal
        
    End If
    
    Traduzir = SUCESSO

    Exit Function

Erro_Traduzir:

    Traduzir = gErr

    Select Case gErr
    
        Case ERRO_SEM_MENSAGEM
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 202853)

    End Select

    Exit Function
End Function
