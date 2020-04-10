VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdmChamaTela"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Property Let lErro_Chama_Tela(ByVal vData As Long)
    lErro_Chama_Tela1 = vData
End Property

Public Property Get lErro_Chama_Tela() As Long
    lErro_Chama_Tela = lErro_Chama_Tela1
End Property

Function Chama_Tela_Nova_Instancia1(sTela As String, objTela As Object) As Long
'chama a tela passada como parametro criando uma nova instancia sem mostra-la

Dim X As Object, sProjeto As String, sClasse As String, lErro As Long
Dim Y As Object
Dim objForm As Object, objFormOcx As Object, iTelaNova As Integer, objAux1 As Object, objAux2 As Object, objAux3 As Object, objAux4 As Object, objAux5 As Object
Dim bAcrescentaOcx As Boolean

On Error GoTo Erro_Chama_Tela_Nova_Instancia1

    lErro_Chama_Tela = -1

    sProjeto = String$(NOME_PROJETO + 1, 0)
    sClasse = String$(NOME_CLASSE + 1, 0)

    lErro = Tela_ObterFuncao(sTela, sProjeto, sClasse)
    If (lErro <> AD_BOOL_TRUE) Then gError 83638

    If Tela_EmUserControl(sProjeto, bAcrescentaOcx) Then
            
        iTelaNova = 1
        
        Set X = CreateObject(gsNomePrinc & ".ClassTelaMDIChild")
        lErro = X.ObtemTelaMdiChild(sProjeto & "." & sTela & IIf(bAcrescentaOcx, "Ocx", ""), objForm, iTelaNova)
        If lErro <> SUCESSO Then gError 83639
        
        Set objTela = objForm.objFormOcx
                
    End If

    Chama_Tela_Nova_Instancia1 = SUCESSO

    Exit Function

Erro_Chama_Tela_Nova_Instancia1:

    Chama_Tela_Nova_Instancia1 = gErr

    Select Case gErr

        Case 429
            lErro = Rotina_Erro(vbOKOnly, sTela & " nao pode ser criada.", gErr, sTela)

        Case 83638
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_NAO_DISPONIVEL", gErr, sTela)

        Case 83639
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142255)

    End Select

    Exit Function

End Function

Function Chama_Tela_Nova_Instancia(sTela As String, Optional vp1 As Variant, Optional vp2 As Variant, Optional vP3 As Variant, Optional vP4 As Variant, Optional vP5 As Variant) As Long
'chama a tela passada como parametro criando uma nova instancia

Dim X As Object, sProjeto As String, sClasse As String, lErro As Long
Dim Y As Object
Dim objForm As Object, objFormOcx As Object, iTelaNova As Integer, objAux1 As Object, objAux2 As Object, objAux3 As Object, objAux4 As Object, objAux5 As Object
Dim bAcrescentaOcx As Boolean
Dim colBrowseCampo As New Collection

On Error GoTo Erro_Chama_Tela_Nova_Instancia

    lErro_Chama_Tela = -1

    sProjeto = String$(NOME_PROJETO + 1, 0)
    sClasse = String$(NOME_CLASSE + 1, 0)

    lErro = Tela_ObterFuncao(sTela, sProjeto, sClasse)
    If (lErro <> AD_BOOL_TRUE) Then gError 71820

    If Tela_EmUserControl(sProjeto, bAcrescentaOcx) Then
                
        iTelaNova = 1
        
        Set X = CreateObject(gsNomePrinc & ".ClassTelaMDIChild")
        lErro = X.ObtemTelaMdiChild(sProjeto & "." & sTela & IIf(bAcrescentaOcx, "Ocx", ""), objForm, iTelaNova)
        If lErro <> SUCESSO Then gError 71821
        
        'se estiver com algum parametro preenchido ou se estiver fazendo o form load (carga da tela)
        If Not (IsMissing(vp1)) Or iTelaNova = 1 Then
        
            Set objFormOcx = objForm.objFormOcx
            
            If VarType(vp1) = vbObject Then Set objAux1 = vp1
            If VarType(vp2) = vbObject Then Set objAux2 = vp2
            If VarType(vP3) = vbObject Then Set objAux3 = vP3
            If VarType(vP4) = vbObject Then Set objAux4 = vP4
            If VarType(vP5) = vbObject Then Set objAux5 = vP5
            
            'se for uma tela de browser
            If InStr(sTela, "Lista") <> 0 Then
            
                'mario. retirar o codigo abaixo quando tiver colocado todos os browsers no novo formato sem tela
                lErro = CF("BrowseCampo_Le", sTela, colBrowseCampo)
                If lErro <> SUCESSO Then gError 89983
            
                If colBrowseCampo.Count = 0 Then gError 92978
            
                If IsMissing(vp1) Then
                    lErro = objFormOcx.Trata_Parametros(sTela)
                ElseIf IsMissing(vp2) Then
                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1))
                ElseIf IsMissing(vP3) Then
                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2))
                ElseIf IsMissing(vP4) Then
                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3))
                ElseIf IsMissing(vP5) Then
                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4))
                Else
                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5))
                End If
            
            Else
            
Label_Chama_Tela_Nova_Instancia_Tradicional:

                If IsMissing(vp1) Then
                    lErro = objFormOcx.Trata_Parametros()
                ElseIf IsMissing(vp2) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1))
                ElseIf IsMissing(vP3) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2))
                ElseIf IsMissing(vP4) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3))
                ElseIf IsMissing(vP5) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4))
                Else
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5))
                End If
                            
            End If
                            
            If lErro <> SUCESSO And iTelaNova = 1 Then gError 71822
            If lErro <> SUCESSO Then gError 71823
            
        End If
        
        objForm.Show vbModeless
        objForm.ZOrder 0
            
    Else
    
        Set X = CreateObject(sProjeto & "." & sClasse)
    
        Select Case sTela
    
            Case "Grafico"
                lErro = Grafico_Chama_Tela(X, vp1)
                If lErro <> SUCESSO Then gError 81653
        
        End Select
    
    End If
    
    Chama_Tela_Nova_Instancia = SUCESSO

    Exit Function

Erro_Chama_Tela_Nova_Instancia:

    Chama_Tela_Nova_Instancia = gErr

    Select Case gErr

        Case 71820
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_NAO_DISPONIVEL", gErr, sTela)

        Case 71821, 71823, 81653
        
        Case 71822
            Unload objForm
        
        Case 92978
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BROWSER_SEM_BROWSECAMPO", gErr, sTela)
            Unload objForm
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142256)

    End Select

    Exit Function

End Function


Function Chama_Tela(sTela As String, Optional vp1 As Variant, Optional vp2 As Variant, Optional vP3 As Variant, Optional vP4 As Variant, Optional vP5 As Variant, Optional vP6 As Variant) As Long

Dim X As Object, sProjeto As String, sClasse As String, lErro As Long
Dim Y As Object
Dim objForm As Object, objFormOcx As Object, iTelaNova As Integer, objAux1 As Object, objAux2 As Object, objAux3 As Object, objAux4 As Object, objAux5 As Object, objAux6 As Object
Dim bAcrescentaOcx As Boolean
Dim colBrowseCampo As New Collection

On Error GoTo Erro_Chama_Tela

    If giLocalOperacao = LOCALOPERACAO_ECF Then
        
        If sTela = "BrowseConfigura" Then
        
                Set X = CreateObject("TelasAdm.ClassTelasAdm")
        
        
                If IsMissing(vp1) Then
                    lErro = BrowseConfigura_Chama_Tela(X)
                Else
                    lErro = BrowseConfigura_Chama_Tela(X, vp1)
                End If
                If lErro <> SUCESSO Then gError 9244
                
        End If
        
        
    Else
    

    lErro_Chama_Tela = -1

    sProjeto = String$(NOME_PROJETO + 1, 0)
    sClasse = String$(NOME_CLASSE + 1, 0)

    lErro = Tela_ObterFuncao(sTela, sProjeto, sClasse)
    If (lErro <> AD_BOOL_TRUE) Then gError 9125

    If Tela_EmUserControl(sProjeto, bAcrescentaOcx) Then
            
        Set X = CreateObject(gsNomePrinc & ".ClassTelaMDIChild")
        lErro = X.ObtemTelaMdiChild(sProjeto & "." & sTela & IIf(bAcrescentaOcx, "Ocx", ""), objForm, iTelaNova)
        If lErro <> SUCESSO Then gError 32218
        
        'se estiver com algum parametro preenchido ou se estiver fazendo o form load (carga da tela)
        If Not (IsMissing(vp1)) Or iTelaNova = 1 Then
        
            Set objFormOcx = objForm.objFormOcx
            
            If VarType(vp1) = vbObject Then Set objAux1 = vp1
            If VarType(vp2) = vbObject Then Set objAux2 = vp2
            If VarType(vP3) = vbObject Then Set objAux3 = vP3
            If VarType(vP4) = vbObject Then Set objAux4 = vP4
            If VarType(vP5) = vbObject Then Set objAux5 = vP5
            If VarType(vP6) = vbObject Then Set objAux6 = vP6
            
            
            'se for uma tela de browser
            If InStr(sTela, "Lista") <> 0 Then
            
                'mario. retirar o codigo abaixo quando tiver colocado todos os browsers no novo formato sem tela
                lErro = CF("BrowseCampo_Le", sTela, colBrowseCampo)
                If lErro <> SUCESSO Then gError 89983
            
                If colBrowseCampo.Count = 0 Then gError 92979
            
'                If IsMissing(vp1) Then
'                    lErro = objFormOcx.Trata_Parametros(sTela)
'                ElseIf IsMissing(vp2) Then
'                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1))
'                ElseIf IsMissing(vP3) Then
'                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2))
'                ElseIf IsMissing(vP4) Then
'                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3))
'                ElseIf IsMissing(vP5) Then
'                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4))
'                Else
'                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5))
'                End If
                            
                lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5), IIf(VarType(vP5) = vbObject, objAux6, vP6))
                            
                            
                            
            Else
            
Label_Chama_Tela_Tradicional:

                If IsMissing(vp1) Then
                    lErro = objFormOcx.Trata_Parametros()
                ElseIf IsMissing(vp2) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1))
                ElseIf IsMissing(vP3) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2))
                ElseIf IsMissing(vP4) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3))
                ElseIf IsMissing(vP5) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4))
                Else
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5))
                End If
            
            End If
            
            If lErro <> SUCESSO And iTelaNova = 1 Then Error 32219
            If lErro <> SUCESSO Then gError 32220
            
        End If
        
        objForm.Show vbModeless
'            If iTelaNova = 1 Then objForm.SetFocus
        objForm.ZOrder 0
            
    Else

        Set X = CreateObject(sProjeto & "." & sClasse)
    
        Select Case sTela
    
            Case "frmWizardEmpresa"
                lErro = frmWizardEmpresa_Chama_Tela(X, vp1)
                If lErro <> SUCESSO Then gError 20725
    
            Case "frmWizardFilial"
                lErro = frmWizardFilial_Chama_Tela(X, vp1)
                If lErro <> SUCESSO Then gError 20726
            
            Case "RelSelecionar"
                If IsMissing(vp1) Then
                    lErro = RelSelecionar_Chama_Tela(X)
                ElseIf IsMissing(vp2) Then
                    lErro = RelSelecionar_Chama_Tela(X, vp1)
                Else
                    lErro = RelSelecionar_Chama_Tela(X, vp1, vp2)
                End If
    
                If lErro <> SUCESSO Then gError 7043
    
            Case "RelCadastro"
                If IsMissing(vp1) Then
                    lErro = RelCadastro_Chama_Tela(X)
                Else
                    lErro = RelCadastro_Chama_Tela(X, vp1)
                End If
                If lErro <> SUCESSO Then gError 7044
    
            Case "BrowseConfigura"
                If IsMissing(vp1) Then
                    lErro = BrowseConfigura_Chama_Tela(X)
                Else
                    lErro = BrowseConfigura_Chama_Tela(X, vp1)
                End If
                If lErro <> SUCESSO Then gError 9244
    
            Case Else
                
                If IsMissing(vp1) Then
                    lErro = Chama_Tela2(X, sTela)
                ElseIf IsMissing(vp2) Then
                    lErro = Chama_Tela2(X, sTela, vp1)
                ElseIf IsMissing(vP3) Then
                    lErro = Chama_Tela2(X, sTela, vp1, vp2)
                ElseIf IsMissing(vP4) Then
                    lErro = Chama_Tela2(X, sTela, vp1, vp2, vP3)
                ElseIf IsMissing(vP5) Then
                    lErro = Chama_Tela2(X, sTela, vp1, vp2, vP3, vP4)
                Else
                    lErro = Chama_Tela2(X, sTela, vp1, vp2, vP3, vP4, vP5)
                End If
    
                If lErro <> SUCESSO Then gError 16811
                
        End Select

    End If
    
    End If
    
    Chama_Tela = SUCESSO

    Exit Function

Erro_Chama_Tela:

    Chama_Tela = gErr

    Select Case gErr

        Case 32218, 32220
        
        Case 32219
            Unload objForm
        
        Case 9125
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_NAO_DISPONIVEL", gErr, sTela)

        Case 9129, 9130, 9135, 9207, 9211, 9215, 9219, 7043, 7044, 7053, 14094, 14095, 14096, 14172, 14873

        Case 9223, 9227, 9231, 9238, 9239, 9243, 9244, 14044, 14053, 14065, 14092, 14093, 14097, 14761

        Case 9277, 9281, 9295, 9360, 9367, 9415, 9453, 9454, 9469, 9473, 9505, 9509, 14717, 14718, 14754

        Case 9534, 9552, 9553, 9629, 9633, 9683, 14427, 14428, 14429, 14430, 14431, 14625, 14736, 14753, 14993

        Case 9759, 9789, 9793, 14407, 14408, 14409, 14419, 14426, 14546, 14603, 14607, 14687, 14688, 14872

        Case 10037, 10264, 10904, 14483, 14494, 14509, 14534, 14535, 14542, 14552, 14572, 14573, 14624, 14916

        Case 14084, 14085, 14086, 14087, 14088, 14089, 14090, 14091, 14091, 14148, 14160, 14203, 14204, 14903

        Case 14281, 14282, 14366, 14382, 14383, 14384, 14385, 14386, 14387, 14388, 14395, 14396, 14397, 14889

        Case 14611, 14612, 14613, 14614, 14615, 14616, 14617, 14618, 14619, 14620, 14621, 14622, 14623, 14929

        Case 14797, 14798, 14799, 14800, 14801, 14802, 14803, 14804, 14805, 14806, 14807, 14808, 14809, 14992

        Case 16811, 19769, 19770, 20263, 20264, 20265, 20266, 20307, 20308, 20309, 20324, 20325, 20468, 20643
        
        Case 20725, 20726

        Case 24001, 24002, 24035, 24055, 24059, 24060, 24075, 24103, 24104, 24105, 24115, 24144, 24163, 24164

        Case 24165, 24183, 24195, 24215, 24219, 24220, 24227, 24228, 24229, 24239, 24240, 24241, 24252, 24264

        Case 24275, 24287, 24301, 24307, 24308, 24316, 24329, 24331, 24335, 24336, 24344, 24345, 24346

        Case 24347, 24348, 24349, 24358, 24363, 24367, 24387, 24391, 24395, 24402, 24425, 24445, 24446, 24453

        Case 24465, 24490, 24491, 24498, 24499, 24500, 24514, 24522, 24530, 24550, 24551, 24557, 24585

        Case 24586, 24602, 24611, 24627, 24637, 24642, 24643, 24650, 24662, 24667, 24672, 24676, 24677

        Case 24683, 24688, 24694, 24707, 24719, 24736, 24737, 24750, 24768, 24769, 24784, 24796, 24814, 24830

        Case 24831, 24834, 24842, 24854, 24866, 24908, 24912, 24925, 24937, 24951, 24955, 24959, 24964, 24980

        Case 24981, 24982, 24983, 24984, 24991, 29510, 29522, 29534, 29546, 29565, 29566, 29567, 29568, 29569

        Case 29570, 29594, 29595, 29596, 29597, 29598, 29614, 29628, 29630, 29632, 29636, 29637, 29640, 29661

        Case 29662, 29669, 29670, 29681, 29691, 29703, 29714, 29718, 29730, 29750, 29751, 29773, 29774, 29788

        Case 29816, 29817, 29818, 29837, 29857, 29858, 29876, 29877, 54798

        Case 92979
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BROWSER_SEM_BROWSECAMPO", gErr, sTela)
            Unload objForm
            
        Case 401
            If sTela = "AvisoWFW" Then Resume Next
            Call Rotina_Erro(vbOKOnly, "ERRO_TELA_NAO_MODAL", gErr)

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142257)

    End Select

    Exit Function

End Function

Function Chama_Tela2(X As Object, sTela As String, Optional vp1 As Variant, Optional vp2 As Variant, Optional vP3 As Variant, Optional vP4 As Variant, Optional vP5 As Variant) As Long
'Alterada Por Tulio em 09/01/2002
'1)Inclusao do case para "EnviarEmail"
'2)Troca de Err pra gErr
'3)Linha "Dim Y As Object" comentada (parece uma declaracao inutil)
'

Dim lErro As Long
'Dim Y As Object

On Error GoTo Erro_Chama_Tela2

    Select Case sTela
    
        Case "Moedas"
            lErro = Moedas_Chama_Tela(X)
            If lErro <> SUCESSO Then gError 108840
        
        Case "Feriados"
            lErro = Feriados_Chama_Tela(X)
            If lErro <> SUCESSO Then gError 45292
            
        Case "CotacaoMoeda"
            lErro = CotacaoMoeda_Chama_Tela(X)
            If lErro <> SUCESSO Then gError 45292
            
        Case "Modulo"
            If IsMissing(vp1) Then
                lErro = Modulo_Chama_Tela(X)
            Else
                lErro = Modulo_Chama_Tela(X, vp1)
            End If
            If lErro <> SUCESSO Then gError 44428
        
        Case "Calendario"
            If IsMissing(vp1) Then
                lErro = Calendario_Chama_Tela(X)
            Else
                lErro = Calendario_Chama_Tela(X, vp1)
            End If
            If lErro <> SUCESSO Then gError 45542
            
        Case "TrataAlteracao"
            lErro = TrataAlteracao_Chama_Tela(X, vp1)
            If lErro <> SUCESSO Then gError 45542
            
        Case "Propriedades"
            lErro = Propriedades_Chama_Tela(X)
            If lErro <> SUCESSO Then gError 60851
            
        Case "CamposInvisiveis"
            lErro = CamposInvisiveis_Chama_Tela(X)
            If lErro <> SUCESSO Then gError 60852
    
        Case "EnviarEmail"
            lErro = EnviarEmail_Chama_Tela(X, vp1, vp2)
            If lErro <> SUCESSO Then gError 97116
    
        Case "RelDlgExec"
            lErro = RelDlgExec_Chama_Tela(X, vp1)
            If lErro <> SUCESSO Then gError 97116
            
        '########################################
        'Inserido por Wagner
        Case "PlanilhasSelecionar"
            lErro = PlanilhasSelecionar_Chama_Tela(X, vp1)
            If lErro <> SUCESSO Then gError 97116
            
        Case "OLAP"
            lErro = OLAP_Chama_Tela(X)
            If lErro <> SUCESSO Then gError 97116
        
        Case "Wait"
            lErro = Wait_Chama_Tela(X, vp1)
            If lErro <> SUCESSO Then gError 97116
        
        Case "LocalizarGrid"
            lErro = LocalizarGrid_Chama_Tela(X, vp1)
            If lErro <> SUCESSO Then gError 97116
            
        Case "BackupConfig"
            lErro = BackupConfig_Chama_Tela(X)
            If lErro <> SUCESSO Then gError 45292

        Case "Logo"
            lErro = Logo_Chama_Tela(X)
            If lErro <> SUCESSO Then gError 45292

        Case "Arquivamento"
            lErro = Arquivamento_Chama_Tela(X)
            If lErro <> SUCESSO Then gError 45292

        '########################################
    
    End Select

    Chama_Tela2 = SUCESSO

    Exit Function

Erro_Chama_Tela2:

    Chama_Tela2 = gErr

    Select Case gErr
        
        Case 16800, 16801, 16802, 16803, 16804, 16805, 16806, 16807, 16808, 16809, 16810, 16822, 16834, 16838, 16850

        Case 16858, 16891, 16907, 16924, 16925, 16948, 16949, 16950, 18527, 18528
        
        Case 19505, 19506, 19507, 19508, 19554, 19555, 19586, 19587, 19594, 19598, 19635, 19636, 19657, 19658, 19748
        
        Case 19785, 19797, 19817, 19818, 19825, 19837, 18857, 19858, 19873, 19929, 19984, 19994, 20535, 20539, 20547
        
        Case 20602, 20543, 24315, 24466, 24666, 108840
        
        Case 29985, 29986, 29987, 29988, 29989, 29990, 29991, 29992, 29993, 29994, 29995, 29996, 29997, 29998, 29999
        
        Case 41792, 44428, 44498, 44639, 44643, 44700, 44906
        
        Case 45024, 45025, 45026, 45027, 45042, 45046, 45047, 45068, 45075, 45083, 45087, 45093, 45108, 45112, 45141
        
        Case 45142, 45143, 45144, 45157, 45158, 45169, 45173, 45174, 45175, 45176, 45177, 45178, 45208, 45210, 45211
        
        Case 45212, 45222, 45223, 45224, 45234, 45235, 45236, 45237, 45238, 45248, 45249, 45252, 45253, 45254, 45255
        
        Case 45276, 45288, 45292, 45302, 45307, 45311, 45312, 45313, 45314, 45335, 45405, 45416, 45437, 45438, 45466
        
        Case 45478, 45490, 45503, 45515, 45521, 45534, 45542, 45551, 45565, 45635, 45647, 45738, 45792, 45852, 45853
        
        Case 45867, 45871, 45876, 45877, 45881, 45891, 45904, 45931, 45944, 45957, 45961, 51204, 51205, 51212, 51214

        Case 51285, 51286, 54657, 54658, 51306, 51307, 51308, 51309, 51314, 51315, 51316, 51317, 51318, 51319, 51320
        
        Case 57217, 57218, 57219, 57220, 57221, 57222, 57223, 57224, 57225, 57226, 57227, 57228, 57229, 57230, 54719, 51342
        
        Case 60851, 60852, 97116
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142258)
        
    End Select
        
    Exit Function
    
End Function

Function BrowseConfigura_Chama_Tela(X As Object, Optional vp1 As Variant) As Long

Dim lErro As Long
Dim objBrowseConfigura1 As AdmBrowseConfigura

On Error GoTo Erro_BrowseConfigura_Chama_Tela

    Load X.objBrowseConfigura
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 9245
            
    If IsMissing(vp1) Then
    
        Error 9248
        
    Else
    
        Set objBrowseConfigura1 = vp1
        
    End If
    
    lErro = X.objBrowseConfigura.Trata_Parametros(objBrowseConfigura1)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 9246
            
    If lErro <> SUCESSO Then Error 9247
        
    X.objBrowseConfigura.Show vbModal
            
    BrowseConfigura_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BrowseConfigura_Chama_Tela:

    BrowseConfigura_Chama_Tela = Err

    Select Case Err
    
        Case 9247
    
        Case 9246, 9245
            Unload X.objBrowseConfigura
            
        Case 9248
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BROWSECONFIGURA_SEM_PARAMETRO", Err)
            
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142259)
        
    End Select
        
    Exit Function

End Function

Function Modulo_Chama_Tela(X As Object, Optional vp1 As Variant) As Long

Dim lErro As Long
Dim objTelaPrincipal As Form

On Error GoTo Erro_Modulo_Chama_Tela

    Load X.objModulo
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 44425
            
    If IsMissing(vp1) Then
    
        Set objTelaPrincipal = Nothing
        
    Else
    
        Set objTelaPrincipal = vp1
        
    End If
    
    lErro = X.objModulo.Trata_Parametros(objTelaPrincipal)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 44426
            
    If lErro <> SUCESSO Then Error 44427
        
    X.objModulo.Show vbModal
            
    Modulo_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Modulo_Chama_Tela:

    Modulo_Chama_Tela = Err

    Select Case Err
    
        Case 44427
    
        Case 44426, 44425
            Unload X.objModulo
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142260)
        
    End Select
        
    Exit Function

End Function


Function RelCadastro_Chama_Tela(X As Object, Optional vp1 As Variant) As Long

Dim lErro As Long
Dim objRelatorio As AdmRelatorio

On Error GoTo Erro_RelCadastro_Chama_Tela

    Load X.objRelCadastro
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 7045
            
    If IsMissing(vp1) Then
    
        Set objRelatorio = Nothing
        
    Else
    
        Set objRelatorio = vp1
        
    End If
    
    lErro = X.objRelCadastro.Trata_Parametros(objRelatorio)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 7046
            
    If lErro <> SUCESSO Then Error 7047
        
    X.objRelCadastro.Show vbModal
            
    RelCadastro_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelCadastro_Chama_Tela:

    RelCadastro_Chama_Tela = Err

    Select Case Err
        
        Case 7046, 7045
            Unload X.objRelCadastro
            
        Case 7047
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
            
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142261)
        
    End Select
        
    Exit Function

End Function

Function RelSelecionar_Chama_Tela(X As Object, Optional vp1 As Variant, Optional vp2 As Variant) As Long

Dim lErro As Long
Dim objRelSel As AdmRelSel, sSiglaModulo As String

On Error GoTo Erro_RelSelecionar_Chama_Tela

    Load X.objRelSelecionar
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 7049
            
    If IsMissing(vp1) Or IsMissing(vp2) Then
    
        Error 7052
        
    Else
    
        Set objRelSel = vp1
        sSiglaModulo = vp2
        
    End If
    
    lErro = X.objRelSelecionar.Trata_Parametros(objRelSel, sSiglaModulo)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 7050
            
    If lErro <> SUCESSO Then Error 7051
        
    X.objRelSelecionar.Show vbModal
            
    RelSelecionar_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelSelecionar_Chama_Tela:

    RelSelecionar_Chama_Tela = Err

    Select Case Err
    
        Case 7051
    
        Case 7050, 7049
            Unload X.objRelSelecionar
            
        Case 7052
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142262)
        
    End Select
        
    Exit Function

End Function

Function Reimpressao_Chama_Tela(X As Object, Optional vp1 As Variant, Optional vp2 As Variant) As Long

Dim lErro As Long
Dim objRelSel As AdmRelSel, sSiglaModulo As String

On Error GoTo Erro_Reimpressao_Chama_Tela

    Load X.objReimpressao
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45204
            
    If IsMissing(vp1) Or IsMissing(vp2) Then
    
        Error 45207
        
    Else
    
        Set objRelSel = vp1
        sSiglaModulo = vp2
        
    End If
    
    lErro = X.objReimpressao.Trata_Parametros(objRelSel, sSiglaModulo)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45205
            
    If lErro <> SUCESSO Then Error 45206
        
    X.objReimpressao.Show vbModal
            
    Reimpressao_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Reimpressao_Chama_Tela:

    Reimpressao_Chama_Tela = Err

    Select Case Err
    
        Case 45206
    
        Case 45205, 45204
            Unload X.objReimpressao
            
        Case 45207
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_SEM_PARAMETRO", Err)
            
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142263)
        
    End Select
        
    Exit Function

End Function

Function Feriados_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Feriados_Chama_Tela

    Load X.objFeriados
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45293
    
    X.objFeriados.Show
            
    Feriados_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Feriados_Chama_Tela:

    Feriados_Chama_Tela = Err

    Select Case Err
    
        Case 45293
            Unload X.objFeriados
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142264)
        
    End Select
        
    Exit Function

End Function

Function Calendario_Chama_Tela(X As Object, Optional vp1 As Variant) As Long

Dim lErro As Long
Dim objData As AdmGenerico

On Error GoTo Erro_Calendario_Chama_Tela

    Load X.objCalendario
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45539
                
    If IsMissing(vp1) Then
    
        Set objData = Nothing
        
    Else
    
        Set objData = vp1
        
    End If
    
    lErro = X.objCalendario.Trata_Parametros(objData)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45540
            
    If lErro <> SUCESSO Then Error 45541

    X.objCalendario.Show vbModal
            
    Calendario_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Calendario_Chama_Tela:

    Calendario_Chama_Tela = Err

    Select Case Err
    
        Case 45539, 45540
            Unload X.objCalendario
            
        Case 45541
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142265)
        
    End Select
        
    Exit Function

End Function

Function frmWizardEmpresa_Chama_Tela(X As Object, vp1 As Variant) As Long

Dim lErro As Long
Dim objConfiguraADM As ClassConfiguraADM

On Error GoTo Erro_frmWizardEmpresa_Chama_Tela

    Load X.objfrmWizardEmpresa
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 44878
            
    Set objConfiguraADM = vp1
        
    lErro = X.objfrmWizardEmpresa.Trata_Parametros(objConfiguraADM)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 44879
            
    If lErro <> SUCESSO Then Error 44880
        
    X.objfrmWizardEmpresa.Show vbModal
            
    frmWizardEmpresa_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_frmWizardEmpresa_Chama_Tela:

    frmWizardEmpresa_Chama_Tela = Err

    Select Case Err
    
        Case 44878
    
        Case 44879, 44880
            Unload X.objfrmWizardEmpresa
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142266)
        
    End Select
        
    Exit Function

End Function

Function frmWizardFilial_Chama_Tela(X As Object, vp1 As Variant) As Long

Dim lErro As Long
Dim objConfiguraADM As ClassConfiguraADM

On Error GoTo Erro_frmWizardFilial_Chama_Tela

    Load X.objfrmWizardFilial
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 44881
            
    Set objConfiguraADM = vp1
        
    lErro = X.objfrmWizardFilial.Trata_Parametros(objConfiguraADM)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 44882
            
    If lErro <> SUCESSO Then Error 44883
        
    X.objfrmWizardFilial.Show vbModal
            
    frmWizardFilial_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_frmWizardFilial_Chama_Tela:

    frmWizardFilial_Chama_Tela = Err

    Select Case Err
    
        Case 44881
    
        Case 44882, 44883
            Unload X.objfrmWizardFilial
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142267)
        
    End Select
        
    Exit Function

End Function

Function Chama_Tela_Modal(sTela As String, Optional vp1 As Variant, Optional vp2 As Variant, Optional vP3 As Variant, Optional vP4 As Variant, Optional vP5 As Variant) As Long

Dim X As Object, sProjeto As String, sClasse As String, lErro As Long
Dim Y As Object, objForm As Object, objFormOcx As Object, iTelaNova As Integer
Dim objAux1 As Object, objAux2 As Object, objAux3 As Object, objAux4 As Object, objAux5 As Object
Dim bAcrescentaOcx As Boolean, sNomeTela As String
Dim colBrowseCampo As New Collection

On Error GoTo Erro_Chama_Tela_Modal

    lErro_Chama_Tela = -1

    sProjeto = String$(NOME_PROJETO + 1, 0)
    sClasse = String$(NOME_CLASSE + 1, 0)
    
    If sTela = "RomaneioGrade" Then
    
        lErro = CF("RomaneioGrade_Le_Tela", sNomeTela, vp1)
        If lErro <> SUCESSO Then gError 89983
        
        sTela = sNomeTela
    
    End If

    lErro = Tela_ObterFuncao(sTela, sProjeto, sClasse)
    If (lErro <> AD_BOOL_TRUE) Then Error 56578

    If Tela_EmUserControl(sProjeto, bAcrescentaOcx) Then
            
        Set X = CreateObject(gsNomePrinc & ".ClassTelaPopUp")
        lErro = X.ObtemTela(sProjeto & "." & sTela & IIf(bAcrescentaOcx, "Ocx", ""), objForm, iTelaNova)
        If lErro <> SUCESSO Then Error 56579
        
        'se estiver com algum parametro preenchido ou se estiver fazendo o form load (carga da tela)
        If Not (IsMissing(vp1)) Or iTelaNova = 1 Then
        
            Set objFormOcx = objForm.objFormOcx
            
            If VarType(vp1) = vbObject Then Set objAux1 = vp1
            If VarType(vp2) = vbObject Then Set objAux2 = vp2
            If VarType(vP3) = vbObject Then Set objAux3 = vP3
            If VarType(vP4) = vbObject Then Set objAux4 = vP4
            If VarType(vP5) = vbObject Then Set objAux5 = vP5
            
            'se for uma tela de browser
            If InStr(sTela, "Lista") <> 0 Then
            
                'mario. retirar o codigo abaixo quando tiver colocado todos os browsers no novo formato sem tela
                lErro = CF("BrowseCampo_Le", sTela, colBrowseCampo)
                If lErro <> SUCESSO Then gError 89983
            
                If colBrowseCampo.Count = 0 Then gError 92979
            
                If IsMissing(vp1) Then
                    lErro = objFormOcx.Trata_Parametros(sTela)
                ElseIf IsMissing(vp2) Then
                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1))
                ElseIf IsMissing(vP3) Then
                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2))
                ElseIf IsMissing(vP4) Then
                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3))
                ElseIf IsMissing(vP5) Then
                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4))
                Else
                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5))
                End If
                            
            Else
            
Label_Chama_Tela_Tradicional:
            
                If IsMissing(vp1) Then
                    lErro = objFormOcx.Trata_Parametros()
                ElseIf IsMissing(vp2) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1))
                ElseIf IsMissing(vP3) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2))
                ElseIf IsMissing(vP4) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3))
                ElseIf IsMissing(vP5) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4))
                Else
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5))
                End If
                            
            End If
                            
            If lErro <> SUCESSO And iTelaNova = 1 Then Error 56580
            If lErro <> SUCESSO Then Error 56581
            
        End If
        
        objForm.Show vbModal
    
    End If
    
    Chama_Tela_Modal = SUCESSO

    Exit Function

Erro_Chama_Tela_Modal:

    Chama_Tela_Modal = Err

    Select Case Err

        Case 56578
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_NAO_DISPONIVEL", Err, sTela)

        Case 56579, 56581
        
        Case 56580
            Unload objForm
                
        Case 429
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_NAO_DISPONIVEL", Err, sTela)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142268)

    End Select

    Exit Function

End Function

Function Chama_Tela_Popup(sTela As String, Optional vp1 As Variant, Optional vp2 As Variant, Optional vP3 As Variant, Optional vP4 As Variant, Optional vP5 As Variant) As Long

Dim X As Object, sProjeto As String, sClasse As String, lErro As Long
Dim Y As Object, objForm As Object, objFormOcx As Object, iTelaNova As Integer
Dim objAux1 As Object, objAux2 As Object, objAux3 As Object, objAux4 As Object, objAux5 As Object
Dim bAcrescentaOcx As Boolean

On Error GoTo Erro_Chama_Tela_Popup

    lErro_Chama_Tela = -1

    sProjeto = String$(NOME_PROJETO + 1, 0)
    sClasse = String$(NOME_CLASSE + 1, 0)

    lErro = Tela_ObterFuncao(sTela, sProjeto, sClasse)
    If (lErro <> AD_BOOL_TRUE) Then Error 54714

    If Tela_EmUserControl(sProjeto, bAcrescentaOcx) Then
    
        Set X = CreateObject(gsNomePrinc & ".ClassTelaPopUp")
        lErro = X.ObtemTela(sProjeto & "." & sTela & IIf(bAcrescentaOcx, "Ocx", ""), objForm, iTelaNova)
        If lErro <> SUCESSO Then Error 54715
        
        'se estiver com algum parametro preenchido ou se estiver fazendo o form load (carga da tela)
        If Not (IsMissing(vp1)) Or iTelaNova = 1 Then
        
            Set objFormOcx = objForm.objFormOcx
            
            If VarType(vp1) = vbObject Then Set objAux1 = vp1
            If VarType(vp2) = vbObject Then Set objAux2 = vp2
            If VarType(vP3) = vbObject Then Set objAux3 = vP3
            If VarType(vP4) = vbObject Then Set objAux4 = vP4
            If VarType(vP5) = vbObject Then Set objAux5 = vP5
            
            If IsMissing(vp1) Then
                lErro = objFormOcx.Trata_Parametros()
            ElseIf IsMissing(vp2) Then
                lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1))
            ElseIf IsMissing(vP3) Then
                lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2))
            ElseIf IsMissing(vP4) Then
                lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3))
            ElseIf IsMissing(vP5) Then
                lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4))
            Else
                lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5))
            End If
                            
            If lErro <> SUCESSO And iTelaNova = 1 Then Error 54716
            If lErro <> SUCESSO Then Error 54717
            
        End If
        
        objForm.Show
            
    End If
    
    Chama_Tela_Popup = SUCESSO

    Exit Function

Erro_Chama_Tela_Popup:

    Chama_Tela_Popup = Err

    Select Case Err

        Case 54714
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_NAO_DISPONIVEL", Err, sTela)

        Case 54715, 54717
        
        Case 54716
            Unload objForm
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142269)

    End Select

    Exit Function

End Function

Function Propriedades_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Propriedades_Chama_Tela

    Load X.objPropriedades
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 60853
    
    X.objPropriedades.Show
            
    Propriedades_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Propriedades_Chama_Tela:

    Propriedades_Chama_Tela = Err

    Select Case Err
    
        Case 60853
            Unload X.objPropriedades
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142270)
        
    End Select
        
    Exit Function

End Function

Function CamposInvisiveis_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_CamposInvisiveis_Chama_Tela

    Load X.objCamposInvisiveis
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 60854
    
    X.objCamposInvisiveis.Show
            
    CamposInvisiveis_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CamposInvisiveis_Chama_Tela:

    CamposInvisiveis_Chama_Tela = Err

    Select Case Err
    
        Case 60854
            Unload X.objCamposInvisiveis
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142271)
        
    End Select
        
    Exit Function

End Function

Function CotacaoMoeda_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_CotacaoMoeda_Chama_Tela

    Load X.objCotacaoMoeda
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45293
    
    X.objCotacaoMoeda.Show
            
    CotacaoMoeda_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_CotacaoMoeda_Chama_Tela:

    CotacaoMoeda_Chama_Tela = Err

    Select Case Err
    
        Case 45293
            Unload X.objCotacaoMoeda
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142272)
        
    End Select
        
    Exit Function

End Function

Function EnviarEmail_Chama_Tela(X As Object, vp1 As Variant, vp2 As Variant) As Long
'Criada por Tulio em 09/01/2002

Dim lErro As Long
Dim objRelOp As AdmRelOpcoes
Dim sOpcao As String

On Error GoTo Erro_EnviarEmail_Chama_Tela

    Load X.objEnviarEmail
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then gError 97117
    
    Set objRelOp = vp1
    sOpcao = CStr(vp2)
    
    lErro = X.objEnviarEmail.Trata_Parametros(objRelOp, sOpcao)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then gError 97118
    
    If lErro <> SUCESSO Then gError 97119
    
    X.objEnviarEmail.Show vbModal
            
    EnviarEmail_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_EnviarEmail_Chama_Tela:

    EnviarEmail_Chama_Tela = gErr

    Select Case gErr
    
        Case 97117, 97118
            Unload X.objEnviarEmail
                    
        Case 97119
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142273)
        
    End Select
        
    Exit Function

End Function

Function TrataAlteracao_Chama_Tela(X As Object, vp1 As Variant) As Long

Dim lErro As Long, objObjetoBD As ClassObjetoBD

On Error GoTo Erro_TrataAlteracao_Chama_Tela

    Load X.objTrataAlteracao
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45539
                    
    Set objObjetoBD = vp1
        
    lErro = X.objTrataAlteracao.Trata_Parametros(objObjetoBD)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then Error 45540
            
    If lErro <> SUCESSO Then Error 45541

    X.objTrataAlteracao.Show vbModal
            
    TrataAlteracao_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_TrataAlteracao_Chama_Tela:

    TrataAlteracao_Chama_Tela = Err

    Select Case Err
    
        Case 45539, 45540
            Unload X.objTrataAlteracao
            
        Case 45541
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142274)
        
    End Select
        
    Exit Function

End Function

Function Grafico_Chama_Tela(X As Object, vp1 As Variant) As Long

Dim lErro As Long
Dim objGrafico As ClassGrafico

On Error GoTo Erro_Grafico_Chama_Tela

    Load X.objGrafico
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then gError 81650
            
    Set objGrafico = vp1
        
    lErro = X.objGrafico.Trata_Parametros(objGrafico)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then gError 81651
            
    If lErro <> SUCESSO Then gError 81652
        
    X.objGrafico.Show
            
    Grafico_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Grafico_Chama_Tela:

    Grafico_Chama_Tela = gErr

    Select Case gErr
    
        Case 81650
    
        Case 81651, 81652
            Unload X.objGrafico
            
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142275)
        
    End Select
        
    Exit Function

End Function

Function Chama_TelaECF(sTela As String, Optional vp1 As Variant, Optional vp2 As Variant, Optional vP3 As Variant, Optional vP4 As Variant, Optional vP5 As Variant) As Long

Dim X As Object, sProjeto As String, sClasse As String, lErro As Long
Dim Y As Object
Dim objForm As Object, objFormOcx As Object, iTelaNova As Integer, objAux1 As Object, objAux2 As Object, objAux3 As Object, objAux4 As Object, objAux5 As Object
Dim bAcrescentaOcx As Boolean
Dim colBrowseCampo As New Collection

On Error GoTo Erro_Chama_TelaECF



'    lErro_Chama_Tela = -1
'
'    sProjeto = String$(NOME_PROJETO + 1, 0)
'    sClasse = String$(NOME_CLASSE + 1, 0)
'
'    lErro = Tela_ObterFuncao(sTela, sProjeto, sClasse)
'    If (lErro <> AD_BOOL_TRUE) Then gError 9125
'
'    If Tela_EmUserControl(sProjeto, bAcrescentaOcx) Then
'
'        Set X = CreateObject(gsNomePrinc & ".ClassTelaMDIChild")
'        lErro = X.ObtemTelaMdiChild(sProjeto & "." & sTela & IIf(bAcrescentaOcx, "Ocx", ""), objForm, iTelaNova)
'        If lErro <> SUCESSO Then gError 32218
'
'        'se estiver com algum parametro preenchido ou se estiver fazendo o form load (carga da tela)
'        If Not (IsMissing(vp1)) Or iTelaNova = 1 Then
'
'            Set objFormOcx = objForm.objFormOcx
'
'            If VarType(vp1) = vbObject Then Set objAux1 = vp1
'            If VarType(vp2) = vbObject Then Set objAux2 = vp2
'            If VarType(vP3) = vbObject Then Set objAux3 = vP3
'            If VarType(vP4) = vbObject Then Set objAux4 = vP4
'            If VarType(vP5) = vbObject Then Set objAux5 = vP5
'            If VarType(vP6) = vbObject Then Set objAux6 = vP6
'
'
'            'se for uma tela de browser
'            If InStr(sTela, "Lista") <> 0 Then
'
'                'mario. retirar o codigo abaixo quando tiver colocado todos os browsers no novo formato sem tela
'                lErro = CF("BrowseCampo_Le", sTela, colBrowseCampo)
'                If lErro <> SUCESSO Then gError 89983
'
'                If colBrowseCampo.Count = 0 Then gError 92979
'
''                If IsMissing(vp1) Then
''                    lErro = objFormOcx.Trata_Parametros(sTela)
''                ElseIf IsMissing(vp2) Then
''                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1))
''                ElseIf IsMissing(vP3) Then
''                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2))
''                ElseIf IsMissing(vP4) Then
''                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3))
''                ElseIf IsMissing(vP5) Then
''                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4))
''                Else
''                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5))
''                End If
'
'                lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5), IIf(VarType(vP5) = vbObject, objAux6, vP6))
'
'
'
'            Else
'
'Label_Chama_Tela_Tradicional:
'
'                If IsMissing(vp1) Then
'                    lErro = objFormOcx.Trata_Parametros()
'                ElseIf IsMissing(vp2) Then
'                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1))
'                ElseIf IsMissing(vP3) Then
'                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2))
'                ElseIf IsMissing(vP4) Then
'                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3))
'                ElseIf IsMissing(vP5) Then
'                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4))
'                Else
'                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5))
'                End If
'
'            End If
'
'            If lErro <> SUCESSO And iTelaNova = 1 Then Error 32219
'            If lErro <> SUCESSO Then gError 32220
'
'        End If
'
'        objForm.Show vbModeless
''            If iTelaNova = 1 Then objForm.SetFocus
'        objForm.ZOrder 0







'        Set X = CreateObject(gsNomePrinc & ".ClassTelaMDIChild")
'        lErro = X.ObtemTelaMdiChild(sProjeto & "." & sTela & IIf(bAcrescentaOcx, "Ocx", ""), objForm, iTelaNova)
'        If lErro <> SUCESSO Then gError 32218




    lErro_Chama_Tela = -1

    Set X = CreateObject(gsNomePrinc & ".ClassTelaMDIChild")
    
    lErro = Rotina_ObterObjetoTela_ECF(sTela)
    If lErro <> SUCESSO Then gError 199480
    
    lErro = X.ObtemTelaMdiChild("TelasECF" & "." & sTela, objForm, iTelaNova)
    If lErro <> SUCESSO Then gError 32218
    
    'se estiver com algum parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vp1)) Or iTelaNova = 1 Then
    
        Set objFormOcx = objForm.objFormOcx
        
        If VarType(vp1) = vbObject Then Set objAux1 = vp1
        If VarType(vp2) = vbObject Then Set objAux2 = vp2
        If VarType(vP3) = vbObject Then Set objAux3 = vP3
        If VarType(vP4) = vbObject Then Set objAux4 = vP4
        If VarType(vP5) = vbObject Then Set objAux5 = vP5
        
        'se for uma tela de browser
        If InStr(sTela, "ProdutosLista") <> 0 Then

                    'mario. retirar o codigo abaixo quando tiver colocado todos os browsers no novo formato sem tela
                    lErro = CF("BrowseCampo_Le", sTela, colBrowseCampo)
                    If lErro <> SUCESSO Then gError 214924

                    If colBrowseCampo.Count = 0 Then gError 214925

                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5))
        Else


'                    If IsMissing(vp1) Then
'                        lErro = objFormOcx.Trata_Parametros(sTela)
'                    ElseIf IsMissing(vp2) Then
'                        lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1))
'                    ElseIf IsMissing(vP3) Then
'                        lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2))
'                    ElseIf IsMissing(vP4) Then
'                        lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3))
'                    ElseIf IsMissing(vP5) Then
'                        lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4))
'                    Else
'                        lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5))
'                    End If
'
'                Else
'
'Label_Chama_Tela_Tradicional:
'
'                    If IsMissing(vp1) Then
'                        lErro = objFormOcx.Trata_Parametros()
'                    ElseIf IsMissing(vp2) Then
'                        lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1))
'                    ElseIf IsMissing(vP3) Then
'                        lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2))
'                    ElseIf IsMissing(vP4) Then
'                        lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3))
'                    ElseIf IsMissing(vP5) Then
'                        lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4))
'                    Else
'                        lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5))
'                    End If
'
'                End If
'
'                If lErro <> SUCESSO And iTelaNova = 1 Then Error 32219
'                If lErro <> SUCESSO Then gError 32220
'
'            End If
            
'            objForm.Show vbModeless
''            If iTelaNova = 1 Then objForm.SetFocus
'            objForm.ZOrder 0
'
'            Exit Function
'
   
            
            If IsMissing(vp1) Then
                lErro = Chama_Tela2(X, sTela)
            ElseIf IsMissing(vp2) Then
                lErro = Chama_Tela2(X, sTela, vp1)
            ElseIf IsMissing(vP3) Then
                lErro = Chama_Tela2(X, sTela, vp1, vp2)
            ElseIf IsMissing(vP4) Then
                lErro = Chama_Tela2(X, sTela, vp1, vp2, vP3)
            ElseIf IsMissing(vP5) Then
                lErro = Chama_Tela2(X, sTela, vp1, vp2, vP3, vP4)
            Else
                lErro = Chama_Tela2(X, sTela, vp1, vp2, vP3, vP4, vP5)
            End If
        
            If lErro <> SUCESSO Then gError 16811
    
        End If
    End If
    
    objForm.Show vbModeless
'    If iTelaNova = 1 Then objForm.SetFocus
    objForm.ZOrder 0
    
    Chama_TelaECF = SUCESSO

    Exit Function

Erro_Chama_TelaECF:

    Chama_TelaECF = gErr

    Select Case gErr

        Case 32218, 32220, 199480, 214924
        
        Case 32219
            Unload objForm
        
        Case 9125
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_NAO_DISPONIVEL", gErr, sTela)

        Case 9129, 9130, 9135, 9207, 9211, 9215, 9219, 7043, 7044, 7053, 14094, 14095, 14096, 14172, 14873

        Case 9223, 9227, 9231, 9238, 9239, 9243, 9244, 14044, 14053, 14065, 14092, 14093, 14097, 14761

        Case 9277, 9281, 9295, 9360, 9367, 9415, 9453, 9454, 9469, 9473, 9505, 9509, 14717, 14718, 14754

        Case 9534, 9552, 9553, 9629, 9633, 9683, 14427, 14428, 14429, 14430, 14431, 14625, 14736, 14753, 14993

        Case 9759, 9789, 9793, 14407, 14408, 14409, 14419, 14426, 14546, 14603, 14607, 14687, 14688, 14872

        Case 10037, 10264, 10904, 14483, 14494, 14509, 14534, 14535, 14542, 14552, 14572, 14573, 14624, 14916

        Case 14084, 14085, 14086, 14087, 14088, 14089, 14090, 14091, 14091, 14148, 14160, 14203, 14204, 14903

        Case 14281, 14282, 14366, 14382, 14383, 14384, 14385, 14386, 14387, 14388, 14395, 14396, 14397, 14889

        Case 14611, 14612, 14613, 14614, 14615, 14616, 14617, 14618, 14619, 14620, 14621, 14622, 14623, 14929

        Case 14797, 14798, 14799, 14800, 14801, 14802, 14803, 14804, 14805, 14806, 14807, 14808, 14809, 14992

        Case 16811, 19769, 19770, 20263, 20264, 20265, 20266, 20307, 20308, 20309, 20324, 20325, 20468, 20643
        
        Case 20725, 20726

        Case 24001, 24002, 24035, 24055, 24059, 24060, 24075, 24103, 24104, 24105, 24115, 24144, 24163, 24164

        Case 24165, 24183, 24195, 24215, 24219, 24220, 24227, 24228, 24229, 24239, 24240, 24241, 24252, 24264

        Case 24275, 24287, 24301, 24307, 24308, 24316, 24329, 24331, 24335, 24336, 24344, 24345, 24346

        Case 24347, 24348, 24349, 24358, 24363, 24367, 24387, 24391, 24395, 24402, 24425, 24445, 24446, 24453

        Case 24465, 24490, 24491, 24498, 24499, 24500, 24514, 24522, 24530, 24550, 24551, 24557, 24585

        Case 24586, 24602, 24611, 24627, 24637, 24642, 24643, 24650, 24662, 24667, 24672, 24676, 24677

        Case 24683, 24688, 24694, 24707, 24719, 24736, 24737, 24750, 24768, 24769, 24784, 24796, 24814, 24830

        Case 24831, 24834, 24842, 24854, 24866, 24908, 24912, 24925, 24937, 24951, 24955, 24959, 24964, 24980

        Case 24981, 24982, 24983, 24984, 24991, 29510, 29522, 29534, 29546, 29565, 29566, 29567, 29568, 29569

        Case 29570, 29594, 29595, 29596, 29597, 29598, 29614, 29628, 29630, 29632, 29636, 29637, 29640, 29661

        Case 29662, 29669, 29670, 29681, 29691, 29703, 29714, 29718, 29730, 29750, 29751, 29773, 29774, 29788

        Case 29816, 29817, 29818, 29837, 29857, 29858, 29876, 29877, 54798

        Case 214925
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BROWSER_SEM_BROWSECAMPO", gErr, sTela)
            Unload objForm

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142276)

    End Select

    Exit Function

End Function

Function Chama_TelaECF_Modal(sTela As String, Optional vp1 As Variant, Optional vp2 As Variant, Optional vP3 As Variant, Optional vP4 As Variant, Optional vP5 As Variant) As Long

Dim X As Object, sProjeto As String, sClasse As String, lErro As Long
Dim Y As Object, objForm As Object, objFormOcx As Object, iTelaNova As Integer
Dim objAux1 As Object, objAux2 As Object, objAux3 As Object, objAux4 As Object, objAux5 As Object
Dim bAcrescentaOcx As Boolean
Dim colBrowseCampo As New Collection
Dim objAux6 As Object


On Error GoTo Erro_Chama_TelaECF_Modal

    lErro_Chama_Tela = -1
    
    Set X = CreateObject(gsNomePrinc & ".ClassTelaPopUp")
        
    lErro = Rotina_ObterObjetoTela_ECF(sTela)
    If lErro <> SUCESSO Then gError 199479
    
    lErro = X.ObtemTela("TelasECF" & "." & sTela, objForm, iTelaNova)
    If lErro <> SUCESSO Then gError 56579
    
    
        Set X = CreateObject(gsNomePrinc & ".ClassTelaPopUp")
        lErro = X.ObtemTela(sProjeto & "." & sTela & IIf(bAcrescentaOcx, "Ocx", ""), objForm, iTelaNova)
        If lErro <> SUCESSO Then Error 56579
    
    
    
    
    'se estiver com algum parametro preenchido ou se estiver fazendo o form load (carga da tela)
    If Not (IsMissing(vp1)) Or iTelaNova = 1 Then
    
        Set objFormOcx = objForm.objFormOcx
        
        If VarType(vp1) = vbObject Then Set objAux1 = vp1
        If VarType(vp2) = vbObject Then Set objAux2 = vp2
        If VarType(vP3) = vbObject Then Set objAux3 = vP3
        If VarType(vP4) = vbObject Then Set objAux4 = vP4
        If VarType(vP5) = vbObject Then Set objAux5 = vP5
        
        
        'se for uma tela de browser
        If InStr(sTela, "ProdutosLista") <> 0 Or InStr(sTela, "ClientesLista") <> 0 Or InStr(sTela, "VendedoresLista") <> 0 Or InStr(sTela, "OrcamentoLista") <> 0 Then
        
            'mario. retirar o codigo abaixo quando tiver colocado todos os browsers no novo formato sem tela
            lErro = CF("BrowseCampo_Le", sTela, colBrowseCampo)
            If lErro <> SUCESSO Then gError 214926
        
            If colBrowseCampo.Count = 0 Then gError 214927
        
            lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5))
        
        Else
        
            If IsMissing(vp1) Then
                lErro = objFormOcx.Trata_Parametros()
            ElseIf IsMissing(vp2) Then
                lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1))
            ElseIf IsMissing(vP3) Then
                lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2))
            ElseIf IsMissing(vP4) Then
                lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3))
            ElseIf IsMissing(vP5) Then
                lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4))
            Else
                lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5))
            End If
                            
                            
            If lErro <> SUCESSO And iTelaNova = 1 Then gError 56580
            If lErro <> SUCESSO Then gError 56581
        
        End If
        
    End If
    
    objForm.Show vbModal

    Chama_TelaECF_Modal = SUCESSO

    Exit Function

Erro_Chama_TelaECF_Modal:

    Chama_TelaECF_Modal = gErr

    Select Case gErr

        Case 56578
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_NAO_DISPONIVEL", gErr, sTela)

        Case 56579, 56581, 199479, 214926
        
        Case 56580
            Unload objForm
        
        Case 214927
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BROWSER_SEM_BROWSECAMPO", gErr, sTela)
            Unload objForm
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142277)

    End Select

    Exit Function

End Function

Function Moedas_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Moedas_Chama_Tela

    Load X.objMoedas
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then Error 45293
    
    X.objMoedas.Show
            
    Moedas_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Moedas_Chama_Tela:

    Moedas_Chama_Tela = Err

    Select Case Err
    
        Case 45293
            Unload X.objMoedas
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142278)
        
    End Select
        
    Exit Function

End Function

Function RelDlgExec_Chama_Tela(X As Object, vp1 As Variant) As Long

Dim lErro As Long
Dim objRelOp As AdmRelOpcoes

On Error GoTo Erro_RelDlgExec_Chama_Tela

    Load X.objRelDlgExec
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then gError 97117
    
    Set objRelOp = vp1
    
    lErro = X.objRelDlgExec.Trata_Parametros(objRelOp)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then gError 97118
    
    If lErro <> SUCESSO Then gError 97119
    
    X.objRelDlgExec.Show vbModal
            
    RelDlgExec_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_RelDlgExec_Chama_Tela:

    RelDlgExec_Chama_Tela = gErr

    Select Case gErr
    
        Case 97117, 97118
            Unload X.objRelDlgExec
                    
        Case 97119
                    
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142279)
        
    End Select
        
    Exit Function

End Function

Private Function Tela_EmUserControl(ByVal sProjeto As String, bAcrescentaOcx As Boolean) As Boolean
'retorna TRUE se a tela está implementada em UserControl e FALSE caso contrario.
'se for TRUE ainda retorna se a string "Ocx" deve ser acrescentada ao nome lógico da tela

    Select Case UCase(sProjeto)
    
        Case "TELASADM", "INSTALACAO"
            Tela_EmUserControl = False
            
        Case Else
            Tela_EmUserControl = True
        
            Select Case UCase(sProjeto)
                Case "TELAS", "TELASCPR", "TELASFAT", "TELASTRB", "RELATORIOSCPR", "RELATORIOSFAT", "RELATORIOSEST", "TELASCOM", "TELASFIS", "RELATORIOSCOM", "RELATORIOSFIS", "TELASUSU", "TELASCPRUSU", "TELASFATUSU", "TELASTRBUSU", "RELATORIOSCPRUSU", "RELATORIOSFATUSU", "RELATORIOSESTUSU", "TELASESTUSU", "TELASCOMUSU", "TELASFISUSU", "RELATORIOSCOMUSU", "RELATORIOSFISUSU", "RELATORIOSFATGR", "RELATORIOSPCP", "RELATORIOSPCPINPAL", "RELATORIOSCONTAB", "TELASSRV"
                    bAcrescentaOcx = True
                Case Else
                    bAcrescentaOcx = False
            End Select

    End Select

End Function

'#####################################################
'Inserido por Wagner
Function PlanilhasSelecionar_Chama_Tela(X As Object, vp1 As Variant) As Long

Dim lErro As Long
Dim sModulo As String

On Error GoTo Erro_PlanilhasSelecionar_Chama_Tela

    Load X.objPlanilhasSelecionar
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then gError 45539
                           
    sModulo = vp1
    
    lErro = X.objPlanilhasSelecionar.Trata_Parametros(sModulo)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then gError 45540
            
    If lErro <> SUCESSO Then gError 45541

    X.objPlanilhasSelecionar.Show vbModal
            
    PlanilhasSelecionar_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_PlanilhasSelecionar_Chama_Tela:

    PlanilhasSelecionar_Chama_Tela = gErr

    Select Case gErr
    
        Case 45539, 45540
            Unload X.objPlanilhasSelecionar
            
        Case 45541
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142280)
        
    End Select
        
    Exit Function

End Function

Function OLAP_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_OLAP_Chama_Tela

    Load X.objOLAP
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then gError 45539
                           
    lErro = X.objOLAP.Trata_Parametros()
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then gError 45540
            
    If lErro <> SUCESSO Then gError 45541

    X.objOLAP.Show vbModal
            
    OLAP_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_OLAP_Chama_Tela:

    OLAP_Chama_Tela = gErr

    Select Case gErr
    
        Case 45539, 45540
            Unload X.objOLAP
            
        Case 45541
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142281)
        
    End Select
        
    Exit Function

End Function

Function Wait_Chama_Tela(X As Object, vp1 As Variant) As Long

Dim lErro As Long
Dim sModulo As String

On Error GoTo Erro_Wait_Chama_Tela

    Load X.objWait
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then gError 45539
                           
    sModulo = vp1
    
    lErro = X.objWait.Trata_Parametros(sModulo)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then gError 45540
            
    If lErro <> SUCESSO Then gError 45541

    X.objWait.Show vbModal
    
    Wait_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Wait_Chama_Tela:

    Wait_Chama_Tela = gErr

    Select Case gErr
    
        Case 45539, 45540
            Unload X.objWait
            
        Case 45541
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142280)
        
    End Select
        
    Exit Function

End Function

Function LocalizarGrid_Chama_Tela(X As Object, vp1 As Variant) As Long

Dim lErro As Long

On Error GoTo Erro_LocalizarGrid_Chama_Tela

    Load X.objLocalizarGrid
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then gError 45539
                           
    lErro = X.objLocalizarGrid.Trata_Parametros(vp1)
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then gError 45540
            
    If lErro <> SUCESSO Then gError 45541

    X.objLocalizarGrid.Show
    
    LocalizarGrid_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_LocalizarGrid_Chama_Tela:

    LocalizarGrid_Chama_Tela = gErr

    Select Case gErr
    
        Case 45539, 45540
            Unload X.objLocalizarGrid
            
        Case 45541
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142280)
        
    End Select
        
    Exit Function

End Function

Function BackupConfig_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_BackupConfig_Chama_Tela

    Load X.objBackupConfig
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then gError 45539
                           
    lErro = X.objBackupConfig.Trata_Parametros()
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then gError 45540
            
    If lErro <> SUCESSO Then gError 45541

    X.objBackupConfig.Show
    
    BackupConfig_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_BackupConfig_Chama_Tela:

    BackupConfig_Chama_Tela = gErr

    Select Case gErr
    
        Case 45539, 45540
            Unload X.objBackupConfig
            
        Case 45541
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142280)
        
    End Select
        
    Exit Function

End Function

Function Logo_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Logo_Chama_Tela

    Load X.objLogo
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then gError 45539
                           
    lErro = X.objLogo.Trata_Parametros()
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then gError 45540
            
    If lErro <> SUCESSO Then gError 45541

    X.objLogo.Show
    
    Logo_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Logo_Chama_Tela:

    Logo_Chama_Tela = gErr

    Select Case gErr
    
        Case 45539, 45540
            Unload X.objLogo
            
        Case 45541
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142280)
        
    End Select
        
    Exit Function

End Function
'########################################################

Function Arquivamento_Chama_Tela(X As Object) As Long

Dim lErro As Long

On Error GoTo Erro_Arquivamento_Chama_Tela

    Load X.objArquivamento
    If lErro_Chama_Tela <> SUCESSO And lErro_Chama_Tela <> -1 Then gError 45539
                           
    lErro = X.objArquivamento.Trata_Parametros()
    If lErro <> SUCESSO And lErro_Chama_Tela = SUCESSO Then gError 45540
            
    If lErro <> SUCESSO Then gError 45541

    X.objArquivamento.Show
    
    Arquivamento_Chama_Tela = SUCESSO
            
    Exit Function
    
Erro_Arquivamento_Chama_Tela:

    Arquivamento_Chama_Tela = gErr

    Select Case gErr
    
        Case 45539, 45540
            Unload X.objArquivamento
            
        Case 45541
                    
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 142280)
        
    End Select
        
    Exit Function

End Function

Function Chama_Tela_Relatorio(sTela As String, Optional vp1 As Variant, Optional vp2 As Variant, Optional vP3 As Variant, Optional vP4 As Variant, Optional vP5 As Variant) As Long

Dim X As Object, sProjeto As String, sClasse As String, lErro As Long
Dim Y As Object
Dim objForm As Object, objFormOcx As Object, iTelaNova As Integer, objAux1 As Object, objAux2 As Object, objAux3 As Object, objAux4 As Object, objAux5 As Object
Dim bAcrescentaOcx As Boolean
Dim colBrowseCampo As New Collection
Dim objControl As Object
Dim iIndice As Integer

On Error GoTo Erro_Chama_Tela_Relatorio

    lErro_Chama_Tela = -1

    sProjeto = String$(NOME_PROJETO + 1, 0)
    sClasse = String$(NOME_CLASSE + 1, 0)

    lErro = Tela_ObterFuncao(sTela, sProjeto, sClasse)
    If (lErro <> AD_BOOL_TRUE) Then gError 9125

    If Tela_EmUserControl(sProjeto, bAcrescentaOcx) Then
            
        Set X = CreateObject(gsNomePrinc & ".ClassTelaMDIChild")
        lErro = X.ObtemTelaMdiChild(sProjeto & "." & sTela & IIf(bAcrescentaOcx, "Ocx", ""), objForm, iTelaNova)
        If lErro <> SUCESSO Then gError 32218
        
        'se estiver com algum parametro preenchido ou se estiver fazendo o form load (carga da tela)
        If Not (IsMissing(vp1)) Or iTelaNova = 1 Then
        
            Set objFormOcx = objForm.objFormOcx
            
            If VarType(vp1) = vbObject Then Set objAux1 = vp1
            If VarType(vp2) = vbObject Then Set objAux2 = vp2
            If VarType(vP3) = vbObject Then Set objAux3 = vP3
            If VarType(vP4) = vbObject Then Set objAux4 = vP4
            If VarType(vP5) = vbObject Then Set objAux5 = vP5
            
            
            'se for uma tela de browser
            If InStr(sTela, "Lista") <> 0 Then
            
                'mario. retirar o codigo abaixo quando tiver colocado todos os browsers no novo formato sem tela
                lErro = CF("BrowseCampo_Le", sTela, colBrowseCampo)
                If lErro <> SUCESSO Then gError 89983
            
                If colBrowseCampo.Count = 0 Then gError 92979
            
'                If IsMissing(vp1) Then
'                    lErro = objFormOcx.Trata_Parametros(sTela)
'                ElseIf IsMissing(vp2) Then
'                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1))
'                ElseIf IsMissing(vP3) Then
'                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2))
'                ElseIf IsMissing(vP4) Then
'                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3))
'                ElseIf IsMissing(vP5) Then
'                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4))
'                Else
'                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5))
'                End If
                            
                lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5))
                            
                            
                            
            Else
            
Label_Chama_Tela_Tradicional:

                If IsMissing(vp1) Then
                    lErro = objFormOcx.Trata_Parametros()
                ElseIf IsMissing(vp2) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1))
                ElseIf IsMissing(vP3) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2))
                ElseIf IsMissing(vP4) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3))
                ElseIf IsMissing(vP5) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4))
                Else
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5))
                End If
            
            End If
            
            If lErro <> SUCESSO And iTelaNova = 1 Then Error 32219
            If lErro <> SUCESSO Then gError 32220
            
        End If
        
        For Each objControl In objFormOcx.Controls
        
            If objControl.Name = "ComboOpcoes" Then
            
                For iIndice = 0 To objControl.ListCount - 1
                    If objControl.List(iIndice) = objAux1.sOpcao Then
                        objControl.ListIndex = iIndice
                        Exit For
                    End If
                Next
                
                Exit For
                
            End If
                
        Next
                
        For Each objControl In objFormOcx.Controls
        
            If objControl.Name = "BotaoExecutar" Then
                objAux1.iDispositivoDeSaida = REL_SAIDA_PREVIA
                objControl.Value = True
                
                Call objFormOcx.Unload(objFormOcx) 'Inserido por Wagner (Libera a tela)
                
                Exit For
                
            End If
                
        Next
                
                
        
'        objForm.Show vbModeless
'            If iTelaNova = 1 Then objForm.SetFocus
'        objForm.ZOrder 0
            
    Else

        Set X = CreateObject(sProjeto & "." & sClasse)
    
        Select Case sTela
    
            Case "frmWizardEmpresa"
                lErro = frmWizardEmpresa_Chama_Tela(X, vp1)
                If lErro <> SUCESSO Then gError 20725
    
            Case "frmWizardFilial"
                lErro = frmWizardFilial_Chama_Tela(X, vp1)
                If lErro <> SUCESSO Then gError 20726
            
            Case "RelSelecionar"
                If IsMissing(vp1) Then
                    lErro = RelSelecionar_Chama_Tela(X)
                ElseIf IsMissing(vp2) Then
                    lErro = RelSelecionar_Chama_Tela(X, vp1)
                Else
                    lErro = RelSelecionar_Chama_Tela(X, vp1, vp2)
                End If
    
                If lErro <> SUCESSO Then gError 7043
    
            Case "RelCadastro"
                If IsMissing(vp1) Then
                    lErro = RelCadastro_Chama_Tela(X)
                Else
                    lErro = RelCadastro_Chama_Tela(X, vp1)
                End If
                If lErro <> SUCESSO Then gError 7044
    
            Case "BrowseConfigura"
                If IsMissing(vp1) Then
                    lErro = BrowseConfigura_Chama_Tela(X)
                Else
                    lErro = BrowseConfigura_Chama_Tela(X, vp1)
                End If
                If lErro <> SUCESSO Then gError 9244
    
            Case Else
                
                If IsMissing(vp1) Then
                    lErro = Chama_Tela2(X, sTela)
                ElseIf IsMissing(vp2) Then
                    lErro = Chama_Tela2(X, sTela, vp1)
                ElseIf IsMissing(vP3) Then
                    lErro = Chama_Tela2(X, sTela, vp1, vp2)
                ElseIf IsMissing(vP4) Then
                    lErro = Chama_Tela2(X, sTela, vp1, vp2, vP3)
                ElseIf IsMissing(vP5) Then
                    lErro = Chama_Tela2(X, sTela, vp1, vp2, vP3, vP4)
                Else
                    lErro = Chama_Tela2(X, sTela, vp1, vp2, vP3, vP4, vP5)
                End If
    
                If lErro <> SUCESSO Then gError 16811
                
        End Select

    End If
    
    Chama_Tela_Relatorio = SUCESSO

    Exit Function

Erro_Chama_Tela_Relatorio:

    Chama_Tela_Relatorio = gErr

    Select Case gErr

        Case 32218, 32220
        
        Case 32219
            Unload objForm
        
        Case 9125
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_NAO_DISPONIVEL", gErr, sTela)

        Case 9129, 9130, 9135, 9207, 9211, 9215, 9219, 7043, 7044, 7053, 14094, 14095, 14096, 14172, 14873

        Case 9223, 9227, 9231, 9238, 9239, 9243, 9244, 14044, 14053, 14065, 14092, 14093, 14097, 14761

        Case 9277, 9281, 9295, 9360, 9367, 9415, 9453, 9454, 9469, 9473, 9505, 9509, 14717, 14718, 14754

        Case 9534, 9552, 9553, 9629, 9633, 9683, 14427, 14428, 14429, 14430, 14431, 14625, 14736, 14753, 14993

        Case 9759, 9789, 9793, 14407, 14408, 14409, 14419, 14426, 14546, 14603, 14607, 14687, 14688, 14872

        Case 10037, 10264, 10904, 14483, 14494, 14509, 14534, 14535, 14542, 14552, 14572, 14573, 14624, 14916

        Case 14084, 14085, 14086, 14087, 14088, 14089, 14090, 14091, 14091, 14148, 14160, 14203, 14204, 14903

        Case 14281, 14282, 14366, 14382, 14383, 14384, 14385, 14386, 14387, 14388, 14395, 14396, 14397, 14889

        Case 14611, 14612, 14613, 14614, 14615, 14616, 14617, 14618, 14619, 14620, 14621, 14622, 14623, 14929

        Case 14797, 14798, 14799, 14800, 14801, 14802, 14803, 14804, 14805, 14806, 14807, 14808, 14809, 14992

        Case 16811, 19769, 19770, 20263, 20264, 20265, 20266, 20307, 20308, 20309, 20324, 20325, 20468, 20643
        
        Case 20725, 20726

        Case 24001, 24002, 24035, 24055, 24059, 24060, 24075, 24103, 24104, 24105, 24115, 24144, 24163, 24164

        Case 24165, 24183, 24195, 24215, 24219, 24220, 24227, 24228, 24229, 24239, 24240, 24241, 24252, 24264

        Case 24275, 24287, 24301, 24307, 24308, 24316, 24329, 24331, 24335, 24336, 24344, 24345, 24346

        Case 24347, 24348, 24349, 24358, 24363, 24367, 24387, 24391, 24395, 24402, 24425, 24445, 24446, 24453

        Case 24465, 24490, 24491, 24498, 24499, 24500, 24514, 24522, 24530, 24550, 24551, 24557, 24585

        Case 24586, 24602, 24611, 24627, 24637, 24642, 24643, 24650, 24662, 24667, 24672, 24676, 24677

        Case 24683, 24688, 24694, 24707, 24719, 24736, 24737, 24750, 24768, 24769, 24784, 24796, 24814, 24830

        Case 24831, 24834, 24842, 24854, 24866, 24908, 24912, 24925, 24937, 24951, 24955, 24959, 24964, 24980

        Case 24981, 24982, 24983, 24984, 24991, 29510, 29522, 29534, 29546, 29565, 29566, 29567, 29568, 29569

        Case 29570, 29594, 29595, 29596, 29597, 29598, 29614, 29628, 29630, 29632, 29636, 29637, 29640, 29661

        Case 29662, 29669, 29670, 29681, 29691, 29703, 29714, 29718, 29730, 29750, 29751, 29773, 29774, 29788

        Case 29816, 29817, 29818, 29837, 29857, 29858, 29876, 29877, 54798

        Case 92979
            lErro = Rotina_Erro(vbOKOnly, "ERRO_BROWSER_SEM_BROWSECAMPO", gErr, sTela)
            Unload objForm

        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 178420)

    End Select

    Exit Function

End Function

'Inserido por Wagner em 24/05/2007
Function Wait(lMilissegundos As Long) As Long
    Call Chama_Tela("Wait", lMilissegundos)
End Function


Function Chama_Tela_Modal_Invisivel(sTela As String, Optional vp1 As Variant, Optional vp2 As Variant, Optional vP3 As Variant, Optional vP4 As Variant, Optional vP5 As Variant) As Long

Dim X As Object, sProjeto As String, sClasse As String, lErro As Long
Dim Y As Object, objForm As Object, objFormOcx As Object, iTelaNova As Integer
Dim objAux1 As Object, objAux2 As Object, objAux3 As Object, objAux4 As Object, objAux5 As Object
Dim bAcrescentaOcx As Boolean, sNomeTela As String
Dim colBrowseCampo As New Collection

On Error GoTo Erro_Chama_Tela_Modal_Invisivel

    lErro_Chama_Tela = -1

    sProjeto = String$(NOME_PROJETO + 1, 0)
    sClasse = String$(NOME_CLASSE + 1, 0)
    
    If sTela = "RomaneioGrade" Then
    
        lErro = CF("RomaneioGrade_Le_Tela", sNomeTela, vp1)
        If lErro <> SUCESSO Then gError 89983
        
        sTela = sNomeTela
    
    End If

    lErro = Tela_ObterFuncao(sTela, sProjeto, sClasse)
    If (lErro <> AD_BOOL_TRUE) Then Error 56578

    If Tela_EmUserControl(sProjeto, bAcrescentaOcx) Then
            
        Set X = CreateObject(gsNomePrinc & ".ClassTelaPopUp")
        lErro = X.ObtemTela(sProjeto & "." & sTela & IIf(bAcrescentaOcx, "Ocx", ""), objForm, iTelaNova)
        If lErro <> SUCESSO Then Error 56579
        
        'se estiver com algum parametro preenchido ou se estiver fazendo o form load (carga da tela)
        If Not (IsMissing(vp1)) Or iTelaNova = 1 Then
        
            Set objFormOcx = objForm.objFormOcx
            
            If VarType(vp1) = vbObject Then Set objAux1 = vp1
            If VarType(vp2) = vbObject Then Set objAux2 = vp2
            If VarType(vP3) = vbObject Then Set objAux3 = vP3
            If VarType(vP4) = vbObject Then Set objAux4 = vP4
            If VarType(vP5) = vbObject Then Set objAux5 = vP5
            
            'se for uma tela de browser
            If InStr(sTela, "Lista") <> 0 Then
            
                'mario. retirar o codigo abaixo quando tiver colocado todos os browsers no novo formato sem tela
                lErro = CF("BrowseCampo_Le", sTela, colBrowseCampo)
                If lErro <> SUCESSO Then gError 89983
            
                If colBrowseCampo.Count = 0 Then gError 92979
            
                If IsMissing(vp1) Then
                    lErro = objFormOcx.Trata_Parametros(sTela)
                ElseIf IsMissing(vp2) Then
                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1))
                ElseIf IsMissing(vP3) Then
                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2))
                ElseIf IsMissing(vP4) Then
                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3))
                ElseIf IsMissing(vP5) Then
                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4))
                Else
                    lErro = objFormOcx.Trata_Parametros(sTela, IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5))
                End If
                            
            Else
            
Label_Chama_Tela_Tradicional:
            
                If IsMissing(vp1) Then
                    lErro = objFormOcx.Trata_Parametros()
                ElseIf IsMissing(vp2) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1))
                ElseIf IsMissing(vP3) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2))
                ElseIf IsMissing(vP4) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3))
                ElseIf IsMissing(vP5) Then
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4))
                Else
                    lErro = objFormOcx.Trata_Parametros(IIf(VarType(vp1) = vbObject, objAux1, vp1), IIf(VarType(vp2) = vbObject, objAux2, vp2), IIf(VarType(vP3) = vbObject, objAux3, vP3), IIf(VarType(vP4) = vbObject, objAux4, vP4), IIf(VarType(vP5) = vbObject, objAux5, vP5))
                End If
                            
            End If
                            
            If lErro <> SUCESSO And iTelaNova = 1 Then Error 56580
            If lErro <> SUCESSO Then Error 56581
            
        End If
        
        'objForm.Show vbModal
    
    End If
    
    Chama_Tela_Modal_Invisivel = SUCESSO

    Exit Function

Erro_Chama_Tela_Modal_Invisivel:

    Chama_Tela_Modal_Invisivel = Err

    Select Case Err

        Case 56578
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_NAO_DISPONIVEL", Err, sTela)

        Case 56579, 56581
        
        Case 56580
            Unload objForm
                
        Case 429
            lErro = Rotina_Erro(vbOKOnly, "ERRO_TELA_NAO_DISPONIVEL", Err, sTela)
        
        Case Else
            lErro = Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", Err, Error$, 142268)

    End Select

    Exit Function

End Function
