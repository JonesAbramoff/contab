VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTRegiaoVendaTRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim gobjCT As Object

Private WithEvents objEventoRegiao As AdmEvento
Attribute objEventoRegiao.VB_VarHelpID = -1

Public Property Get CodRegiaoPai(ByVal objCT As Object) As Object
     Set CodRegiaoPai = objCT.objUserControl.Controls("CodRegiaoPai")
End Property

Public Property Get DescRegiaoPai(ByVal objCT As Object) As Object
     Set DescRegiaoPai = objCT.objUserControl.Controls("DescRegiaoPai")
End Property

Public Function RG_Grava_Registro(ByVal objCT As Object, objRegiaoVenda As ClassRegiaoVenda) As Long

Dim lErro As Long
Dim iIndice As Integer
Dim objRegiaoVendaTRP As New ClassRegiaoVendaTRP

On Error GoTo Erro_RG_Grava_Registro
    
    objRegiaoVendaTRP.iCodRegiaoPai = StrParaInt(CodRegiaoPai(objCT).Text)
    
    If objRegiaoVendaTRP.iCodRegiaoPai = objRegiaoVenda.iCodigo Then gError 197177

    Set objRegiaoVenda.objInfoUsu = objRegiaoVendaTRP

    RG_Grava_Registro = SUCESSO
    
    Exit Function
    
Erro_RG_Grava_Registro:
    
    RG_Grava_Registro = gErr
    
    Select Case gErr
    
        Case 197177
            Call Rotina_Erro(vbOKOnly, "ERRO_REGIAOPAI_IGUAL_REGIAO", gErr)
        
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197178)

    End Select

    Exit Function

End Function

Public Function RG_Traz_Regiao_Tela(ByVal objCT As Object, objRegiaoVenda As ClassRegiaoVenda) As Long

Dim lErro As Long
Dim objRegiaoVendaTRP As ClassRegiaoVendaTRP
Dim objRegiaoVenda1 As New ClassRegiaoVenda

On Error GoTo Erro_RG_Traz_Regiao_Tela

    CodRegiaoPai(objCT).PromptInclude = False
    CodRegiaoPai(objCT).Text = ""
    CodRegiaoPai(objCT).PromptInclude = True
    
    DescRegiaoPai(objCT).Caption = ""

    lErro = CF("RegiaoVenda_Le_Customizado", objRegiaoVenda)
    If lErro <> SUCESSO Then gError 197184
    
    If Not (objRegiaoVenda.objInfoUsu Is Nothing) Then
    
        Set objRegiaoVendaTRP = objRegiaoVenda.objInfoUsu
        
        If objRegiaoVendaTRP.iCodRegiaoPai <> 0 Then
        
            'Coloca Tipo no Text
            CodRegiaoPai(objCT).PromptInclude = False
            CodRegiaoPai(objCT).Text = CStr(objRegiaoVendaTRP.iCodRegiaoPai)
            CodRegiaoPai(objCT).PromptInclude = True
    
            objRegiaoVenda1.iCodigo = objRegiaoVendaTRP.iCodRegiaoPai
    
            'Le a Região de Venda para Verificar se esta no BD
            lErro = CF("RegiaoVenda_Le", objRegiaoVenda1)
            If lErro <> SUCESSO And lErro <> 16137 Then gError 197185
            
            DescRegiaoPai(objCT).Caption = objRegiaoVenda1.sDescricao
            
        End If
        
    End If
    
    RG_Traz_Regiao_Tela = SUCESSO
     
    Exit Function
    
Erro_RG_Traz_Regiao_Tela:

    RG_Traz_Regiao_Tela = gErr
     
    Select Case gErr
    
        Case 197184, 197185
          
        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197186)
     
    End Select
     
    Exit Function

End Function

Public Function RG_Limpa_Tela_RegiaoVenda(ByVal objCT As Object) As Long
    
    DescRegiaoPai(objCT).Caption = ""
    
    RG_Limpa_Tela_RegiaoVenda = SUCESSO
    
End Function



Public Sub UserControl_KeyDown(ByVal objCT As Object, KeyCode As Integer, Shift As Integer)
    
    If KeyCode = KEYCODE_BROWSER Then
        
        If objCT.objUserControl.ActiveControl Is CodRegiaoPai(objCT) Then
            Call LabelRegiaoPai_Click(objCT)
        End If
    
    End If

End Sub

Public Sub CodRegiaoPai_Change(ByVal objCT As Object)
    objCT.iAlterado = REGISTRO_ALTERADO
End Sub

Public Sub CodRegiaoPai_GotFocus(ByVal objCT As Object)
   
    Call MaskEdBox_TrataGotFocus(CodRegiaoPai(objCT), objCT.iAlterado)

End Sub

Public Sub CodRegiaoPai_Validate(ByVal objCT As Object, Cancel As Boolean)

Dim lErro As Long
Dim objRegiaoVenda1 As New ClassRegiaoVenda

On Error GoTo Erro_CodRegiaoPai_Validate
    
    'verifica se codigo está preenchido
    If Len(Trim(CodRegiaoPai(objCT).Text)) > 0 Then
    
        'verifica se codigo é numérico
        If Not IsNumeric(CodRegiaoPai(objCT).Text) Then gError 197173
        
        'verifica se codigo é menor que um
        If CInt(CodRegiaoPai(objCT).Text) < 1 Then gError 197174
        
        objRegiaoVenda1.iCodigo = StrParaInt(CodRegiaoPai(objCT).Text)

        'Le a Região de Venda para Verificar se esta no BD
        lErro = CF("RegiaoVenda_Le", objRegiaoVenda1)
        If lErro <> SUCESSO And lErro <> 16137 Then gError 197197
        
        DescRegiaoPai(objCT).Caption = objRegiaoVenda1.sDescricao
        
    End If
        
    Exit Sub
    
Erro_CodRegiaoPai_Validate:

    Cancel = True

    Select Case gErr

        Case 197173
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_NAO_NUMERICO", gErr, CodRegiaoPai(objCT).Text)
    
        Case 197174
            Call Rotina_Erro(vbOKOnly, "ERRO_VALOR_MENOR_QUE_UM", gErr, CodRegiaoPai(objCT).Text)

        Case 197197

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error, 197175)
    
    End Select

    Exit Sub
    
End Sub

Public Sub LabelRegiaoPai_Click(ByVal objCT As Object)

Dim objRegiao As New ClassRegiaoVenda
Dim colSelecao As New Collection
    
    'Se o Vendedor estiver preenchido move seu codigo para objVendedor
    If Len(Trim(CodRegiaoPai(objCT).Text)) > 0 Then objRegiao.iCodigo = StrParaInt(CodRegiaoPai(objCT).Text)
    
    Set gobjCT = objCT
    
    Set objEventoRegiao = New AdmEvento
    
    'Chama a tela que lista os vendedores
    Call Chama_Tela("RegiaoVendaLista", colSelecao, objRegiao, objEventoRegiao)

End Sub

Private Sub objEventoRegiao_evSelecao(obj1 As Object)

Dim objRegiao As ClassRegiaoVenda
Dim iIndice As Integer

On Error GoTo Erro_objEventoRegiao_evSelecao

    Set objRegiao = obj1

    CodRegiaoPai(gobjCT).PromptInclude = False
    CodRegiaoPai(gobjCT).Text = objRegiao.iCodigo
    CodRegiaoPai(gobjCT).PromptInclude = True
    
    DescRegiaoPai(gobjCT).Caption = objRegiao.sDescricao

    gobjCT.objUserControl.Show

    Exit Sub

Erro_objEventoRegiao_evSelecao:

    Select Case gErr

        Case Else
            Call Rotina_Erro(vbOKOnly, "ERRO_FORNECIDO_PELO_VB", gErr, Error$, 197176)

    End Select

    Exit Sub

End Sub


