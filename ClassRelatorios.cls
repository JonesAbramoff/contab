VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassRelatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function Rel_Razao(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
    Rel_Razao = Chama_Tela("RelOpRazao", objRelatorio, objRelOpcoes)
    
End Function

Function Rel_BalPatr(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
    Rel_BalPatr = Chama_Tela("RelOpBalPatr", objRelatorio, objRelOpcoes)
    
End Function

Function Rel_BalPatrPer(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
   Rel_BalPatrPer = Chama_Tela("RelOpBalPatrPer", objRelatorio, objRelOpcoes)
    
End Function

Function Rel_BalVerif(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
    Rel_BalVerif = Chama_Tela("RelOpBalVerif", objRelatorio, objRelOpcoes)
    
End Function

Function Rel_DemResEx(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
    Rel_DemResEx = Chama_Tela("RelOpDemResEx", objRelatorio, objRelOpcoes)
    
End Function

Function Rel_DemResPer(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
    Rel_DemResPer = Chama_Tela("RelOpDemResPer", objRelatorio, objRelOpcoes)
    
End Function

Function Rel_DespPerCcl(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
    Rel_DespPerCcl = Chama_Tela("RelOpDespPerCcl", objRelatorio, objRelOpcoes)
    
End Function

Function Rel_Diario(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
    Rel_Diario = Chama_Tela("RelOpDiario", objRelatorio, objRelOpcoes)
    
End Function

Function Rel_LancCcl(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
    Rel_LancCcl = Chama_Tela("RelOpLancCcl", objRelatorio, objRelOpcoes)
    
End Function

Function Rel_LancData(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
    Rel_LancData = Chama_Tela("RelOpLancData", objRelatorio, objRelOpcoes)
    
End Function

Function Rel_LancLote(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
    Rel_LancLote = Chama_Tela("RelOpLancLote", objRelatorio, objRelOpcoes)
    
End Function

Function Rel_Lote(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
    Rel_Lote = Chama_Tela("RelOpLote", objRelatorio, objRelOpcoes)
    
End Function

Function Rel_OrcReal(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
    Rel_OrcReal = Chama_Tela("RelOpOrcReal", objRelatorio, objRelOpcoes)
    
End Function

Function Rel_OrcRealCcl(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
    Rel_OrcRealCcl = Chama_Tela("RelOpOrcRealCcl", objRelatorio, objRelOpcoes)
    
End Function

Function Rel_PlanoContas(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
    Rel_PlanoContas = Chama_Tela("RelOpPlanoContas", objRelatorio, objRelOpcoes)
    
End Function

Function Rel_PlanSaldos(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
    Rel_PlanSaldos = Chama_Tela("RelOpPlanSaldos", objRelatorio, objRelOpcoes)
    
End Function

Function Rel_MutacaoPL(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
    Rel_MutacaoPL = Chama_Tela("RelOpMutacaoPL", objRelatorio, objRelOpcoes)
    
End Function

Function Rel_BalPatrCom(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
    Rel_BalPatrCom = Chama_Tela("RelOpBalPatrCom", objRelatorio, objRelOpcoes)

End Function

Function Rel_LotePend(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
    Rel_LotePend = Chama_Tela("RelOpLotePend", objRelatorio, objRelOpcoes)
    
End Function

Function Rel_LanPend(objRelatorio As AdmRelatorio, objRelOpcoes As AdmRelOpcoes) As Long
    
    Rel_LanPend = Chama_Tela("RelOpLanPend", objRelatorio, objRelOpcoes)
    
End Function


