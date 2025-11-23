VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ==============================================================================
' INICIALIZAÇÃO DO SISTEMA (WORKBOOK_OPEN)
' Executado automaticamente assim que o arquivo Excel é aberto.
' Responsável por preparar o ambiente, definir variáveis e configurar a interface.
' ==============================================================================

Private Sub Workbook_Open()
    
    ' Declaração de objetos para manipular as abas principais
    Dim ws As Worksheet
    Dim wsa As Worksheet
    Dim wsb As Worksheet
  
    ' Referencia as abas pelo nome para facilitar o uso no código
    Set ws = Sheets("simular")
    Set wsa = Sheets("dashboard")
    Set wsb = Sheets("sobre")

    ' 1. LIMITAÇÃO DA ÁREA DE ROLAGEM (SCROLL AREA)
    ' --------------------------------------------------------------------------
    ' As linhas abaixo definem os limites de navegação do usuário.
    ' impedem que o usuário role a tela para áreas vazias ou de rascunho,
    ' garantindo que ele veja apenas a interface do sistema.
    
    ws.ScrollArea = "A1:T80" ' Limita a aba Simular
    wsa.ScrollArea = "A1:T40" ' Limita o Dashboard
    wsb.ScrollArea = "A1:T40" ' Limita a aba Sobre

    ' 2. AJUSTE AUTOMÁTICO DE ZOOM
    ' --------------------------------------------------------------------------
    ' Chama o módulo 'AjustarZoomPlanilhas' para adaptar o tamanho da interface
    ' à resolução do monitor do usuário assim que o arquivo abre.
    ' (para facilitar a edição/manutenção lembre de comentar o trecho -> "1. LIMITAÇÃO DA ÁREA DE ROLAGEM...").
    
    Call AjustarZoomPlanilhas

End Sub
