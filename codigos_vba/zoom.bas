Attribute VB_Name = "zoom"
' ==============================================================================
' MÓDULO DE INTERFACE E ZOOM
' Ajusta automaticamente o nível de zoom das planilhas para garantir que
' o conteúdo principal caiba perfeitamente na tela do usuário, independente
' da resolução do monitor.
' ==============================================================================

Sub AjustarZoomPlanilhas()
    
    ' Declaração de variáveis para manipulação de abas e intervalos
    Dim planilhas As Variant
    Dim ws As Worksheet
    Dim intervalo As Range
    
    ' 1. DEFINIÇÃO DOS GRUPOS DE VISUALIZAÇÃO
    ' --------------------------------------------------------------------------

    planilhas = Array("sobre", "dashboard", "simular")
    simuladores = Array("simular")
    
    ' 2. AJUSTE NO GRUPO
    ' --------------------------------------------------------------------------
    For Each planilhas In simuladores
        On Error Resume Next ' Evita travar o código caso a aba não seja encontrada
        Set ws = ThisWorkbook.Sheets(planilhas)
        
        If Not ws Is Nothing Then
            ' Define a área exata que deve ficar visível (Da célula A1 até U25)
            Set intervalo = ws.Range("A1:U25")
            
            ' Processo de Zoom:
            ws.Activate          ' 1. Entra na planilha
            intervalo.Select     ' 2. Seleciona a área definida
            ActiveWindow.zoom = True ' 3. zoom para encaixar a seleção na tela
        End If
    Next planilhas
    
End Sub
