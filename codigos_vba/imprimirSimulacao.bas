Attribute VB_Name = "imprimirSimulacao"
' ==============================================================================
' MÓDULO DE IMPRESSÃO E EXPORTAÇÃO
' Gerencia a visualização de impressão selecionando a aba correta com base
' na quantidade de grupos simulados e no estilo desejado.
' ==============================================================================

Sub imprimir()
    
    ' Declarando a variável
    Dim tpImpressao As Integer
    
    
    ' 2. PROCESSAMENTO DO MODO IMPRESSÃO (OPÇÃO 1)
    ' --------------------------------------------------------------------------
    ' Se escolhido 1, imprime no preto e branco.
    ' A variável 'qtdGrupo' é Pública e vem do módulo Principal (simular).
    tpImpressao = CInt(InputBox("Escolha o tipo de gráfico desejado(Apenas números):" & vbCrLf & "" _
                & vbCrLf & " 1 - Para Impressão(cores escuras)" & vbCrLf & _
                " 2 - Para PDF(colorido)", "Tipo de gráfico"))
                
    If tpImpressao = 1 Then
        If qtdGrupo = 1 Then
            Sheets("impressaoFolha1G").PrintPreview
        ElseIf qtdGrupo = 2 Then
            Sheets("impressaoFolha2G").PrintPreview
        Else
            Sheets("impressaoFolha3G").PrintPreview
        End If
    End If
    
    ' 3. PROCESSAMENTO DO MODO PDF/COLORIDO (OPÇÃO 2)
    ' --------------------------------------------------------------------------
    ' Se escolhido 2, seleciona as abas com gráfico de barras colorido.
        If tpImpressao = 2 Then
            If qtdGrupo = 1 Then
                Sheets("impressaoFolha1GColor").PrintPreview
            ElseIf qtdGrupo = 2 Then
                Sheets("impressaoFolha2GColor").PrintPreview
            Else
                Sheets("impressaoFolha3GColor").PrintPreview
            End If
        End If
     
    ' 4. TRATAMENTO DE ERROS
    ' --------------------------------------------------------------------------
    ' Verifica se o usuário digitou um número fora das opções permitidas
    If tpImpressao <> 1 And tpImpressao <> 2 Then
        MsgBox "Entrada inválida. Por favor, digite 1 ou 2"
        Exit Sub
    End If
    
End Sub
    

