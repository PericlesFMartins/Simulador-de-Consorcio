Attribute VB_Name = "simular"
' ==============================================================================
' MÓDULO PRINCIPAL: SIMULAÇÃO DE CONSÓRCIO
' Responsável por coletar inputs do usuário, buscar dados na base e gerar o relatório.
' ==============================================================================

Public qtdGrupo As Integer ' Variável pública para ser usada em outros módulos (ex: impressão)

Sub Simulacao()
    
    ' 1. CONFIGURAÇÕES INICIAIS
    ' --------------------------------------------------------------------------
    Application.ScreenUpdating = False  ' Otimização: Congela a tela para o código rodar mais rápido sem piscar
    
    ' Declaração de variáveis locais
    Dim percentuaisArray() As String
    Dim grupo As String, percentual As String, grupoPorcentagem As String
    Dim contemplados As Integer
    Dim fundoReserva As Double
    Dim taxaAdm_A As Double
    
    ' Definição das planilhas de trabalho (Origem dos dados e Destino)
    Set wsdados = ThisWorkbook.Sheets("BaseDados")
    Set wsSimulador = ThisWorkbook.Sheets("simular")
    
    ' 2. COLETA DE DADOS (INPUTS DO USUÁRIO)
    ' --------------------------------------------------------------------------
    qtdGrupo = InputBox("Vão ser informados quantos grupos?", "Quantidade de grupos")
    
    ' Inicializa as variáveis acumuladoras de texto
    grupos = ""
    percentuais = ""
    
    ' Loop para solicitar o número e a % de cada grupo sequencialmente
    If qtdGrupo >= 1 Then
        For inserido = 0 To qtdGrupo - 1
            ' Pergunta o número do grupo
            grupo = InputBox("Informe o número do " & (inserido + 1) & "°" & " Grupo", "Grupos")
            
            If grupo <> "" Then
                ' Pergunta a porcentagem do lance (ex: 50)
                percentual = InputBox("Digite a Faixa de valor da carta " & grupo & " (apenas o número, ex: 50):", "Porcentagem da carta")
                
                If percentual <> "" Then
                    ' Constrói uma string separada por vírgulas (ex: "1005,1006")
                    If grupos <> "" Then
                        grupos = grupos & "," & grupo
                        percentuais = percentuais & "," & percentual
                    Else
                        grupos = grupo
                        percentuais = percentual
                    End If
                End If
            End If
        Next
    End If
    
    ' Tratamento de erro: Se não houver grupos válidos, encerra a rotina
    If grupos = "" Then
        MsgBox "Entradas inválidas. Por favor, tente novamente.", vbExclamation, "Erro"
        Exit Sub
    End If
    
    ' 3. TRATAMENTO DOS DADOS DE ENTRADA
    ' --------------------------------------------------------------------------
    ' Cria arrays separando os dados pelas vírgulas
    percentuaisArray = Split(percentuais, ",")
    grupoArray = Split(grupos, ",")
    
    ' Converte as porcentagens de inteiro (50) para decimal (0.5) para cálculo
    For i = LBound(percentuaisArray) To UBound(percentuaisArray)
        percentuaisArray(i) = percentuaisArray(i) / 100
    Next
    
    ' 4. PREPARAÇÃO DA PLANILHA DE SIMULAÇÃO
    ' --------------------------------------------------------------------------
    wsSimulador.Unprotect "123" ' Retira a proteção para permitir a escrita
    
    ' Limpa os dados da simulação anterior (mantendo o cabeçalho principal)
    wsSimulador.Range("C7:T51").ClearContents
    
    ' Recria os cabeçalhos das colunas (Garante que o texto esteja correto)
    wsSimulador.Cells(7, 3).Value = "Grupo "
    wsSimulador.Cells(7, 4).Value = "% da Carta  "
    wsSimulador.Cells(7, 5).Value = "Valor do Bem    "
    wsSimulador.Cells(7, 6).Value = "Tx adm a.m   "
    wsSimulador.Cells(7, 7).Value = "Total da Tx R$  "
    wsSimulador.Cells(7, 8).Value = "Divida Total R$  "
    wsSimulador.Cells(7, 9).Value = "Até 6º pcl  "
    wsSimulador.Cells(7, 10).Value = "Demais pcl"
    wsSimulador.Cells(7, 11).Value = "Qtda de pcl(Lance) "
    wsSimulador.Cells(7, 12).Value = "Lance do bolso R$ "
    wsSimulador.Cells(7, 13).Value = "Prazo Restante"
    wsSimulador.Cells(7, 14).Value = "Embutido"
    wsSimulador.Cells(7, 15).Value = "Pcl média após o lance "
    wsSimulador.Cells(7, 16).Value = "Média de contemplados mês "
    wsSimulador.Cells(7, 17).Value = "% de lance do bolso  "
    wsSimulador.Cells(7, 18).Value = "Valor crédito após o lance"
    wsSimulador.Cells(7, 20).Value = "Taxa ADM"
    
    ' Identifica até onde vai a base de dados para limitar a busca
    ultimaLinha = wsdados.Cells(wsdados.Rows.Count, 2).End(xlUp).Row
    
    ' Zera variáveis da tabela total
    somaValorBens = 0: somaTotalTx = 0: somaDividaTotal = 0
    somaAtepcl = 0: somaDemaispcl = 0: somaLance = 0
    somaEmbutido = 0: somaPclmediapos = 0: somaCreditoPosLance = 0
    
    ' Configuração visual (Formatação da tabela de resultados)
    With wsSimulador.Range("C7:Q" & ultimaLinha)
        .Font.Name = "Calibri"
        .Font.Size = 17
    End With
    wsSimulador.Columns("C:Q").AutoFit
    
    ' 5. BUSCA E PROCESSAMENTO (CORE DO SISTEMA)
    ' --------------------------------------------------------------------------
    ' Loop principal: Para cada grupo solicitado pelo usuário
    For j = LBound(grupoArray) To UBound(grupoArray)
        grupoAtual = Trim(grupoArray(j))
        lancePercentual = percentuaisArray(j)
        encontrado = False
        
        'Varre a base de dados linha por linha procurando correspondência
        For i = 2 To ultimaLinha
            grupoPorcentagem = wsdados.Cells(i, 3).Value ' Coluna chave na base (Ex: "1005... 50%")
            
            ' Verifica se a linha contém TANTO o número do grupo QUANTO a porcentagem informada
            If InStr(grupoPorcentagem, grupoAtual) > 0 And InStr(grupoPorcentagem, CStr(lancePercentual * 100) & "%") > 0 Then
                
                ' --- Captura dos dados da Base ---
                valorTotal = wsdados.Cells(i, 7).Value
                primeiraParcela = wsdados.Cells(i, 10).Value
                demaisParcelas = wsdados.Cells(i, 11).Value
                valorLance = wsdados.Cells(i, 23).Value
                taxaAdm_A = wsdados.Cells(i, 16).Value
                prazoRestante = wsdados.Cells(i, 6).Value
                embutido = wsdados.Cells(i, 21).Value
                parcelaMediaPos_Lance = wsdados.Cells(i, 25).Value
                TotalTx = wsdados.Cells(i, 26).Value
                quantidadeLance = wsdados.Cells(i, 13).Value
                dividaTotal = wsdados.Cells(i, 7).Value
                contemplados = wsdados.Cells(i, 27).Value
                vagas = wsdados.Cells(i, 1).Value
                porcLance = wsdados.Cells(i, 22).Value
                porcTaxaADM = wsdados.Cells(i, 16).Value
                
                ' --- Escrita dos dados na Planilha Simulador ---
                With wsSimulador
                    .Cells(j + 8, 3).Value = grupoAtual
                    .Cells(j + 8, 4).Value = lancePercentual * 100 & "%"
                    .Cells(j + 8, 5).Value = valorTotal
                    .Cells(j + 8, 6).Value = Round(taxaAdm_A / prazoRestante, 2) & "%" ' Calcula taxa mensal média
                    .Cells(j + 8, 7).Value = TotalTx
                    .Cells(j + 8, 8).Value = dividaTotal + (((taxaAdm_A + fundoReserva) / 100) * valorTotal)
                    .Cells(j + 8, 9).Value = primeiraParcela
                    .Cells(j + 8, 10).Value = demaisParcelas
                    .Cells(j + 8, 11).Value = quantidadeLance
                    .Cells(j + 8, 11).NumberFormat = "0"
                    .Cells(j + 8, 12).Value = Round(valorLance, 1)
                    .Cells(j + 8, 12).NumberFormat = "$ #,##0.00_);Red"
                    .Cells(j + 8, 13).Value = prazoRestante
                    .Cells(j + 8, 14).Value = embutido
                    .Cells(j + 8, 15).Value = parcelaMediaPos_Lance
                    .Cells(j + 8, 16).Value = contemplados
                    .Cells(j + 8, 17).Value = porcLance
                    .Cells(j + 8, 17).NumberFormat = "0%"
                    .Cells(j + 8, 18).Value = valorTotal - embutido
                    .Cells(j + 8, 18).NumberFormat = "$ #,##0.00_);Red"
                    .Cells(j + 8, 20).Value = porcTaxaADM & "%"
                    
                    ' Formatação específica para a linha inserida
                    .Range(.Cells(j + 8, 3), .Cells(j + 8, 18)).Font.Name = "Calibri"
                    .Range(.Cells(j + 8, 3), .Cells(j + 8, 18)).Font.Size = 16
                    .Columns("H:H").NumberFormat = "$ #,##0.00_);Red"
                    
                ' --- Acumulação para Totais ---
                somaValorBens = somaValorBens + valorTotal
                somaTotalTx = somaTotalTx + TotalTx
                somaDividaTotal = somaDividaTotal + (dividaTotal + (((taxaAdm_A + fundoReserva) / 100) * valorTotal))
                somaAtepcl = somaAtepcl + primeiraParcela
                somaDemaispcl = somaDemaispcl + demaisParcelas
                somaLance = somaLance + valorLance
                somaEmbutido = somaEmbutido + embutido
                somaPclmediapos = somaPclmediapos + parcelaMediaPos_Lance
                somaCreditoPosLance = (somaCreditoPosLance + valorTotal) - embutido
                
                encontrado = True
                Exit For ' Sai do loop da base de dados pois já achou o grupo
            End If
        Next i
        
        ' Feedback se o grupo não for encontrado
        If Not encontrado Then
            MsgBox "Grupo " & grupoAtual & " com a porcentagem de " & (lancePercentual * 100) & "% não encontrado. Verifique os dados.", vbExclamation, "Erro"
        End If
    Next j
    
    ' 6. EXIBIÇÃO DOS TOTAIS (SE HOUVER MAIS DE 1 GRUPO)
    ' --------------------------------------------------------------------------
    If qtdGrupo > 1 Then
        ' Preenche as células de resumo no final da lista de grupos
        wsSimulador.Cells(30, 12).Value = "Soma Total % do Bem:"
        wsSimulador.Cells(30, 14).Value = somaValorBens
        wsSimulador.Cells(31, 12).Value = "Soma Total de Tx:"
        wsSimulador.Cells(31, 14).Value = somaTotalTx
        wsSimulador.Cells(32, 12).Value = "Soma Divida Total:"
        wsSimulador.Cells(32, 14).Value = somaDividaTotal
        wsSimulador.Cells(33, 12).Value = "Soma Até 6° pcl:"
        wsSimulador.Cells(33, 14).Value = somaAtepcl
        wsSimulador.Cells(34, 12).Value = "Soma Demais pcl:"
        wsSimulador.Cells(34, 14).Value = somaDemaispcl
        wsSimulador.Cells(35, 12).Value = "Soma Lance do bolso:"
        wsSimulador.Cells(35, 14).Value = somaLance
        wsSimulador.Cells(36, 12).Value = "Soma Embutido:"
        wsSimulador.Cells(36, 14).Value = somaEmbutido
        wsSimulador.Cells(37, 12).Value = "Soma Pcl média após o lance:"
        wsSimulador.Cells(37, 14).Value = somaPclmediapos
        wsSimulador.Cells(38, 12).Value = "Soma Crédito após o lance"
        wsSimulador.Cells(38, 14).Value = somaCreditoPosLance
        
        ' Aplica formatação monetária aos totais
        For linha = 29 To 37
            wsSimulador.Cells(UBound(grupoArray) + linha, 14).NumberFormat = "$ #,##0.00_);Red"
        Next linha
    End If
    
    ' 7. FINALIZAÇÃO E SEGURANÇA
    ' --------------------------------------------------------------------------
    ' Reativa a proteção da planilha, mas libera colunas específicas para uso
    wsSimulador.Unprotect "123"
    wsSimulador.Columns("B:T").Locked = False
    wsSimulador.Protect "123"
    
    MsgBox "Simulação concluída!", vbInformation, "Encerrado"
    
    Application.ScreenUpdating = True  ' Reativa a atualização de tela
    
End Sub
