Attribute VB_Name = "Módulo1"
Sub CompararTextosEntreParenteses()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim textoOriginal As String
    Dim pos1 As Long, pos2 As Long
    Dim posAtual As Long
    Dim textoEntreParenteses As String
    Dim textosEncontrados As Collection
    Dim textoDuplicado As String
    Dim duplicados As Boolean
    Dim i As Integer
    Dim textoForaParenteses As String
    Dim textoAntes As String
    Dim textoDepois As String
    Dim novoTexto As String
    
    ' Define a planilha
    Set ws = ThisWorkbook.Sheets("Planilha1") ' Altere "Planilha1" para o nome real da sua planilha
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Por favor, selecione um intervalo de células antes de executar a macro."
        Exit Sub
    End If
    
    ' Define o intervalo da seleção
    Set rng = Selection
    
    ' Itera sobre cada célula não vazia na coluna A
    For Each cell In rng
        If IsEmpty(cell.Value) Then Exit For
        
        ' Obtém o texto da célula
        textoOriginal = cell.Value
        
        ' Inicializa a posição atual e a coleção de textos encontrados
        posAtual = 1
        duplicados = False
        Set textosEncontrados = New Collection
        
        ' Procura por múltiplos pares de parênteses
        Do
            ' Encontra o próximo parêntese de abertura e fechamento
            pos1 = InStr(posAtual, textoOriginal, "(")
            pos2 = InStr(pos1 + 1, textoOriginal, ")")
            
            ' Se ambos os parênteses forem encontrados
            If pos1 > 0 And pos2 > pos1 Then
                ' Extrai o texto entre os parênteses
                textoEntreParenteses = Mid(textoOriginal, pos1 + 1, pos2 - pos1 - 1)
                
                ' Verifica se o texto já foi encontrado
                On Error Resume Next
                textosEncontrados.Add textoEntreParenteses, CStr(textoEntreParenteses)
                If Err.Number <> 0 Then
                    ' Se houve erro ao adicionar, significa que o texto já existe
                    textoDuplicado = textoEntreParenteses
                    duplicados = True
                    Exit Do
                End If
                On Error GoTo 0
                
                ' Atualiza a posição atual para procurar o próximo par de parênteses
                posAtual = pos2 + 1
            Else
                ' Se não houver mais parênteses, encerra o loop
                Exit Do
            End If
        Loop
        
        ' Se textos duplicados foram encontrados, modifica o texto
        If duplicados Then
            ' Inicializa variáveis
            textoForaParenteses = ""
            textoAntes = textoOriginal
            textoDepois = textoOriginal
            
            ' Procura por parênteses e atualiza o texto fora dos parênteses
            Do
                ' Encontra o próximo parêntese de abertura e fechamento
                pos1 = InStr(textoAntes, "(")
                pos2 = InStr(pos1 + 1, textoAntes, ")")
                
                ' Se ambos os parênteses forem encontrados
                If pos1 > 0 And pos2 > pos1 Then
                    ' Adiciona o texto antes do parêntese atual
                    textoForaParenteses = textoForaParenteses & Mid(textoAntes, 1, pos1 - 1)
                    
                    ' Atualiza o texto que resta depois do parêntese atual
                    textoDepois = Mid(textoAntes, pos2 + 1)
                    
                    ' Atualiza o texto antes do próximo parêntese
                    textoAntes = textoDepois
                Else
                    ' Se não houver mais parênteses, adiciona o texto restante
                    textoForaParenteses = textoForaParenteses & textoAntes
                    Exit Do
                End If
            Loop
            
            ' Imprime o texto fora dos parênteses
            novoTexto = textoForaParenteses & "(" & textoEntreParenteses & ")"
            cell.Value = novoTexto
        End If
    Next cell
    
    MsgBox "Processamento concluído."
End Sub

