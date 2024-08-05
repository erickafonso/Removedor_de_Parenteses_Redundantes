Attribute VB_Name = "M�dulo1"
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
        MsgBox "Por favor, selecione um intervalo de c�lulas antes de executar a macro."
        Exit Sub
    End If
    
    ' Define o intervalo da sele��o
    Set rng = Selection
    
    ' Itera sobre cada c�lula n�o vazia na coluna A
    For Each cell In rng
        If IsEmpty(cell.Value) Then Exit For
        
        ' Obt�m o texto da c�lula
        textoOriginal = cell.Value
        
        ' Inicializa a posi��o atual e a cole��o de textos encontrados
        posAtual = 1
        duplicados = False
        Set textosEncontrados = New Collection
        
        ' Procura por m�ltiplos pares de par�nteses
        Do
            ' Encontra o pr�ximo par�ntese de abertura e fechamento
            pos1 = InStr(posAtual, textoOriginal, "(")
            pos2 = InStr(pos1 + 1, textoOriginal, ")")
            
            ' Se ambos os par�nteses forem encontrados
            If pos1 > 0 And pos2 > pos1 Then
                ' Extrai o texto entre os par�nteses
                textoEntreParenteses = Mid(textoOriginal, pos1 + 1, pos2 - pos1 - 1)
                
                ' Verifica se o texto j� foi encontrado
                On Error Resume Next
                textosEncontrados.Add textoEntreParenteses, CStr(textoEntreParenteses)
                If Err.Number <> 0 Then
                    ' Se houve erro ao adicionar, significa que o texto j� existe
                    textoDuplicado = textoEntreParenteses
                    duplicados = True
                    Exit Do
                End If
                On Error GoTo 0
                
                ' Atualiza a posi��o atual para procurar o pr�ximo par de par�nteses
                posAtual = pos2 + 1
            Else
                ' Se n�o houver mais par�nteses, encerra o loop
                Exit Do
            End If
        Loop
        
        ' Se textos duplicados foram encontrados, modifica o texto
        If duplicados Then
            ' Inicializa vari�veis
            textoForaParenteses = ""
            textoAntes = textoOriginal
            textoDepois = textoOriginal
            
            ' Procura por par�nteses e atualiza o texto fora dos par�nteses
            Do
                ' Encontra o pr�ximo par�ntese de abertura e fechamento
                pos1 = InStr(textoAntes, "(")
                pos2 = InStr(pos1 + 1, textoAntes, ")")
                
                ' Se ambos os par�nteses forem encontrados
                If pos1 > 0 And pos2 > pos1 Then
                    ' Adiciona o texto antes do par�ntese atual
                    textoForaParenteses = textoForaParenteses & Mid(textoAntes, 1, pos1 - 1)
                    
                    ' Atualiza o texto que resta depois do par�ntese atual
                    textoDepois = Mid(textoAntes, pos2 + 1)
                    
                    ' Atualiza o texto antes do pr�ximo par�ntese
                    textoAntes = textoDepois
                Else
                    ' Se n�o houver mais par�nteses, adiciona o texto restante
                    textoForaParenteses = textoForaParenteses & textoAntes
                    Exit Do
                End If
            Loop
            
            ' Imprime o texto fora dos par�nteses
            novoTexto = textoForaParenteses & "(" & textoEntreParenteses & ")"
            cell.Value = novoTexto
        End If
    Next cell
    
    MsgBox "Processamento conclu�do."
End Sub

