# README - Macro VBA: Removedor de parenteses redundantes tabajara
## Descrição

Este código VBA é uma macro desenvolvida para ser utilizada no Microsoft Excel. A macro realiza a análise de células selecionadas, busca por pares de parênteses e verifica se há textos duplicados dentro desses pares. Se encontrar textos duplicados, a macro ajusta o conteúdo da célula para remover esses duplicados.

## Funcionalidades

1. **Seleção de Intervalo**: A macro opera no intervalo de células atualmente selecionado pelo usuário na planilha ativa, tanto no mouse quanto por fórmula.

2. **Verificação de Textos Duplicados**:
   - A macro procura por pares de parênteses em cada célula do intervalo selecionado.
   - Dentro desses pares de parênteses, verifica se há textos duplicados.
   - Se um texto duplicado for encontrado, a macro realiza uma modificação no conteúdo da célula.

3. **Modificação do Texto**:
   - Quando textos duplicados são encontrados, o texto dentro dos parênteses é ajustado para remover duplicações.
   - O texto fora dos parênteses é preservado e combinado com o texto modificado dentro dos parênteses.

4. **Mensagem de Conclusão**:
   - Ao final do processamento, uma mensagem é exibida para informar que o processo foi concluído.

## Como Utilizar

1. **Preparação**:
   - Abra o Microsoft Excel e a planilha onde você deseja executar a macro.
   - Selecione o intervalo de células que você deseja processar.

2. **Execução**:
   - Abra o Editor VBA (pressione `ALT + F11`).
   - Insira um novo módulo (clicando com o botão direito no Projeto VBA -> `Inserir` -> `Módulo`).
   - Cole o código VBA fornecido no módulo.
   - (OPCIONAL) Baixe o arquivo desse repositório e importe.

3. **Rodar a Macro**:
   - Feche o Editor VBA e volte para o Excel.
   - Execute a macro (pressione `ALT + F8`, selecione `CompararTextosEntreParenteses` e clique em `Executar`).

4. **Resultado**:
   - Verifique as células no intervalo selecionado. Se houver textos duplicados dentro dos parênteses, eles serão ajustados conforme descrito acima.
   - Uma mensagem de conclusão aparecerá quando o processamento estiver terminado.



## Bugs
  - A macro ativa e limpa o conteúdo de todos os parenteses e mantem apenas o primeiro, portanto, se houver 2 parenteses iguais e um diferente, a macro irá bugar.
