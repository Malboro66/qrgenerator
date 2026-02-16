# 5 Sugestões de Melhoria para UI/UX

1. **Adicionar seletor explícito de formato de saída (PDF/PNG/ZIP/SVG)**
   - Hoje o fluxo depende de `self.formato_saida`, mas não há controles visíveis no layout para o usuário escolher o formato com clareza.
   - Sugestão: incluir `Radiobuttons` ou `Combobox` no topo para reduzir ambiguidade.

2. **Exibir resumo pós-processamento no lugar de apenas modal de sucesso**
   - Atualmente o feedback final é apenas `messagebox` com caminho.
   - Sugestão: mostrar também total processado, total ignorado por validação e duração da geração em um painel de status persistente.

3. **Melhorar feedback de progresso durante carregamento x geração**
   - O mesmo componente de progresso é reaproveitado para carregamento e geração, mas o texto pode ficar pouco descritivo.
   - Sugestão: separar visualmente estados (ex.: badge “Carregando arquivo” vs “Gerando códigos”) e ETA aproximado em lotes grandes.

4. **Adicionar botão de cancelamento de operação longa**
   - O app usa threads e mantém UI responsiva, porém não há ação de cancelamento durante geração.
   - Sugestão: implementar token de cancelamento e botão “Cancelar” habilitado durante tarefas em andamento.

5. **Fortalecer prevenção de erros com microcopy contextual na UI**
   - Existem validações importantes (tamanho, limites, registros inválidos), mas o usuário só as vê quando tenta gerar.
   - Sugestão: exibir dicas inline próximas aos campos (ex.: faixa válida de tamanho, limite por lote e regras de barcode) para reduzir tentativa/erro.
