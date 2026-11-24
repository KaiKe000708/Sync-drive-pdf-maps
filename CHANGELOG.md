# Histórico de versões

## 2.2 (24/11/2026)

### Alterado
- O Drive passa a ficar com exatamente os mesmos PDFs da pasta principal do servidor. PDFs que saíram do servidor são retirados da pasta do Drive e guardados em `_ARQUIVADOS`. A opção se chama "Manter o Drive igual ao servidor" e vem ligada.
- A configuração `ArquivarRemovidos` foi substituída por `EspelharDrive`.
- Desempenho: estilos e cores são carregados uma única vez para o programa inteiro, e a lista recicla as linhas ao rolar. Caixas de mensagem, revisão e relatório abrem bem mais rápido.

### Adicionado
- Trava de segurança: nada é retirado do Drive se mais da metade dos PDFs fosse sair de uma vez.
- Janela de revisão antes de sincronizar, listando cada arquivo que será adicionado, atualizado ou retirado do Drive.
- Colunas da lista ordenáveis com um clique (arquivo, tamanho, data e situação). Por padrão, o que precisa de atenção aparece primeiro.
- Os números de novos, alterados, em dia e só no Drive funcionam como filtros.
- Busca por nome na lista principal.
- Caixas de confirmação e aviso próprias, no mesmo visual e tema da ferramenta, no lugar das caixas padrão do Windows.

## 2.1.1 (24/11/2026)

### Corrigido
- A janela não abria com o erro "não é um valor válido para a propriedade Foreground". As cores do tema eram criadas com `New-Object`, que embrulha o objeto num `PSObject` que o WPF não reconhece no dicionário de recursos. Agora são criadas com `[SolidColorBrush]::new()`.

## 2.1 (24/11/2026)

### Adicionado
- Tema claro, escuro ou igual ao Windows, com barra de título escura no modo escuro.
- Histórico de alterações em `historico.csv`: adicionados, atualizados, saídas do servidor, arquivados, falhas e resumo de cada execução.
- Tela de relatório com filtros por período e tipo, busca por nome, totais e exportação para planilha.
- Execuções automáticas e manuais identificadas no histórico.

### Alterado
- Textos da interface, do log e da documentação revisados.

## 2.0 (24/11/2026)

Reescrita completa em PowerShell com interface WPF.

### Adicionado
- Interface com indicadores de conexão, lista de arquivos com situação e barra de progresso.
- Cópia somente dos `.pdf` da pasta principal, com conferência exata da extensão.
- Cópia em duas etapas com conferência de tamanho e novas tentativas.
- Verificação de servidor e Google Drive antes de começar.
- Arquivamento opcional dos PDFs que saíram do servidor, com trava de segurança.
- Trava contra execuções simultâneas.
- Modo `-Auto` e `-Simular` para o Agendador de Tarefas.
- Configuração em `config.json` e limpeza automática de logs antigos.
- Layout responsivo.

### Removido
- Exclusão de arquivos no destino (`robocopy /MIR`).
- Cópia de subpastas.

## 1.0

- Script `.bat` com `robocopy /MIR` e log com data no nome. Mantido em [`docs/versao-1`](docs/versao-1) para comparação.
