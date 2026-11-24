# Versão 1

O script original, guardado para comparação. Os caminhos reais foram trocados por exemplos.

**Não use esta versão.** O `robocopy /MIR` apaga do destino tudo o que não estiver na origem. Se o servidor responder vazio por uma falha de rede, os arquivos são removidos do Google Drive.

## Problemas que levaram à versão 2

| Problema | Consequência |
|---|---|
| `/MIR` espelha exclusões | Uma falha de acesso ao servidor podia apagar os mapas do Drive |
| Copiava subpastas | Arquivos que não precisavam ir para o Drive iam junto |
| Sem verificação prévia | O `robocopy` começava e falhava no meio se o servidor ou o Drive estivessem fora |
| Data do log por `%date%` | O nome do arquivo quebrava se o formato regional do Windows mudasse |
| Código de retorno só exibido | Não distinguia sucesso de falha (no `robocopy`, 1 é sucesso e 8 ou mais é erro) |
| Log sem Unicode | Nomes de arquivo com acento ficavam ilegíveis |
| Sem limpeza de logs | Os arquivos de log se acumulavam sem limite |
