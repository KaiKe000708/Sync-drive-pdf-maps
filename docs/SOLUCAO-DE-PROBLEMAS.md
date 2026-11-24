# Solução de problemas

## A janela não abre

**Nada acontece ao clicar no `.bat`**
O Windows pode ter bloqueado o script baixado da internet. Clique com o botão direito em `SincronizadorMapas.ps1` > **Propriedades** > marque **Desbloquear** > **OK**.

Para ver a mensagem de erro, abra o PowerShell na pasta da ferramenta e rode:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File .\SincronizadorMapas.ps1
```

**Aparece "O arquivo config.json está com erro de formatação"**
O `config.json` foi editado e ficou com algum erro de sintaxe (vírgula sobrando, aspas faltando). Corrija o arquivo ou apague-o: ele é recriado com os valores padrão e você escolhe as pastas de novo pela interface. Lembre que no JSON a barra invertida é escrita dobrada (`\\`).

**Acentos aparecem como símbolos estranhos**
O `SincronizadorMapas.ps1` foi salvo sem BOM por algum editor. Abra-o no VS Code, clique na codificação no canto inferior direito e escolha **Salvar com codificação > UTF-8 com BOM**.

## Origem ou destino indisponível

**Origem: Indisponível**
- Confirme que o computador está na rede da empresa (ou na VPN).
- Abra o caminho do servidor no Explorador de Arquivos para ver se pede senha.
- Confira se o caminho no campo Origem está correto.

**Destino: Indisponível**
- O Google Drive para computador precisa estar aberto e conectado.
- Confira se a letra da unidade continua a mesma (ela pode mudar depois de reinstalar o Google Drive).

**Destino: Será criada**
A pasta ainda não existe, mas a pasta acima dela existe. Ela será criada na primeira sincronização.

## Durante a sincronização

**Um arquivo ficou como "Falhou"**
Passe o mouse sobre a etiqueta para ver o motivo. Os mais comuns:

| Motivo | O que fazer |
|---|---|
| O arquivo está sendo usado por outro processo | Feche o PDF no computador que o tem aberto e sincronize de novo |
| O arquivo foi modificado no servidor durante a cópia | Alguém estava salvando o PDF. Basta sincronizar de novo |
| Acesso negado | Falta permissão de leitura no servidor ou de escrita no Drive |

**"Já existe uma sincronização em andamento"**
Outra janela ou a tarefa agendada está sincronizando. Aguarde terminar.

**O mesmo arquivo é copiado em toda sincronização**
Acontece se o arquivo no servidor tiver a data de modificação no futuro (relógio do servidor adiantado). Confira a data na coluna Modificado.

## Onde ficam os registros

| Arquivo | Local padrão | Conteúdo |
|---|---|---|
| `Sincronizacao_*.log` | Pasta de logs | Registro técnico de cada execução |
| `historico.csv` | Pasta de logs | O que foi adicionado, atualizado, arquivado ou falhou |
| `estado.json` | Pasta da ferramenta | Última sincronização e PDFs já registrados como fora do servidor |
| `config.json` | Pasta da ferramenta | Configurações |

Se a pasta de logs estiver inacessível (Drive desconectado), os registros vão para `%LOCALAPPDATA%\SincronizadorMapas\Logs`.
