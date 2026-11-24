# Agendamento no Windows

O modo automático roda a sincronização sem abrir janela. É o modo indicado para o Agendador de Tarefas.

## Antes de agendar

1. Abra a interface pelo menos uma vez e confirme que origem e destino aparecem como **Conectado**.
2. Faça uma sincronização manual para conferir o resultado.
3. Teste o modo automático sem alterar nada:

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Ferramentas\SincronizadorMapas\SincronizadorMapas.ps1" -Auto -Simular
   ```

## Criar a tarefa

1. Abra o **Agendador de Tarefas** (`taskschd.msc`).
2. Clique em **Criar Tarefa** (não "Criar Tarefa Básica", para ter acesso a todas as opções).
3. Aba **Geral**
   - Nome: `Sincronizador de Mapas PDF`
   - Marque **Executar somente quando o usuário estiver conectado**. O Google Drive para computador só monta a unidade quando o usuário está logado.
4. Aba **Disparadores** > **Novo**
   - Exemplo: diariamente às 7h, repetindo a cada 1 hora por 12 horas.
5. Aba **Ações** > **Novo**
   - Ação: **Iniciar um programa**
   - Programa: `C:\Ferramentas\SincronizadorMapas\Sincronizar sem janela.bat`
   - Iniciar em: `C:\Ferramentas\SincronizadorMapas`
6. Aba **Configurações**
   - Marque **Se a tarefa já estiver em execução: Não iniciar uma nova instância**.
   - Marque **Parar a tarefa se ela for executada por mais de: 1 hora**.

## Como saber se deu certo

- A coluna **Resultado da Última Execução** do Agendador mostra o código de saída:

  | Código | Significado |
  |---|---|
  | `0x0` | Concluído sem falhas |
  | `0x1` | Concluído com falhas ou interrompido |
  | `0x2` | Não executou (servidor ou Drive indisponível, ou outra sincronização em andamento) |

- O **Relatório** da interface mostra as execuções automáticas no filtro **Execuções**, marcadas como "Automática".
- Cada execução gera um log `Sincronizacao_AAAA-MM-DD_HH-MM-SS.log` na pasta de logs.

## Se a janela estiver aberta na hora agendada

Não há conflito. Se a janela estiver sincronizando, a tarefa agendada encontra a trava de execução, registra que não executou e termina com código 2. A próxima execução segue normalmente.
