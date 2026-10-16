# Política de segurança

## Como reportar

Use o **private vulnerability reporting** do GitHub: aba **Security** do repositório → **Report a vulnerability**. O relato fica visível só para o mantenedor. Não abra issue pública para falhas de segurança nem para relatar dado pessoal exposto.

Correções entram no `main`; não há versões antigas mantidas em paralelo.

## Escopo

Vale para o código deste repositório: o pipeline (`main.py`, `scripts/`), os testes e o workflow de CI. Exemplos do que reportar:

- leitura de planilha, PDF ou ODT malicioso que execute código ou grave fora da pasta de saída;
- caminho em que o arquivo original de entrada seja alterado ou sobrescrito;
- dependência com vulnerabilidade conhecida que o `pip-audit` do CI não pegou;
- qualquer dado real (de pacientes, profissionais ou da instituição) que tenha entrado no histórico do git.

Fica fora do escopo a segurança das máquinas e das pastas onde o pipeline roda com dados reais. Isso é responsabilidade de quem opera.

## Política de dados

O pipeline processa dados de saúde. Por isso o repositório só versiona código, testes e o gerador de dados fictícios.

| Ignorado pelo git | Por quê |
|---|---|
| `dados_entrada/*` | planilhas, PDF e ODT reais do mês |
| `relatorios_gerados/*` | relatórios Word e logs de auditoria gerados a partir dos dados reais |
| `templates/*` | o documento-base do modelo é um relatório real e contém nomes de pacientes |
| `mapa_variaveis.txt` | gabarito com todos os valores que foram para o relatório |
| `exemplo/dados_entrada/*`, `exemplo/relatorios_gerados/*` | gerados na hora por `scripts/gerar_dados_exemplo.py`; não precisam ser versionados |

Regras:

- **Nunca use `git add -f`** nessas pastas e não remova as entradas do `.gitignore`. Se um arquivo real entrar num commit, não basta apagá-lo num commit seguinte, porque o histórico precisa ser reescrito. Reporte pelo canal acima.
- Testes e exemplos usam só dados sintéticos, gerados em pasta temporária ou pelo gerador. Nenhum teste lê `dados_entrada/`, e o gerador se recusa a gravar lá.
- O log de auditoria (`execucao_<MÊS>_<ANO>_*.log`) grava o usuário do sistema, o nome da máquina, o comando e o hash de cada arquivo de entrada. Ele fica só na pasta de saída local, que é ignorada, e nunca é enviado a lugar nenhum. O pipeline não faz chamadas de rede.
