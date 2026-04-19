# Design — Sistema de Planilhas para Professora de 1º Ano

**Data:** 2026-04-19  
**Contexto:** Professora de ensino fundamental público, 1º ano, 27 alunos (1 aluno especial incluído)

---

## Contexto e Objetivos

Criar um conjunto de planilhas `.xlsx` para apoiar o dia a dia de uma professora de 1º ano do ensino fundamental público, cobrindo:

- Controle de presença mensal
- Registro e acompanhamento de notas por bimestre
- Planejamento semanal de aulas (estrutura pronta; conteúdo será preenchido após entrega do livro didático)
- Controles adicionais essenciais: recuperação, contato com responsáveis, ocorrências

**Disciplinas:** Português, Matemática, Ciências, História, Geografia, Artes  
**Sistema de avaliação:** Notas numéricas (0–10)  
**Estrutura letiva:** 4 bimestres  
**Avaliações:** 1 prova por disciplina por bimestre  

---

## Arquivos

### 1. `diario-de-classe.xlsx`

Uso diário. Contém presença, notas, recuperação, ocorrências e resumo anual.

#### Aba: `Turma`
- Colunas: Nº, Nome completo, Data de nascimento, Responsável, Contato, Aluno Especial (flag Sim/Não), Observações
- 27 linhas de alunos
- Aluno especial destacado com formatação diferenciada

#### Abas: `Presença - Jan` a `Presença - Dez`
- Uma aba por mês (apenas meses letivos ativos)
- Linhas = alunos (puxados da aba Turma)
- Colunas = dias do mês (apenas dias úteis letivos)
- Valores válidos: **P** (Presente), **F** (Falta), **J** (Justificada)
- Colunas automáticas ao final: Total P, Total F, Total J, % Frequência
- Alerta visual: células de % Frequência abaixo de 75% em vermelho

#### Abas: `Notas - B1`, `Notas - B2`, `Notas - B3`, `Notas - B4`
- Linhas = alunos
- Colunas = Português, Matemática, Ciências, História, Geografia, Artes, **Média Geral**
- 1 nota de prova por disciplina (0–10)
- Formatação condicional: notas abaixo de 5 em vermelho, 5–6,9 em amarelo, 7+ em verde
- Média Geral calculada automaticamente

#### Aba: `Recuperação`
- Gerada a partir das notas: lista alunos com nota < 5 por bimestre e disciplina
- Colunas: Bimestre, Aluno, Disciplina, Nota Original, Nota Recuperação, Situação Final
- Situação Final calculada automaticamente (Aprovado / Em recuperação)

#### Aba: `Contatos e Reuniões`
- Registro de interações com responsáveis
- Colunas: Data, Aluno, Responsável, Tipo (reunião / recado / ligação), Assunto, Encaminhamento

#### Aba: `Ocorrências`
- Registro breve de ocorrências relevantes
- Colunas: Data, Aluno, Descrição, Providência tomada

#### Aba: `Resumo Anual`
- Painel consolidado: média por aluno por bimestre e por disciplina
- Frequência anual por aluno
- Situação final: Aprovado / Recuperação / Reprovado
- Formatação visual com cores para facilitar leitura rápida

---

### 2. `planejamento-semanal.xlsx`

Uso semanal. Estrutura pronta para receber conteúdo quando o livro didático for disponibilizado.

#### Aba: `Calendário Letivo`
- Datas de início e fim de cada bimestre
- Feriados nacionais e municipais (campos editáveis)
- Datas de reunião de pais, conselho de classe, eventos da escola

#### Abas: `B1 - Planejamento`, `B2 - Planejamento`, `B3 - Planejamento`, `B4 - Planejamento`
- Uma aba por bimestre
- Linhas = semanas letivas numeradas (com data de início e fim)
- Colunas = 6 disciplinas
- Cada célula contém:
  - Conteúdo / tema da semana
  - Objetivo de aprendizagem
  - Recurso / atividade prevista
  - Campo "Realizado?" (Sim / Não / Parcial)
- Linha de "Semana de Recuperação" ao final de cada bimestre

---

## Controles Adicionais Incluídos

Com base em boas práticas para professores de ensino fundamental público:

| Controle | Onde | Justificativa |
|----------|------|---------------|
| Frequência com alerta de 75% | `Presença - [Mês]` | Obrigação legal (LDB) |
| Aluno especial destacado | `Turma` + todas as abas | Facilita acompanhamento diferenciado |
| Recuperação paralela | `Recuperação` | Exigência pedagógica bimestral |
| Contato com responsáveis | `Contatos e Reuniões` | Evidência para conselho de classe |
| Ocorrências | `Ocorrências` | Registro institucional |
| Resumo consolidado anual | `Resumo Anual` | Facilita relatórios e conselho de classe |

---

## Fora do Escopo Atual

- Conteúdo real das aulas no planejamento semanal (aguardando livro didático)
- Integração com sistemas da secretaria escolar
- Versão digital de entrega de atividades

---

## Estrutura de Arquivos do Projeto

```
school/
├── diario-de-classe.xlsx
├── planejamento-semanal.xlsx
├── .claude/
│   └── ROADMAP.md
└── docs/
    └── superpowers/
        └── specs/
            └── 2026-04-19-diario-professor-design.md
```
