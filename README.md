# Diário de Classe — 1º Ano

Sistema de planilhas para controle de presença, notas e planejamento semanal de aulas do 1º ano do ensino fundamental público.

## Arquivos gerados

| Arquivo | Abas | Uso |
|---------|------|-----|
| `diario-de-classe.xlsx` | 19 abas | Uso diário |
| `planejamento-semanal.xlsx` | 5 abas | Uso semanal |

## diario-de-classe.xlsx

| Aba | Descrição |
|-----|-----------|
| `Turma` | Lista de 27 alunos com dados cadastrais e flag de aluno especial |
| `Presença - Fev` ... `Presença - Nov` | Controle mensal P/F/J com totais automáticos e alerta de frequência < 75% |
| `Notas - B1` ... `Notas - B4` | Notas das 6 disciplinas por bimestre — azul (≥ 5 Aprovado), vermelho (< 5 Reprovado) |
| `Recuperação` | Controle de alunos em recuperação com situação final automática |
| `Contatos e Reuniões` | Registro de interações com responsáveis |
| `Ocorrências` | Registro de ocorrências e providências |
| `Resumo Anual` | Painel consolidado com médias, frequência e situação final de cada aluno |

## planejamento-semanal.xlsx

| Aba | Descrição |
|-----|-----------|
| `Calendário Letivo` | Datas dos bimestres, feriados e eventos da escola |
| `B1 - Planejamento` ... `B4 - Planejamento` | Planejamento semanal por disciplina (Conteúdo / Objetivo / Atividade / Realizado?) |

## Disciplinas

Português · Matemática · Ciências · História · Geografia · Artes

## Como regenerar as planilhas

```bash
python scripts/criar_diario.py
python scripts/criar_planejamento.py
```

**Requisito:** Python 3.8+ com `openpyxl` instalado.

```bash
pip install openpyxl
```

## Próximos passos

- [ ] Preencher nomes dos alunos na aba `Turma`
- [ ] Ajustar datas no `Calendário Letivo` conforme calendário municipal
- [ ] Adicionar livro didático ao projeto para preenchimento do planejamento semanal
