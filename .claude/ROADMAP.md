# Roadmap — Sistema de Planilhas da Professora

## Contexto
Professora de 1º ano do ensino fundamental público.  
27 alunos (1 aluno especial). Avaliação numérica (0–10). 4 bimestres. 1 prova por disciplina por bimestre.  
Disciplinas: Português, Matemática, Ciências, História, Geografia, Artes.

---

## FEITO

- [x] Levantamento de requisitos com a professora
- [x] Design aprovado das duas planilhas
- [x] Documento de spec criado em `docs/superpowers/specs/2026-04-19-diario-professor-design.md`
- [x] Plano de implementação criado em `docs/superpowers/plans/2026-04-19-diario-professor.md`
- [x] Criação da pasta `.claude` com este ROADMAP
- [x] Setup Python 3.11 + openpyxl 3.1.2
- [x] `diario-de-classe.xlsx` — Aba Turma (27 alunos, headers, formatação)
- [x] `diario-de-classe.xlsx` — Abas Presença mensais Fev-Nov (P/F/J, totalizadores, alerta 75%)
- [x] `diario-de-classe.xlsx` — Abas Notas B1-B4 (6 disciplinas, Média Geral, cores condicional)
- [x] `diario-de-classe.xlsx` — Aba Recuperação (40 linhas, Situação Final automática)
- [x] `diario-de-classe.xlsx` — Aba Contatos e Reuniões
- [x] `diario-de-classe.xlsx` — Aba Ocorrências
- [x] `diario-de-classe.xlsx` — Aba Resumo Anual (médias, frequência, situação)
- [x] `planejamento-semanal.xlsx` — Calendário Letivo com bimestres e feriados
- [x] `planejamento-semanal.xlsx` — Planejamento B1-B4 (semanas × disciplinas × conteúdo/objetivo/atividade/realizado)

---

## PRÓXIMAS FASES (aguardando)

- [ ] Professora adiciona o livro didático à pasta do projeto
- [ ] Preencher conteúdo das aulas no planejamento semanal conforme livro e metodologia da escola
- [ ] Ajustar datas no Calendário Letivo conforme calendário municipal real
- [ ] (Opcional) Adicionar nomes reais dos alunos na aba Turma

---

## ARQUIVOS GERADOS

| Arquivo | Abas | Uso |
|---------|------|-----|
| `diario-de-classe.xlsx` | 19 abas | Uso diário |
| `planejamento-semanal.xlsx` | 5 abas | Uso semanal |

## COMO REGENERAR

Se precisar regenerar os arquivos após alterações no script:

```bash
python scripts/criar_diario.py
python scripts/criar_planejamento.py
```
