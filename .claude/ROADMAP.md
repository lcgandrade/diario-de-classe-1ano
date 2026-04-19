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
- [x] Criação da pasta `.claude` com este ROADMAP

---

## EM ANDAMENTO

- [ ] Criar `diario-de-classe.xlsx` com todas as abas
- [ ] Criar `planejamento-semanal.xlsx` com estrutura base

---

## TODO

### diario-de-classe.xlsx
- [ ] Aba `Turma` — lista de 27 alunos com campos e flag de aluno especial
- [ ] Abas `Presença - [Mês]` — controle mensal P/F/J com totais e alerta de 75%
- [ ] Aba `Notas - B1` a `B4` — notas por disciplina com formatação condicional (vermelho/amarelo/verde)
- [ ] Aba `Recuperação` — lista automática de alunos abaixo de 5
- [ ] Aba `Contatos e Reuniões` — registro de interações com responsáveis
- [ ] Aba `Ocorrências` — registro de ocorrências relevantes
- [ ] Aba `Resumo Anual` — painel consolidado com médias e situação final

### planejamento-semanal.xlsx
- [ ] Aba `Calendário Letivo` — datas de bimestres, feriados, eventos
- [ ] Abas `B1` a `B4 - Planejamento` — estrutura por semana/disciplina com campo "Realizado?"

### Próximas fases (aguardando)
- [ ] Preencher conteúdo do planejamento semanal conforme livro didático (aguardando entrega do livro)
- [ ] Adaptar metodologia conforme orientações da escola

---

## NOTAS
- O livro didático será adicionado à pasta do projeto futuramente para geração do planejamento semanal com conteúdo real
- O aluno especial deve ser destacado em todas as abas relevantes
- Frequência mínima legal: 75% (LDB) — alertas visuais incluídos
