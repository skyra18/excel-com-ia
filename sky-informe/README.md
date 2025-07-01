# SKY INFORME

**Sky Informe** é uma planilha desenvolvida em Excel com o objetivo de auxiliar na organização das informações necessárias para a declaração de Imposto de Renda de Pessoa Física.
---

## Sobre o projeto

Esta planilha foi desenvolvida como parte de um projeto proposto no bootcamp *Excel com IA*. A proposta era aplicar conhecimentos de Excel avançado em um projeto funcional, documentado e compartilhável.

A solução escolhida foi a criação de um agregador de dados focado na preparação para o Imposto de Renda. A ideia central é fornecer ao usuário um ambiente que permita controlar suas informações de forma clara, validada e centralizada.

---

## Funcionalidades

- Navegação por botões entre seções (Titular, Informes e Notas)
- Validação automática de dados (ex: campos obrigatórios, formatos)
- Campos otimizados para preenchimento rápido
- Área dedicada para anexos (links de arquivos em PDF)
- Interface visual amigável com temática personalizada

---

## VBA aplicado

A planilha utiliza uma macro em **VBA (Visual Basic for Applications)** para movimentar elementos gráficos, como ícones, dentro da interface da planilha.

A função `MoverIconeParaPosicao()` percorre os objetos visuais da aba ativa e reposiciona o item desejado em uma nova coordenada, conforme definido no código.

> **Nota:** o código VBA foi fornecido pelo professor como parte do conteúdo do curso. A função foi compreendida, testada e adaptada dentro do contexto do projeto.

---

## Organização do repositório

- `Sky_Informe.xlsx` — Arquivo principal da planilha
- `README.md` — Arquivo de documentação
- `/images/` — Pasta com capturas de tela da planilha

---

## Como usar

1. Baixe o arquivo `Sky_Informe.xlsx`
2. Abra no Excel e habilite o uso de macros (caso solicitado)
3. Navegue entre as seções usando os botões laterais
4. Preencha os dados solicitados com atenção aos formatos
5. Utilize os campos de "Anexo" para referenciar PDFs relacionados (como informes de rendimento)
6. A planilha pode ser personalizada conforme a necessidade do usuário

---

## Aprendizados

Durante o desenvolvimento deste projeto, os seguintes pontos foram trabalhados:

- Aplicação prática de Excel em um cenário real
- Uso de formatações avançadas e validações
- Criação de uma interface funcional e limpa
- Entendimento e aplicação de código VBA no Excel
- Células bloqueadas estrategicamente para evitar edições acidentais em fórmulas e estruturas críticas.

---

## Créditos

Projeto desenvolvido como parte do bootcamp **Excel com IA**.

Design e estrutura da planilha por Maria Eduarda Coutinho  
Macro VBA fornecida pelo professor e adaptada para fins educacionais.

