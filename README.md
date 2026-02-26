# Automação de Validação Cruzada – Google Sheets (Apps Script)

## Objetivo

Desenvolver automação para validar dados entre duas abas ("BASE" e "PEDIDOS") e atualizar status automaticamente, reduzindo inconsistências operacionais.

## Problema

Processos manuais geravam divergências entre:

- Código de produto
- Cliente
- Quantidade
- Preço
- Datas
- Identificadores internos

Isso exigia conferência manual linha a linha.

## Solução

Implementação de:

- Trigger onEdit
- Normalização de valores numéricos
- Conversão de datas no padrão brasileiro
- Cruzamento completo entre bases
- Atualização automática de status
- Destaque visual de inconsistências

## Regras Aplicadas

Se todos os critérios coincidirem:
→ Status atualizado para "ENCERRADO"

Se houver divergência:
→ Destaque automático das colunas inconsistentes

## Competências Demonstradas

- JavaScript
- Automação de processos
- Lógica condicional
- Validação de dados
- Tratamento de datas
- Integração e consistência de informações
