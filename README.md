# Sistema de Informação para Gestão de Work In Process e Monitoramento de Ordens de Produção

Este repositório contém o código-fonte e a implementação de um Sistema de Informação desenvolvido para a gestão eficiente de Work In Process (WIP) e monitoramento em tempo real de ordens de produção em uma empresa do ramo automobilístico.

## Resumo do Projeto

O processo de Planejamento e Controle da Produção (PCP) desempenha um papel crucial na manufatura, alinhando diferentes partes da cadeia produtiva. Este projeto visa melhorar a assertividade e confiabilidade do fluxo de informações no PCP, reduzindo desperdícios e proporcionando monitoramento em tempo real das ordens de produção.

## Resumo Técnico e Prático

### Detalhes Técnicos
- **Ferramentas Utilizadas:** Microsoft Access, VBA, SQL (DML e DQL)
- **Estrutura do Banco de Dados:** Normalizado, com uma tabela principal ("Fluxo") e tabelas auxiliares.
- **Funcionalidades Principais:**
  - Controle em tempo real das ordens de produção.
  - Apontamento de produção facilitado por códigos de barra.
  - Utilização de consultas SQL para manipulação de dados.

### Aplicação Prática
- **Interface do Usuário:** Desenvolvida em Microsoft Access.
- **Controle em Tempo Real:** Registro do último apontamento de cada ordem de produção.
- **Auditoria e Rastreabilidade:** Utilização da tabela "Histórico" para fins de auditoria e análise.
- **Automatização de Relatórios:** Relatórios diários enviados automaticamente à alta gerência.
- **Painel do Gestor:** Formulário coletando estatísticas em tempo real sobre metas de produção e produção por família de produtos, operação e área.

### Funcionamento e Aplicação do Sistema

- **Desenvolvimento no MS Access com VBA e SQL:**
O sistema foi inteiramente implementado no Microsoft Access, utilizando VBA para funcionalidades específicas e consultas SQL para manipulação eficiente de tabelas.

- **Objetivo e Contexto:**
O objetivo principal é exercer controle em tempo real sobre as ordens de produção na fábrica, fornecendo informações imediatas para a tomada de decisões pelos gestores. A escolha por ferramentas de baixo custo, como Microsoft Access, evitou a necessidade de recorrer a alternativas mais dispendiosas, como ERPs renomados.

- **Estrutura do Banco de Dados:**
O banco de dados, embora normalizado, adota uma abordagem semelhante à modelagem "one big table" na tabela principal "Fluxo". Esta tabela permite obter, em tempo real, a quantidade líquida presente em cada etapa do fluxo produtivo (Work in Process), mantendo apenas o último apontamento de cada ordem de produção.

Outras tabelas, como "Histórico", são processadas em tempo real e servem para auditoria, rastreabilidade e análise. A tabela "Histórico" atua como uma tabela fato no Power BI, alimentando o painel de produção atualizado em intervalos de 10 minutos.

### Considerações Importantes

- **Escalabilidade:** Há a perspectiva de comprometimento da escalabilidade do projeto após alguns meses devido ao aumento expressivo de dados históricos. O Microsoft Access, utilizado inicialmente, tem um limite de armazenamento de alguns gigabytes. Para contornar essa limitação, está planejada a migração das tabelas para um servidor local do Microsoft SQL Server. O sistema continuará a funcionar como um "front-end", mantendo os formulários de apontamento e o painel gerencial.

- **Resultados Expressivos:** Após a implementação plena do sistema na rotina dos colaboradores, os resultados foram notáveis, destacando-se a redução significativa de Work In Process (WIP) e lead time, conforme mencionado anteriormente no resumo do artigo.

- **Utilização de Ferramentas de Apoio à Decisão:**
  - **Painel do Gestor:** Um formulário específico foi desenvolvido para coletar estatísticas em tempo real sobre o atingimento de metas de produção, quantidade produzida por família de produtos, por operação, por área, entre outros.
  
  - **Relatórios One-Page Automatizados:** Relatórios diários eram enviados automaticamente para a alta gerência. Esses relatórios eram gerados de forma automática por meio da integração entre o Outlook, Microsoft Access e o sistema operacional Windows, utilizando VBA e VB (PowerShell) para automação do processo.

- **Relacionamento entre Tabelas:** A tabela principal "Fluxo" e a tabela "Histórico" estabeleciam relações por chaves (SKU do produto), criando uma relação um-para-muitos. Essa estrutura permitia não apenas visualizar a posição e o status da ordem de produção em tempo real, mas também examinar todo o histórico de apontamentos desde a primeira operação.

Essas considerações foram cruciais para a eficiência operacional do sistema, proporcionando não apenas resultados imediatos, mas também preparando-o para enfrentar desafios futuros relacionados ao volume crescente de dados históricos.

## Como Utilizar

1. **Requisitos:**
   - Computadores interconectados em rede.
   - Microsoft Access instalado nos computadores da fábrica.

2. **Implementação:**
   - Clone ou faça o download do repositório.
   - Execute o arquivo do projeto no Microsoft Access.
   - Preencha os dados através do formulário para realizar apontamentos de produção.

3. **Observações:**
   - Acesse o "Painel do Gestor" para estatísticas em tempo real.
   - Consulte os relatórios automatizados para análises diárias.


## Resultados Obtidos

O sistema contribuiu para uma redução significativa de 44,78% no Work In Process (WIP) na empresa em estudo, demonstrando sua eficácia na gestão da produção.

**Nota:** Este projeto foi inicialmente implementado no Microsoft Access para facilitar a adoção pela equipe de produção. A migração para um servidor MS SQL Server foi considerada para garantir a escalabilidade do sistema com o aumento de dados históricos.

---

Este README fornece uma visão geral do projeto. Para mais detalhes técnicos, consulte o artigo publicado. Contribuições são bem-vindas!
