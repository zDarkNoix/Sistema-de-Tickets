# Sistema de Tickets com Google Apps Script

# Visão Geral
Este projeto é um sistema de gerenciamento de tickets desenvolvido para ser executado diretamente de uma Planilha Google. Ele utiliza o Google Apps Script como backend e o serviço HTML para criar uma interface de usuário web.

O sistema é projetado para equipes de atendimento ou suporte, permitindo que os agentes visualizem, atualizem e resolvam tickets que são, na verdade, linhas em uma planilha.

# Funcionalidades
 - Painel de Tickets: Uma interface web para visualizar todos os tickets "Em Andamento" ou "Resolvidos".
 - Criação e Edição: Permite que um agente "abra" um ticket a partir de um número de linha da planilha e edite suas informações.
 - Filtros Avançados: Filtra tickets por status (Em Andamento, Resolvido, Sem Sucesso) e por estágio de atendimento.
 - Relatório de Desempenho: Uma tela de relatório que exibe estatísticas de resolução de tickets por agente e por data.

# Tecnologias
 - Google Apps Script (.gs): Lógica do backend, manipulação da planilha e exposição de dados para o frontend.
 - HTML/CSS/JavaScript: Frontend do sistema, incluindo os painéis de tickets e relatórios.

# Configuração
Para usar este sistema, basta copiar os três arquivos (Code.gs, index.html, relatorio.html) para um novo projeto do Google Apps Script associado a uma Planilha Google. A estrutura da planilha deve corresponder às colunas definidas no objeto COLUNAS_DADOS em Code.gs.
