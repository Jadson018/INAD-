Nome do Projeto: INAD+ - Sistema de Análise de Big Data para Polo Imóveis

Desenvolvedor: Jadson Pamplona Viana

Descrição Geral:
O INAD+ é um software desenvolvido para a empresa Polo Imóveis, com o objetivo de processar e analisar dados financeiros de planilhas Excel, transformando-os em gráficos informativos e interativos. O sistema facilita a visualização dos dados financeiros de forma clara e eficiente, permitindo aos usuários identificar facilmente informações críticas como valores devidos por obra, parcelas atrasadas e valores devidos dentro de intervalos de datas especificados.

Funcionalidades Principais:

Interface Gráfica com Tkinter:

O software utiliza a biblioteca Tkinter para criar uma interface gráfica amigável e intuitiva, facilitando a interação do usuário com o sistema.

Seleção de Planilhas Excel:

Os usuários podem selecionar arquivos Excel para análise através de uma janela de diálogo de arquivo. O sistema suporta arquivos com extensão .xlsx.

Análise de Dados Financeiros:

O software permite três tipos principais de análise:
Total Devido por Obra:
Agrupa e soma os valores devidos por cada obra, exibindo os resultados em um gráfico de barras horizontais.

Valor Devido por Obra + Data:
Filtra os dados por um intervalo de datas especificado pelo usuário e agrupa os valores devidos por cada obra dentro desse intervalo.

Total Devido por Parcelas + Obra:
Agrupa os dados pelo número de parcelas atrasadas e soma os valores devidos para cada obra específica.

Visualização Gráfica com Matplotlib:

Os dados processados são exibidos em gráficos de barras horizontais, permitindo uma fácil visualização e interpretação das informações financeiras. Os gráficos são gerados com a biblioteca Matplotlib e incluem rótulos detalhados com valores monetários e percentuais.

Tratamento de Erros e Interface Responsiva:

O sistema inclui tratamento de erros para garantir que problemas comuns, como falhas ao carregar arquivos ou formatação de dados inválida, sejam comunicados ao usuário de forma clara. Além disso, a interface é configurada para ser responsiva e maximizar a janela automaticamente para melhor visualização.
Detalhes Técnicos:

Linguagem de Programação: Python
Bibliotecas Utilizadas:
pandas: Para manipulação e análise de dados.
matplotlib: Para criação de gráficos.
tkinter: Para interface gráfica.
tkcalendar: Para seleção de datas.
Pillow: Para manipulação de imagens.
Formato de Arquivo Suportado: .xlsx
Localidade: Configurado para português do Brasil (pt_BR.UTF-8).
Processo de Uso:

Iniciar o Sistema:

Ao iniciar o software, uma janela principal é exibida com três botões, cada um representando um tipo de análise disponível.
Selecionar Análise:

O usuário seleciona o tipo de análise desejada:
Total Devido por Obra: Abre uma janela de diálogo para seleção de uma planilha Excel e gera um gráfico dos valores devidos por obra.
Valor Devido por Obra + Data: Abre uma nova janela para seleção de um intervalo de datas e, em seguida, permite a seleção de uma planilha Excel. Gera um gráfico dos valores devidos por obra dentro do intervalo especificado.
Total Devido por Parcelas + Obra: Abre uma janela de diálogo para seleção de uma planilha Excel, seguida de uma nova janela para seleção de uma obra específica. Gera um gráfico dos valores devidos por número de parcelas atrasadas para a obra selecionada.
Visualização dos Resultados:

Após a seleção dos dados e execução da análise, os resultados são exibidos em gráficos detalhados, maximizando a janela para melhor visualização. Os gráficos incluem rótulos de valores monetários e percentuais, facilitando a compreensão dos dados apresentados.
