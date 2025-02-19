## Manipulação e Visualização de Despesas em Excel  
Este projeto utiliza Python e as bibliotecas `openpyxl`, `pandas` e `matplotlib` para manipular e visualizar dados financeiros de um arquivo Excel em um ambiente Jupyter Notebook.  

### Funcionalidades  
- Manipulação de planilhas com `openpyxl`  
- Adição de colunas e formatação (alinhamento, bordas e largura das células)  
- Leitura de dados CSV e inserção no Excel  
- Geração de gráfico comparativo de despesas  

### Dependências  
Antes de executar o código, instale as bibliotecas necessárias:  
pip install openpyxl pandas matplotlib jupyter

### Estrutura do Projeto  
projeto_despesas
 ├── despesas.xlsx (arquivo base)
 ├── despesas_geradas.csv (dados para importar)
 ├── despesas_atualizada.xlsx (arquivo atualizado)
 ├── notebook.ipynb (código principal em Jupyter Notebook)
 ├── README.md (documentação)

### Como Executar  
1. Certifique-se de que os arquivos `despesas.xlsx` e `despesas_geradas.csv` estejam na pasta correta.  
2. Abra o Jupyter Notebook
3. Execute as células do arquivo `notebook.ipynb`.  
4. Após a execução:  
   - O arquivo **"despesas_atualizada.xlsx"** será gerado com os dados formatados.  
   - Um gráfico de barras comparando os tipos de despesas será exibido dentro do Jupyter Notebook.  


