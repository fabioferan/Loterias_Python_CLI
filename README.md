# Loterias (Python CLI)

Este projeto utiliza dados históricos dos resultados da Mega-Sena para fornecer ferramentas de análise e sugestão de números através de uma interface de linha de comando (CLI) em Python.

## Funcionalidades

O projeto é dividido em dois scripts principais:

### 1. Verificador de Resultados (`resultado_mega_sena.py`)

Este script permite que você verifique se um conjunto de 6 números já foi premiado em algum concurso anterior da Mega-Sena.

- **Entrada**: Você informa 6 números de sua escolha.
- **Processamento**: O script lê o arquivo `MEGA_SENA.xlsx` e compara seus números com todos os resultados históricos.
- **Saída**: Lista todos os concursos em que seus números teriam recebido algum prêmio (Quadra, Quina ou Sena), mostrando a data do sorteio, os números sorteados e a quantidade de acertos.

### 2. Gerador de Sugestões (`sugestoes_mega_sena.py`)

Para quem busca inspiração, este script analisa todos os resultados para gerar sugestões de jogos com base em critérios estatísticos.

- **Análise de Frequência**: Calcula a frequência com que cada número (de 1 a 60) foi sorteado.
- **Números "Esquecidos"**: Identifica números que nunca foram sorteados ou os que apareceram com menos frequência.
- **Sugestões Balanceadas**: Gera combinações de 6 números que são estatisticamente equilibradas (ex: mistura de pares/ímpares, baixos/altos) a partir dos números menos frequentes.
- **Estatísticas Gerais**: Exibe os números mais e menos comuns na história dos sorteios.
- **Exportação**: Salva as sugestões geradas em um novo arquivo Excel chamado `sugestoes_mega_sena.xlsx`.

## Como Usar

1. **Pré-requisitos**:
   - Certifique-se de ter o Python instalado.
   - Instale as dependências necessárias:

     ```bash
     pip install -r requirements.txt
     ```

   - Tenha o arquivo `MEGA_SENA.xlsx` com os resultados históricos na pasta raiz do projeto.

2. **Para verificar um resultado**:

   ```bash
   python src/resultado_mega_sena.py
   ```

   E siga as instruções para digitar os números.

3. **Para gerar sugestões**:

   ```bash
   python src/sugestoes_mega_sena.py
   ```

   As estatísticas serão exibidas no terminal e as sugestões salvas em um arquivo Excel.

## Estrutura do Arquivo `MEGA_SENA.xlsx`

Para o correto funcionamento, o arquivo Excel deve conter as seguintes colunas:

- `Concurso`: Número do concurso.
- `Data do Sorteio`: Data em que o sorteio ocorreu.
- `Bola1` a `Bola6`: Os seis números sorteados.

---

Este projeto foi criado como uma ferramenta para estudo de Python, manipulação de dados com a biblioteca `pandas` e análise estatística.
