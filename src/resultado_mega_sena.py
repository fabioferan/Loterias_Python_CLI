import pandas as pd

# === CONFIGURAÇÕES ===
ARQUIVO_EXCEL = "MEGA_SENA.xlsx"
COLUNAS_BOLAS = ['Bola1', 'Bola2', 'Bola3', 'Bola4', 'Bola5', 'Bola6']
COLUNA_CONCURSO = 'Concurso'
COLUNA_DATA = 'Data do Sorteio'

# === PEDIR OS NÚMEROS AO USUÁRIO ===
print("🔢 Digite os 6 números da Mega-Sena que deseja consultar (separados por espaço):")
while True:
    entrada = input().strip()
    try:
        numeros_informados = set(map(int, entrada.split()))
        if len(numeros_informados) != 6:
            print("❌ Você deve digitar exatamente 6 números distintos. Tente novamente:")
            continue
        if any(n < 1 or n > 60 for n in numeros_informados):
            print("❌ Todos os números devem estar entre 1 e 60. Tente novamente:")
            continue
        break
    except ValueError:
        print("❌ Entrada inválida. Digite apenas números inteiros separados por espaço:")

print(f"\n✅ Números informados: {sorted(numeros_informados)}\n")

# === LER O ARQUIVO EXCEL ===
try:
    df = pd.read_excel(ARQUIVO_EXCEL)
except FileNotFoundError:
    print(f"Erro: Arquivo '{ARQUIVO_EXCEL}' não encontrado.")
    exit(1)

# Verificar se as colunas existem
colunas_esperadas = [COLUNA_CONCURSO, COLUNA_DATA] + COLUNAS_BOLAS
for col in colunas_esperadas:
    if col not in df.columns:
        print(f"Erro: Coluna '{col}' não encontrada no arquivo.")
        print("Colunas disponíveis:", df.columns.tolist())
        exit(1)

# Remover linhas com dados faltantes nas bolas ou na data
df = df.dropna(subset=COLUNAS_BOLAS + [COLUNA_DATA])

# Converter as colunas das bolas para inteiros
for col in COLUNAS_BOLAS:
    df[col] = df[col].astype(int)


# === CONVERSÃO ROBUSTA DA DATA ===
def parse_data_flexible(date_val):
    if pd.isna(date_val):
        return None
    if isinstance(date_val, pd.Timestamp):
        return date_val.strftime('%d/%m/%Y')
    if isinstance(date_val, str):
        formats = [
            '%d/%m/%Y',  # 25/03/2024
            '%d-%m-%Y',  # 25-03-2024
            '%Y-%m-%d',  # 2024-03-25
            '%d.%m.%Y',  # 25.03.2024
            '%m/%d/%Y',  # 03/25/2024
            '%m-%d-%Y'  # 03-25-2024
        ]
        for fmt in formats:
            try:
                parsed = pd.to_datetime(date_val, format=fmt)
                return parsed.strftime('%d/%m/%Y')
            except ValueError:
                continue
        try:
            parsed = pd.to_datetime(date_val, dayfirst=True, errors='raise')
            return parsed.strftime('%d/%m/%Y')
        except Exception:
            pass
    # Se for número (serial do Excel), tenta converter como data
    if isinstance(date_val, (int, float)):
        try:
            # Excel usa 1899-12-30 como base (diferente do Unix)
            parsed = pd.to_datetime(date_val, origin='1899-12-30', unit='D')
            return parsed.strftime('%d/%m/%Y')
        except Exception:
            pass
    # Se nada der certo, retorna como string
    return str(date_val)


# Aplica a conversão
df[COLUNA_DATA] = df[COLUNA_DATA].apply(parse_data_flexible)

# Verifica e avisa sobre datas inválidas
invalidas = df[COLUNA_DATA].isnull() | (df[COLUNA_DATA] == "None")
if invalidas.any():
    print(f"⚠️  {invalidas.sum()} registro(s) com data inválida ou não reconhecida:")
    print(df.loc[invalidas, [COLUNA_CONCURSO, COLUNA_DATA]])

# === BUSCAR CONCURSOS COM PELO MENOS 4 COINCIDÊNCIAS ===
resultados = []

for index, row in df.iterrows():
    numeros_sorteados = set(row[COLUNAS_BOLAS])
    intersecao = numeros_informados & numeros_sorteados
    qtd_coincidencias = len(intersecao)

    if qtd_coincidencias >= 4:
        resultados.append({
            'Concurso': row[COLUNA_CONCURSO],
            'Data do Sorteio': row[COLUNA_DATA],
            'Números Sorteados': list(numeros_sorteados),
            'Coincidências': sorted(intersecao),
            'Quantidade': qtd_coincidencias
        })

# === EXIBIR RESULTADOS ===
if resultados:
    print(f"🔍 Foram encontrados {len(resultados)} concursos com pelo menos 4 acertos:\n")
    for r in resultados:
        print(f"🎯 Concurso {r['Concurso']} - Data: {r['Data do Sorteio']}")
        print(f"   Números sorteados: {sorted(r['Números Sorteados'])}")
        print(f"   Coincidências: {r['Coincidências']} ({r['Quantidade']} acertos)\n")
else:
    print("❌ Nenhum concurso encontrado com pelo menos 4 acertos.")