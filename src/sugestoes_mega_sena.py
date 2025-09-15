import pandas as pd
import random
from collections import Counter

# === CONFIGURAÇÕES ===
ARQUIVO_EXCEL = "MEGA_SENA.xlsx"
COLUNAS_BOLAS = ['Bola1', 'Bola2', 'Bola3', 'Bola4', 'Bola5', 'Bola6']

# === LER O ARQUIVO EXCEL ===
try:
    df = pd.read_excel(ARQUIVO_EXCEL)
except FileNotFoundError:
    print(f"Erro: Arquivo '{ARQUIVO_EXCEL}' não encontrado.")
    exit(1)

# Verificar se as colunas existem
for col in COLUNAS_BOLAS:
    if col not in df.columns:
        print(f"Erro: Coluna '{col}' não encontrada no arquivo.")
        print("Colunas disponíveis:", df.columns.tolist())
        exit(1)

# Remover linhas com dados faltantes nas bolas
df = df.dropna(subset=COLUNAS_BOLAS)

# Converter todas as bolas para inteiros
for col in COLUNAS_BOLAS:
    df[col] = df[col].astype(int)

# === CONTAGEM DE FREQUÊNCIA DE TODOS OS NÚMEROS (1 a 60) ===
todos_numeros = []
for col in COLUNAS_BOLAS:
    todos_numeros.extend(df[col].dropna().astype(int))

frequencia = Counter(todos_numeros)

# Todos os números possíveis na Mega-Sena (1 a 60)
todos_os_numeros = set(range(1, 61))
numeros_sorteados = set(frequencia.keys())
numeros_nunca_sorteados = todos_os_numeros - numeros_sorteados

print("=" * 60)
print("📊 ANÁLISE ESTATÍSTICA DA MEGA-SENA")
print("=" * 60)

if numeros_nunca_sorteados:
    print("🎉 WOW! Existem números que NUNCA foram sorteados:")
    print(sorted(numeros_nunca_sorteados))
    print("💡 Dica: Jogar nesses números é extremamente raro — mas possível!")
    # Sugerir 6 dos nunca sorteados (se houver)
    sugestao = sorted(numeros_nunca_sorteados)[:6]
    print(f"\n✨ SUGESTÃO BASEADA EM NÚMEROS NUNCA SORTEADOS: {sorted(sugestao)}")
else:
    print("❌ Todos os números de 1 a 60 já foram sorteados pelo menos uma vez.")
    print("\n🔍 Buscando os números MENOS sorteados (mais 'esquecidos')...")

    # Pegar os 12 números com menor frequência
    menos_sorteados = sorted(frequencia.items(), key=lambda x: x[1])[:12]
    menos_sorteados_nums = [num for num, freq in menos_sorteados]
    print(f"🏆 Top 12 números menos sorteados (frequência):")
    for num, freq in menos_sorteados:
        print(f"   {num} → {freq} vez(es)")

    # === CRITÉRIOS PARA SUGESTÃO BALANCEADA ===
    def is_balanced(numbers):
        """Verifica se a combinação tem boa distribuição"""
        pares = sum(1 for n in numbers if n % 2 == 0)
        impares = 6 - pares
        baixos = sum(1 for n in numbers if n <= 30)
        altos = 6 - baixos
        return 2 <= pares <= 4 and 2 <= impares <= 4 and 2 <= baixos <= 4 and 2 <= altos <= 4

    # Gerar combinações candidatas a partir dos 12 menos sorteados
    candidatos = menos_sorteados_nums
    melhores_sugestoes = []

    # Tentar combinar 6 números dos 12 menos sorteados, com boa distribuição
    attempts = 0
    max_attempts = 10000
    while len(melhores_sugestoes) < 3 and attempts < max_attempts:
        attempts += 1
        sample = random.sample(candidatos, 6)
        if is_balanced(sample) and sample not in melhores_sugestoes:
            melhores_sugestoes.append(sorted(sample))

    if melhores_sugestoes:
        print(f"\n✨ SUGESTÕES ESTATISTICAMENTE EQUILIBRADAS (com base nos menos sorteados):")
        for i, sug in enumerate(melhores_sugestoes, 1):
            pares = sum(1 for n in sug if n % 2 == 0)
            impares = 6 - pares
            baixos = sum(1 for n in sug if n <= 30)
            altos = 6 - baixos
            print(f"   {i}. {sug}")
            print(f"      → Pares: {pares} | Ímpares: {impares} | Baixos: {baixos} | Altos: {altos}")
    else:
        print(f"\n⚠️  Não encontramos combinações balanceadas entre os 12 menos sorteados.")
        print("💡 Sugestão alternativa: Escolha 6 números dos 12 abaixo, evitando sequências:")
        print(f"   {candidatos}")

# === INFORMAÇÕES ADICIONAIS ===
print("\n" + "=" * 60)
print("📈 ESTATÍSTICAS GERAIS")
print("=" * 60)
total_sorteios = len(df)
print(f"Total de concursos analisados: {total_sorteios}")
print(f"Números únicos sorteados: {len(numeros_sorteados)} / 60")
print(f"Frequência média por número: {sum(frequencia.values()) / 60:.2f}")

mais_sorteados = frequencia.most_common(3)
menos_sorteados = frequencia.most_common()[-3:]
print(f"\n🔝 3 mais sorteados: {[(n,f) for n,f in mais_sorteados]}")
print(f"📉 3 menos sorteados: {[(n,f) for n,f in menos_sorteados]}")

print("\n💡 DICA FINAL: Números com baixa frequência têm a mesma chance teórica, mas jogar neles reduz a probabilidade de dividir o prêmio com muitas pessoas!")

# Salvar sugestões em Excel
if melhores_sugestoes:
    df_sugestoes = pd.DataFrame({
        "Sugestão": [f"Sugestão {i+1}" for i in range(len(melhores_sugestoes))],
        "Números": [", ".join(map(str, sug)) for sug in melhores_sugestoes]
    })
    df_sugestoes.to_excel("sugestoes_mega_sena.xlsx", index=False)
    print("\n💾 Sugestões salvas em 'sugestoes_mega_sena.xlsx'")