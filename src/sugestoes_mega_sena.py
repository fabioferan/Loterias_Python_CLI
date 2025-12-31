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

# === INPUTS DO USUÁRIO ===
try:
    print("\n--- CONFIGURAÇÃO DA GERAÇÃO ---")
    qtde_numeros = int(input("Quantos números por jogo (6 a 20)? "))
    if not (6 <= qtde_numeros <= 20):
        print("⚠️ Quantidade inválida! Usando padrão: 6")
        qtde_numeros = 6
    
    qtde_sugestoes = int(input("Quantas sugestões deseja gerar? "))
    if qtde_sugestoes < 1:
        print("⚠️ Quantidade inválida! Usando padrão: 3")
        qtde_sugestoes = 3
except ValueError:
    print("⚠️ Entrada inválida! Usando valores padrão (6 números, 3 sugestões).")
    qtde_numeros = 6
    qtde_sugestoes = 3


if numeros_nunca_sorteados:
    print("\n🎉 WOW! Existem números que NUNCA foram sorteados:")
    print(sorted(numeros_nunca_sorteados))
    print("💡 Dica: Jogar nesses números é extremamente raro — mas possível!")
else:
    print("\n❌ Todos os números de 1 a 60 já foram sorteados pelo menos uma vez.")
    print("\n🔍 Buscando os números MENOS sorteados (mais 'esquecidos')...")

    # Ajustar tamanho do pool de candidatos com base na quantidade de números
    # Precisamos de pelo menos o dobro de números para ter variabilidade
    tamanho_pool = max(15, qtde_numeros * 2)
    tamanho_pool = min(tamanho_pool, 60) # Limite máximo
    
    menos_sorteados = sorted(frequencia.items(), key=lambda x: x[1])[:tamanho_pool]
    menos_sorteados_nums = [num for num, freq in menos_sorteados]
    
    print(f"🏆 Top {tamanho_pool} números menos sorteados (frequência):")
    for num, freq in menos_sorteados[:12]: # Mostrar só os top 12 pra não poluir
        print(f"   {num} → {freq} vez(es)")
    if tamanho_pool > 12:
        print(f"   ... e mais {tamanho_pool - 12} números.")

    # === CRITÉRIOS PARA SUGESTÃO BALANCEADA ===
    def is_balanced(numbers, total_nums):
        """Verifica se a combinação tem boa distribuição"""
        pares = sum(1 for n in numbers if n % 2 == 0)
        impares = total_nums - pares
        baixos = sum(1 for n in numbers if n <= 30)
        altos = total_nums - baixos
        
        # Margem de aceitação (aprox 1/3 a 2/3)
        min_v = int(total_nums * 0.3)
        max_v = int(total_nums * 0.7) + 1 # +1 para garantir margem em nums pequenos
        
        return (min_v <= pares <= max_v) and (min_v <= baixos <= max_v)

    # Gerar combinações candidatas a partir dos números menos sorteados
    candidatos = menos_sorteados_nums
    melhores_sugestoes = []

    # Tentar combinar números dos menos sorteados, com boa distribuição
    attempts = 0
    max_attempts = 20000
    while len(melhores_sugestoes) < qtde_sugestoes and attempts < max_attempts:
        attempts += 1
        # Se o pool for menor que a quantidade pedida (pode acontecer se filtrarmos demais), usa todos
        if len(candidatos) < qtde_numeros:
             sample = candidatos
        else:
             sample = random.sample(candidatos, qtde_numeros)
             
        if is_balanced(sample, qtde_numeros) and sorted(sample) not in melhores_sugestoes:
            melhores_sugestoes.append(sorted(sample))
            
    # Se não conseguiu balanceados o suficiente, preenche com aleatórios do pool
    while len(melhores_sugestoes) < qtde_sugestoes:
        if len(candidatos) >= qtde_numeros:
            s = sorted(random.sample(candidatos, qtde_numeros))
            if s not in melhores_sugestoes:
                melhores_sugestoes.append(s)
        else:
            break

    if melhores_sugestoes:
        print(f"\n✨ {len(melhores_sugestoes)} SUGESTÕES GERADAS COM BASE NOS MENOS SORTEADOS:")
        for i, sug in enumerate(melhores_sugestoes, 1):
            pares = sum(1 for n in sug if n % 2 == 0)
            impares = qtde_numeros - pares
            baixos = sum(1 for n in sug if n <= 30)
            altos = qtde_numeros - baixos
            print(f"   {i}. {sug}")
            print(f"      → Pares: {pares} | Ímpares: {impares} | Baixos: {baixos} | Altos: {altos}")
    else:
        print(f"\n⚠️  Não foi possível gerar combinações adequadas.")

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