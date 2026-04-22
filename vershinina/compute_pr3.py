"""Расчёты для ПР3 Вершининой — нечёткий восходящий вывод.

Предметная область: подбор мощности искусственной подсветки рабочего
места в офисе (Вт) по уровню естественной освещённости (лк).
"""

# Универсумы
X = [100, 300, 500, 700, 900]   # естественная освещённость, лк
Y = [10, 20, 30, 40, 50]        # мощность подсветки, Вт

# Нечёткие множества на X (лингвистическая переменная «естественная освещённость»)
A = {
    "A1_очень_низкая":   {100: 1.0, 300: 0.5, 500: 0.1, 700: 0.0, 900: 0.0},
    "A2_низкая":         {100: 0.5, 300: 1.0, 500: 0.5, 700: 0.1, 900: 0.0},
    "A3_средняя":        {100: 0.1, 300: 0.5, 500: 1.0, 700: 0.5, 900: 0.1},
    "A4_высокая":        {100: 0.0, 300: 0.1, 500: 0.5, 700: 1.0, 900: 0.5},
    "A5_очень_высокая":  {100: 0.0, 300: 0.0, 500: 0.1, 700: 0.5, 900: 1.0},
}

# Нечёткие множества на Y (лингвистическая переменная «мощность подсветки»)
B = {
    "B1_максимальная":   {10: 0.0, 20: 0.0, 30: 0.1, 40: 0.5, 50: 1.0},
    "B2_высокая":        {10: 0.0, 20: 0.1, 30: 0.5, 40: 1.0, 50: 0.5},
    "B3_средняя":        {10: 0.1, 20: 0.5, 30: 1.0, 40: 0.5, 50: 0.1},
    "B4_низкая":         {10: 0.5, 20: 1.0, 30: 0.5, 40: 0.1, 50: 0.0},
    "B5_минимальная":    {10: 1.0, 20: 0.5, 30: 0.1, 40: 0.0, 50: 0.0},
}

# 5 продукционных правил: IF X есть A_k THEN Y есть B_k
rules = [
    ("A1_очень_низкая",  "B1_максимальная"),
    ("A2_низкая",        "B2_высокая"),
    ("A3_средняя",       "B3_средняя"),
    ("A4_высокая",       "B4_низкая"),
    ("A5_очень_высокая", "B5_минимальная"),
]


def build_relation(A_name, B_name):
    """Матрица нечёткого отношения импликации Мамдани
    μ_R(x, y) = min(μ_A(x), μ_B(y))."""
    a, b = A[A_name], B[B_name]
    return [[min(a[x], b[y]) for y in Y] for x in X]


def union(*matrices):
    """Объединение отношений: μ_R(x, y) = max_i μ_{R_i}(x, y)."""
    out = [[0.0 for _ in Y] for _ in X]
    for m in matrices:
        for i in range(len(X)):
            for j in range(len(Y)):
                if m[i][j] > out[i][j]:
                    out[i][j] = m[i][j]
    return out


def max_min_composition(a_prime, R):
    """B'(y) = max_x min(A'(x), R(x, y))."""
    result = {}
    for j, y in enumerate(Y):
        vals = [min(a_prime[X[i]], R[i][j]) for i in range(len(X))]
        result[y] = max(vals)
    return result


def centroid(b_prime):
    """Дефаззификация методом центра тяжести."""
    num = sum(y * mu for y, mu in b_prime.items())
    den = sum(mu for mu in b_prime.values())
    return num / den if den > 0 else 0.0


def fmt(v):
    return f"{v:.2f}".rstrip("0").rstrip(".") if v else "0"


def print_matrix(name, M, row_labels, col_labels):
    print(f"\n{name}")
    header = "        " + "  ".join(f"{y:>5}" for y in col_labels)
    print(header)
    for i, x in enumerate(row_labels):
        row = f"{x:>5} | " + "  ".join(f"{M[i][j]:>5.2f}" for j in range(len(col_labels)))
        print(row)


if __name__ == "__main__":
    # 1. Матрицы отношений R1..R5
    Rs = []
    for k, (a_name, b_name) in enumerate(rules, start=1):
        Rk = build_relation(a_name, b_name)
        Rs.append(Rk)
        print_matrix(f"R{k}: {a_name} -> {b_name}", Rk, X, Y)

    # 2. Общее отношение
    R = union(*Rs)
    print_matrix("R = R1 ∪ R2 ∪ R3 ∪ R4 ∪ R5", R, X, Y)

    # 3. Наблюдение A': "естественная освещённость ~420 лк"
    A_prime = {100: 0.1, 300: 0.8, 500: 0.7, 700: 0.2, 900: 0.0}
    print("\nA' =", A_prime)

    # 4. Восходящий вывод
    B_prime = max_min_composition(A_prime, R)
    print("\nB' =", B_prime)

    # 5. Дефаззификация
    y_star = centroid(B_prime)
    print(f"\ny* (центр тяжести) = {y_star:.2f} Вт")

    # Подробный расчёт B'(y) для каждого y
    print("\nПодробный расчёт:")
    for j, y in enumerate(Y):
        parts = []
        for i, x in enumerate(X):
            parts.append(f"min({A_prime[x]:.1f}, {R[i][j]:.2f})={min(A_prime[x], R[i][j]):.2f}")
        vals = [min(A_prime[X[i]], R[i][j]) for i in range(len(X))]
        print(f"  B'({y}) = max[{', '.join(parts)}] = {max(vals):.2f}")
