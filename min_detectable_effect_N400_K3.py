#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import numpy as np
from datetime import datetime, UTC

# -----------------------
# CONFIG
# -----------------------
K = 3
ORDER = 3
N = 400
N_SHUFFLES = 80          # shuffle reps per trial
N_TRIALS = 200           # Monte Carlo trials per delta
Z_THRESH = 2.0           # detection threshold
TARGET_POWER = 0.60

DELTAS = np.linspace(0.0, 0.30, 16)  # tune range if needed
rng = np.random.default_rng(123)

# -----------------------
# MI tools
# -----------------------
def entropy(counts):
    tot = counts.sum()
    if tot <= 0:
        return 0.0
    p = counts / tot
    p = p[p > 0]
    return float(-(p * np.log2(p)).sum())

def mutual_info(digits_youngest_first: str):
    chrono = digits_youngest_first[::-1]
    n = len(chrono)
    if n <= ORDER:
        return 0.0, 0.0

    state_map = {}
    joint = []
    next_counts = np.zeros(K)

    def idx(s):
        if s not in state_map:
            state_map[s] = len(joint)
            joint.append(np.zeros(K))
        return state_map[s]

    for i in range(ORDER, n):
        s = chrono[i-ORDER:i]
        x = int(chrono[i]) - 1
        j = idx(s)
        joint[j][x] += 1
        next_counts[x] += 1

    Hx = entropy(next_counts)
    total = next_counts.sum()

    HxS = 0.0
    for row in joint:
        rs = row.sum()
        if rs > 0:
            HxS += (rs / total) * entropy(row)

    return Hx - HxS, Hx

def shuffle_corrected_MI(digits: str):
    MI_real, Hx = mutual_info(digits)

    arr = np.array(list(digits), dtype="<U1")
    sh = np.empty(N_SHUFFLES, dtype=float)

    for i in range(N_SHUFFLES):
        rng.shuffle(arr)
        MI_s, _ = mutual_info("".join(arr))
        sh[i] = MI_s

    bias = float(sh.mean())
    sd = float(sh.std(ddof=0))
    MI_corr = MI_real - bias
    Z = MI_corr / sd if sd > 0 else 0.0
    return MI_real, bias, MI_corr, Z, Hx

# -----------------------
# Tunable generator
# -----------------------
def generate_seq(delta: float) -> str:
    """
    Base is uniform.
    If newest digit is '1', increase P(next='1') by delta, take from others equally.
    """
    seq = [str(rng.integers(1, 4)) for _ in range(ORDER)]
    while len(seq) < N:
        newest = seq[-1]
        p = np.array([1/3, 1/3, 1/3], dtype=float)
        if newest == "1":
            p[0] += delta
            p[1] -= delta/2
            p[2] -= delta/2
            p = np.clip(p, 1e-12, 1.0)
            p = p / p.sum()
        nxt = str(rng.choice([1, 2, 3], p=p))
        seq.append(nxt)
    return "".join(seq[::-1])  # youngest-first

# -----------------------
# Main
# -----------------------
def main():
    print(f"N={N} K={K} ORDER={ORDER} trials={N_TRIALS} shuffles={N_SHUFFLES}")
    print("delta,power(Z>2),MIcorr_mean,MIcorr_p10,MIcorr_p90,Z_mean")

    best = None

    for d in DELTAS:
        Z_hits = 0
        corr_vals = []
        z_vals = []

        for _ in range(N_TRIALS):
            digits = generate_seq(d)
            _, _, corr, z, _ = shuffle_corrected_MI(digits)
            corr_vals.append(corr)
            z_vals.append(z)
            if z > Z_THRESH:
                Z_hits += 1

        power = Z_hits / N_TRIALS
        corr_vals = np.array(corr_vals)
        z_vals = np.array(z_vals)

        line = (
            f"{d:.3f},{power:.3f},"
            f"{corr_vals.mean():.4f},"
            f"{np.percentile(corr_vals,10):.4f},"
            f"{np.percentile(corr_vals,90):.4f},"
            f"{z_vals.mean():.2f}"
        )
        print(line)

        if best is None and power >= TARGET_POWER:
            best = d

    print("\nRESULT")
    if best is None:
        print(f"No delta in grid reached power {TARGET_POWER:.0%}. Increase DELTAS max.")
    else:
        print(f"Approx min detectable delta for power {TARGET_POWER:.0%} at Z>{Z_THRESH}: {best:.3f}")
    print("GeneratedUTC:", datetime.now(UTC).isoformat())

if __name__ == "__main__":
    main()