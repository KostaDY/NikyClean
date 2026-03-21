#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Synthetic K=3 experiment.

Rule:
If last 3 digits == "123", then next digit = 1 (forced).
Otherwise next digit random in {1,2,3}.

Then perform shuffle-corrected MI analysis.
"""

import numpy as np
from datetime import datetime, UTC
import math

# ============================================================
# SETTINGS
# ============================================================

K = 3
ORDER = 3
N_DIGITS = 400
N_SHUFFLES = 100

rng = np.random.default_rng()

# ============================================================
# GENERATE SYNTHETIC SEQUENCE
# ============================================================

def generate_sequence():
    seq = [str(rng.integers(1, 4)) for _ in range(ORDER)]

    while len(seq) < N_DIGITS:
        last3 = "".join(seq[-3:])

        if last3 == "123":
            nxt = "1"
        else:
            nxt = str(rng.integers(1, 4))

        seq.append(nxt)

    return "".join(seq[::-1])  # youngest-first format


# ============================================================
# INFORMATION THEORY
# ============================================================

def entropy(counts):
    tot = counts.sum()
    if tot == 0:
        return 0.0
    p = counts / tot
    p = p[p > 0]
    return float(-(p * np.log2(p)).sum())


def mutual_info(digits):
    chrono = digits[::-1]  # oldest -> youngest
    n = len(chrono)

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

    HxS = 0
    total = next_counts.sum()
    for row in joint:
        rs = row.sum()
        if rs > 0:
            HxS += (rs/total) * entropy(row)

    MI = Hx - HxS
    return MI, Hx


# ============================================================
# MAIN
# ============================================================

def main():

    digits = generate_sequence()

    MI_real, Hx = mutual_info(digits)

    # Shuffle baseline
    arr = np.array(list(digits))
    shuffle_vals = []

    for _ in range(N_SHUFFLES):
        rng.shuffle(arr)
        shuffled = "".join(arr)
        MI_s, _ = mutual_info(shuffled)
        shuffle_vals.append(MI_s)

    bias = np.mean(shuffle_vals)
    std = np.std(shuffle_vals)

    MI_corr = MI_real - bias
    z = MI_corr / std if std > 0 else 0

    print("============== SYNTHETIC TEST ==============")
    print(f"MI_real  = {MI_real:.4f} bits")
    print(f"Bias     = {bias:.4f} bits")
    print(f"MI_corr  = {MI_corr:.4f} bits")
    print(f"Z-score  = {z:.2f}")
    print(f"Hx       = {Hx:.4f} bits")
    print("GeneratedUTC:", datetime.now(UTC).isoformat())


if __name__ == "__main__":
    main()