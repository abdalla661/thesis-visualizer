import pandas as pd

baseline = pd.read_csv(
    "data/C3_18_2_v2_3/baseline/baseline_schedule.csv"
)

affected = pd.read_csv(
    "data/C3_18_2_v2_3/disruptions/D1/"
    "affected_fractions_D1_machine_window.csv"
)

baseline_slot = baseline[
    (baseline["day"] == 26)
    & (baseline["machine"] == 2)
    & (baseline["window"] == 1)
]

affected_slot = affected[
    (affected["day"] == 26)
    & (affected["machine"] == 2)
    & (affected["window"] == 1)
]

print("Baseline appointments:", len(baseline_slot))
print("Affected appointments:", len(affected_slot))

print("\nBaseline keys:")
print(baseline_slot[["patient", "fraction"]].to_string(index=False))

print("\nAffected keys:")
print(affected_slot[["patient", "fraction"]].to_string(index=False))

baseline_keys = set(
    zip(
        baseline_slot["patient"].astype(int),
        baseline_slot["fraction"].astype(int),
    )
)

affected_keys = set(
    zip(
        affected_slot["patient"].astype(int),
        affected_slot["fraction"].astype(int),
    )
)

print("Matching:", len(baseline_keys & affected_keys))
print("Missing from baseline:", affected_keys - baseline_keys)