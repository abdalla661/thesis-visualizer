from config import (
    BASELINE,
    DISRUPTIONS,
    STRATEGY_FOLDER_NAMES,
    strategy_paths,
)


def check_file(label, path):
    status = "OK" if path.exists() else "MISSING"
    print(f"{status:8} {label}: {path}")


check_file("Baseline schedule", BASELINE["schedule_csv"])

for disruption_id, disruption in DISRUPTIONS.items():
    check_file(
        f"{disruption_id} affected",
        disruption["affected"],
    )

    for strategy in STRATEGY_FOLDER_NAMES:
        paths = strategy_paths(disruption_id, strategy)

        check_file(
            f"{disruption_id} {strategy} schedule",
            paths["schedule_csv"],
        )
        check_file(
            f"{disruption_id} {strategy} summary",
            paths["summary"],
        )