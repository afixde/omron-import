from pathlib import Path

from csv_reader import OmronCsvReader


def main() -> None:

    print("=" * 45)
    print(" Omron Import v0.2.0")
    print("=" * 45)

    csv_dir = Path("data") / "csv"

    print(f"CSV-Verzeichnis: {csv_dir.resolve()}")

    reader = OmronCsvReader()

    print("CSV-Reader initialisiert.")


if __name__ == "__main__":
    main()