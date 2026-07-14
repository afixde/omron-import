from pathlib import Path

def main():
    print("=" * 40)
    print("Omron Import v0.1.0")
    print("=" * 40)

    print("Projektverzeichnis:")
    print(Path.cwd())

    print("\nBereit.")

if __name__ == "__main__":
    main()