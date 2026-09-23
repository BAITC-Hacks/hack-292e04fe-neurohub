from pathlib import Path
import pandas as pd

# Папка с Excel-файлами
INPUT_DIR = Path("systeme_electric")

# Папка, куда сохранятся CSV
OUTPUT_DIR = Path("systeme_electric_csv")
OUTPUT_DIR.mkdir(exist_ok=True)

# Находим все .xlsx
excel_files = list(INPUT_DIR.glob("*.xlsx"))

print(f"Найдено Excel-файлов: {len(excel_files)}")

for file in excel_files:
    print(f"\nОбрабатываю: {file.name}")

    # Читаем Excel
    df = pd.read_excel(file)

    # Имя будущего CSV
    output_file = OUTPUT_DIR / f"{file.stem}.csv"

    # Сохраняем
    df.to_csv(
        output_file,
        index=False,
        encoding="utf-8-sig"
    )

    print(f"Сохранено: {output_file}")

print("\nГотово!")