import os
import shutil
from template_utils import IDEAL_DOCX_PATH, REPORTS_DIR

OUTPUT_DOCX_PATH = os.path.join(REPORTS_DIR, "Проблема_синхронизации_данных.docx")


def generate_report_docx(output_filename: str = OUTPUT_DOCX_PATH) -> None:
    """
    Копирует эталонный академический доклад в папку готовых докладов.
    """
    os.makedirs(REPORTS_DIR, exist_ok=True)
    if os.path.exists(IDEAL_DOCX_PATH):
        shutil.copy2(IDEAL_DOCX_PATH, output_filename)
        print(f"[+] Идеальный доклад успешно скопирован в '{output_filename}' из эталона '{IDEAL_DOCX_PATH}'.")
    else:
        print(f"[!] Файл эталона '{IDEAL_DOCX_PATH}' не найден.")


if __name__ == "__main__":
    generate_report_docx(OUTPUT_DOCX_PATH)
