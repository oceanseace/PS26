import json
import re
from pathlib import Path

from openpyxl import load_workbook


# =========================================================
# AYARLAR
# =========================================================

SOURCE = Path("data.xlsx")
OUTPUT = Path("data")

# İlk satır başlık olarak kabul edilir.
HEADER_ROW = 1


# =========================================================
# DOSYA ADI TEMİZLEME
# =========================================================

def safe_filename(name):
    name = str(name)

    # Windows/GitHub için problem oluşturabilecek karakterleri temizle
    name = re.sub(r'[<>:"/\\|?*]', '_', name)

    # Nokta/boşluk problemlerini azalt
    name = name.strip(" .")

    return name or "sheet"


# =========================================================
# DEĞERİ JSON'A UYGUN HALE GETİR
# =========================================================

def clean_value(value):

    if value is None:
        return ""

    # Tarih vb. değerleri string olarak kaydet
    if hasattr(value, "isoformat"):
        try:
            return value.isoformat()
        except Exception:
            pass

    # JSON'un desteklemediği tipleri string yap
    if not isinstance(
        value,
        (str, int, float, bool)
    ):
        return str(value)

    return value


# =========================================================
# ANA İŞLEM
# =========================================================

def main():

    if not SOURCE.exists():

        print("HATA: data.xlsx bulunamadı.")
        print(f"Aranan dosya: {SOURCE.resolve()}")
        return

    print()
    print("Excel okunuyor...")
    print(f"Kaynak: {SOURCE.resolve()}")
    print()


    # -----------------------------------------------------
    # Workbook'u read_only modunda açıyoruz.
    # Böylece RAM kullanımı azalır.
    # -----------------------------------------------------

    workbook = load_workbook(
        SOURCE,
        read_only=True,
        data_only=True
    )


    # -----------------------------------------------------
    # data klasörünü oluştur
    # -----------------------------------------------------

    OUTPUT.mkdir(
        parents=True,
        exist_ok=True
    )


    sheet_index = []


    # -----------------------------------------------------
    # Her sheet'i ayrı JSON yap
    # -----------------------------------------------------

    for sheet in workbook.worksheets:

        sheet_name = sheet.title

        print(
            f"İşleniyor: {sheet_name}"
        )

        rows = sheet.iter_rows(
            values_only=True
        )

        try:
            header_values = next(rows)
        except StopIteration:
            header_values = []


        headers = [
            clean_value(value)
            for value in header_values
        ]


        data_rows = []


        for row in rows:

            cleaned = [
                clean_value(value)
                for value in row
            ]

            # Tamamen boş satırları at
            if not any(
                value != ""
                for value in cleaned
            ):
                continue

            # Başlık kadar sütun garanti et
            if len(cleaned) < len(headers):

                cleaned.extend(
                    [""] *
                    (
                        len(headers) -
                        len(cleaned)
                    )
                )

            # Fazla sütun varsa başlık sayısına göre kes
            elif len(cleaned) > len(headers):

                cleaned = cleaned[
                    :len(headers)
                ]

            data_rows.append(cleaned)


        # -------------------------------------------------
        # Sheet JSON
        # -------------------------------------------------

        sheet_data = {
            "name": sheet_name,
            "headers": headers,
            "rows": data_rows
        }


        filename = (
            safe_filename(sheet_name)
            + ".json"
        )

        output_file =OUTPUT / filename


        # UTF-8 + kompakt JSON
        with open(
            output_file,
            "w",
            encoding="utf-8"
        ) as file:

            json.dump(
                sheet_data,
                file,
                ensure_ascii=False,
                separators=(",", ":")
            )


        sheet_index.append({
            "name": sheet_name,
            "file": filename,
            "rows": len(data_rows),
            "columns": len(headers)
        })


        print(
            f"  -> {filename}"
        )

        print(
            f"  -> {len(data_rows):,} satır"
        )

        print(
            f"  -> {len(headers):,} sütun"
        )

        print()


    # -----------------------------------------------------
    # INDEX JSON
    # -----------------------------------------------------

    index_data = {
        "defaultSheet": (
            "TICKET"
            if any(
                item["name"] == "TICKET"
                for item in sheet_index
            )
            else (
                sheet_index[0]["name"]
                if sheet_index
                else ""
            )
        ),
        "sheets": sheet_index
    }


    with open(
        OUTPUT / "index.json",
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            index_data,
            file,
            ensure_ascii=False,
            separators=(",", ":")
        )


    workbook.close()


    print()
    print("=" * 60)
    print("TAMAMLANDI")
    print("=" * 60)
    print()
    print(
        f"{len(sheet_index)} sheet oluşturuldu."
    )
    print()
    print(
        f"Çıktı klasörü: {OUTPUT.resolve()}"
    )
    print()
    print(
        "Varsayılan sheet: "
        + index_data["defaultSheet"]
    )
    print()


if __name__ == "__main__":
    main()