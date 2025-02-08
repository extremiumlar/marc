from pymarc import MARCReader
import pandas as pd

# ISO formatdagi MARC fayli
iso_file = "library_records.iso"
mrc_file = "library_records.mrc"
output_file = "library_records.xlsx"

# ISO faylni .mrc formatiga o‘tkazish
with open(iso_file, "rb") as iso_f, open(mrc_file, "wb") as mrc_f:
    reader = MARCReader(iso_f)
    for record in reader:
        mrc_f.write(record.as_marc())

print(f"Fayl {iso_file} dan {mrc_file} ga muvaffaqiyatli o‘tkazildi!")

# .mrc faylni .xlsx formatiga o‘tkazish
records_list = []

with open(mrc_file, "rb") as marc_file:
    reader = MARCReader(marc_file)
    for record in reader:
        title = record.get('245', {}).get('a', "Noma'lum")
        author = record.get('100', {}).get('a', "Noma'lum")
        publisher = record.get('260', {}).get('b', "Noma'lum")
        year = record.get('260', {}).get('c', "Noma'lum")
        isbn = record.get('020', {}).get('a', "Noma'lum")

        records_list.append([title, author, publisher, year, isbn])

# DataFrame yaratish va Excel faylga saqlash
df = pd.DataFrame(records_list, columns=["Sarlavha", "Muallif", "Nashriyot", "Nashr yili", "ISBN"])
df.to_excel(output_file, index=False)

print(f"✅ Fayl muvaffaqiyatli {output_file} ga o‘tkazildi!")