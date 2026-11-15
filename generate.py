# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, r"C:\Users\ANCL\Desktop\20 gun")
exec(open(r"C:\Users\ANCL\Desktop\20 gun\_data.py", encoding="utf-8").read())

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

header_fill = PatternFill("solid", fgColor="111827")
header_font = Font(bold=True, color="B8963E", size=11)
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical="top")

VAXT = "18:00-21:00"
AUDIO = "Səs + fon musiqi"
SUBTITR = "Açar sözlər"
KANAL = "DM (mesaj)"

# ---- File 1 ----
wb1 = openpyxl.Workbook()
ws1 = wb1.active
ws1.title = "Postplan - General"
headers1 = ["Tarix", "Həftə", "Platforma", "Vaxt", "Audio", "Subtitr", "Satış kanalı", "Məzmun növü", "Mövzu", "Başlıq (Hook)", "Vizual", "Səs (rəvayət)", "Ekranda mətn", "Çağırış (CTA)", "Hashtag-lar", "Məqsəd", "Status"]
for c, h in enumerate(headers1, start=1):
    cell = ws1.cell(row=1, column=c, value=h); cell.fill = header_fill; cell.font = header_font
    cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center"); cell.border = border
widths1 = [12, 7, 10, 14, 22, 16, 16, 16, 26, 34, 40, 50, 26, 30, 34, 20, 16]
for i, w in enumerate(widths1, start=1):
    ws1.column_dimensions[get_column_letter(i)].width = w
ws1.row_dimensions[1].height = 30
ws1.freeze_panes = "A2"

def week_of(d_str):
    d, m, y = d_str.split(".")
    from datetime import date
    return date(int(y), int(m), int(d)).isocalendar()[1]

r = 2
for p in posts1:
    row = [p['date'], week_of(p['date']), "TikTok", p.get('vaxt', VAXT), p.get('audio', AUDIO), p.get('subtitr', SUBTITR), p.get('kanal', KANAL), p['type'], p['topic'], p['hook'], p['visual'], p['speech'], p['onscreen'], p['cta'], p['hashtags'], p['goal'], p['status']]
    for c, v in enumerate(row, start=1):
        cell = ws1.cell(row=r, column=c, value=v); cell.alignment = wrap; cell.border = border
    ws1.cell(row=r, column=1).number_format = "dd.mm.yyyy"
    r += 1
ws1.auto_filter.ref = f"A1:Q{r-1}"
rng1 = f"Q2:Q{r-1}"
for val, fg, fc in [("Planlaşdırılıb","374151","D1D5DB"),("Hazırlanır","78350F","FDE68A"),("Yayımlandı","065F46","A7F3D0"),("Ləğv edildi","7F1D1D","FECACA")]:
    ws1.conditional_formatting.add(rng1, CellIsRule(operator='equal', formula=[f'"{val}"'], fill=PatternFill("solid", fgColor=fg), font=Font(color=fc)))

ws1p = wb1.create_sheet("Parametrlər")
ws1p.column_dimensions['A'].width = 34; ws1p.column_dimensions['B'].width = 52
params1 = [("Kampaniya","20 günlük dil öyrənmə - TikTok kampaniyası"),("Başlanğıc tarixi","15.11.2026"),("Bitmə tarixi","28.02.2027"),("Platforma","TikTok"),("Hədəf auditoriya","Geniş auditoriya"),("Ton","Motivasiyalı"),("Əsas məqsəd","Kitab satışı"),("Tezlik","Həftədə 3-4 post"),("Məzmun formatı","Qarışıq (qısa + orta + uzun)"),("Yayım vaxtı","18:00-21:00"),("Audio","Səs + fon musiqi"),("Subtitr","Açar sözlər"),("Satış kanalı","DM (mesaj)"),("Reklam","Yalnız orqanik"),("Bonus","Yoxdur"),("Zəmanət","Yoxdur"),("Vizual stil","Minimalist"),("Telegram kanalı","Yox"),("Landing səhifə","Yox"),("Email marketinq","Yox"),("Canlı efir (live)","Daha sonra"),("Çağırış (challenge)","Yox"),("Kolaborasiya","Yox"),("Dillər","İngilis + Rus")]
for i,(k,v) in enumerate(params1, start=1):
    a = ws1p.cell(row=i, column=1, value=k); a.font = Font(bold=True); a.alignment = wrap; a.border = border
    b = ws1p.cell(row=i, column=2, value=v); b.alignment = wrap; b.border = border

ws1h = wb1.create_sheet("Hashtag-lar")
ws1h.column_dimensions['A'].width = 26; ws1h.column_dimensions['B'].width = 60
hrows = [("Əsas","#20gun #diloyrenme #ingilisdili #rusdili #kitab"),("Motivasiya","#motivasiya #bacararsan #hedef #ugur"),("Təhsil","#tehsil #meslehet #oyrenme #metodika"),("Sosial sübut","#ugur #hekaye #netice #transformasiya"),("Satış","#satış #kitab #sifaris #endirim"),("Brend","#muellif #brend #kitab #azerbaijan")]
for c,h in enumerate(["Kateqoriya","Hashtag-lar"], start=1):
    cell = ws1h.cell(row=1, column=c, value=h); cell.fill = header_fill; cell.font = header_font; cell.border = border
for i,(k,v) in enumerate(hrows, start=2):
    ws1h.cell(row=i, column=1, value=k).border = border
    ws1h.cell(row=i, column=2, value=v).border = border

# ---- File 2 ----
wb2 = openpyxl.Workbook()
ws2 = wb2.active
ws2.title = "Postplan - Xarakter"
headers2 = ["Gün", "Tarix", "Platforma", "Vaxt", "Audio", "Subtitr", "Satış kanalı", "Məzmun növü", "Mövzu", "Başlıq (Hook)", "Vizual", "Səs (rəvayət)", "Ekranda mətn", "Çağırış (CTA)", "Hashtag-lar", "Məqsəd", "Status"]
for c, h in enumerate(headers2, start=1):
    cell = ws2.cell(row=1, column=c, value=h); cell.fill = header_fill; cell.font = header_font
    cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center"); cell.border = border
widths2 = [6, 12, 10, 14, 22, 16, 16, 16, 26, 34, 40, 50, 26, 30, 34, 20, 16]
for i, w in enumerate(widths2, start=1):
    ws2.column_dimensions[get_column_letter(i)].width = w
ws2.row_dimensions[1].height = 30
ws2.freeze_panes = "A2"

r = 2
for p in char_posts:
    row = [p['day'], p['date'], "TikTok", p.get('vaxt', VAXT), p.get('audio', AUDIO), p.get('subtitr', SUBTITR), p.get('kanal', KANAL), p['type'], p['topic'], p['hook'], p['visual'], p['speech'], p['onscreen'], p['cta'], p['hashtags'], p['goal'], p['status']]
    for c, v in enumerate(row, start=1):
        cell = ws2.cell(row=r, column=c, value=v); cell.alignment = wrap; cell.border = border
    ws2.cell(row=r, column=2).number_format = "dd.mm.yyyy"
    r += 1
ws2.auto_filter.ref = f"A1:Q{r-1}"
rng2 = f"Q2:Q{r-1}"
for val, fg, fc in [("Planlaşdırılıb","374151","D1D5DB"),("Hazırlanır","78350F","FDE68A"),("Yayımlandı","065F46","A7F3D0"),("Ləğv edildi","7F1D1D","FECACA")]:
    ws2.conditional_formatting.add(rng2, CellIsRule(operator='equal', formula=[f'"{val}"'], fill=PatternFill("solid", fgColor=fg), font=Font(color=fc)))

ws2p = wb2.create_sheet("Parametrlər")
ws2p.column_dimensions['A'].width = 34; ws2p.column_dimensions['B'].width = 52
params2 = [("Kampaniya","20 günlük dil öyrənmə - Xarakter səyahəti"),("Başlanğıc tarixi","15.11.2026"),("Bitmə tarixi","06.12.2026"),("Platforma","TikTok"),("Xarakter","22 yaşlı tələbə qız"),("Məqsəd","İngilis dilini karyera üçün öyrənmək"),("Öyrənilən dil","İngilis"),("Müddət","20 gün"),("Tezlik","Hər gün 1 post"),("Yayım vaxtı","18:00-21:00"),("Audio","Səs + fon musiqi"),("Subtitr","Açar sözlər"),("Satış kanalı","DM (mesaj)"),("Reklam","Yalnız orqanik"),("Vizual stil","Minimalist"),("Canlı efir (live)","Daha sonra"),("Çağırış (challenge)","Yox"),("Kolaborasiya","Yox")]
for i,(k,v) in enumerate(params2, start=1):
    a = ws2p.cell(row=i, column=1, value=k); a.font = Font(bold=True); a.alignment = wrap; a.border = border
    b = ws2p.cell(row=i, column=2, value=v); b.alignment = wrap; b.border = border

ws2c = wb2.create_sheet("Xarakter profili")
ws2c.column_dimensions['A'].width = 26; ws2c.column_dimensions['B'].width = 60
prof = [("Ad","Aysel (nümunə)"),("Yaş","22"),("Status","Tələbə"),("Məqsəd","Karyera üçün ingilis dili"),("Başlanğıc səviyyəsi","A1 (başlanğıc)"),("Hədəf səviyyə","B1 (orta)"),("Gündəlik vaxt","20 dəqiqə"),("Metod","20 günlük metodika"),("Nəticə","20 gün sonra B1 səviyyə")]
for c,h in enumerate(["Parametr","Dəyər"], start=1):
    cell = ws2c.cell(row=1, column=c, value=h); cell.fill = header_fill; cell.font = header_font; cell.border = border
for i,(k,v) in enumerate(prof, start=2):
    ws2c.cell(row=i, column=1, value=k).border = border
    ws2c.cell(row=i, column=2, value=v).border = border

wb1.save(r"C:\Users\ANCL\Desktop\20 gun\Postplan_General.xlsx")
wb2.save(r"C:\Users\ANCL\Desktop\20 gun\Postplan_Xarakter.xlsx")
print("Saved Excel files")