#!/usr/bin/env python3
"""
Areal lingvistika fani bo'yicha mustaqil ta'lim soatlari
uchun uslubiy ko'rsatma — DOCX formatida.
"""
import zipfile, os

output = "/projects/sandbox/Desirtatsiya/Areal_lingvistika_Mustaqil_talim_uslubiy_korsatma.docx"

def esc(t):
    return t.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;")

def p(text, b=False, i=False, c=False, h1=False, h2=False, ni=False):
    pp=""
    if h1: pp='<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
    elif h2: pp='<w:pPr><w:pStyle w:val="Heading2"/></w:pPr>'
    elif c: pp='<w:pPr><w:jc w:val="center"/><w:ind w:firstLine="0"/></w:pPr>'
    elif ni: pp='<w:pPr><w:ind w:firstLine="0"/></w:pPr>'
    rp=""
    if b and i: rp="<w:rPr><w:b/><w:i/></w:rPr>"
    elif b: rp="<w:rPr><w:b/></w:rPr>"
    elif i: rp="<w:rPr><w:i/></w:rPr>"
    return f'<w:p>{pp}<w:r>{rp}<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'

def mp(parts, c=False, ni=False):
    pp=""
    if c: pp='<w:pPr><w:jc w:val="center"/><w:ind w:firstLine="0"/></w:pPr>'
    elif ni: pp='<w:pPr><w:ind w:firstLine="0"/></w:pPr>'
    runs=""
    for text, bold, italic in parts:
        rp=""
        if bold and italic: rp="<w:rPr><w:b/><w:i/></w:rPr>"
        elif bold: rp="<w:rPr><w:b/></w:rPr>"
        elif italic: rp="<w:rPr><w:i/></w:rPr>"
        runs+=f'<w:r>{rp}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'
    return f'<w:p>{pp}{runs}</w:p>'

E='<w:p><w:pPr><w:ind w:firstLine="0"/><w:spacing w:after="0"/></w:pPr></w:p>'
PB='<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

B=[]


# ====== TITUL VARAG'I ======
for _ in range(3): B.append(E)
B.append(p("O'ZBEKISTON RESPUBLIKASI", b=True, c=True))
B.append(p("OLIY TA'LIM, FAN VA INNOVATSIYALAR VAZIRLIGI", b=True, c=True))
B.append(E)
B.append(p("_________________________ UNIVERSITETI", b=True, c=True))
B.append(p("FILOLOGIYA FAKULTETI", b=True, c=True))
B.append(p("O'ZBEK TILSHUNOSLIGI KAFEDRASI", b=True, c=True))
for _ in range(3): B.append(E)
B.append(p('"TASDIQLAYMAN"', b=True, ni=True))
B.append(p("Kafedra mudiri: _______________", ni=True))
B.append(p('"___" _____________ 2025-yil', ni=True))
for _ in range(3): B.append(E)
B.append(p("AREAL LINGVISTIKA", b=True, c=True))
B.append(p("fani bo'yicha", c=True))
B.append(E)
B.append(p("TALABALARNING MUSTAQIL TA'LIM SOATLARINI", b=True, c=True))
B.append(p("O'TKAZISH BO'YICHA", b=True, c=True))
B.append(p("USLUBIY KO'RSATMA", b=True, c=True))
for _ in range(4): B.append(E)
B.append(p("Mutaxassislik: 60111300 — Filologiya va tillarni o'qitish (o'zbek tili)", ni=True))
B.append(p("Ta'lim bosqichi: Bakalavr", ni=True))
B.append(p("Kurs: III-IV", ni=True))
for _ in range(4): B.append(E)
B.append(p("Tuzuvchi: ________________________", ni=True))
for _ in range(3): B.append(E)
B.append(p("Toshkent — 2025", b=True, c=True))
B.append(PB)


# ====== MUNDARIJA ======
B.append(p("MUNDARIJA", h1=True))
B.append(E)
toc = [
    ("Kirish","3"),
    ("I. Mustaqil ta'limning maqsad va vazifalari","5"),
    ("II. Mustaqil ta'lim mavzulari va soatlar taqsimoti","7"),
    ("III. Mustaqil ta'lim topshiriqlari va ularni bajarish bo'yicha ko'rsatmalar","10"),
    ("  1-mavzu. Areal lingvistikaning predmeti va vazifalari","10"),
    ("  2-mavzu. Areal lingvistikaning shakllanish tarixi","12"),
    ("  3-mavzu. Lingvistik geografiya va dialektologiya","14"),
    ("  4-mavzu. Til aloqalari va til ittifoqlari","16"),
    ("  5-mavzu. Til areallarini aniqlash metodlari","18"),
    ("  6-mavzu. Turkiy tillar areali","20"),
    ("  7-mavzu. O'zbek tili dialektlarining areal xususiyatlari","22"),
    ("  8-mavzu. Leksik izoglossalar va ularning tahlili","24"),
    ("  9-mavzu. Fonetik va grammatik izoglossalar","26"),
    ("  10-mavzu. Markaziy Osiyo til ittifoqi","28"),
    ("IV. Mustaqil ta'lim ishlarini baholash mezonlari","30"),
    ("V. Tavsiya etiladigan adabiyotlar ro'yxati","32"),
    ("Ilovalar","34"),
]
for item, pg in toc:
    B.append(p(f"{item} {'.' * max(3, 55-len(item))} {pg}", ni=True))
B.append(PB)


# ====== KIRISH ======
B.append(p("KIRISH", h1=True))
B.append(E)
B.append(p("Ushbu uslubiy ko'rsatma \"Areal lingvistika\" fani bo'yicha talabalarning mustaqil ta'lim soatlarini samarali tashkil etish maqsadida ishlab chiqilgan. Uslubiy ko'rsatmada mustaqil ta'limning maqsad va vazifalari, mavzular rejasi, har bir mavzu bo'yicha topshiriqlar va ularni bajarish tartibi, baholash mezonlari hamda tavsiya etiladigan adabiyotlar ro'yxati keltirilgan."))
B.append(p("Areal lingvistika — tilshunoslikning muhim sohalaridan biri bo'lib, u tillarning geografik tarqalishi, o'zaro ta'siri, til aloqalari natijasida yuzaga keladigan umumiy xususiyatlar va izoglossalarni o'rganadi. Bu fan tillarning tarixiy rivojlanishini, dialektlarning shakllanishini va til ittifoqlarining yuzaga kelish qonuniyatlarini tushunishda muhim ahamiyatga ega."))
B.append(p("O'zbekiston Respublikasi Prezidentining ta'lim sohasidagi islohotlarga qaratilgan farmon va qarorlari talabalarning mustaqil fikrlash qobiliyatini rivojlantirish, ilmiy-tadqiqot ko'nikmalarini shakllantirish zaruratini belgilab bergan. Shu munosabat bilan, mustaqil ta'lim oliy ta'lim tizimining ajralmas va muhim tarkibiy qismi sifatida alohida e'tibor talab etadi."))
B.append(p("Mustaqil ta'lim — talabaning o'qituvchi rahbarligida yoki mustaqil ravishda bilim olish, ko'nikma va malakalarni shakllantirish jarayoni bo'lib, u talabaning ilmiy salohiyatini oshirish, mustaqil fikrlash va tadqiqot olib borish qobiliyatlarini rivojlantirishga qaratilgan."))
B.append(p("Ushbu uslubiy ko'rsatma \"Areal lingvistika\" fanining o'quv dasturi asosida tuzilgan bo'lib, har bir mustaqil ta'lim mavzusi bo'yicha aniq topshiriqlar, ularni bajarish bo'yicha uslubiy tavsiyalar va baholash mezonlari keltirilgan. Bu talabaga mustaqil ta'lim jarayonini rejali va samarali tashkil etishda yordam beradi."))
B.append(PB)


# ====== I. MAQSAD VA VAZIFALAR ======
B.append(p("I. MUSTAQIL TA'LIMNING MAQSAD VA VAZIFALARI", h1=True))
B.append(E)
B.append(p("Mustaqil ta'limning maqsadi:", b=True))
B.append(p("— talabalarning areal lingvistika sohasidagi nazariy bilimlarini mustahkamlash va chuqurlashtirish;"))
B.append(p("— talabalarning mustaqil ilmiy-tadqiqot olib borish ko'nikmalarini shakllantirish;"))
B.append(p("— lingvistik xaritalar bilan ishlash malakasini oshirish;"))
B.append(p("— til areallarini tahlil qilish va izoglossalarni aniqlash bo'yicha amaliy ko'nikmalarni rivojlantirish;"))
B.append(p("— ilmiy adabiyotlar bilan ishlash va ularni tanqidiy tahlil qilish qobiliyatini shakllantirish;"))
B.append(p("— o'zbek tili dialektlari va turkiy tillar arealini mustaqil o'rganish ko'nikmalarini hosil qilish."))
B.append(E)
B.append(p("Mustaqil ta'limning vazifalari:", b=True))
B.append(p("— talabalar ma'ruza va amaliy mashg'ulotlarda olgan bilimlarini mustaqil ta'lim topshiriqlarini bajarish orqali mustahkamlaydilar;"))
B.append(p("— ilmiy manbalar, darsliklar va qo'shimcha adabiyotlarni mustaqil o'rganadilar;"))
B.append(p("— tayanch tushuncha va atamalarni o'zlashtirib, ulardan ilmiy nutqda foydalanish ko'nikmalarini shakllantiradilar;"))
B.append(p("— lingvistik xaritalar tuzish va tahlil qilish amaliyotini bajараdilar;"))
B.append(p("— referat, taqdimot, esse va ilmiy maqola yozish ko'nikmalarini rivojlantiradilar;"))
B.append(p("— guruh va individual tadqiqot loyihalarini amalga oshiradilar."))
B.append(E)
B.append(p("Mustaqil ta'lim natijasida talaba:", b=True))
B.append(p("— areal lingvistikaning asosiy tushunchalari, metodlari va nazariyalarini bilishi kerak;"))
B.append(p("— lingvistik xaritalar bilan ishlash va ularni tahlil qilish malakasiga ega bo'lishi kerak;"))
B.append(p("— til areallarini aniqlash va izoglossalarni belgilash ko'nikmasini egallashi kerak;"))
B.append(p("— ilmiy adabiyotlarni tanqidiy tahlil qilish va ilmiy yozuv uslubida fikr bayon etish qobiliyatiga ega bo'lishi kerak;"))
B.append(p("— olingan bilimlarni amaliy tadqiqotlarda qo'llay olishi kerak."))
B.append(PB)


# ====== II. MAVZULAR VA SOATLAR TAQSIMOTI ======
B.append(p("II. MUSTAQIL TA'LIM MAVZULARI VA SOATLAR TAQSIMOTI", h1=True))
B.append(E)
B.append(p("Jami mustaqil ta'lim soatlari: 30 soat", b=True))
B.append(E)
# Table header
B.append(p("T/r | Mavzu nomi | Soat | Topshiriq turi", b=True, ni=True))
B.append(p("————————————————————————————————————————————————————————————————", ni=True))
table_rows = [
    ("1", "Areal lingvistikaning predmeti, maqsadi va vazifalari", "3", "Referat"),
    ("2", "Areal lingvistikaning shakllanish tarixi va rivojlanish bosqichlari", "3", "Taqdimot"),
    ("3", "Lingvistik geografiya va dialektologiya munosabati", "3", "Konspekt"),
    ("4", "Til aloqalari va til ittifoqlari nazariyasi", "3", "Esse"),
    ("5", "Til areallarini aniqlash metodlari va tamoyillari", "3", "Amaliy ish"),
    ("6", "Turkiy tillar areali va ularning o'ziga xos xususiyatlari", "3", "Referat"),
    ("7", "O'zbek tili dialektlarining areal xususiyatlari", "4", "Loyiha"),
    ("8", "Leksik izoglossalar va ularning lingvistik xaritalarda ifodalanishi", "3", "Amaliy ish"),
    ("9", "Fonetik va grammatik izoglossalar tahlili", "3", "Taqqoslash jadvali"),
    ("10", "Markaziy Osiyo til ittifoqi va uning xususiyatlari", "2", "Ilmiy maqola"),
]
for num, mavzu, soat, turi in table_rows:
    B.append(p(f"  {num}.  {mavzu}  |  {soat}  |  {turi}", ni=True))
B.append(p("————————————————————————————————————————————————————————————————", ni=True))
B.append(p("  JAMI:  30 soat", b=True, ni=True))
B.append(PB)


# ====== III. TOPSHIRIQLAR ======
B.append(p("III. MUSTAQIL TA'LIM TOPSHIRIQLARI VA ULARNI BAJARISH BO'YICHA KO'RSATMALAR", h1=True))
B.append(E)

# --- MAVZU 1 ---
B.append(p("1-MAVZU. AREAL LINGVISTIKANING PREDMETI VA VAZIFALARI", h2=True))
B.append(p("Mustaqil ta'lim soati: 3 soat", b=True, ni=True))
B.append(p("Topshiriq turi: Referat", b=True, ni=True))
B.append(E)
B.append(p("Topshiriq mazmuni:", b=True))
B.append(p("Talaba quyidagi mavzulardan birini tanlab, 8-10 betlik referat yozishi kerak:"))
B.append(p("1. Areal lingvistikaning tilshunoslikdagi o'rni va boshqa fanlar bilan aloqasi."))
B.append(p("2. Areal lingvistikaning predmeti, maqsadi va asosiy tushunchalari."))
B.append(p("3. Areal lingvistika va sotsiolingvistika: umumiy va farqli jihatlari."))
B.append(p("4. Areal lingvistikaning zamonaviy tilshunoslikdagi ahamiyati."))
B.append(E)
B.append(p("Referatni bajarish tartibi:", b=True))
B.append(p("1. Mavzu bo'yicha kamida 5-7 ta ilmiy manbalarni o'rganish."))
B.append(p("2. Reja tuzish (kirish, asosiy qism — 2-3 bo'lim, xulosa)."))
B.append(p("3. Ilmiy uslubda bayon etish, iqtiboslar va havolalar ko'rsatish."))
B.append(p("4. Foydalanilgan adabiyotlar ro'yxatini ilova qilish."))
B.append(p("5. Referat hajmi: 8-10 bet (Times New Roman, 14pt, 1.5 interval)."))
B.append(E)
B.append(p("Nazorat savollari:", b=True))
B.append(p("1. Areal lingvistika qanday masalalarni o'rganadi?"))
B.append(p("2. Areal lingvistikaning asosiy tushunchalari qaysilar?"))
B.append(p("3. Bu fan qanday boshqa tilshunoslik sohalari bilan bog'liq?"))
B.append(p("4. Areal lingvistikaning amaliy ahamiyati nimada?"))
B.append(E)
B.append(p("Tayanch tushunchalar:", i=True))
B.append(p("areal, izoglossa, lingvistik xarita, til aloqasi, til ittifoqi, dialekt, sheva, substrat, superstrat, adstrat.", i=True))
B.append(E)

# --- MAVZU 2 ---
B.append(p("2-MAVZU. AREAL LINGVISTIKANING SHAKLLANISH TARIXI", h2=True))
B.append(p("Mustaqil ta'lim soati: 3 soat", b=True, ni=True))
B.append(p("Topshiriq turi: Taqdimot (prezentatsiya)", b=True, ni=True))
B.append(E)
B.append(p("Topshiriq mazmuni:", b=True))
B.append(p("Talaba areal lingvistikaning shakllanish tarixi va rivojlanish bosqichlari haqida 15-20 slaydli taqdimot tayyorlashi kerak."))
B.append(E)
B.append(p("Taqdimotda quyidagi masalalar yoritilishi kerak:", b=True))
B.append(p("1. Areal lingvistikaning paydo bo'lish sabablari va sharт-sharoitlari."))
B.append(p("2. G.Venker, J.Jileron va boshqa olimlarning lingvistik geografiya sohasidagi ishlari."))
B.append(p("3. Neogrammatiklar va ularning dialektologik tadqiqotlari."))
B.append(p("4. XX asrda areal lingvistikaning rivojlanishi (Praga lingvistik to'garagi, N.S.Trubetskoy)."))
B.append(p("5. Hozirgi zamon areal tadqiqotlari va ularning yo'nalishlari."))
B.append(E)
B.append(p("Taqdimotni tayyorlash bo'yicha ko'rsatmalar:", b=True))
B.append(p("— har bir slaydda matn minimal bo'lishi, asosiy fikrlar tezis shaklida berilishi kerak;"))
B.append(p("— rasmlar, xaritalar, jadvallar va sxemalardan foydalanish tavsiya etiladi;"))
B.append(p("— taqdimot oxirida foydalanilgan manbalar ko'rsatilishi shart;"))
B.append(p("— taqdimot vaqti: 10-12 daqiqa."))
B.append(E)

# --- MAVZU 3 ---
B.append(p("3-MAVZU. LINGVISTIK GEOGRAFIYA VA DIALEKTOLOGIYA", h2=True))
B.append(p("Mustaqil ta'lim soati: 3 soat", b=True, ni=True))
B.append(p("Topshiriq turi: Konspekt", b=True, ni=True))
B.append(E)
B.append(p("Topshiriq mazmuni:", b=True))
B.append(p("Talaba quyidagi mavzular bo'yicha ilmiy adabiyotlardan konspekt tayyorlashi kerak:"))
B.append(p("1. Lingvistik geografiyaning predmeti va vazifalari."))
B.append(p("2. Dialektologiya va lingvistik geografiya o'rtasidagi munosabat."))
B.append(p("3. Lingvistik atlaslar: turlari, tuzilish tamoyillari."))
B.append(p("4. O'zbek dialektologiyasida lingvistik geografiya usullarining qo'llanilishi."))
B.append(E)
B.append(p("Konspekt tuzish bo'yicha ko'rsatmalar:", b=True))
B.append(p("— konspekt kamida 2 ta ilmiy manbadan tayyorlanishi kerak;"))
B.append(p("— asosiy tushunchalar, ta'riflar, misollar ajratib ko'rsatilishi lozim;"))
B.append(p("— konspekt hajmi: 5-6 bet;"))
B.append(p("— o'z fikr-mulohazalari alohida qayd etilishi kerak."))
B.append(E)


# --- MAVZU 4 ---
B.append(p("4-MAVZU. TIL ALOQALARI VA TIL ITTIFOQLARI", h2=True))
B.append(p("Mustaqil ta'lim soati: 3 soat", b=True, ni=True))
B.append(p("Topshiriq turi: Esse", b=True, ni=True))
B.append(E)
B.append(p("Topshiriq mazmuni:", b=True))
B.append(p("Talaba quyidagi mavzulardan birini tanlab, 4-5 betlik ilmiy esse yozishi kerak:"))
B.append(p('1. "Til ittifoqi tushunchasi va uning areal lingvistikadagi o\'rni".'))
B.append(p('2. "Substrat, superstrat va adstrat hodisalari: misollar tahlili".'))
B.append(p('3. "Til aloqalarining leksik va grammatik darajalardagi namoyon bo\'lishi".'))
B.append(p('4. "Balqon til ittifoqi va Markaziy Osiyo til ittifoqi: qiyosiy tahlil".'))
B.append(E)
B.append(p("Esseni yozish bo'yicha ko'rsatmalar:", b=True))
B.append(p("— esse erkin bayon shaklidagi ilmiy ish bo'lib, unda muallif o'z nuqtai nazarini ilmiy dalillar asosida bayon etadi;"))
B.append(p("— esseda kirish (muammo qo'yilishi), asosiy qism (tahlil va dalillar) va xulosa bo'lishi kerak;"))
B.append(p("— kamida 3-4 ta ilmiy manbaga murojaat qilinishi lozim;"))
B.append(p("— esse hajmi: 4-5 bet (Times New Roman, 14pt, 1.5 interval)."))
B.append(E)

# --- MAVZU 5 ---
B.append(p("5-MAVZU. TIL AREALLARINI ANIQLASH METODLARI", h2=True))
B.append(p("Mustaqil ta'lim soati: 3 soat", b=True, ni=True))
B.append(p("Topshiriq turi: Amaliy ish", b=True, ni=True))
B.append(E)
B.append(p("Topshiriq mazmuni:", b=True))
B.append(p("Talaba quyidagi amaliy ishlarni bajarishi kerak:"))
B.append(p("1. Berilgan til materiallariga asoslanib, sodda lingvistik xarita tuzish."))
B.append(p("2. Xaritada izoglossalarni belgilash va ularni izohlash."))
B.append(p("3. Areal identifikatsiyasi usullarini qo'llash (leksik, fonetik, grammatik mezonlar bo'yicha)."))
B.append(E)
B.append(p("Amaliy ishni bajarish tartibi:", b=True))
B.append(p("1-qadam: O'zbek tilining 3-4 ta dialekt guruhidan (qorluq, qipchoq, o'g'uz) bir xil ma'noli so'zlarni tanlash (masalan, 'non', 'suv', 'bola' so'zlarining dialektal variantlari)."))
B.append(p("2-qadam: O'zbekiston xaritasi ustiga ushbu so'zlarning tarqalish hududini belgilash."))
B.append(p("3-qadam: Izoglossa chiziqlarini chizish — bir xil xususiyatga ega hududlarni ajratish."))
B.append(p("4-qadam: Olingan natijalarni tahlil qilish va yozma hisobot tayyorlash (3-4 bet)."))
B.append(E)
B.append(p("Tayanch tushunchalar:", i=True))
B.append(p("izoglossa, lingvistik xarita, areal, dialekt, sheva, geolingvistika, leksik differensiatsiya.", i=True))
B.append(E)

# --- MAVZU 6 ---
B.append(p("6-MAVZU. TURKIY TILLAR AREALI", h2=True))
B.append(p("Mustaqil ta'lim soati: 3 soat", b=True, ni=True))
B.append(p("Topshiriq turi: Referat", b=True, ni=True))
B.append(E)
B.append(p("Topshiriq mazmuni:", b=True))
B.append(p("Talaba quyidagi mavzulardan birini tanlab referat yozishi kerak:"))
B.append(p("1. Turkiy tillar oilasining geografik tarqalishi va areal tasnifi."))
B.append(p("2. Turkiy tillar orasidagi areal aloqalar va ularning lingvistik xaritalarda aks etishi."))
B.append(p("3. O'zbek tili va qo'shni turkiy tillar (qozoq, qirg'iz, turkman) o'rtasidagi areal munosabatlar."))
B.append(p("4. Turkiy tillarning fors-tojik va arab tillari bilan areal kontaktlari."))
B.append(E)
B.append(p("Referatga qo'yiladigan talablar:", b=True))
B.append(p("— hajmi: 8-10 bet;"))
B.append(p("— kamida 5 ta ilmiy manbadan foydalanish;"))
B.append(p("— konkret misollar va jadvallar keltirish;"))
B.append(p("— ilmiy uslubda yozish, havolalar va iqtiboslar to'g'ri rasmiylashtirish."))
B.append(E)


# --- MAVZU 7 ---
B.append(p("7-MAVZU. O'ZBEK TILI DIALEKTLARINING AREAL XUSUSIYATLARI", h2=True))
B.append(p("Mustaqil ta'lim soati: 4 soat", b=True, ni=True))
B.append(p("Topshiriq turi: Loyiha ishi", b=True, ni=True))
B.append(E)
B.append(p("Topshiriq mazmuni:", b=True))
B.append(p("Talaba o'zbek tili dialektlarining areal xususiyatlarini o'rganish bo'yicha mini-loyiha tayyorlashi kerak."))
B.append(E)
B.append(p("Loyiha bosqichlari:", b=True))
B.append(p("1-bosqich. Tadqiqot mavzusini tanlash:", b=True))
B.append(p("  a) biror viloyat yoki tuman shevalarining leksik xususiyatlari;"))
B.append(p("  b) qorluq, qipchoq yoki o'g'uz lahjalarining muayyan xususiyatini tahlil qilish;"))
B.append(p("  c) ikki qo'shni hududdagi shevalар orasidagi farqlarni aniqlash."))
B.append(E)
B.append(p("2-bosqich. Material to'plash:", b=True))
B.append(p("  — ilmiy adabiyotlardan (dialektologik atlaslar, monografiyalar, maqolalar);"))
B.append(p("  — imkoniyat bo'lsa, jonli nutq namunalaridan (oilaviy muhit, kuzatish)."))
B.append(E)
B.append(p("3-bosqich. Tahlil va natijalarni rasmiylashtirish:", b=True))
B.append(p("  — to'plangan materiallarni tizimlashtirish;"))
B.append(p("  — xarita ustida dialekt xususiyatlarini belgilash;"))
B.append(p("  — 6-8 betlik yozma hisobot tayyorlash."))
B.append(E)
B.append(p("4-bosqich. Taqdimot:", b=True))
B.append(p("  — loyiha natijalarini 10-15 daqiqalik taqdimot shaklida guruhga taqdim etish."))
B.append(E)

# --- MAVZU 8 ---
B.append(p("8-MAVZU. LEKSIK IZOGLOSSALAR VA ULARNING TAHLILI", h2=True))
B.append(p("Mustaqil ta'lim soati: 3 soat", b=True, ni=True))
B.append(p("Topshiriq turi: Amaliy ish", b=True, ni=True))
B.append(E)
B.append(p("Topshiriq mazmuni:", b=True))
B.append(p("Talaba o'zbek tili dialektlaridagi leksik farqlanishlarni tahlil qilish bo'yicha amaliy ish bajarishi kerak."))
B.append(E)
B.append(p("Bajarish tartibi:", b=True))
B.append(p("1. Quyidagi leksik-semantik guruhlardan birini tanlash:"))
B.append(p("   a) qarindoshlik atamalari (ota, ona, aka, opa — turli dialektlardagi variantlari);"))
B.append(p("   b) ovqat nomlari (non, palov, manti — dialektal nomlanishlar);"))
B.append(p("   c) tabiat hodisalari (yomg'ir, qor, shamol — dialektal variantlar);"))
B.append(p("   d) kiyim-kechak nomlari (ko'ylak, shim, to'n — turli hududlardagi nomlanishlar)."))
B.append(E)
B.append(p("2. Tanlangan guruh bo'yicha kamida 10-15 ta so'zning turli dialektlardagi variantlarini jadval shaklida yozish."))
B.append(p("3. Leksik izoglossalarni xaritada belgilash."))
B.append(p("4. Tahliliy xulosa yozish (2-3 bet)."))
B.append(E)
B.append(p("Namuna jadval:", b=True, ni=True))
B.append(p("So'z | Adabiy shakl | Qorluq lahjasi | Qipchoq lahjasi | O'g'uz lahjasi", b=True, ni=True))
B.append(p("1. Non | non | non | nan | etmak", ni=True))
B.append(p("2. Suv | suv | suv | suw | su", ni=True))
B.append(p("3. Bola | bola | bola | bala | oglan", ni=True))
B.append(E)

# --- MAVZU 9 ---
B.append(p("9-MAVZU. FONETIK VA GRAMMATIK IZOGLOSSALAR", h2=True))
B.append(p("Mustaqil ta'lim soati: 3 soat", b=True, ni=True))
B.append(p("Topshiriq turi: Taqqoslash jadvali va tahlil", b=True, ni=True))
B.append(E)
B.append(p("Topshiriq mazmuni:", b=True))
B.append(p("Talaba o'zbek tili dialektlaridagi fonetik va grammatik farqlanishlarni taqqoslab, jadval tuzishi va tahliliy xulosa yozishi kerak."))
B.append(E)
B.append(p("A. Fonetik izoglossalar bo'yicha jadval:", b=True))
B.append(p("1. Unli tovushlar tizimidagi farqlar (sungiy, kenglik, lablanish bo'yicha)."))
B.append(p("2. Undosh tovushlardagi o'zgarishlar (masalan, j/y, k/x almashuvlari)."))
B.append(p("3. Assimilyatsiya va dissimilyatsiya hodisalari."))
B.append(E)
B.append(p("B. Grammatik izoglossalar bo'yicha jadval:", b=True))
B.append(p("1. Fe'l shakllaridagi farqlar (masalan, hozirgi zamon fe'l qo'shimchalari)."))
B.append(p("2. Ot-kesim qo'shimchalaridagi farqlar."))
B.append(p("3. So'z tartibidagi farqlanishlar."))
B.append(E)
B.append(p("Topshiriqni bajarish tartibi:", b=True))
B.append(p("1. Har bir izoglossa turi bo'yicha kamida 5 ta misol topish."))
B.append(p("2. Misollarni jadval shaklida rasmiylashtirish."))
B.append(p("3. Izoglossa xaritasini chizish."))
B.append(p("4. Tahliliy xulosa yozish (3-4 bet)."))
B.append(E)


# --- MAVZU 10 ---
B.append(p("10-MAVZU. MARKAZIY OSIYO TIL ITTIFOQI", h2=True))
B.append(p("Mustaqil ta'lim soati: 2 soat", b=True, ni=True))
B.append(p("Topshiriq turi: Ilmiy maqola loyihasi", b=True, ni=True))
B.append(E)
B.append(p("Topshiriq mazmuni:", b=True))
B.append(p("Talaba Markaziy Osiyo til ittifoqi mavzusida kichik ilmiy maqola (5-6 bet) yozishi kerak."))
B.append(E)
B.append(p("Maqolada quyidagi masalalar yoritilishi kerak:", b=True))
B.append(p("1. Markaziy Osiyo til ittifoqi tushunchasi va uni isbotlovchi dalillar."))
B.append(p("2. O'zbek, tojik, qozoq, qirg'iz va turkman tillaridagi umumiy xususiyatlar."))
B.append(p("3. Fonetik, leksik va grammatik darajadagi umumiyliklar."))
B.append(p("4. Til ittifoqining shakllanish sabablari (tarixiy, geografik, madaniy omillar)."))
B.append(p("5. Markaziy Osiyo til ittifoqini boshqa til ittifoqlari bilan qiyoslash."))
B.append(E)
B.append(p("Ilmiy maqola yozish bo'yicha ko'rsatmalar:", b=True))
B.append(p("— maqola tuzilishi: kirish, asosiy qism, xulosa, adabiyotlar;"))
B.append(p("— ilmiy uslubda yozish, havolalar va iqtiboslar to'g'ri rasmiylashtirish;"))
B.append(p("— kamida 5-7 ta ilmiy manbaga murojaat qilish;"))
B.append(p("— konkret til materiallari va misollar keltirish;"))
B.append(p("— hajmi: 5-6 bet (Times New Roman, 14pt, 1.5 interval)."))
B.append(PB)

# ====== IV. BAHOLASH MEZONLARI ======
B.append(p("IV. MUSTAQIL TA'LIM ISHLARINI BAHOLASH MEZONLARI", h1=True))
B.append(E)
B.append(p("Mustaqil ta'lim ishlari quyidagi mezonlar asosida baholanadi:"))
B.append(E)
B.append(p("A'lo baho (86-100 ball) uchun:", b=True))
B.append(p("— mavzu to'liq va chuqur yoritilgan;"))
B.append(p("— ilmiy manbalardan samarali foydalanilgan (5 va undan ortiq manba);"))
B.append(p("— mustaqil fikrlash va tahlil qobiliyati namoyon etilgan;"))
B.append(p("— ish to'g'ri rasmiylashtirmilgan (talablar to'liq bajarilgan);"))
B.append(p("— konkret misollar va dalillar keltirilgan;"))
B.append(p("— ish belgilangan muddatda topshirilgan."))
B.append(E)
B.append(p("Yaxshi baho (71-85 ball) uchun:", b=True))
B.append(p("— mavzu yetarli darajada yoritilgan;"))
B.append(p("— ilmiy manbalardan foydalanilgan (3-4 manba);"))
B.append(p("— tahlil va mustaqil fikrlash elementlari mavjud;"))
B.append(p("— rasmiylashtirish talablari asosan bajarilgan;"))
B.append(p("— ba'zi kamchiliklar mavjud, lekin ular muhim emas."))
B.append(E)
B.append(p("Qoniqarli baho (55-70 ball) uchun:", b=True))
B.append(p("— mavzu qisman yoritilgan;"))
B.append(p("— ilmiy manbalardan cheklangan darajada foydalanilgan (1-2 manba);"))
B.append(p("— mustaqil fikrlash va tahlil etarli emas;"))
B.append(p("— rasmiylashtirish talablarida sezilarli kamchiliklar mavjud;"))
B.append(p("— ish kechikib topshirilgan."))
B.append(E)
B.append(p("Qoniqarsiz baho (55 balldan past) uchun:", b=True))
B.append(p("— mavzu yoritilmagan yoki noto'g'ri yoritilgan;"))
B.append(p("— ilmiy manbalardan foydalanilmagan;"))
B.append(p("— plagiat holatlari aniqlangan;"))
B.append(p("— ish rasmiylashtirish talablariga javob bermaydi;"))
B.append(p("— ish topshirilmagan."))
B.append(E)
B.append(p("Mustaqil ta'lim ishlarining umumiy bahoda ulushi:", b=True))
B.append(p("— Joriy baholashning 30% ni tashkil etadi."))
B.append(p("— Har bir mustaqil ta'lim ishi alohida baholanadi va o'rtacha ball hisoblanadi."))
B.append(PB)


# ====== V. ADABIYOTLAR ======
B.append(p("V. TAVSIYA ETILADIGAN ADABIYOTLAR RO'YXATI", h1=True))
B.append(E)
B.append(p("Asosiy adabiyotlar:", b=True))
B.append(E)
B.append(p("1. Ashirboev S. O'zbek dialektologiyasi. — Toshkent: Fan, 2019. — 240 b.", ni=True))
B.append(p("2. Begmatov E., Ulug'ov N. O'zbek tilining areal leksikasi. — Toshkent: Fan, 2015. — 192 b.", ni=True))
B.append(p("3. Borovkov A.K. O'zbek tili dialektlarining leksikasi. — Toshkent: Fan, 1963. — 168 b.", ni=True))
B.append(p("4. Ishoqov F.G. O'zbek dialektologiyasi. — Toshkent: O'qituvchi, 2014. — 224 b.", ni=True))
B.append(p("5. Mirzayev M. O'zbek tili shevalarining areal xususiyatlari. — Toshkent: Fan, 2010. — 180 b.", ni=True))
B.append(p("6. Rasulov R. O'zbek dialektologiyasi asoslari. — Toshkent: Fan va texnologiyalar, 2018. — 256 b.", ni=True))
B.append(p("7. Reshetov V.V. O'zbek dialektologiyasi. — Toshkent: O'qituvchi, 1966. — 320 b.", ni=True))
B.append(p("8. Shodmonov Q. O'zbek tilining areal tadqiqi. — Samarqand: SamDU, 2012. — 144 b.", ni=True))
B.append(E)
B.append(p("Qo'shimcha adabiyotlar:", b=True))
B.append(E)
B.append(p("9. Baskakov N.A. Turkiy tillar tasnifi. — Moskva: Nauka, 1988. — 248 s.", ni=True))
B.append(p("10. Deshiriyev Yu.D. Sotsiolingvistika asoslari. — Moskva: Nauka, 1977. — 382 s.", ni=True))
B.append(p("11. Johanson L. Turkic Language Contacts. — Amsterdam: John Benjamins, 2002. — 420 p.", ni=True))
B.append(p("12. Menges K. Turkiy tillarning areal aloqalari. — Wiesbaden: Harrassowitz, 1995. — 340 s.", ni=True))
B.append(p("13. Serebrennikov B.A. Umumiy tilshunoslik. — Moskva: Nauka, 1970. — 604 s.", ni=True))
B.append(p("14. Trubetskoy N.S. Fonologiya asoslari. — Moskva: Inostrannaya literatura, 1960. — 372 s.", ni=True))
B.append(p("15. Weinreich U. Languages in Contact. — The Hague: Mouton, 1953. — 148 p.", ni=True))
B.append(E)
B.append(p("Internet resurslari:", b=True))
B.append(E)
B.append(p("16. www.ziyonet.uz — O'zbekiston ta'lim portali", ni=True))
B.append(p("17. www.natlib.uz — O'zbekiston Milliy kutubxonasi", ni=True))
B.append(p("18. www.wals.info — World Atlas of Language Structures", ni=True))
B.append(p("19. www.ethnologue.com — Dunyo tillari ma'lumotnomasi", ni=True))
B.append(p("20. www.glottolog.org — Tillar va dialektlar ma'lumotlar bazasi", ni=True))
B.append(PB)

# ====== ILOVALAR ======
B.append(p("ILOVALAR", h1=True))
B.append(E)
B.append(p("1-ilova. Mustaqil ta'lim ishlarini rasmiylashtirish namunasi", b=True))
B.append(E)
B.append(p("Referat titul varag'i namunasi:", b=True, ni=True))
B.append(E)
B.append(p("_________________________ UNIVERSITETI", c=True))
B.append(p("Filologiya fakulteti", c=True))
B.append(p("O'zbek tilshunosligi kafedrasi", c=True))
B.append(E)
B.append(p("AREAL LINGVISTIKA fanidan", c=True))
B.append(p("MUSTAQIL TA'LIM ISHI", b=True, c=True))
B.append(E)
B.append(p("Mavzu: ________________________________________", c=True))
B.append(E)
B.append(p("Bajardi: _________________ guruh talabasi", ni=True))
B.append(p("F.I.Sh.: _________________________________", ni=True))
B.append(p("Tekshirdi: ________________________________", ni=True))
B.append(E)
B.append(p("Toshkent — 2025", c=True))
B.append(E)
B.append(E)
B.append(p("2-ilova. Mustaqil ta'lim ishlarini topshirish jadvali", b=True))
B.append(E)
B.append(p("Mavzu raqami | Topshirish muddati | Topshiriq turi | Ball", b=True, ni=True))
B.append(p("1-mavzu | 3-hafta | Referat | 10 ball", ni=True))
B.append(p("2-mavzu | 5-hafta | Taqdimot | 10 ball", ni=True))
B.append(p("3-mavzu | 6-hafta | Konspekt | 10 ball", ni=True))
B.append(p("4-mavzu | 8-hafta | Esse | 10 ball", ni=True))
B.append(p("5-mavzu | 9-hafta | Amaliy ish | 10 ball", ni=True))
B.append(p("6-mavzu | 11-hafta | Referat | 10 ball", ni=True))
B.append(p("7-mavzu | 13-hafta | Loyiha | 15 ball", ni=True))
B.append(p("8-mavzu | 14-hafta | Amaliy ish | 10 ball", ni=True))
B.append(p("9-mavzu | 15-hafta | Jadval + tahlil | 10 ball", ni=True))
B.append(p("10-mavzu | 16-hafta | Ilmiy maqola | 15 ball", ni=True))
B.append(E)
B.append(E)
B.append(p("3-ilova. Lingvistik xarita tuzish bo'yicha ko'rsatma", b=True))
B.append(E)
B.append(p("Lingvistik xarita tuzishda quyidagi qoidalarga amal qilish lozim:"))
B.append(p("1. Xarita miqyosi aniq ko'rsatilishi kerak."))
B.append(p("2. Har bir til hodisasi uchun alohida belgi (simvol) tanlash va uning izohi (legenda) berilishi shart."))
B.append(p("3. Izoglossa chiziqlari aniq va ravshan chizilishi kerak."))
B.append(p("4. Xaritada viloyat, tuman nomlari ko'rsatilishi lozim."))
B.append(p("5. Xarita tagida manba ko'rsatilishi shart."))
B.append(p("6. Izoglossalar turli ranglar yoki chiziq turlari (uzluksiz, nuqtali, chiziqli) bilan farqlanishi kerak."))
B.append(p("7. Xarita A4 formatda tayyorlanishi tavsiya etiladi."))
B.append(E)
B.append(E)
B.append(p("4-ilova. Areal lingvistika bo'yicha asosiy atamalar lug'ati", b=True))
B.append(E)
B.append(mp([("Areal ",True,False),("— muayyan til hodisasi tarqalgan geografik hudud.",False,False)], ni=True))
B.append(mp([("Izoglossa ",True,False),("— bitta til xususiyatining tarqalish chegarasini ko'rsatuvchi chiziq.",False,False)], ni=True))
B.append(mp([("Lingvistik xarita ",True,False),("— til hodisalarining geografik tarqalishini ko'rsatuvchi xarita.",False,False)], ni=True))
B.append(mp([("Til ittifoqi (Sprachbund) ",True,False),("— geografik yaqinlik natijasida umumiy xususiyatlarga ega bo'lgan tillar guruhi.",False,False)], ni=True))
B.append(mp([("Substrat ",True,False),("— yangi kelgan til tomonidan siqib chiqarilgan mahalliy tilning iz-qoldiqlari.",False,False)], ni=True))
B.append(mp([("Superstrat ",True,False),("— mahalliy aholi tilini siqib chiqara olmagan bosqinchilar tilining iz-qoldiqlari.",False,False)], ni=True))
B.append(mp([("Adstrat ",True,False),("— ikki qo'shni tilning o'zaro ta'siri (ikkalasi ham saqlanib qolgan holda).",False,False)], ni=True))
B.append(mp([("Dialekt ",True,False),("— tilning muayyan hududda qo'llaniladigan mahalliy ko'rinishi.",False,False)], ni=True))
B.append(mp([("Sheva ",True,False),("— dialektning kichik hududiy bo'linmasi.",False,False)], ni=True))
B.append(mp([("Lingvistik geografiya ",True,False),("— til hodisalarining geografik tarqalishini o'rganuvchi fan.",False,False)], ni=True))
B.append(mp([("Geolingvistika ",True,False),("— lingvistik geografiyaning zamonaviy nomi.",False,False)], ni=True))
B.append(mp([("Dialektologik atlas ",True,False),("— biror til dialektlarining xaritalar to'plami.",False,False)], ni=True))
B.append(mp([("Konvergensiya ",True,False),("— tillarning o'zaro yaqinlashishi, umumiy xususiyatlar hosil bo'lishi.",False,False)], ni=True))
B.append(mp([("Divergensiya ",True,False),("— tillarning o'zaro uzoqlashishi, farqlanish jarayoni.",False,False)], ni=True))
B.append(mp([("Kreolizatsiya ",True,False),("— pidjin tilining kreol tilga aylanishi jarayoni.",False,False)], ni=True))


# ====== ASSEMBLE DOCX ======
content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''

rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

word_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''

styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/>
      <w:sz w:val="28"/><w:szCs w:val="28"/>
      <w:lang w:val="uz-Latn-UZ"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:line="360" w:lineRule="auto" w:after="0" w:before="0"/>
      <w:jc w:val="both"/>
      <w:ind w:firstLine="709"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:spacing w:before="240" w:after="120"/><w:jc w:val="center"/><w:ind w:firstLine="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:pPr><w:spacing w:before="200" w:after="100"/><w:jc w:val="center"/><w:ind w:firstLine="0"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
  </w:style>
</w:styles>'''

settings = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="708"/>
</w:settings>'''

sect_props = '''<w:sectPr>
  <w:pgSz w:w="11906" w:h="16838"/>
  <w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/>
  <w:cols w:space="708"/>
</w:sectPr>'''

body_xml = ''.join(B)
document = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    {body_xml}
    {sect_props}
  </w:body>
</w:document>'''

with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as zf:
    zf.writestr('[Content_Types].xml', content_types)
    zf.writestr('_rels/.rels', rels)
    zf.writestr('word/_rels/document.xml.rels', word_rels)
    zf.writestr('word/document.xml', document)
    zf.writestr('word/styles.xml', styles)
    zf.writestr('word/settings.xml', settings)

file_size = os.path.getsize(output)
print(f"DOCX yaratildi!")
print(f"Fayl: {output}")
print(f"Hajmi: {file_size:,} bayt ({file_size/1024:.1f} KB)")
print(f"Jami paragraflar: {len(B)}")
