# Slovenská verzia

## 🏴 English version below 🇸🇰 / 🇬🇧 [Switch to English](#english-version)

Makro na export **sledovaných zmien** a **komentárov** z dokumentov MS Word do Excelu. Výstupom je prehľadná tabuľka s informáciami o autoroch, dátumoch, typoch zmien, obsahu, kapitole, odstavci, strane a štruktúrou vlákien komentárov.

---

## 🔧 Hlavné funkcie

- 📑 Exportuje **všetky sledované zmeny a komentáre**
- 📆 Zaznamenáva **dátum a čas** úprav
- 🧠 Identifikuje **najbližší nadpis a odstavec** alebo obrázok
- 📌 Zobrazuje číslo strany (voliteľne – pre rýchlejší režim vypnite)
- 💬 Vytvára prehľad komentárov a **vlákien odpovedí** (Parent ID)
- 🔁 **Automatická korekcia Parent ID** pomocou spätného porovnávania
- 📉 **Minimalizuje Excel aj VBA** editor počas behu makra
- ⚙️ Efektívne hospodári s pamäťou – vhodné aj pre **veľké dokumenty**
- 📊 Možnosť spustenia v **rýchlom režime** bez strán pre extrémny výkon

---

## 🖼 Ukážka výstupu (Excel)

| **Autor**  | **Dátum**        | **Typ**     | **Obsah**                 | **Kapitola**             | **Odstavec/Obrázok**       | **Strana** | **ID** | **Parent ID** |
|------------|------------------|-------------|----------------------------|--------------------------|-----------------------------|------------|--------|---------------|
| J. Novak   | 2024-02-14 14:02 | Zmena       | „Aktualizovaný text...“   | 2.1 Proces               | „Text pred zmenou...“       | 5          |        |               |
| P. Kováč   | 2024-02-14 14:05 | Komentár    | „Treba to preformulovať.“ | 2.1 Proces               | „Obrázok: Diagram (str. 5)“ | 5          | 12     |               |
| A. Svitek  | 2024-02-14 14:06 | Reakcia     | „Súhlasím.“                | 2.1 Proces               | „Obrázok: Diagram (str. 5)“ | 5          | 13     | 12            |

---

## ⚙️ Ako spustiť makro?

1. Otvorte Word a stlačte `ALT + F11`
2. Vložte nový modul (Insert → Module)
3. Skopírujte kód makra (pozri súbor `ExportToExcelUltraFast.vba`)
4. Spustite `ExportToExcelUltraFast`

---

## 🛠 Technické poznámky

- Makro používa **WinAPI funkcie** na minimalizáciu okien (`FindWindowA`, `ShowWindow`)
- Nepotrebuje žiadne externé knižnice
- Funguje v MS Word 2010+ na Windows (VBA 7 aj staršie)

---

✅ Výsledný Excel súbor sa uloží po dokončení exportu automaticky.

---

## **Nastavenie parametrov v kóde / Macro parameters**

| Parameter | Predvolená hodnota | Popis 🇸🇰 | Description 🇬🇧 |
|:----------|:--------------------|:---------|:---------------|
| `FastMode` | `True` | Ak `True`, čísla strán sa nevypisujú pre rýchlejší export. Ak `False`, dopĺňajú sa aj čísla strán. | If `True`, page numbers are not exported (faster). If `False`, page numbers are included. |
| `StatusUpdateFrequency` | `500` | Počet položiek medzi aktualizáciami stavového riadku. | How many items between status bar updates. |
| `MaxBackwardSearch` | `50` | Maximálny počet riadkov pri spätnom hľadaní Parent Comment ID. | Maximum number of rows to search backwards for Parent Comment ID. |

---

## **Priebežný stav spracovania / Progress tracking**

- ✅ Počas spracovania dokumentu sa priebežne aktualizuje **stavový riadok Wordu** s počtom spracovaných zmien a komentárov.
- ✅ **Excel je minimalizovaný** počas spracovania.
- ✅ **VBA okno je minimalizované** počas spracovania.

---

# English version

## 🇸🇰 **[Prepnúť na slovenčinu](#Slovenská-verzia)**

This macro exports **tracked changes** and **comments** from MS Word documents into a structured Excel table. Output includes author, date, type, content, chapter, paragraph/image, page number, comment ID, and parent comment references.

---

## 🔧 Features

- 📑 Exports **all tracked changes and comments**
- 📆 Captures **date and time**
- 🧠 Finds nearest **heading and paragraph or image**
- 📌 Optionally shows page number (disable for fast mode)
- 💬 Tracks **comment threads** with Parent ID
- 🔁 **Auto-corrects missing Parent IDs**
- 📉 Minimizes both **Excel and VBA editor** while running
- ⚙️ Efficient memory usage – suitable for **large documents**
- 🚀 Optional **fast mode** without page numbers for speed

---

## 🖼 Sample Excel output

| **Author** | **Date**         | **Type**  | **Content**               | **Chapter**            | **Paragraph/Image**       | **Page** | **ID** | **Parent ID** |
|------------|------------------|-----------|----------------------------|------------------------|----------------------------|----------|--------|---------------|
| J. Novak   | 2024-02-14 14:02 | Change    | "Updated text..."          | 2.1 Process            | "Previous paragraph text"  | 5        |        |               |
| P. Kováč   | 2024-02-14 14:05 | Comment   | "Needs rephrasing."        | 2.1 Process            | "Image: Diagram (page 5)"  | 5        | 12     |               |
| A. Svitek  | 2024-02-14 14:06 | Reply     | "Agreed."                  | 2.1 Process            | "Image: Diagram (page 5)"  | 5        | 13     | 12            |

---

## ⚙️ How to run the macro?

1. Open Word and press `ALT + F11`
2. Insert a new module (Insert → Module)
3. Paste the macro code (see `ExportToExcelUltraFast.vba`)
4. Run `ExportToExcelUltraFast`

---

## 🛠 Technical Notes

- Uses **WinAPI functions** (`FindWindowA`, `ShowWindow`) to manage windows
- No external libraries needed
- Compatible with MS Word 2010+ (VBA 6 / 7)

---

✅ Excel file with results saves automatically after data export is finished.

---

## **Macro parameters (in code)**

| Parameter | Default Value | Description |
|:----------|:---------------|:------------|
| `FastMode` | `True` | If `True`, page numbers are skipped for faster processing. |
| `StatusUpdateFrequency` | `500` | How many items between status bar updates. |
| `MaxBackwardSearch` | `50` | Maximum number of rows to search backwards to find Parent Comment ID. |

---

## **Progress tracking**

- ✅ Word's status bar shows **number of processed changes and comments** live during processing.
- ✅ **Excel remains minimalized** during processing.
- ✅ **VBA window remains minimalized** during processing.

---

## 🇸🇰 **[Prepnúť na slovenčinu](#Slovenská-verzia)**
