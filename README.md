# 🇸🇰 ms_word_report_changes – Verzia 2.0

## 🏴 English version below 🇸🇰 / 🇬🇧 **[Switch to English](#english-version)**

Makro **ExportToExcelUltraFast** exportuje všetky **sledované zmeny** a **komentáre vrátane odpovedí** z dokumentu Word do tabuľky v Exceli.  
Nová verzia automaticky opravuje Parent Comment ID podľa kapitoly a odstavca/obrázka.

---

## **Ako vyzerá výsledná tabuľka v Exceli?**

| **Autor**  | **Dátum**  | **Typ**    | **Obsah**                  | **Kapitola**                     | **Odstavec/Obrázok**     | **Strana** | **Comment ID** | **Parent Comment ID** |
|------------|------------|------------|-----------------------------|----------------------------------|--------------------------|------------|----------------|------------------------|
| J. Novak   | 2024-02-14 | Zmena      | "Aktualizovaný text..."     | 2.1 Schvaľovací proces           | "Predošlý text v odstavci..." | 5 |   |   |
| P. Kováč   | 2024-02-13 | Komentár   | "Treba upraviť formátovanie." | 3.2 Výstupy                    | "Obrázok: Diagram schémy (strana 5)" | 8 | 5 |   |
| M. Horváth | 2024-02-13 | Reakcia    | "Doplnené podľa odporúčania." | 3.2 Výstupy                   | "Obrázok: Diagram schémy (strana 5)" | 8 | 6 | 5 |

---

## **Čo je nové vo verzii 2.0?**

- 🔥 Automatická korekcia **Parent Comment ID** aj pri veľkých dokumentoch
- 🚀 **Ultra-fast** režim (aj pri 1000+ komentároch)
- 🛠 Nastaviteľný parameter spätného hľadania `MaxBackwardSearch`
- 💾 Predvolený názov Excel súboru `Exported_Changes_YYYYMMDD_HHMM.xlsx`
- 🗂 Dvojjazyčné komentáre v kóde (🇸🇰 / 🇬🇧)

---

## **Ako spustiť makro?**

1. **Otvorte Word** a stlačte `ALT + F11`.
2. **Vložte nový modul** (Insert > Module).
3. **Importujte alebo vložte kód** z `ExportToExcelUltraFast.bas`.
4. **Spustite makro `ExportToExcelUltraFast`**.

✅ Výsledný Excel súbor sa uloží automaticky.

---

## **[🇬🇧 Switch to English](#english-version)**

---

# 🇬🇧 English version

## 🇬🇧 **[Prepnúť na slovenčinu](#ms_word_report_changes--verzia-20)**

The **ExportToExcelUltraFast** macro exports all **tracked changes** and **comments including replies** from a Word document into an Excel table.  
The new version automatically corrects Parent Comment IDs based on the chapter and paragraph/image context.

---

## **What does the Excel output look like?**

| **Author** | **Date**    | **Type**  | **Content**                | **Chapter**                   | **Paragraph/Image**        | **Page** | **Comment ID** | **Parent Comment ID** |
|------------|-------------|-----------|-----------------------------|--------------------------------|-----------------------------|----------|----------------|------------------------|
| J. Novak   | 2024-02-14  | Change    | "Updated text..."           | 2.1 Approval Process           | "Previous paragraph text..." | 5        |                |                        |
| P. Kovac   | 2024-02-13  | Comment   | "Formatting needs adjustment." | 3.2 Outputs                  | "Image: Diagram scheme (page 5)" | 8    | 5              |                        |
| M. Horvath | 2024-02-13  | Reply     | "Updated according to suggestion." | 3.2 Outputs             | "Image: Diagram scheme (page 5)" | 8    | 6              | 5                      |

---

## **What's new in version 2.0?**

- 🔥 Automatic correction of **Parent Comment IDs** even in large documents
- 🚀 **Ultra-fast** processing (even with 1000+ comments)
- 🛠 Configurable backward search via `MaxBackwardSearch`
- 💾 Default Excel filename `Exported_Changes_YYYYMMDD_HHMM.xlsx`
- 🗂 Dual-language (🇸🇰 / 🇬🇧) code comments

---

## **How to run the macro?**

1. **Open Word** and press `ALT + F11`.
2. **Insert a new module** (Insert > Module).
3. **Import or paste** the code from `ExportToExcelUltraFast.bas`.
4. **Run the macro `ExportToExcelUltraFast`**.

✅ The resulting Excel file will be saved automatically.

