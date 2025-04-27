# **🇸🇰 Slovenská-verzia**

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
- 🛠 Nastaviteľné správanie pomocou parametrov v kóde
- 📈 Priebežný stav spracovania zobrazený v stavovom riadku Wordu
- 🛡 Excel je počas spracovania skrytý – otvorí sa až po ukončení
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

## **Nastavenie parametrov v kóde / Macro parameters**

| Parameter | Predvolená hodnota | Popis 🇸🇰 | Description 🇬🇧 |
|:----------|:--------------------|:---------|:---------------|
| `FastMode` | `True` | Ak `True`, čísla strán sa nevypisujú pre rýchlejší export. Ak `False`, dopĺňajú sa aj čísla strán. | If `True`, page numbers are not exported (faster). If `False`, page numbers are included. |
| `StatusUpdateFrequency` | `500` | Počet položiek medzi aktualizáciami stavového riadku. | How many items between status bar updates. |
| `MaxBackwardSearch` | `50` | Maximálny počet riadkov pri spätnom hľadaní Parent Comment ID. | Maximum number of rows to search backwards for Parent Comment ID. |

---

## **Priebežný stav spracovania / Progress tracking**

- ✅ Počas spracovania dokumentu sa priebežne aktualizuje **stavový riadok Wordu** s počtom spracovaných zmien a komentárov.
- ✅ **Excel sa nezobrazuje** počas spracovania – pre vyššiu rýchlosť exportu.
- ✅ **Excel sa otvorí automaticky** po dokončení exportu dát.


---


# **🇬🇧 English-version**

## 🇸🇰 **[Prepnúť na slovenčinu](#-slovensk%C3%A1-verzia)**

The **ExportToExcelUltraFast** macro exports all **tracked changes** and **comments including replies** from a Word document into an Excel table.  
The new version automatically corrects Parent Comment IDs based on chapter and paragraph/image context.

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
- 🛠 Behavior customizable through macro parameters
- 📈 Live progress shown in Word's status bar
- 🛡 Excel hidden during processing – shown after export
- 💾 Default Excel filename `Exported_Changes_YYYYMMDD_HHMM.xlsx`
- 🗂 Dual-language (🇸🇰 / 🇬🇧) code comments

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
- ✅ **Excel remains hidden** during processing – boosting speed.
- ✅ **Excel opens automatically** after data export is finished.

---

✅ **Thank you for using version 2.0!**

## 🇸🇰 **[Prepnúť na slovenčinu](#-slovensk%C3%A1-verzia)**
