# 📜 VBA Word Contract Generator & Corporate Entity Automator

Automatyczne narzędzie w **Excel VBA**, które generuje gotowe do podpisu drafty umów w formacie **Microsoft Word (.docx)** dla grupy 12 podmiotów powiązanych, wykorzystując wspólny szablon bazowy.

---

## 🎯 Problem Biznesowy
Zarządzanie dokumentacją prawną dla 12 różnych podmiotów w ramach jednej grupy kapitałowej generowało ogromne ryzyko błędów. Ręczne uzupełnianie danych rejestrowych, reprezentacji, kontaktów oraz podmienianie logotypów w Wordzie było czasochłonne, a zmiana układu stron często powodowała „rozjeżdżanie się” struktury załączników.

## 💡 Rozwiązanie
Formularz w Excelu zoptymalizowany pod kątem automatyzacji Worda, który po wybraniu konkretnego podmiotu z listy rozwijanej wykonuje pełny montaż dokumentu:

1. **Branding & Nagłówki:** Dynamicznie wstawia dedykowane logo podmiotu w nagłówku każdej strony dokumentu Word.
2. **Synchronizacja Metadanych:** Automatycznie wstrzykuje w odpowiednie zakładki/placeholdery:
   * Numery i daty umów głównych oraz umów z operatorem systemów dystrybucyjnych (OSD).
   * Dane rejestrowe: KRS, REGON, NIP.
   * Dane reprezentacji (Prezes Zarządu składający podpis).
   * Dedykowane adresy e-mail (operacyjne oraz do e-faktur).
3. **Formatowanie Strukturalne:** Automatycznie wstawia twarde podziały stron przed każdym załącznikiem, gwarantując idealną spójność wizualną tekstu i braki przesunięć formatowania.

---

## 🛠️ Stos Technologiczny
* **Środowisko:** Microsoft Excel + Microsoft Word Automation
* **Język:** VBA (Visual Basic for Applications)
* **Biblioteki:** Word Object Library (`Word.Application`, `Word.Document`)
* **Mechanizmy Worda:** Bookmarks / Find & Replace, Header Sections, Page Breaks

---

## 🏗️ Architektura Przepływu

```text
[ User Selects Entity from Dropdown (1 of 12) ]
                      │
                      ▼
[ Execute VBA Macro (Excel -> Word Automation) ]
                      │
   ┌──────────────────┼──────────────────┐
   ▼                  ▼                  ▼
[ Inject Logo ]  [ Replace Text ]  [ Apply Page Breaks ]
(Headers)        (KRS, NIP, Rep)   (Attachments)
   │                  │                  │
   └──────────────────┼──────────────────┘
                      ▼
[ Generated Word Draft (.docx) Ready for Signature ]
