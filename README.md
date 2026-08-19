# 📜 Word VBA Contract Draft Generator & Entity Automator

Automatyczne narzędzie zbudowane w całości w **Microsoft Word (VBA)**, które generuje gotowe do podpisu drafty umów w formacie `.docx` dla grupy 12 podmiotów powiązanych za pomocą interaktywnego formularza **UserForm**.

---

## 🎯 Problem Biznesowy
Ręczne edytowanie szablonów umów w Wordzie dla 12 różnych podmiotów w ramach grupy kapitałowej wiązało się z dużym ryzykiem błędów. Uzupełnianie danych rejestrowych, reprezentacji, dedykowanych adresów e-mail czy podmienianie logotypów w nagłówkach było czasochłonne. Dodatkowo dopisywanie treści często powodowało „rozjeżdżanie się” układu stron i załączników.

## 💡 Rozwiązanie
Formularz **VBA UserForm** wbudowany bezpośrednio w szablon Worda (`.dotm` / `.docx`), który prowadzi użytkownika krok po kroku:

1. **Interaktywny Wybór Podmiotu:** Po uruchomieniu makra pojawia się okienko dialogowe (UserForm), z którego użytkownik wybiera 1 z 12 podmiotów.
2. **Automatyczny Branding & Nagłówki:** Skrypt dynamicznie wstawia dedykowane logo wybranej spółki w nagłówku każdej strony dokumentu.
3. **Synchronizacja Danych i Zmiennych:** Automatycznie uzupełnia i wstrzykuje w odpowiednie miejsca:
   * Numery oraz daty umów głównych i umów z operatorem (TAURON).
   * Dane rejestrowe spółki: KRS, REGON, NIP.
   * Personalizację reprezentacji (Prezes Zarządu podpisujący umowę).
   * Dedykowane adresy e-mail (kontaktowe oraz do e-faktur).
4. **Formatowanie Strukturalne Załączników:** Makro wymusza rozpoczynanie każdego załącznika od nowej strony (twardy podział stron), zapobiegając przesuwaniu się tekstu i psuciu układu graficznego.

---

## 🛠️ Stos Technologiczny
* **Środowisko:** Microsoft Word
* **Język:** VBA (Visual Basic for Applications)
* **Komponenty:** UserForms, Word Bookmarks / Content Controls, Headers & Footers API, Page Break Engine

---

## 🏗️ Architektura Przepływu

```text
[ User Opens Word Template ]
             │
             ▼
[ VBA UserForm Pops Up ] ──> [ Select 1 of 12 Entities from Dropdown ]
             │
             ▼
[ Click "Generate Contract" ]
             │
   ┌─────────┼──────────────────┐
   ▼         ▼                  ▼
[ Inject ] [ Replace Text ] [ Apply Page Breaks ]
[ Logo   ] (KRS, NIP, Rep,  (Attachments forced
(Header)   TAURON Agreement) to New Pages)
   │         │                  │
   └─────────┼──────────────────┘
             ▼
[ Completed Word Draft Ready to Sign ]
