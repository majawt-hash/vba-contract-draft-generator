# 📜 VBA Contract Draft Generator & Corporate Data Sync

Automatyczne narzędzie w **Excel VBA**, które generuje dostosowane drafty umów handlowych na podstawie danych rejestrowych wybranej spółki, predefiniowanych szablonów oraz historii umów.

---

## 🎯 Problem Biznesowy
Ręczne przygotowywanie draftów umów dla wielu podmiotów/spółek córki wiązało się z ryzykiem pomyłek w numerach KRS, NIP, adresach rejestrowych oraz sygnaturach wcześniejszych porozumień. Proces był czasochłonny i wymagał każdorazowej weryfikacji danych formalnych.

## 💡 Rozwiązanie
Makro VBA wbudowane w arkusz menedżerski, które:
1. **Pobiera dane spółki:** Po wypraniu podmiotu z listy rozwijanej automatycznie uzupełnia dane rejestrowe, numery umów bazowych oraz powiązane logotypy.
2. **Generuje spersonalizowany dokument:** Dynamicznie podstawia zmienne do szablonu umowy.
3. **Standaryzuje proces:** Skraca czas przygotowania draftu z kilkunastu minut do kilku sekund, eliminując błędy literowe i formalne.

---

## 🛠️ Stos Technologiczny
* **Język:** VBA (Visual Basic for Applications)
* **Środowisko:** Microsoft Excel
* **Modyfikowane obiekty:** Excel Forms, Dynamic Shapes/Images, Text Placeholders

---

## ⚙️ Instrukcja Uruchomienia / Szablon
1. Otwórz plik `.xlsm` z włączoną obsługą makr.
2. Wybierz spółkę z listy rozwijanej w sekcji **"Formularz Umowy"**.
3. Kliknij przycisk **"Generuj Draft Umowy"**.
