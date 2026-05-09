# 🔧 Baza Narzędzi Dynamometrycznych (VBA Tool Database)

Profesjonalne i w pełni zautomatyzowane rozwiązanie oparte na **MS Excel (VBA)** do zarządzania bazą danych narzędzi dynamometrycznych (kluczy/wkrętaków). Projekt umożliwia błyskawiczne i bezbłędne rejestrowanie nowych narzędzi, z naciskiem na zaawansowaną walidację danych "w locie" i świetne doświadczenie użytkownika (UX).

![Wersja](https://img.shields.io/badge/Wersja-1.0-blue)
![Język](https://img.shields.io/badge/Język-VBA%20%7C%20VBScript-brightgreen)
![Środowisko](https://img.shields.io/badge/Środowisko-MS%20Excel-orange)

## 📖 O projekcie

Aplikacja ma na celu usprawnienie procesu wprowadzania narzędzi z rygorystycznymi zakresami pracy (min/max) oraz tolerancjami procentowymi, co ma kluczowe znaczenie w kontroli jakości procesów produkcyjnych. 

Dzięki sprytnemu plikowi rozruchowemu `.vbs`, cały uciążliwy interfejs programu MS Excel jest ukrywany w tle, przez co użytkownik otrzymuje **czystą, przypominającą natywną aplikację formę (UserForm)** umiejscowioną zawsze na środku ekranu.

## ✨ Kluczowe funkcje

- **🚀 Tryb "Kiosk" (Standalone):** Dzięki plikowi `StartDodawanieNarzedziTYP3.vbs` interfejs Excela zostaje całkowicie ukryty w tle (`Visible = False`), a proces kończy się bezpiecznie przy użyciu `Application.Quit`.
- **🎨 Dynamiczny UX i podświetlanie (Focus Tracking):** Aktywne pole tekstowe automatycznie podświetla się na delikatny, pastelowy zielony kolor, ułatwiając nawigację. Zablokowane pola pozostają dyskretnie zablokowane. Przyciski zyskują kontrastowe kolory (cyjanowy, koralowy) w momencie wejścia w Focus.
- **🛡️ Bezpieczna walidacja "w locie":**
  - Automatyczne powiększanie liter (Uppercase) dla nazw.
  - Ostrzeganie o duplikatach *zanim* użytkownik wypełni resztę danych – makro po wciśnięciu `Enter` szuka w tle numeru narzędzia i blokuje przejście dalej.
  - Sprawdzanie i formatowanie przecinków, walidacja błędnych znaków (blokada liter w polach liczbowych).
  - Weryfikacja logiczna: czy `Tolerancja %` mieści się w dopuszczalnym zakresie (np. 1-30%) oraz czy wartość `MAX` nie jest przypadkiem mniejsza niż `MIN`.
- **⚡ Nawigacja wyłącznie klawiaturą (Enter-Flow):** Płynne przechodzenie z pola do pola z podświetlaniem następnego "KROKU" i płynne uaktywnianie przycisku Zapisz – myszka jest wręcz niepotrzebna.


👨‍💻 Technologie
VBA (Visual Basic for Applications) – Logika biznesowa i Formularze.
VBScript – Bezproblemowy start i integracja powłoki systemu operacyjnego (ukryty instancjonowany obiekt Excel).
