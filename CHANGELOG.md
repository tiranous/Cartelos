# Changelog

## [0.1.0] - 2024-03-18
### Προσθήκες
- Αναδιοργάνωση της λογικής σε modules (`Module_Engine`, `Module_IO`, `Module_Utils`, `Module_Errors`, `Module_Logging`).
- Προσθήκη μηχανισμού logging και test harness (`RunAllTests`).
- Προσθήκη `Option Explicit` σε όλα τα modules και τυποποιημένο error handling.

### Αλλαγές
- Τα γεγονότα στα φύλλα περιορίστηκαν σε λεπτούς χειριστές που καλούν τον κινητήρα.
- Καθαρισμός διπλού/νεκρού κώδικα και κατάργηση `Select/Activate`.

### Διορθώσεις
- Προστασία της κατάστασης της εφαρμογής (ScreenUpdating/Calculation/Events) πριν και μετά την εκτέλεση.
- Δημιουργία φύλλων `Logs` και `Tests` όταν δεν υπάρχουν.
