# Sakarya Weekly Report Dashboards

Dashboard interattivi **per singolo report settimanale (DPR)** del progetto
Sakarya Gas Field Development – Phase 3, produzione **Tenaris Confab**
(Pindamonhangaba, Brasile). IQS S.r.l. obo Saipem.

> Diverso dal repository **`sakarya-dashboard`** (fanatics-hue.github.io/sakarya-dashboard),
> che è il dashboard *Overall Status* combinato. Qui ogni cartella è **un report** di una
> settimana, con il suo dashboard HTML autosufficiente.

🔗 **Pagina pubblica:** https://fanatics-hue.github.io/sakarya-weekly-reports/

## Struttura

```
index.html                                  landing che elenca i report (GitHub Pages)
SC26-3-SPM-CNF-DPR-XXX (WEEK NN)/
   ├─ SC26-3-SPM-CNF-DPR-XXX (WEEK NN).doc  report Word (NON versionato: gitignore)
   └─ SC26-3-SPM-CNF-DPR-XXX_WEEKNN_dashboard.html   dashboard pubblicato
prepara_report.py                           parte automatica del flusso (vedi sotto)
PREPARA_REPORT.bat / PUBBLICA.bat           wrapper 1-clic
```

Le cartelle sono organizzate **per disciplina**:
- **3LPP Coating** (serie DPR-013C → 014C → 017C → …)
- **Pipe Mill** (serie DPR-016 → 017 → 020 → …)

Il nome della cartella segue sempre **l'ultimo `.doc` caricato** dentro di essa.

## Flusso settimanale ("ogni volta con i nuovi file")

1. Metti il nuovo `.doc` del report nella cartella della **sua disciplina**
   (3LPP con 3LPP, Pipe Mill con Pipe Mill — non scambiarli).
2. Doppio clic su **`PREPARA_REPORT.bat`**. Lo script:
   - prende il `.doc` più recente di ogni cartella,
   - lo estrae in `.txt` (Word) così è leggibile,
   - rileva la disciplina,
   - rinomina la cartella in base all'ultimo `.doc`,
   - rigenera le card della `index.html`,
   - **segnala con `*` i dashboard da ricostruire** (settimana del `.doc` ≠ settimana del dashboard).
3. Apri Claude e chiedi *"ricostruisci i dashboard da ricostruire"*: legge il `.txt`
   già estratto e rigenera l'HTML del dashboard (contenuto narrativo: KPI, NCR, attendees,
   tabelle test). **Questo passo lo fa Claude** — i report Word non si convertono in modo
   affidabile con una regex.
4. Doppio clic su **`PUBBLICA.bat`** per commit + push (GitHub Pages si aggiorna da solo).

## Note tecniche

- I `.doc` e i `.txt` sono **gitignored**: su GitHub vanno solo i dashboard HTML + l'index.
- L'estrazione testo richiede **Microsoft Word** installato (usata via PowerShell COM).
- Ogni dashboard è un singolo file HTML autosufficiente (CSS/JS inline, font da Google Fonts).
