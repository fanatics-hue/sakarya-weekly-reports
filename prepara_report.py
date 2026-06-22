#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
==========================================================
  SAKARYA WEEKLY REPORTS - PREPARA REPORT
  Parte automatica del flusso per-report:
   1. Trova le cartelle  SC26-3-SPM-CNF-DPR-XXX (WEEK NN)
   2. Per ognuna prende il .doc piu' recente (ultimo caricato)
   3. Lo estrae in .txt con Word (cosi' e' leggibile/diffabile)
   4. Rileva la disciplina (3LPP Coating / Pipe Mill)
   5. Allinea il NOME della cartella all'ultimo .doc
   6. Rigenera le card della landing index.html
   7. Segnala quali dashboard HTML sono da RICOSTRUIRE
      (settimana del .doc diversa da quella del dashboard)

  NB: la RICOSTRUZIONE del contenuto del dashboard HTML dal
  report Word la fa Claude (i report sono narrativi: tabelle,
  NCR, attendees... non si estraggono in modo affidabile con
  una regex). Questo script fa solo la parte meccanica e dice
  a Claude/Rino dove intervenire.
==========================================================
"""
import re, sys, subprocess
from pathlib import Path

BASE = Path(__file__).parent
FOLDER_RE = re.compile(r"^SC26-3-SPM-CNF-DPR-[\w-]+ \(WEEK \d+\)$", re.I)
WEEK_RE   = re.compile(r"WEEK\s*(\d+)", re.I)

def log(msg=""):
    print(msg)

def estrai_testo_doc(doc_path: Path) -> Path | None:
    """Estrae il testo di un .doc/.docx in .txt usando Word via PowerShell COM."""
    out = doc_path.with_suffix(".txt")
    ps = (
        "$w = New-Object -ComObject Word.Application; $w.Visible=$false; "
        f"$d = $w.Documents.Open('{doc_path}', $false, $true); "
        f"Set-Content -Path '{out}' -Value $d.Content.Text -Encoding UTF8; "
        "$d.Close($false); $w.Quit()"
    )
    try:
        subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                       check=True, capture_output=True, timeout=120)
        return out if out.exists() else None
    except Exception as e:
        log(f"   [!] Estrazione Word fallita: {e}")
        return None

def rileva_disciplina(testo: str) -> tuple[str, str]:
    """Ritorna (codice, etichetta) della disciplina dal contenuto del report."""
    t = testo.upper()
    if "3LPP COATING" in t and "UOE SAWL PIPE" not in t:
        return ("coat", "3LPP Coating")
    if "UOE SAWL PIPE" in t or "PIPE MILL MASS PRODUCTION" in t or "SAWL UOE" in t:
        return ("mill", "Pipe Mill")
    # fallback: cerca nel nome dashboard esistente
    return ("?", "Da confermare")

def settimana_da_nome(nome: str) -> int | None:
    m = WEEK_RE.search(nome)
    return int(m.group(1)) if m else None

def scan():
    cartelle = sorted([d for d in BASE.iterdir()
                       if d.is_dir() and FOLDER_RE.match(d.name)])
    if not cartelle:
        log("Nessuna cartella report trovata (pattern: 'SC26-3-SPM-CNF-DPR-XXX (WEEK NN)').")
        return []

    risultati = []
    log("=" * 58)
    log("  SAKARYA WEEKLY REPORTS - PREPARA REPORT")
    log("=" * 58)

    for d in cartelle:
        log(f"\n[ {d.name} ]")
        docs = sorted([f for f in d.glob("*.doc")] + [f for f in d.glob("*.docx")],
                      key=lambda f: f.stat().st_mtime, reverse=True)
        if not docs:
            log("   [!] Nessun .doc/.docx nella cartella - salto.")
            continue
        ultimo = docs[0]
        log(f"   Ultimo file caricato : {ultimo.name}")

        txt = estrai_testo_doc(ultimo)
        disc_code, disc_lbl = ("?", "Da confermare")
        if txt:
            testo = txt.read_text(encoding="utf-8", errors="ignore")
            disc_code, disc_lbl = rileva_disciplina(testo)
            log(f"   Testo estratto       : {txt.name} ({len(testo)} caratteri)")
        log(f"   Disciplina rilevata  : {disc_lbl}")

        # allinea nome cartella all'ultimo .doc
        nome_atteso = ultimo.stem  # es. 'SC26-3-SPM-CNF-DPR-020 (WEEK 25)'
        if d.name != nome_atteso and FOLDER_RE.match(nome_atteso + ""):
            nuovo = d.with_name(nome_atteso)
            if not nuovo.exists():
                d.rename(nuovo)
                log(f"   Cartella rinominata  : -> {nome_atteso}")
                d = nuovo

        # dashboard html presente + settimana
        dash = sorted(d.glob("*_dashboard.html"), key=lambda f: f.stat().st_mtime,
                      reverse=True)
        wk_doc = settimana_da_nome(ultimo.stem)
        if dash:
            wk_dash = settimana_da_nome(dash[0].name)
            stato = "OK" if wk_dash == wk_doc else "DA RICOSTRUIRE"
            log(f"   Dashboard            : {dash[0].name}  [W{wk_dash}]  -> {stato}")
        else:
            stato = "DA CREARE"
            log(f"   Dashboard            : nessuno  -> {stato}")

        risultati.append({
            "folder": d.name, "doc": ultimo.name, "disc": disc_code,
            "disc_lbl": disc_lbl, "week": wk_doc,
            "dashboard": dash[0].name if dash else None, "stato": stato,
        })

    return risultati

def rigenera_index(risultati):
    index = BASE / "index.html"
    if not index.exists():
        log("\n[!] index.html non trovato - salto rigenerazione card.")
        return
    html = index.read_text(encoding="utf-8")
    cards = []
    for r in risultati:
        if not r["dashboard"]:
            continue
        folder_enc = r["folder"].replace(" ", "%20")
        dash_enc = r["dashboard"].replace(" ", "%20")
        href = f"{folder_enc}/{dash_enc}"
        cls = "rcard pm" if r["disc"] == "mill" else "rcard"
        disc_cls = "mill" if r["disc"] == "mill" else "coat"
        code = r["folder"].split(" (")[0]
        cards.append(
            f'    <a class="{cls}" href="{href}">\n'
            f'      <div class="rtag">{code}</div>\n'
            f'      <div class="rdisc {disc_cls}">{r["disc_lbl"].upper()}</div>\n'
            f'      <div class="rtitle">{r["disc_lbl"]} — Weekly Report</div>\n'
            f'      <div class="rweek">Week {r["week"]} · 2026</div>\n'
            f'      <div class="ropen">OPEN DASHBOARD →</div>\n'
            f'    </a>'
        )
    blocco = (
        '  <!-- AUTO:REPORT_CARDS -->\n'
        '  <div class="sl">Available Reports</div>\n'
        '  <div class="grid">\n\n' + "\n\n".join(cards) + "\n\n  </div>\n"
        '  <!-- /AUTO:REPORT_CARDS -->'
    )
    nuovo, n = re.subn(r"  <!-- AUTO:REPORT_CARDS -->.*?<!-- /AUTO:REPORT_CARDS -->",
                       blocco, html, flags=re.DOTALL)
    if n:
        index.write_text(nuovo, encoding="utf-8")
        log(f"\n[+] index.html rigenerato con {len(cards)} card.")
    else:
        log("\n[!] Marker AUTO:REPORT_CARDS non trovato in index.html (card non rigenerate).")

def main():
    risultati = scan()
    if not risultati:
        return
    rigenera_index(risultati)

    log("\n" + "=" * 58)
    log("  RIEPILOGO")
    log("=" * 58)
    da_fare = [r for r in risultati if r["stato"] != "OK"]
    for r in risultati:
        flag = "   " if r["stato"] == "OK" else " * "
        log(f"{flag}{r['folder']}  [{r['disc_lbl']}]  -> {r['stato']}")
    if da_fare:
        log("\n  AZIONE: i report con ' * ' hanno un .doc piu' recente del dashboard.")
        log("  Apri Claude e chiedi: \"ricostruisci i dashboard da ricostruire\"")
        log("  (il .txt estratto e' gia' nella cartella, pronto da leggere).")
    else:
        log("\n  Tutti i dashboard sono allineati all'ultimo report. Niente da ricostruire.")
    log("")

if __name__ == "__main__":
    main()
