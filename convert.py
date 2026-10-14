"""Convert a Saxo Bank French transaction report (.xlsx) to Portfolio Performance CSV files.

Usage:
  python convert.py              # rebuild the full history
  python convert.py pea          # one account only
  python convert.py --new        # write only the lines you have not imported
  python convert.py --confirm    # record those lines as imported
  python convert.py --mark-all   # record the whole export as imported

One folder per account. Each folder holds the Saxo export as tx.xlsx. With no
folder argument, the script converts every subfolder that holds a tx.xlsx.

A full run writes the three CSV files next to the export. A --new run writes
them to <folder>/nouveau/ instead, so the full files stay unchanged.

Always export the COMPLETE history from Saxo. The script derives dividend
share counts and dividend corrections from the whole history. A partial
export computes them from partial data.
"""
import collections, csv, json, os, re, sys
from decimal import Decimal as D, ROUND_HALF_UP
import openpyxl

CENT = D("0.01")
TRADE_RE = re.compile(r"^(Acheter|Vendre)\s+(-?[\d.]+)\s+@\s+([\d.]+)\s+([A-Z]{3})$")
# Saxo wordings that credit cash as investment income.
DIVIDEND_EVENTS = ("Dividende", "Réinvestissement de dividendes")

# Saxo appends a MIC to each symbol ("ESE:xpar"). Yahoo Finance needs its own
# exchange suffix instead. Without it a European listing silently resolves to an
# unrelated US ticker: bare "ESE" is ESCO Technologies, bare "AI" is C3.ai and
# bare "TTE" is the TotalEnergies ADR in USD. Unknown MICs stop the run.
YAHOO_SUFFIX = {
    "xnys": "", "xnas": "", "arcx": "", "bats": "",   # US listings need no suffix
    "xpar": ".PA", "xams": ".AS", "xbru": ".BR", "xlis": ".LS",
    "xetr": ".DE", "xmil": ".MI", "xmad": ".MC", "xlon": ".L",
    "xswx": ".SW", "xvtx": ".SW", "xcse": ".CO", "xsto": ".ST",
    "xhel": ".HE", "xosl": ".OL", "xwbo": ".VI", "xdub": ".IR",
}


# Some tickers differ outright, not just by suffix. Keyed by ISIN, verified
# against Yahoo's ISIN lookup. Saxo calls ASM International "ASMI"; Yahoo uses
# "ASM.AS" ("ASM.F" is ASML, a different company).
SYMBOL_BY_ISIN = {
    "NL0000334118": "ASM.AS",   # ASM International NV, Euronext Amsterdam
}


def yahoo_symbol(saxo_symbol, place, isin=""):
    """Turn Saxo's "TICKER:mic" into a Yahoo Finance symbol."""
    if isin in SYMBOL_BY_ISIN:
        return SYMBOL_BY_ISIN[isin]
    if not saxo_symbol:
        return ""
    ticker, _, mic = saxo_symbol.partition(":")
    if mic.lower() not in YAHOO_SUFFIX:
        sys.exit(f"unknown exchange {mic!r} ({place}) for {ticker}: "
                 f"add it to YAHOO_SUFFIX with the right Yahoo suffix")
    return ticker + YAHOO_SUFFIX[mic.lower()]


# French Portfolio Performance labels (PP 0.89, locale fr_FR)
T_BUY, T_SELL = "Achat", "Vente"
T_DEPOSIT, T_REMOVAL, T_DIVIDEND = "Dépôt", "Retrait", "Dividendes"


def r2(x):
    return D(x).quantize(CENT, rounding=ROUND_HALF_UP)


def fr(x, places=2):
    """Format a Decimal with a comma decimal separator, as a French PP install expects."""
    return f"{D(x).quantize(D(1).scaleb(-places), rounding=ROUND_HALF_UP)}".replace(".", ",")


def load(path):
    ws = openpyxl.load_workbook(path, data_only=True)["Transactions"]
    rows = []
    for r in list(ws.iter_rows(values_only=True))[1:]:
        if not r[0]:
            continue
        rows.append(dict(date=r[1].date(), vdate=r[2].date() if r[2] else None,
                         kind=r[3], name=(r[4] or "").strip(), isin=r[5] or "",
                         ccy=r[6] or "", ticker=yahoo_symbol(r[8], r[7], r[5] or ""),
                         event=str(r[9]), amount=D(str(r[10])),
                         order=str(r[11] or ""), rate=D(str(r[12]))))
    rows.sort(key=lambda x: (x["date"], x["order"]))
    return rows


def split(rows):
    trades, cash, divs = [], [], []
    for r in rows:
        m = TRADE_RE.match(r["event"])
        if m:
            side, shares, price, ccy = m.group(1), abs(D(m.group(2))), D(m.group(3)), m.group(4)
            gross = r2(shares * price * r["rate"])       # gross converted to account currency
            value = abs(r["amount"])                      # net cash moved
            # Fee is the exact residual, so value == gross +/- fee closes to the cent.
            fee = value - gross if side == "Acheter" else gross - value
            trades.append(dict(r, side=side, shares=shares, price=price, sec_ccy=ccy,
                               gross=gross, value=value, fee=r2(fee)))
        elif r["event"].startswith(DIVIDEND_EVENTS):
            divs.append(r)
        elif r["event"] in ("Dépôts", "Retrait"):
            cash.append(r)
        else:
            sys.exit(f"unhandled event: {r['event']!r}")
    return trades, cash, divs


def net_dividends(divs):
    """Saxo books corrections as extra rows on the same security and value date
    (a negative dividend plus a positive "Contre-passation"). PP has no negative
    dividend type, so collapse each group into one net row."""
    groups = collections.OrderedDict()
    for d in divs:
        groups.setdefault((d["isin"], d["vdate"]), []).append(d)
    out, netted = [], []
    for (isin, vdate), grp in groups.items():
        total = sum(g["amount"] for g in grp)
        if len(grp) > 1:
            netted.append((grp[0]["ticker"], vdate, [g["amount"] for g in grp], total))
        if total > 0:
            out.append(dict(grp[0], amount=total))
        elif total != 0:
            sys.exit(f"dividend group nets to {total} for {isin} {vdate}")
    out.sort(key=lambda d: d["date"])
    return out, netted


def positions(trades, divs):
    """Running share count per security, used for the Shares column on dividends."""
    pos = collections.defaultdict(D)
    at_div = {}
    for e in sorted(trades + divs, key=lambda x: (x["date"], 0 if "side" in x else 1)):
        if "side" in e:
            pos[e["isin"]] += e["shares"] if e["side"] == "Acheter" else -e["shares"]
        else:
            at_div[id(e)] = pos[e["isin"]]
    return pos, at_div


def write(path, header, body):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(header)
        w.writerows(body)
    print(f"  {path}: {len(body)} rows")


LEDGER_NAME = ".imported.json"
PENDING_NAME = ".pending.json"


def fingerprint(item):
    """Return the value that shows whether an imported line changed."""
    if "side" in item:
        return f"{item['side']}:{item['shares']}:{item['value']}"
    return str(item["amount"])


def natural_key(item):
    """Return a stable identity for one Saxo line.

    A trade carries a Saxo order ID, which is unique. A dividend uses the
    security and the value date. A later correction nets onto the same key, so
    the script reports it as a changed amount, not as a new line. A cash
    transfer has no identifier, so it uses the date, the direction and the
    amount.
    """
    if "side" in item:
        if item["order"]:
            return f"trade:{item['order']}"
        return f"trade:{item['date']}:{item['isin']}:{item['side']}:{item['shares']}"
    if item["event"].startswith(DIVIDEND_EVENTS):
        return f"div:{item['isin']}:{item['vdate']}"
    side = "dep" if item["amount"] > 0 else "ret"
    return f"cash:{item['date']}:{side}:{abs(item['amount'])}"


def keyed(items):
    """Pair each item with its key. A repeated key gets an occurrence index."""
    seen = collections.Counter()
    out = []
    for item in items:
        key = natural_key(item)
        seen[key] += 1
        if seen[key] > 1:
            key = f"{key}#{seen[key]}"
        out.append((key, item))
    return out


def read_json(path, default):
    if not os.path.exists(path):
        return default
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def write_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=1, sort_keys=True, ensure_ascii=False)
        f.write("\n")


def catalogue(trades, divs):
    secs = {}
    for t in trades:
        secs.setdefault(t["isin"], dict(isin=t["isin"], ticker=t["ticker"],
                                        name=t["name"], ccy=t["sec_ccy"]))
    for d in divs:
        secs.setdefault(d["isin"], dict(isin=d["isin"], ticker=d["ticker"],
                                        name=d["name"], ccy=d["ccy"]))
    return secs


def main(folder, mode="full"):
    rows = load(f"{folder}/tx.xlsx")
    trades, cash, divs = split(rows)
    divs, netted = net_dividends(divs)
    # The script derives positions, netting and the security list from the
    # full history, even when it emits only a few lines.
    pos, div_shares = positions(trades, divs)
    all_secs = catalogue(trades, divs)

    print(f"\n{folder} [{mode}]:")

    ledger = read_json(f"{folder}/{LEDGER_NAME}", {}).get("entries", {})
    marks = keyed(trades) + keyed(divs) + keyed(cash)
    fresh = {k: fingerprint(i) for k, i in marks if k not in ledger}
    changed = {k: (ledger[k], fingerprint(i)) for k, i in marks
               if k in ledger and ledger[k] != fingerprint(i)}

    if mode == "mark-all":
        write_json(f"{folder}/{LEDGER_NAME}",
                   {"version": 1, "entries": {k: fingerprint(i) for k, i in marks}})
        print(f"  recorded all {len(marks)} lines as imported")
        return

    if mode == "confirm":
        pending = read_json(f"{folder}/{PENDING_NAME}", None)
        if not pending:
            print("  nothing pending. Run --new first.")
            return
        ledger.update(pending["entries"])
        write_json(f"{folder}/{LEDGER_NAME}", {"version": 1, "entries": ledger})
        os.remove(f"{folder}/{PENDING_NAME}")
        print(f"  recorded {len(pending['entries'])} lines as imported "
              f"({len(ledger)} in total)")
        return

    for key, (was, now) in sorted(changed.items()):
        print(f"  CHANGED: {key} was {was}, now {now}. Saxo corrected a line you "
              f"already imported. Edit that transaction in Portfolio Performance "
              f"by hand.")

    out = folder
    if mode == "new":
        if not fresh:
            print("  no new lines. Nothing to import.")
            return
        keep = {id(i) for k, i in marks if k in fresh}
        trades = [t for t in trades if id(t) in keep]
        divs = [d for d in divs if id(d) in keep]
        cash = [c for c in cash if id(c) in keep]
        out = f"{folder}/nouveau"
        os.makedirs(out, exist_ok=True)

    # The script emits an exchange rate only when the securities sit outside
    # the account currency.
    fx = any(r["rate"] != 1 for r in rows)
    secs = catalogue(trades, divs)   # only what the emitted lines refer to

    write(f"{out}/1-titres.csv", ["ISIN", "Symbole boursier", "Nom du titre", "Devise"],
          [[s["isin"], s["ticker"], s["name"], s["ccy"]]
           for s in sorted(secs.values(), key=lambda s: s["name"])])

    hdr = ["Date", "Type", "Valeur", "Devise de l'opération"]
    hdr += ["Taux de change"] if fx else []
    hdr += ["Parts", "Frais", "ISIN", "Symbole boursier", "Nom du titre", "Note"]
    write(f"{out}/2-operations-titres.csv", hdr,
          [[t["date"].isoformat(), T_BUY if t["side"] == "Acheter" else T_SELL,
            fr(t["value"]), "EUR"] + ([fr(t["rate"], 8)] if fx else []) +
           [fr(t["shares"], 4), fr(t["fee"]), t["isin"], t["ticker"], t["name"],
            f"Saxo ordre {t['order']}" if t["order"] else "Saxo"]
           for t in trades])

    acct = []
    for c in cash:
        acct.append([c["date"].isoformat(), T_DEPOSIT if c["amount"] > 0 else T_REMOVAL,
                     fr(abs(c["amount"])), "EUR"] + ([""] if fx else []) +
                    ["", "", "", "", "", "Saxo virement"])
    for d in divs:
        note = "Saxo " + ("réinvestissement de dividendes"
                          if d["event"].startswith("Réinvestissement") else "dividende")
        if "Contre-passation" in d["event"]:
            note += " (contre-passation)"
        sh = div_shares[id(d)]
        acct.append([d["date"].isoformat(), T_DIVIDEND, fr(abs(d["amount"])), "EUR"] +
                    ([fr(d["rate"], 8)] if fx else []) +
                    [fr(sh, 4) if sh > 0 else "", "",
                     d["isin"], d["ticker"], d["name"], note])
    acct.sort(key=lambda r: (r[0], 0 if r[1] == T_DEPOSIT else 1))
    hdr = ["Date", "Type", "Valeur", "Devise de l'opération"]
    hdr += ["Taux de change"] if fx else []
    hdr += ["Parts", "Impôts / Taxes", "ISIN", "Symbole boursier", "Nom du titre", "Note"]
    write(f"{out}/3-operations-especes.csv", hdr, acct)

    if mode == "new":
        write_json(f"{folder}/{PENDING_NAME}", {"version": 1, "entries": fresh})
        print(f"  {len(fresh)} new lines written. Import them, then run --confirm.")

    # ---- reconciliation ----
    if mode == "full":
        for tk, vd, parts, tot in netted:
            print(f"  netted dividend: {tk} value date {vd}: "
                  f"{' + '.join(str(p) for p in parts)} = {tot} EUR")
    dep = sum(c["amount"] for c in cash if c["amount"] > 0)
    wit = -sum(c["amount"] for c in cash if c["amount"] < 0)
    div = sum(d["amount"] for d in divs)
    fees = sum(t["fee"] for t in trades)
    net = dep - wit + div - sum(t["value"] if t["side"] == "Acheter" else -t["value"]
                                for t in trades)
    print(f"  {len(rows)} lines in file -> {len(trades)} trades + {len(acct)} cash emitted")
    print(f"  deposits {dep} | withdrawals {wit} | dividends {div} | fees {fees}")
    if mode == "full":
        assert net == sum(r["amount"] for r in rows), "does not reconcile"
        print(f"  closing cash {net} EUR (reconciles with the Saxo report)")
    else:
        print(f"  net cash effect of the emitted lines: {net} EUR")
    held = [(all_secs[i]["ticker"], q) for i, q in pos.items() if q]
    closed = [all_secs[i]["ticker"] for i, q in pos.items() if not q]
    print("  positions: " + (", ".join(f"{t} {q}" for t, q in sorted(held)) or "none"))
    print("  closed:    " + (", ".join(sorted(closed)) or "none"))


MODES = {"--new": "new", "--confirm": "confirm", "--mark-all": "mark-all"}

if __name__ == "__main__":
    args = sys.argv[1:]
    mode = "full"
    for flag, name in MODES.items():
        if flag in args:
            mode = name
            args.remove(flag)
    unknown = [a for a in args if a.startswith("-")]
    if unknown:
        sys.exit(f"unknown option {unknown[0]}. Options: {', '.join(MODES)}")
    folders = args or sorted(d for d in os.listdir(".")
                             if os.path.isfile(os.path.join(d, "tx.xlsx")))
    if not folders:
        sys.exit("no account folder found. Create one (e.g. cto/) "
                 "that holds the Saxo export as tx.xlsx.")
    for folder in folders:
        main(folder.rstrip("/"), mode)
