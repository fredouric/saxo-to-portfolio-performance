"""Convert a Saxo Bank French transaction report (.xlsx) to Portfolio Performance CSV files.

Usage:
  python convert.py                      # every account folder, full history
  python convert.py pea                  # one account
  python convert.py --since 2026-10-13   # only rows on or after that date

One folder per account, each containing the Saxo export as tx.xlsx. With no
argument, every subfolder holding a tx.xlsx is converted. A full run writes the three CSV files next to
it. A --since run writes them to <folder>/depuis-<date>/ instead, so the full
files are left alone.

Always export the COMPLETE history from Saxo, even for a --since run. Share
counts on dividends and dividend corrections are derived from the whole
history; a partial export would compute them from partial data.
"""
import collections, csv, datetime, os, re, sys
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


def catalogue(trades, divs):
    secs = {}
    for t in trades:
        secs.setdefault(t["isin"], dict(isin=t["isin"], ticker=t["ticker"],
                                        name=t["name"], ccy=t["sec_ccy"]))
    for d in divs:
        secs.setdefault(d["isin"], dict(isin=d["isin"], ticker=d["ticker"],
                                        name=d["name"], ccy=d["ccy"]))
    return secs


def main(folder, since=None):
    rows = load(f"{folder}/tx.xlsx")
    trades, cash, divs = split(rows)
    divs, netted = net_dividends(divs)
    # Positions, netting and the security list always come from the full
    # history, even when only a tail of the rows is emitted.
    pos, div_shares = positions(trades, divs)
    all_secs = catalogue(trades, divs)

    print(f"\n{folder}" + (f" (depuis {since})" if since else "") + ":")
    out = folder
    if since:
        for tk, vd, _parts, tot in netted:
            if vd < since:
                print(f"  WARNING: the {tk} dividend with value date {vd} was corrected "
                      f"by a later line, so its net is now {tot} EUR. That dividend is "
                      f"before the cutoff. If you already imported it, edit it by hand.")
        trades = [t for t in trades if t["date"] >= since]
        cash = [c for c in cash if c["date"] >= since]
        divs = [d for d in divs if d["date"] >= since]
        out = f"{folder}/depuis-{since.isoformat()}"
        os.makedirs(out, exist_ok=True)

    # Only emit an exchange rate when the securities are not in the account currency.
    fx = any(r["rate"] != 1 for r in rows)
    secs = catalogue(trades, divs)   # only what the emitted rows refer to

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

    # ---- reconciliation ----
    for tk, vd, parts, tot in netted:
        if not since:
            print(f"  netted dividend: {tk} value date {vd}: "
                  f"{' + '.join(str(p) for p in parts)} = {tot} EUR")
    dep = sum(c["amount"] for c in cash if c["amount"] > 0)
    wit = -sum(c["amount"] for c in cash if c["amount"] < 0)
    div = sum(d["amount"] for d in divs)
    fees = sum(t["fee"] for t in trades)
    net = dep - wit + div - sum(t["value"] if t["side"] == "Acheter" else -t["value"]
                                for t in trades)
    print(f"  {len(rows)} rows in file -> {len(trades)} trades + {len(acct)} cash emitted")
    print(f"  deposits {dep} | withdrawals {wit} | dividends {div} | fees {fees}")
    if since:
        print(f"  net cash effect of the emitted rows: {net} EUR")
    else:
        assert net == sum(r["amount"] for r in rows), "does not reconcile"
        print(f"  closing cash {net} EUR (reconciles with the Saxo report)")
    held = [(all_secs[i]["ticker"], q) for i, q in pos.items() if q]
    closed = [all_secs[i]["ticker"] for i, q in pos.items() if not q]
    print("  positions: " + (", ".join(f"{t} {q}" for t, q in sorted(held)) or "none"))
    print("  closed:    " + (", ".join(sorted(closed)) or "none"))


if __name__ == "__main__":
    args = sys.argv[1:]
    since = None
    if "--since" in args:
        i = args.index("--since")
        since = datetime.date.fromisoformat(args[i + 1])
        del args[i:i + 2]
    folders = [a.rstrip("/") for a in args]
    if not folders:
        folders = sorted(d for d in os.listdir(".")
                         if os.path.isfile(os.path.join(d, "tx.xlsx")))
        if not folders:
            sys.exit("no account folder found. Create one (e.g. cto/) "
                     "containing the Saxo export as tx.xlsx.")
    for folder in folders:
        main(folder, since)
