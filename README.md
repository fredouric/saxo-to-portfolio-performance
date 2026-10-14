# saxo-to-portfolio-performance

Converts a Saxo Bank **French** transaction report (`.xlsx`) into CSV files that
[Portfolio Performance](https://www.portfolio-performance.info/) imports
directly, with column headers and transaction types already in French.

Portfolio Performance 0.89 ships a `SaxoBankPDFExtractor`, but it only
recognises Saxo's English and German per-event documents (`Trade details`,
`Corporate Action Detail Report`, `Cash Transfer Detail Report`). It does not
read the French combined report (`Comptes rendus des transactions`), which is
what this tool is for.

## Layout

One folder per account, each holding the Saxo export as `tx.xlsx`:

```
cto/tx.xlsx     # ordinary account, foreign-currency securities
pea/tx.xlsx     # French PEA, EUR securities
```

```sh
uv run --with openpyxl python convert.py        # every account folder
uv run --with openpyxl python convert.py pea    # just one
```

Each run writes three files next to the source and prints a reconciliation.
A full run **asserts** that the rebuilt cash balance equals the sum of the
report's booked amounts, and stops if it does not.

| File | Import target ("Type de donnée") |
|---|---|
| `1-titres.csv` | Titres |
| `2-operations-titres.csv` | Opérations sur compte-titres |
| `3-operations-especes.csv` | Opérations du compte |

## Importing

1. Create a securities account ("Compte-titres") and a cash account
   ("Compte espèces") in EUR, and link them.
2. Import the three files **in order** via
   `Fichier > Importer > Fichiers CSV...`. For each one set:
   - **Type de donnée**: per the table above. Set this **first** — the wizard
     keeps the previous selection, and a mismatched target is the one thing
     that reliably breaks an import.
   - **Délimiteur**: `Point-virgule ( ; )`
   - **Encodage**: `UTF-8`
3. Column headers use the exact French Portfolio Performance labels, so they
   map automatically. If one shows as unmapped, double-click it and pick the
   field of the same name.
4. Afterwards, set a quote feed on each security. Symbols are pre-filled.

Import `1-titres.csv` first: it creates the securities with the right
currency. Importing transactions first makes Portfolio Performance create them
in the account currency instead.

## Keeping it up to date

Do not re-import the full files. Portfolio Performance flags a repeat as
`Opération déjà existante`, but only as a warning. You must then clear each
row by hand.

The script tracks what you imported, so it can emit only the new lines.

### First, record what you already imported

Run this once, after your first manual import:

```sh
uv run --with openpyxl python convert.py --mark-all
```

This writes `<folder>/.imported.json`. The file lists one entry per Saxo line.
It stays local, because it holds dates, amounts and ISINs.

### Then, on every update

1. Export the **complete** history again from Saxo.
2. Replace each `tx.xlsx`.
3. Write the new lines:

```sh
uv run --with openpyxl python convert.py --new
```

The script writes only unimported lines to `<folder>/nouveau/`. The full files
stay unchanged. It prints `no new lines` when there is nothing to do.

4. Import the two transaction files. Skip `1-titres.csv` unless a new holding
   appeared.
5. Record them as imported:

```sh
uv run --with openpyxl python convert.py --confirm
```

The script keeps the new keys in `<folder>/.pending.json` until you confirm.
A failed import therefore loses nothing. Run `--new` again to retry.

### How the script identifies a line

Saxo gives no identifier to most lines, so the key depends on the type:

| Line | Key |
|---|---|
| Trade | the Saxo order ID, which is unique |
| Dividend | the security and the **value date** |
| Cash transfer | the date, the direction and the amount |

Two identical lines get an occurrence index. This assumes Saxo keeps the same
row order between exports, which holds for the reports tested so far.

A dividend key uses the value date on purpose. A later correction nets onto
the same key, so the script reports a **changed amount** instead of a new
line:

```
CHANGED: div:US8740391003:2026-01-08 was 3.77, now 3.70.
```

Portfolio Performance cannot import that as a new row. Edit the existing
transaction by hand.

Always export the **complete** history. The script derives dividend share
counts and dividend corrections from the whole history. The ledger controls
only which lines the script writes, never how it calculates them.

## How the figures are built

Saxo books every line in the account currency, net of costs, and never breaks
out the commission. So:

- `Montant comptabilisé` becomes the `Valeur` of the transaction.
- `Frais` is the **residual** between the amount booked and the gross trade
  value converted to the account currency. Deriving it this way guarantees the
  cash balance closes to the cent, which a modelled fee rate would not.
- `Taux de change` is emitted only when the securities are not in the account
  currency. Saxo quotes it as account currency per 1 unit of security
  currency, which is the direction Portfolio Performance expects.
- The Saxo order ID goes into `Note`, which helps spot duplicates later.

Because the report gives only net amounts, the `Impôts / Taxes` column is
always empty and the whole cost sits in `Frais`. Both feed the cost basis, so
totals and performance are correct; only the breakdown is unavailable. For
gross dividends and withholding tax, download the per-event
`Corporate Action Detail Report` PDFs **in English** and use Portfolio
Performance's built-in Saxo importer.

## Ticker symbols

Saxo writes symbols as `TICKER:mic`, e.g. `ESE:xpar`. Yahoo Finance uses its
own exchange suffix, so the MIC is **translated, not dropped**: `xpar` becomes
`.PA`, `xams` becomes `.AS`, and US listings (`xnas`, `xnys`) take no suffix.
An unknown MIC stops the conversion rather than emitting a bare ticker.

This matters because a bare European ticker silently resolves to an unrelated
US security:

| Saxo | Bare ticker resolves to | Correct symbol |
|---|---|---|
| `ESE:xpar` | ESCO Technologies (NYSE, USD) | `ESE.PA` |
| `AI:xpar` | C3.ai Inc (NYSE, USD) | `AI.PA` |
| `TTE:xpar` | TotalEnergies ADR (NYSE, USD) | `TTE.PA` |

The `TTE` case is the nastiest: it loads real TotalEnergies prices, just the
USD ADR rather than the Paris line.

Some tickers differ outright. `SYMBOL_BY_ISIN` holds those overrides — Saxo
calls ASM International `ASMI` while Yahoo uses `ASM.AS` (beware `ASM.F`,
which is ASML, a different company). Verify a new symbol against its ISIN:

```sh
curl -s "https://query2.finance.yahoo.com/v1/finance/search?q=<ISIN>" | jq '.quotes[].symbol'
```

### Repairing a security with wrong prices

Changing the symbol does not delete prices already downloaded under the old
one:

1. Double-click the security in `Titres`, fix `Symbole`, save.
2. In `Cotations historiques`, right-click and choose `Tout supprimer`.
3. Update the quotes.

Transactions stay valid — they match on ISIN, so only the symbol and the
downloaded prices were ever wrong.

## Saxo wordings handled

| Événement | Becomes |
|---|---|
| `Acheter N @ P CCY` | `Achat` |
| `Vendre -N @ P CCY` | `Vente` |
| `Dividende en espèces` | `Dividendes` |
| `Dividende en espèces - Contre-passation` | netted into the original dividend |
| `Réinvestissement de dividendes` | `Dividendes` (noted) |
| `Dépôts` / `Retrait` | `Dépôt` / `Retrait` |

Two wordings need explaining:

- **Corrections.** Saxo books a corrected dividend as extra lines on the same
  security and value date: the original, a negative reversal, and a re-credit.
  Portfolio Performance has no negative dividend type, so each group is
  collapsed into one net row. A `--new` run reports a correction that lands on
  a dividend you already imported, because Portfolio Performance cannot take
  it as a new row. Edit that transaction by hand.
- **`Réinvestissement de dividendes`.** Despite the name this can arrive as
  plain cash with no share purchase, so it is imported as a dividend with the
  original wording kept in `Note`.

An unrecognised `Événement` aborts the run rather than silently dropping the
line. Add it to `DIVIDEND_EVENTS` or to the cash branch of `split()`.

## Notes

- A dividend paid after the position was closed carries no share count. The
  cash is right; only the per-share figure is unavailable.
- Moving cash between two Saxo accounts appears as a withdrawal in one export
  and a deposit in the other. Both import, and each account's cash is correct.
  Replace the pair with a `Transfert` afterwards if you prefer one movement.
- Exports and generated CSVs contain account numbers, balances and your full
  transaction history. They are gitignored; keep them that way.

## Licence

MIT
