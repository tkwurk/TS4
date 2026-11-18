# Data Access Layer (Seam 2)

The database seam that keeps business logic independent of where data is stored.
See [`/AGENTS.md`](../AGENTS.md) for the overall architecture.

## Why this exists

App users must **not** have direct access to the backing spreadsheet, and the app
must stay **free**. The active design (**Option C**) deploys the Web App to
**execute as the owner**, so the script reaches the sheet with the owner's access
while users never get it — using `SpreadsheetApp`, with **no Google Cloud project
and no service account**. All data access goes through a provider-agnostic `Repo`
facade so the store can be swapped later without touching business logic.

```
Workers ──► Repo (provider-agnostic) ──► active adapter
                                          ├─ SpreadsheetAppAdapter  (DEFAULT, Option C)
                                          └─ SheetsAdapter          (Option B, future)
```

## Files

| File | Role |
|---|---|
| [repo.gs](./repo.gs) | Provider-agnostic facade the workers call. Defaults to SpreadsheetAppAdapter; set `Repo.adapter` to switch. |
| [spreadsheetAppAdapter.gs](./spreadsheetAppAdapter.gs) | **Default** adapter — `SpreadsheetApp`, execute-as-owner, free, no Cloud. |
| [sheetsAdapter.gs](./sheetsAdapter.gs) | Option B adapter — Sheets REST API (per-user concurrency). |
| [serviceAccount.gs](./serviceAccount.gs) | Option B only — mints a service-account token. |

## Setup — Option C (active, free)

1. Deploy the Web App with **Execute as: Me (the owner)** and **Who has access:
   Anyone with a Google Account** (or your Workspace domain).
2. Nothing else. `SPREADSHEET_ID` is resolved by `config.js`; the owner already
   has access to the sheet. No Cloud project, no service account, no billing.

## Usage

```js
// Read a whole collection as objects keyed by the header row
var products = Repo.readRows('PRODUCTS');          // SHEETS key or literal name

// Append objects (aligned to the sheet's headers — column order independent)
Repo.appendObjects('ORDERS', { OrderNo: 'A1001', Total: 42.5, Outlet: 'Main' });

// Append pre-ordered rows
Repo.appendRows('AUDIT_LOG', [[ new Date(), 'checkout', 'ok' ]]);

// Overwrite an explicit range
Repo.updateRange('ORDERS', 'G2:G2', [[ 'PAID' ]]);
```

The purchase-save path reads account types and purchase rows through `Repo` and
appends purchase lines with their investment flag in the same write. It preserves
the legacy 20-column purchase layout, adding investment headers as needed, and
locks the duplicate check and append. Other legacy purchasing operations have
not yet been migrated.

## Migration path

- **Incremental:** new features read/write only through `Repo`. Existing workers
  that still call `SpreadsheetApp` keep working and are ported collection by
  collection.
- **Swap the database later:** implement a new adapter exposing the same methods
  as the current adapter (`getValues`, `appendRows`, `updateValues`,
  `clearRange`), then set `Repo.adapter = YourNewAdapter`. No worker or
  business-logic changes.

## Option B — service-account REST (future, still free)

Switch to this only if you need the **per-user 30-concurrency pool** (many tills
hitting the app at once). It is also free — the Sheets API and service accounts
carry no charge — but requires a one-time Google Cloud setup:

1. Create a Cloud **service account** + JSON key; enable the **Google Sheets API**.
2. Share the spreadsheet with the service account's `client_email` as **Editor**
   (not with app users).
3. Set Script Properties: `SA_CLIENT_EMAIL`, `SA_PRIVATE_KEY` (keep the `\n`).
4. Deploy the Web App **Execute as: user accessing**, and set
   `Repo.adapter = SheetsAdapter`.

## Rules (enforced by /AGENTS.md)

- Workers MUST NOT call `SpreadsheetApp` directly — go through `Repo`.
- Only an adapter knows the data is in Google Sheets.
- Any service-account key lives in Script Properties, never in source.
