# Feature modules

Each subfolder is a self-contained feature migrated from the original Google
Apps Script project (for example `orders/`, `checkout/`, `purchasing/`).

Recommended structure per feature:

```
features/
  orders/
    components/      # React components (UI)
    hooks/           # Feature hooks (state/logic)
    api.ts           # Data access — calls the Seam 1 client in src/services/api/
    types.ts         # Feature-specific types
    index.ts         # Public exports
```

Existing examples: `auth/`, `checkout/`, `catalog/`, `reports/`.

## Migration guide (Apps Script -> TypeScript)

The client never calls Google Sheets directly. It calls the Apps Script backend
through the Seam 1 JSON API; the backend owns all data access via the `Repo`
layer (see [`/AGENTS.md`](../../AGENTS.md)).

| Apps Script (old UI)                         | TS4 replacement                                   |
| -------------------------------------------- | ------------------------------------------------- |
| `google.script.run.myFunction(params)`       | `callApi(eventType, payload, auth)` / `api.*`      |
| `SpreadsheetApp...getValues()` (server)      | `Repo.readRows(COLLECTION)` in a worker           |
| `sheet.appendRow([...])` (server)            | `Repo.appendObjects(COLLECTION, {...})`           |
| HTML `<script>` UI logic                     | a React component under `components/`             |

Drop the original `.gs`/`.html` source into `/apps-script-source` for reference
while porting.
