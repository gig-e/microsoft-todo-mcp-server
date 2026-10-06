# Cloud Connector (Azure Container Apps)

The default way to run this server is local stdio: each machine installs it, signs in once,
and keeps its own machine-bound token cache. The cloud connector is an opt-in alternative:
one deployment on Azure Container Apps serving the same tool set over Streamable HTTP, so any
machine can reach your To Do data without a local install or sign-in.

The local stdio path (`dist/cli.js`) is unaffected; nothing here changes it.

## How it fits together

```
Claude Desktop ──stdio──> local-proxy.js ──HTTPS + Bearer──> Container App (/mcp)
Claude Code ─────────────── HTTP + Bearer ─────────────────> Container App (/mcp)
                                                                    │
                                              managed identity      ▼
                                              ────────────> Key Vault secret
                                                            (encrypted MSAL cache)
```

| Piece                     | Role                                                                                                |
| ------------------------- | --------------------------------------------------------------------------------------------------- |
| `src/http-server.ts`      | HTTP entry point. Serves MCP at `/mcp` behind a static bearer token; unauthenticated `GET /health`. |
| `src/token-store.ts`      | `MSTODO_TOKEN_STORE=key-vault` keeps the encrypted token cache in a Key Vault secret, not on disk.  |
| `src/seed-cloud-token.ts` | One-time local step: re-encrypts your existing signed-in cache with the server's key for upload.    |
| `src/local-proxy.ts`      | stdio ↔ HTTP relay for Claude Desktop, which can't send a static bearer token to a remote server.   |
| `Dockerfile`              | Two-stage `node:22-slim` image running `dist/http-server.js`.                                       |

There is no interactive sign-in in the container. You sign in locally once, then seed the
resulting token cache into Key Vault. From then on the server refreshes tokens silently and
writes the updated cache back to the same secret.

## Environment variables

**Server (Container App):**

| Variable                   | Required | Purpose                                                                                                 |
| -------------------------- | -------- | ------------------------------------------------------------------------------------------------------- |
| `CLIENT_ID`                | yes      | App registration client ID, same as local.                                                              |
| `TENANT_ID`                | yes\*    | Same as local (\*defaults to `organizations`).                                                          |
| `MSTODO_BEARER_TOKEN`      | yes      | Shared secret clients must send as `Authorization: Bearer <token>`. Server exits without it.            |
| `MSTODO_ENCRYPTION_KEY`    | yes      | Key material for the token cache. Replaces the machine-ID key, since a container has no stable machine. |
| `MSTODO_TOKEN_STORE`       | yes      | Set to `key-vault`. Anything else falls back to a local file, which is lost on restart.                 |
| `MSTODO_KEYVAULT_URL`      | yes      | e.g. `https://<vault>.vault.azure.net/`                                                                 |
| `MSTODO_TOKEN_SECRET_NAME` | yes      | Name of the Key Vault secret holding the cache.                                                         |
| `MSTODO_ACCESS_MODE`       | no       | `read` / `write` / `full`, same as local. Defaults to `full`.                                           |
| `MSTODO_TIMEZONE`          | no       | **Set this.** The container runs in UTC, so without it `get-agenda` decides "today" in UTC.             |
| `PORT`                     | no       | Injected by Container Apps; defaults to `8080`.                                                         |

Store `MSTODO_BEARER_TOKEN` and `MSTODO_ENCRYPTION_KEY` as Container App secrets and reference
them with `secretref:`, not as plain env values.

**Client (`local-proxy.js`):**

| Variable            | Purpose                                               |
| ------------------- | ----------------------------------------------------- |
| `MSTODO_AUTH_TOKEN` | The same value as the server's `MSTODO_BEARER_TOKEN`. |

## Deployment

Commands below use placeholders: `<rg>`, `<vault>`, `<app>`, `<env>`, `<secret-name>`.

### 1. Generate the two secrets

```bash
openssl rand -base64 32   # -> MSTODO_BEARER_TOKEN
openssl rand -base64 32   # -> MSTODO_ENCRYPTION_KEY
```

Keep both somewhere safe (a password manager). Losing the encryption key means re-seeding;
losing the bearer token means rotating it on the server and every client.

### 2. Sign in locally and seed the token cache

Sign in on any machine with a local checkout (`pnpm run auth`), then re-encrypt that cache
with the server's key:

```bash
MSTODO_ENCRYPTION_KEY='<encryption-key>' pnpm run seed-cloud-token
```

This writes `.cloud-token-cache.bin` (base64 ciphertext) in the current directory.

> **Pass the key inline for this one command only. Don't put `MSTODO_ENCRYPTION_KEY` in your
> local `.env`.** Locally the cache is encrypted with the machine-ID key, and the seed script
> decrypts with that key. If `MSTODO_ENCRYPTION_KEY` is set in your local environment, the
> local server will start encrypting its own cache with it instead, and seeding will fail to
> decrypt.

### 3. Create the Key Vault and upload the cache

```bash
az keyvault create --name <vault> --resource-group <rg> --enable-rbac-authorization true
az keyvault secret set --vault-name <vault> --name <secret-name> --file .cloud-token-cache.bin
```

Then delete `.cloud-token-cache.bin`. It's in `.gitignore` and `.dockerignore`, but it's
still a credential (useless without the encryption key, though).

### 4. Deploy the container

From the repo root, which builds the `Dockerfile` remotely:

```bash
az containerapp up --name <app> --resource-group <rg> --environment <env> \
  --source . --ingress external --target-port 8080
```

### 5. Grant Key Vault access via managed identity

The server authenticates to Key Vault with `DefaultAzureCredential`, which picks up the
Container App's managed identity. It needs to **read and write** the secret, because token refreshes
are written back. Assign _Key Vault Secrets Officer_, not _Secrets User_:

```bash
az containerapp identity assign --name <app> --resource-group <rg> --system-assigned

az role assignment create \
  --assignee <principal-id-from-previous-output> \
  --role "Key Vault Secrets Officer" \
  --scope $(az keyvault show --name <vault> --query id -o tsv)
```

### 6. Configure secrets and environment

```bash
az containerapp secret set --name <app> --resource-group <rg> \
  --secrets bearer-token='<bearer-token>' encryption-key='<encryption-key>'

az containerapp update --name <app> --resource-group <rg> --set-env-vars \
  CLIENT_ID=<client-id> \
  TENANT_ID=<tenant-id> \
  MSTODO_BEARER_TOKEN=secretref:bearer-token \
  MSTODO_ENCRYPTION_KEY=secretref:encryption-key \
  MSTODO_TOKEN_STORE=key-vault \
  MSTODO_KEYVAULT_URL=https://<vault>.vault.azure.net/ \
  MSTODO_TOKEN_SECRET_NAME=<secret-name> \
  MSTODO_TIMEZONE=America/New_York \
  MSTODO_ACCESS_MODE=full
```

### 7. Point the health probes at `/health`

Container Apps' default probes never get a 2xx from this server: `/` is a 404, and `/mcp`
returns 401 because a platform probe can't send the bearer token. Without explicit probes the
platform logs `Probe of StartUp failed` and may never consider a replica healthy.

Export the app's YAML (`az containerapp show -n <app> -g <rg> -o yaml > app.yaml`), add
probes to the container, and apply it with `az containerapp update -n <app> -g <rg> --yaml app.yaml`:

```yaml
properties:
  template:
    containers:
      - name: <app>
        probes:
          - type: Startup
            httpGet: { path: /health, port: 8080 }
            initialDelaySeconds: 3
            periodSeconds: 5
            failureThreshold: 10
          - type: Readiness
            httpGet: { path: /health, port: 8080 }
            periodSeconds: 10
          - type: Liveness
            httpGet: { path: /health, port: 8080 }
            periodSeconds: 30
```

### 8. Verify

```bash
curl https://<app-fqdn>/health          # -> ok
curl -i -X POST https://<app-fqdn>/mcp  # -> 401 (no bearer token)
```

## Connecting clients

The MCP endpoint is `https://<app-fqdn>/mcp`.

### Claude Code

Claude Code speaks Streamable HTTP directly and can send a header, so no proxy is needed:

```bash
claude mcp add --transport http microsoftTodoCloud https://<app-fqdn>/mcp \
  --header "Authorization: Bearer <bearer-token>"
```

### Claude Desktop

Claude Desktop's custom-connector UI only supports OAuth, so it can't send a static bearer
token. Instead, run `local-proxy.js` as a stdio server and let it make the HTTP calls. It
needs a built copy of this repo (or a global install) on the client machine, but no sign-in
and no `.env`.

```json
{
  "mcpServers": {
    "microsoftTodoCloud": {
      "command": "node",
      "args": ["C:\\path\\to\\microsoft-todo-mcp-server\\dist\\local-proxy.js", "https://<app-fqdn>/mcp"],
      "env": { "MSTODO_AUTH_TOKEN": "<bearer-token>" }
    }
  }
}
```

Use `"command": "node"` with an absolute script path, not `npx`. Claude Desktop on Windows
launches `npx` through `cmd.exe`, which mangles Program Files paths. That is why
`mcp-remote` failed every time with `'C:\Program' is not recognized as an internal or
external command`, and why `local-proxy.js` exists.

The bearer token sits in plain text in this config file. Treat the file accordingly.

## Behaviour and limitations

- **One session at a time.** The server holds a single MCP session. A new client's
  `initialize` closes the previous session, so two clients used concurrently will keep
  evicting each other. This is deliberate for a single-user tool.
- **Cold starts take ~26s** if the app scales to zero (measured from replica scheduling to
  container start). Set `--min-replicas 1` if that matters more than cost.
- **The bearer token is the only gate.** Anyone holding it gets whatever `MSTODO_ACCESS_MODE`
  allows on your account. Rotate it by updating the `bearer-token` secret, restarting the
  revision, and updating every client.
- **`local-proxy.js` is request/response only.** It relays each stdio message as a POST and
  doesn't open the standalone GET/SSE stream, so server-initiated notifications aren't
  delivered. No current tool depends on them.
- **Re-seeding.** If the refresh token expires or is revoked (password change, admin revoke,
  long inactivity), the server can no longer get tokens. Sign in locally again and repeat
  steps 2–3; no redeploy is needed.

## Troubleshooting

- **Startup exits with `MSTODO_BEARER_TOKEN is required`**: the secret reference didn't
  resolve; check `az containerapp secret list`.
- **`MSTODO_TOKEN_STORE=key-vault requires ...`**: `MSTODO_KEYVAULT_URL` or
  `MSTODO_TOKEN_SECRET_NAME` is missing.
- **403 from Key Vault**: the managed identity lacks a role on the vault, or has a read-only
  one (refresh write-back needs _Secrets Officer_).
- **Decryption errors at startup or on the first tool call**: the server's `MSTODO_ENCRYPTION_KEY` doesn't match the
  key used when seeding. Re-run step 2 with the server's key.
- **Claude Desktop connector fails**: check
  `%LOCALAPPDATA%\Claude\Logs\mcp-server-<name>.log`; `local-proxy.js` logs non-2xx responses
  from the remote server there.
- **Probe failures**: query `ContainerAppSystemLogs_CL` in the environment's Log Analytics
  workspace for `Probe of ... failed`.
