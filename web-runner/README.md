# Entra Secret Rotate web runner

Browser UI for Entra ID app secret rotation and Barracuda XDR ATR permissions. Runs next to XOA/GA on the lab webapps LXC.

## Lab topology

| App | Host | Port | URL |
|-----|------|------|-----|
| XOA | webapps `10.23.20.110` | 8765 | `https://xoa.lab.river-run.com` |
| GA | same | 8775 | `https://ga.lab.river-run.com` |
| ESR | same | 8785 | `https://esr.lab.river-run.com` |

Caddy on OPNsense terminates TLS and reverse-proxies to the LXC port (same pattern as XOA/GA).

## Shared Graph apps (EOA store)

ESR does **not** create/update/delete Graph app registrations. Use XOA for that.

- Prefix: `EOA` (same as XOA)
- Store file: `/var/lib/exchangeonlineanalyzer/.exchangeonlineanalyzer/graph-apps.json`
- Module: `/opt/exchangeonlineanalyzer/Modules/GraphAppCredential.psm1` (`ESR_XOA_ROOT`)

Set `EOA_CREDENTIAL_STORE` in the systemd unit so Linux does not look under an empty `HOME` path.

## Local run

```powershell
cd web-runner
.\Start-WebRunner.ps1
# http://127.0.0.1:8785/
```

LAN / lab bind:

```powershell
.\Start-WebRunner.ps1 -NoBrowser -ListenLab
```

## Deploy to webapps LXC

From a machine that can `ssh webapps`:

```powershell
# Copy runner (from repo root entrasecretrotate/)
ssh webapps "mkdir -p /opt/entrasecretrotate/web-runner"
scp -r web-runner/Start-WebRunner.ps1 web-runner/www webapps:/opt/entrasecretrotate/web-runner/
scp web-runner/deploy/entrasecretrotate.service webapps:/etc/systemd/system/entrasecretrotate.service
ssh webapps "systemctl daemon-reload && systemctl enable --now entrasecretrotate.service && systemctl status entrasecretrotate.service --no-pager"
```

Health check:

```bash
curl -s http://127.0.0.1:8785/api/health
curl -s http://127.0.0.1:8785/api/tenants | head
```

## Caddy (OPNsense)

Add a site matching XOA/GA, for example:

```
esr.lab.river-run.com {
    reverse_proxy 10.23.20.110:8785
}
```

DNS: `esr.lab.river-run.com` → same VIP / hairpin as `xoa.lab` / `ga.lab`.

## Env vars

| Variable | Purpose |
|----------|---------|
| `EOA_CREDENTIAL_STORE` | Path to shared `graph-apps.json` |
| `ESR_XOA_ROOT` | Path to XOA install (GraphAppCredential module) |
| `ESR_PUBLIC_URL` | Public base URL (behind Caddy) |
| `ESR_SECRET_NAME_TEMPLATE` | Secret display name; default `secret{YEAR}` |
| `ESR_ADD_REVOKE_SESSIONS` | Set `0` to omit `User.RevokeSessions.All` from ATR |

## UI features (v1)

- Connect to a tenant from the shared EOA list
- Find expired secrets / select any application
- Generate new secret (shown once + copy)
- Delete oldest expired secret
- Add Barracuda XDR ATR application permissions + admin consent

App registration lifecycle (Add / Update / Delete / Clear) remains in XOA.
