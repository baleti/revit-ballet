# Security Notes

Short section — the threat model in CLAUDE.md ("local dev/automation only, don't expose,
don't share the token") is honest and the implementation matches it. These notes are for
awareness, ordered by relevance; none are urgent given the stated model.

## Standing observations

1. **The server is arbitrary code execution by design.** Any process (or user) on the
   Windows machine that can read `%APPDATA%\revit-ballet\network\token` can execute
   arbitrary C# inside Revit with full model access — and, since scripts run as the
   Revit user, full user-level access to the machine. That's the intended trust
   boundary (same as any local process), but it means: the token file's protection *is*
   the security of the system. It inherits normal user-profile ACLs, which is adequate
   for a single-user dev machine. Do not run Revit elevated with the server up
   (already documented).

2. **Localhost binding is the load-bearing control.** Confirmed: `127.0.0.1` binding
   plus TLS with self-signed certs. TLS adds little here (localhost traffic), and the
   clients bypass validation anyway (10 call sites — see code-smells.md #2), so the
   real controls are the bind address and the token. If port-forwarding from the dev
   machine (as the current workflow does over SSH), the SSH tunnel is doing the actual
   transport security — that's fine.

3. **Certificate bypass belongs in one place.** Not a vulnerability under this model,
   but centralizing it in `NetworkClient` (duplication.md) enables an easy upgrade
   later: pin the known self-signed cert's thumbprint instead of `=> true`, so a
   different process squatting a port in the 23717–23817 range can't impersonate a
   session to the clients.

4. **The `documents` registry is trusted input.** InNetwork commands read hostname/port
   from a world-writable-by-user CSV and then send the token to whatever it lists. Same
   trust boundary as the token file itself, so consistent — just be aware the registry
   is part of the attack surface if the model ever changes.

5. **Script timeout ≠ cancellation.** The 30s timeout returns an error to the caller,
   but a hung script on the Revit UI thread (via ExternalEvent) can't actually be
   aborted — the timeout abandons it. Known Revit constraint; worth documenting in
   CLAUDE.md so agents don't retry-hammer a wedged session.

## If the model ever changes (multi-user office, exposed network)

Would need: per-session tokens with rotation, cert pinning (see #3), an allowlist of
script capabilities or a review gate, and moving the registry/token out of
world-user-readable AppData. Not worth building until then.
