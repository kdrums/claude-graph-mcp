# Maintenance Plan

This project ships a self-contained Windows executable for Software Center.

## Release Flow

1. Update the version in `src\McpServerGraphApi\McpServerGraphApi.csproj`.
2. Build the release package:

```powershell
.\build.ps1
```

For production, build with Authenticode signing and require a valid signature:

```powershell
.\build.ps1 -SignToolPath "C:\Program Files (x86)\Windows Kits\10\bin\x64\signtool.exe" -CodeSigningCertificateThumbprint "<thumbprint>" -RequireSignature
```

3. Validate the executable and release metadata:

```powershell
.\artifacts\dist\<version>\McpServerGraphApi.exe --help
.\artifacts\dist\<version>\McpServerGraphApi.exe --test-auth --tenant-id "<tenant_id>" --client-id "<client_id>" --national-cloud Global --graph-scopes "User.Read"
Get-AuthenticodeSignature .\artifacts\dist\<version>\McpServerGraphApi.exe
Get-FileHash .\artifacts\dist\<version>\McpServerGraphApi.exe -Algorithm SHA256
Get-Content .\artifacts\dist\<version>\release.json
```

4. Commit source, docs, `version.txt`, `McpServerGraphApi.exe`, `McpServerGraphApi.exe.sha256`, and `release.json`.
5. Copy `artifacts\dist\<version>` to the Software Center source location.
6. Update the Software Center detection rule to the new `version.txt` content.

For clean binary provenance, commit source changes first, then build from that source commit and commit only the generated release artifacts. `release.json` records the source revision, dirty-source flag, hash, signature status, and build time.

## Entra App Registration

Use one central IT-managed app registration.

Do not require users to create app registrations.

Required app registration settings:

- Platform: Mobile and desktop applications.
- Redirect URI: `http://localhost`.
- Public client flow enabled.
- Microsoft Graph delegated permissions matching the deployed `GRAPH_SCOPES`.

Minimum delegated permission:

```text
User.Read
```

Recommended starting delegated permissions:

```text
User.Read
Mail.Read
Calendars.Read
Files.Read.All
```

Review permissions quarterly. Remove scopes that are not needed.

## Software Center Commands

Install command:

```powershell
McpServerGraphApi.exe --install --tenant-id "<tenant_id>" --client-id "<central_app_client_id>" --national-cloud Global --graph-scopes "User.Read Mail.Read Calendars.Read Files.Read.All"
```

Uninstall command:

```powershell
McpServerGraphApi.exe --uninstall
```

Use `--client auto` for normal deployment. It detects Claude Code CLI at `%USERPROFILE%\.local\bin\claude.exe` and Claude Desktop at `%APPDATA%\Claude`. Use `--client claude-code`, `--client desktop`, or `--client both` when a deployment collection should target a specific client.

Detection:

```text
%LOCALAPPDATA%\Programs\McpServerGraphApi\version.txt
```

The application should run in the user context because Claude Desktop configuration is user-scoped.

Claude Code configuration is also user-scoped and is managed by running `claude mcp add --scope user` from `%USERPROFILE%\.local\bin\claude.exe`.

The install location is fixed to `%LOCALAPPDATA%\Programs\McpServerGraphApi` for user installs and `%ProgramFiles%\McpServerGraphApi` for `--machine`. Do not use arbitrary install roots in Software Center.

## Operational Checks

After each release, test:

- `McpServerGraphApi.exe --help`
- `McpServerGraphApi.exe --test-auth ... --graph-scopes "User.Read"`
- Claude Desktop shows the `graphApi` MCP server as running.
- Claude Code `claude mcp list` shows the `graphApi` MCP server when CLI deployment is targeted.
- Prompt: `Use the graph-me tool from the Graph MCP server to get my Microsoft 365 account details.`

## Dependency Maintenance

Monthly:

- Run `dotnet list src\McpServerGraphApi\McpServerGraphApi.csproj package --outdated`.
- Review `Azure.Identity`, `Microsoft.Extensions.Hosting`, and `ModelContextProtocol` updates.
- Rebuild and re-test auth before deploying package updates.

Quarterly:

- Review Graph delegated permissions.
- Confirm the central app registration owner list is current.
- Confirm Software Center detection still targets the intended version file.
- Confirm the code-signing certificate is valid, has a documented renewal owner, and is not near expiry.
- Confirm checked-in release artifacts include a matching `.sha256` file and `release.json`.

## Security Maintenance

Each release:

- Require a valid Authenticode signature before production deployment.
- Verify `release.json` has `sourceDirty=false` for production builds.
- Compare `McpServerGraphApi.exe.sha256` with `Get-FileHash` before copying to the Software Center source share.
- Run a Claude prompt that exercises `graph-me` and confirm failures return sanitized status/code/request IDs only.

## Troubleshooting Pattern

If Claude reports a tool failure, test outside Claude first:

```powershell
McpServerGraphApi.exe --test-auth --tenant-id "<tenant_id>" --client-id "<central_app_client_id>" --national-cloud Global --graph-scopes "User.Read"
```

Common causes:

- Missing delegated Graph permission.
- Admin consent not granted.
- Wrong tenant ID or client ID.
- Redirect URI missing from the public client app registration.
- Conditional Access policy blocking the desktop sign-in.
- Repeated authentication prompts can mean the MCP client is starting multiple server processes. Check the MCP stderr log for different `PID` values.
