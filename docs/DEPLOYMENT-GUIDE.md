# Deployment Guide

The deployment model is simple:

1. Build `McpServerGraphApi.exe`.
2. Deploy that exe through Software Center.
3. The exe installs itself and configures the detected MCP client.

## Build The Exe

From the repository root:

```powershell
.\build.ps1
```

The deployable files are created here:

```text
artifacts\dist\<version>\
```

Software Center only needs these files:

```text
McpServerGraphApi.exe
version.txt
```

The build also creates `McpServerGraphApi.exe.sha256` and `release.json` for release validation. Keep them with the package source even if Software Center only executes the `.exe`.

The generated `artifacts` folder is intentionally ignored by git. Attach release files to a GitHub Release or copy them to the Software Center package source instead of committing build output.

For production releases, sign the executable during build:

```powershell
.\build.ps1 -SignToolPath "C:\Program Files (x86)\Windows Kits\10\bin\x64\signtool.exe" -CodeSigningCertificateThumbprint "<thumbprint>" -RequireSignature
```

## Software Center Deployment

Create a Software Center application that runs in the user context.

Use one central IT-managed Entra app registration for all users. Users do not need permission to create app registrations in the tenant.

Install command:

```powershell
McpServerGraphApi.exe --install --tenant-id "<tenant_id>" --client-id "<public_client_application_id>" --national-cloud Global --graph-scopes "User.Read Mail.Read Calendars.Read Files.Read.All"
```

The installer uses `--client auto` by default:

- If Claude Code CLI exists at `%USERPROFILE%\.local\bin\claude.exe`, it runs `claude mcp add --scope user`.
- If Claude Desktop exists, it writes `%APPDATA%\Claude\claude_desktop_config.json`.
- If both are present, both are configured.

To force a target:

```powershell
McpServerGraphApi.exe --install --client claude-code --tenant-id "<tenant_id>" --client-id "<public_client_application_id>" --national-cloud Global
McpServerGraphApi.exe --install --client desktop --tenant-id "<tenant_id>" --client-id "<public_client_application_id>" --national-cloud Global
McpServerGraphApi.exe --install --client both --tenant-id "<tenant_id>" --client-id "<public_client_application_id>" --national-cloud Global
```

Uninstall command:

```powershell
McpServerGraphApi.exe --uninstall
```

Detection rule:

```text
%LOCALAPPDATA%\Programs\McpServerGraphApi\version.txt
```

The file content should match the version in `src\McpServerGraphApi\McpServerGraphApi.csproj`.

The install command copies the exe to:

```text
%LOCALAPPDATA%\Programs\McpServerGraphApi\McpServerGraphApi.exe
```

Install roots are intentionally fixed. The exe will only install to the per-user folder above, or to `%ProgramFiles%\McpServerGraphApi` when `--machine` is used.

It also configures the MCP client:

```text
%APPDATA%\Claude\claude_desktop_config.json
%USERPROFILE%\.local\bin\claude.exe mcp add --scope user
```

## Local Test

Build the exe first:

```powershell
.\build.ps1
```

Then run the installer command from the output folder:

```powershell
cd .\artifacts\dist\1.0.0
.\McpServerGraphApi.exe --install --tenant-id "<tenant_id>" --client-id "<public_client_application_id>" --national-cloud Global
```

Restart Claude Desktop from the system tray.

To test sign-in without installing:

```powershell
.\McpServerGraphApi.exe --test-auth --tenant-id "<tenant_id>" --client-id "<public_client_application_id>" --national-cloud Global
```

If `--test-auth` fails, fix that first. Claude Desktop should not be used to debug Entra/browser sign-in errors because it only reports that the MCP transport disconnected.

For a minimal account-only test, use only `User.Read`:

```powershell
.\McpServerGraphApi.exe --test-auth --tenant-id "<tenant_id>" --client-id "<public_client_application_id>" --national-cloud Global --graph-scopes "User.Read"
```

## First Use

After Software Center installs the app:

1. Quit Claude Desktop from the system tray, or close the current Claude Code session.
2. Start Claude Desktop or a new Claude Code session.
3. Start a new chat.
4. Ask Claude for Microsoft 365 data, for example:

```text
List my recent unread emails.
```

The MCP server starts automatically when Claude Desktop needs it. Authentication happens when the Graph tool is called. If sign-in fails, the tool returns an error instead of closing the MCP server.

Within one running MCP process, the server reuses the Microsoft Graph access token until it is close to expiry. If you see repeated browser prompts, check the MCP logs for different `PID` values; that means the MCP client is starting a new server process for each request.

## If Claude Says The MCP Disconnected

Run this outside Claude:

```powershell
cd .\artifacts\dist\1.0.0
.\McpServerGraphApi.exe --test-auth --tenant-id "<tenant_id>" --client-id "<public_client_application_id>" --national-cloud Global
```

Common causes:

- The Entra app is not configured as a public client.
- `http://localhost` is missing as a redirect URI.
- The account cannot consent to the requested Graph permissions.
- The wrong tenant ID, client ID, or national cloud was used.

## Graph Permissions

Add Microsoft Graph **Delegated** permissions to the app registration.

Minimum permission:

```text
User.Read
```

Recommended starting set:

```text
User.Read
Mail.Read
Calendars.Read
Files.Read.All
```

The exe writes these scopes into Claude Desktop as `GRAPH_SCOPES`. Keep the app registration permissions and the `--graph-scopes` install argument aligned.

## Notes

Use `Global` for normal Microsoft 365 tenants.

Supported `--national-cloud` values:

- `Global`
- `US_GOV`
- `US_GOV_DOD`
- `China`
- `Germany`

No client secret is used.
