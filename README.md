# MCP Server for Microsoft Graph
Forked from https://github.com/MartinM85/mcp-server-graph-api/.

Model Context Protocol (MCP) stdio server for Microsoft Graph, built as a self-contained Windows executable.

The server exposes one tool, `graph-api`, that can call Microsoft Graph as the signed-in user. It is intended for MCP clients such as Claude Desktop and Claude Code CLI.

For central deployment and local testing, see [docs/DEPLOYMENT-GUIDE.md](docs/DEPLOYMENT-GUIDE.md).

For release ownership and upkeep, see [docs/MAINTENANCE.md](docs/MAINTENANCE.md).

## Authentication Model

This project uses delegated interactive sign-in through `InteractiveBrowserCredential`.

Create an Entra ID app registration as a public client:

- Add a platform under **Authentication**: **Mobile and desktop applications**.
- Add `http://localhost` as a redirect URI.
- Grant delegated Microsoft Graph permissions for the data you want the MCP client to access, for example `User.Read`, `Mail.Read`, `Calendars.Read`, or `Files.Read.All`.
- Admin consent may be required depending on your tenant policy and selected permissions.

No client secret is required or used.

## Build

Prerequisites:

- Windows x64
- PowerShell 5.1 or newer
- .NET 9 SDK

From the repository root:

```powershell
.\build.ps1
```

The deployable package is written to:

```text
artifacts\dist\<version>\
```

The package contains:

- `McpServerGraphApi.exe`
- `version.txt`
- `McpServerGraphApi.exe.sha256`
- `release.json`

The `artifacts\dist\<version>` release payload is generated locally. For public distribution, attach those files to a GitHub Release or copy them to your Software Center package source rather than committing generated binaries to source history.

Production releases should be Authenticode-signed:

```powershell
.\build.ps1 -SignToolPath "C:\Program Files (x86)\Windows Kits\10\bin\x64\signtool.exe" -CodeSigningCertificateThumbprint "<thumbprint>" -RequireSignature
```

## Executable Commands

Default MCP server mode, used by MCP clients:

```powershell
McpServerGraphApi.exe
```

Install and configure an MCP client:

```powershell
McpServerGraphApi.exe --install --tenant-id "<tenant_id>" --client-id "<public_client_application_id>" --national-cloud Global
```

Client selection is automatic by default. On a CLI-only machine, the installer detects Claude Code at:

```text
%USERPROFILE%\.local\bin\claude.exe
```

You can force a target client:

```powershell
McpServerGraphApi.exe --install --client claude-code --tenant-id "<tenant_id>" --client-id "<public_client_application_id>" --national-cloud Global
McpServerGraphApi.exe --install --client desktop --tenant-id "<tenant_id>" --client-id "<public_client_application_id>" --national-cloud Global
McpServerGraphApi.exe --install --client both --tenant-id "<tenant_id>" --client-id "<public_client_application_id>" --national-cloud Global
```

Uninstall:

```powershell
McpServerGraphApi.exe --uninstall
```

Test sign-in without installing:

```powershell
McpServerGraphApi.exe --test-auth --tenant-id "<tenant_id>" --client-id "<public_client_application_id>" --national-cloud Global
```

Optional scope override:

```powershell
McpServerGraphApi.exe --install --tenant-id "<tenant_id>" --client-id "<public_client_application_id>" --graph-scopes "User.Read Mail.Read Calendars.Read Files.Read.All"
```

## MCP Client Config

For Claude Desktop, the install command writes this file automatically:

```text
%APPDATA%\Claude\claude_desktop_config.json
```

For Claude Code CLI, the install command runs:

```powershell
claude mcp add --scope user ...
```

After install, restart Claude Desktop or start a new Claude Code session.

## Supported Clouds

Supported `--national-cloud` values:

- `Global`
- `US_GOV`
- `US_GOV_DOD`
- `China`
- `Germany`

Most tenants should use `Global`.

## Troubleshooting

Claude Desktop logs are in:

```text
%APPDATA%\Claude\logs
```

Claude Code MCP logs are under:

```text
%USERPROFILE%\.local
```

The server writes startup/authentication status to stderr so it does not interfere with MCP stdio JSON messages.
