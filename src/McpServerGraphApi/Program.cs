using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using ModelContextProtocol.Server;
using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

internal class Program
{
    private const string ExeName = "McpServerGraphApi.exe";
    private const string ServerName = "graphApi";
    private const string ClientAuto = "auto";
    private const string ClientDesktop = "desktop";
    private const string ClientClaudeCode = "claude-code";
    private const string ClientBoth = "both";
    private static readonly string[] DefaultGraphScopes = ["User.Read", "Mail.Read", "Calendars.Read", "Files.Read.All"];

    static async Task<int> Main(string[] args)
    {
        if (args.Length > 0)
        {
            return await CliCommand.RunAsync(args);
        }

        var tenantId = GetRequiredEnvironmentVariable("TENANT_ID", Console.Error);
        var clientId = GetRequiredEnvironmentVariable("CLIENT_ID", Console.Error);
        if (tenantId is null || clientId is null)
        {
            return 2;
        }

        GraphCloudSettings graphCloud;
        try
        {
            graphCloud = GraphCloudSettings.FromEnvironment(
                Environment.GetEnvironmentVariable("NATIONAL_CLOUD", EnvironmentVariableTarget.Process));
        }
        catch (InvalidOperationException ex)
        {
            Console.Error.WriteLine($"[MCP Graph API] {ex.Message}");
            return 2;
        }

        var graphScopes = GetGraphScopes(Environment.GetEnvironmentVariable("GRAPH_SCOPES", EnvironmentVariableTarget.Process));
        await RunMcpServerAsync(tenantId, clientId, graphCloud, graphScopes);
        return 0;
    }

    private static async Task RunMcpServerAsync(string tenantId, string clientId, GraphCloudSettings graphCloud, string[] graphScopes)
    {
        var credential = CreateCredential(tenantId, clientId, graphCloud);

        LogInfo($"Starting MCP server for {graphCloud.Name}. PID: {Environment.ProcessId}. Authentication will happen when a Graph tool is called.");

        var builder = Host.CreateEmptyApplicationBuilder(settings: null);

        builder.Services.AddMcpServer()
            .WithStdioServerTransport()
            .WithTools<GraphApiTool>();

        builder.Services.AddSingleton(_ =>
            new HttpClient(new GraphAuthenticationHandler(credential, graphScopes))
            {
                BaseAddress = new Uri(graphCloud.GraphEndpoint)
            });

        var app = builder.Build();

        await app.RunAsync();
    }

    private static TokenCredential CreateCredential(string tenantId, string clientId, GraphCloudSettings graphCloud)
    {
        return new PersistentInteractiveBrowserCredential(tenantId, clientId, graphCloud);
    }

    private static string[] GetGraphScopes(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return DefaultGraphScopes;
        }

        return value
            .Split([' ', ';', ','], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private sealed class GraphAuthenticationHandler : DelegatingHandler
    {
        private static readonly TimeSpan TokenRefreshBuffer = TimeSpan.FromMinutes(5);
        private readonly TokenCredential credential;
        private readonly string[] scopes;
        private readonly SemaphoreSlim tokenLock = new(1, 1);
        private AccessToken? cachedToken;

        public GraphAuthenticationHandler(TokenCredential credential, string[] scopes)
            : base(new HttpClientHandler())
        {
            this.credential = credential;
            this.scopes = scopes;
        }

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            var token = await GetAccessTokenAsync(cancellationToken);
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.Token);
            return await base.SendAsync(request, cancellationToken);
        }

        private async Task<AccessToken> GetAccessTokenAsync(CancellationToken cancellationToken)
        {
            if (cachedToken is { } currentToken && IsUsable(currentToken))
            {
                return currentToken;
            }

            await tokenLock.WaitAsync(cancellationToken);
            try
            {
                if (cachedToken is { } lockedToken && IsUsable(lockedToken))
                {
                    return lockedToken;
                }

                LogInfo("Requesting Microsoft Graph access token.");
                cachedToken = await credential.GetTokenAsync(new TokenRequestContext(scopes), cancellationToken);
                LogInfo($"Access token acquired. Expires: {cachedToken.Value.ExpiresOn:O}");
                return cachedToken.Value;
            }
            finally
            {
                tokenLock.Release();
            }
        }

        private static bool IsUsable(AccessToken token)
        {
            return token.ExpiresOn > DateTimeOffset.UtcNow.Add(TokenRefreshBuffer);
        }
    }

    private sealed class PersistentInteractiveBrowserCredential : TokenCredential
    {
        private const string TokenCacheName = "mcp-graph-server";
        private readonly SemaphoreSlim authenticationLock = new(1, 1);
        private readonly string authenticationRecordPath;
        private readonly InteractiveBrowserCredential credential;
        private bool hasAuthenticationRecord;

        public PersistentInteractiveBrowserCredential(string tenantId, string clientId, GraphCloudSettings graphCloud)
        {
            authenticationRecordPath = GetAuthenticationRecordPath(tenantId, clientId, graphCloud);
            var authenticationRecord = TryLoadAuthenticationRecord(authenticationRecordPath);
            hasAuthenticationRecord = authenticationRecord is not null;

            credential = new InteractiveBrowserCredential(new InteractiveBrowserCredentialOptions
            {
                TenantId = tenantId,
                ClientId = clientId,
                AuthorityHost = graphCloud.AuthorityHost,
                AuthenticationRecord = authenticationRecord,
                TokenCachePersistenceOptions = new TokenCachePersistenceOptions { Name = TokenCacheName }
            });
        }

        public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            return GetTokenAsync(requestContext, cancellationToken).AsTask().GetAwaiter().GetResult();
        }

        public override async ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            if (!hasAuthenticationRecord)
            {
                LogInfo($"No saved authentication account record found at {authenticationRecordPath}.");
                await AuthenticateAndPersistAsync(requestContext, cancellationToken);
            }

            try
            {
                return await credential.GetTokenAsync(requestContext, cancellationToken);
            }
            catch (AuthenticationRequiredException)
            {
                await AuthenticateAndPersistAsync(requestContext, cancellationToken);
                return await credential.GetTokenAsync(requestContext, cancellationToken);
            }
        }

        private async Task AuthenticateAndPersistAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
        {
            await authenticationLock.WaitAsync(cancellationToken);
            try
            {
                LogInfo("Interactive sign-in is required. The signed-in account will be remembered for future MCP processes.");
                var authenticationRecord = await credential.AuthenticateAsync(requestContext, cancellationToken);

                Directory.CreateDirectory(Path.GetDirectoryName(authenticationRecordPath)!);
                await using var stream = File.Open(authenticationRecordPath, FileMode.Create, FileAccess.Write, FileShare.None);
                await authenticationRecord.SerializeAsync(stream, cancellationToken);
                hasAuthenticationRecord = true;
                LogInfo($"Saved authentication account record to {authenticationRecordPath}.");
            }
            finally
            {
                authenticationLock.Release();
            }
        }

        private static AuthenticationRecord? TryLoadAuthenticationRecord(string path)
        {
            if (!File.Exists(path))
            {
                return null;
            }

            try
            {
                using var stream = File.OpenRead(path);
                var authenticationRecord = AuthenticationRecord.Deserialize(stream);
                LogInfo($"Loaded authentication account record from {path}.");
                return authenticationRecord;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
            {
                LogInfo($"Saved authentication account record could not be read from {path}. Interactive sign-in may be required. Error: {ex.GetType().Name}.");
                return null;
            }
        }

        private static string GetAuthenticationRecordPath(string tenantId, string clientId, GraphCloudSettings graphCloud)
        {
            var fileName = string.Join(
                '-',
                SafeFileNamePart(graphCloud.Name),
                SafeFileNamePart(tenantId),
                SafeFileNamePart(clientId)) + ".authrecord.json";

            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "McpServerGraphApi",
                "auth",
                fileName);
        }

        private static string SafeFileNamePart(string value)
        {
            var safe = Regex.Replace(value, @"[^A-Za-z0-9_.-]", "_");
            return string.IsNullOrWhiteSpace(safe) ? "default" : safe;
        }
    }

    private static string? GetRequiredEnvironmentVariable(string name, TextWriter error)
    {
        var value = Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
        if (!string.IsNullOrWhiteSpace(value))
        {
            return value;
        }

        error.WriteLine($"[MCP Graph API] Missing required environment variable: {name}");
        return null;
    }

    private static void LogInfo(string message)
    {
        var line = $"[MCP Graph API] {DateTimeOffset.Now:O} {message}";
        Console.Error.WriteLine(line);

        try
        {
            var logDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "McpServerGraphApi",
                "logs");
            Directory.CreateDirectory(logDirectory);
            File.AppendAllText(Path.Combine(logDirectory, "mcp-server-graphApi.log"), line + Environment.NewLine);
        }
        catch
        {
            // Logging must never interfere with MCP stdio.
        }
    }

    private sealed class CliCommand
    {
        private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

        public static async Task<int> RunAsync(string[] args)
        {
            var options = CliOptions.Parse(args);

            if (options.HasFlag("help") || options.HasFlag("?"))
            {
                WriteHelp(Console.Out);
                return 0;
            }

            if (options.HasFlag("install"))
            {
                return Install(options);
            }

            if (options.HasFlag("uninstall"))
            {
                return Uninstall(options);
            }

            if (options.HasFlag("test-auth"))
            {
                return await TestAuthAsync(options);
            }

            Console.Error.WriteLine("Unknown command. Use --help.");
            return 2;
        }

        private static int Install(CliOptions options)
        {
            var tenantId = options.GetRequired("tenant-id", Console.Error);
            var clientId = options.GetRequired("client-id", Console.Error);
            if (tenantId is null || clientId is null)
            {
                return 2;
            }

            if (!ValidateClientOption(options))
            {
                return 2;
            }

            var nationalCloud = options.Get("national-cloud", "Global");
            var graphScopes = GetGraphScopes(options.Get("graph-scopes", string.Join(' ', DefaultGraphScopes)));
            try
            {
                _ = GraphCloudSettings.FromEnvironment(nationalCloud);
            }
            catch (InvalidOperationException ex)
            {
                Console.Error.WriteLine(ex.Message);
                return 2;
            }

            var installRoot = ResolveInstallRoot(options, Console.Error);
            if (installRoot is null)
            {
                return 2;
            }

            Directory.CreateDirectory(installRoot);

            var sourceExe = GetCurrentExecutablePath();
            var targetExe = Path.Combine(installRoot, ExeName);

            if (!Path.GetFullPath(sourceExe).Equals(Path.GetFullPath(targetExe), StringComparison.OrdinalIgnoreCase))
            {
                File.Copy(sourceExe, targetExe, overwrite: true);
            }

            File.WriteAllText(Path.Combine(installRoot, "version.txt"), GetVersion());
            var configuredClients = ConfigureMcpClients(options, targetExe, tenantId, clientId, nationalCloud, graphScopes);

            Console.WriteLine($"Installed {ExeName} to {installRoot}");
            foreach (var configuredClient in configuredClients)
            {
                Console.WriteLine(configuredClient);
            }

            Console.WriteLine("Restart your MCP client before using the MCP server.");
            return 0;
        }

        private static int Uninstall(CliOptions options)
        {
            if (!ValidateClientOption(options))
            {
                return 2;
            }

            foreach (var removedClient in RemoveMcpClientConfigs(options))
            {
                Console.WriteLine(removedClient);
            }

            var installRoot = ResolveInstallRoot(options, Console.Error);
            if (installRoot is null)
            {
                return 2;
            }

            var targetExe = Path.Combine(installRoot, ExeName);
            var currentExe = GetCurrentExecutablePath();

            if (Path.GetFullPath(currentExe).Equals(Path.GetFullPath(targetExe), StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"Close this process, then remove the install folder manually if needed: {installRoot}");
                return 0;
            }

            if (Directory.Exists(installRoot))
            {
                Directory.Delete(installRoot, recursive: true);
            }

            Console.WriteLine($"Removed {installRoot}");
            return 0;
        }

        private static async Task<int> TestAuthAsync(CliOptions options)
        {
            var tenantId = options.GetRequired("tenant-id", Console.Error);
            var clientId = options.GetRequired("client-id", Console.Error);
            if (tenantId is null || clientId is null)
            {
                return 2;
            }

            GraphCloudSettings graphCloud;
            try
            {
                graphCloud = GraphCloudSettings.FromEnvironment(options.Get("national-cloud", "Global"));
            }
            catch (InvalidOperationException ex)
            {
                Console.Error.WriteLine(ex.Message);
                return 2;
            }

            try
            {
                var graphScopes = GetGraphScopes(options.Get("graph-scopes", string.Join(' ', DefaultGraphScopes)));
                Console.WriteLine($"Testing Microsoft 365 sign-in for {graphCloud.Name} with scopes: {string.Join(' ', graphScopes)}");
                var credential = CreateCredential(tenantId, clientId, graphCloud);
                await credential.GetTokenAsync(new TokenRequestContext(graphScopes), CancellationToken.None);
                Console.WriteLine("Authentication succeeded.");
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Authentication failed.");
                Console.Error.WriteLine(SafeExceptionMessage(ex));
                return 1;
            }
        }

        private static List<string> ConfigureMcpClients(CliOptions options, string command, string tenantId, string clientId, string nationalCloud, string[] graphScopes)
        {
            var messages = new List<string>();
            var targetClients = GetTargetClients(options);

            if (targetClients.Contains(ClientDesktop))
            {
                WriteClaudeDesktopConfig(command, tenantId, clientId, nationalCloud, graphScopes);
                messages.Add($"Updated Claude Desktop MCP config: {GetClaudeDesktopConfigPath()}");
            }

            if (targetClients.Contains(ClientClaudeCode))
            {
                var claudeCodeExe = GetClaudeCodeExePath();
                if (claudeCodeExe is null)
                {
                    messages.Add("Claude Code CLI was selected but was not found at %USERPROFILE%\\.local\\bin\\claude.exe.");
                }
                else if (WriteClaudeCodeConfig(claudeCodeExe, command, tenantId, clientId, nationalCloud, graphScopes, out var error))
                {
                    messages.Add("Updated Claude Code MCP config using 'claude mcp add --scope user'.");
                }
                else
                {
                    messages.Add($"Claude Code MCP config was not updated: {error}");
                }
            }

            if (messages.Count == 0)
            {
                messages.Add("No MCP client was detected. The server exe was installed, but client configuration was not changed.");
            }

            return messages;
        }

        private static List<string> RemoveMcpClientConfigs(CliOptions options)
        {
            var messages = new List<string>();
            var targetClients = GetTargetClients(options, includeExistingConfigs: true);

            if (targetClients.Contains(ClientDesktop))
            {
                RemoveClaudeDesktopConfigEntry();
                messages.Add("Removed graphApi from Claude Desktop MCP config.");
            }

            if (targetClients.Contains(ClientClaudeCode))
            {
                var claudeCodeExe = GetClaudeCodeExePath();
                if (claudeCodeExe is null)
                {
                    messages.Add("Claude Code CLI was selected but was not found at %USERPROFILE%\\.local\\bin\\claude.exe.");
                }
                else if (RemoveClaudeCodeConfig(claudeCodeExe, out var error))
                {
                    messages.Add("Removed graphApi from Claude Code MCP config.");
                }
                else
                {
                    messages.Add($"Claude Code MCP config was not updated: {error}");
                }
            }

            if (messages.Count == 0)
            {
                messages.Add("No MCP client configuration was found to remove.");
            }

            return messages;
        }

        private static bool ValidateClientOption(CliOptions options)
        {
            var value = options.Get("client", ClientAuto).Trim().ToLowerInvariant();
            if (value is ClientAuto or ClientDesktop or "claude-desktop" or ClientClaudeCode or "cli" or "claude-cli" or "code" or ClientBoth)
            {
                return true;
            }

            Console.Error.WriteLine("Unsupported --client value. Use auto, desktop, claude-code, or both.");
            return false;
        }

        private static HashSet<string> GetTargetClients(CliOptions options, bool includeExistingConfigs = false)
        {
            var value = options.Get("client", ClientAuto).Trim().ToLowerInvariant();
            var targets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (value is ClientDesktop or "claude-desktop")
            {
                targets.Add(ClientDesktop);
                return targets;
            }

            if (value is ClientClaudeCode or "cli" or "claude-cli" or "code")
            {
                targets.Add(ClientClaudeCode);
                return targets;
            }

            if (value == ClientBoth)
            {
                targets.Add(ClientDesktop);
                targets.Add(ClientClaudeCode);
                return targets;
            }

            if (IsClaudeCodeInstalled())
            {
                targets.Add(ClientClaudeCode);
            }

            if (Directory.Exists(Path.GetDirectoryName(GetClaudeDesktopConfigPath())!) || File.Exists(GetClaudeDesktopConfigPath()))
            {
                targets.Add(ClientDesktop);
            }

            if (includeExistingConfigs && File.Exists(GetClaudeDesktopConfigPath()))
            {
                targets.Add(ClientDesktop);
            }

            if (targets.Count == 0)
            {
                targets.Add(ClientDesktop);
            }

            return targets;
        }

        private static void WriteClaudeDesktopConfig(string command, string tenantId, string clientId, string nationalCloud, string[] graphScopes)
        {
            var configPath = GetClaudeDesktopConfigPath();
            Directory.CreateDirectory(Path.GetDirectoryName(configPath)!);

            JsonObject config;
            if (File.Exists(configPath))
            {
                var existing = JsonNode.Parse(File.ReadAllText(configPath));
                config = existing as JsonObject ?? [];
            }
            else
            {
                config = [];
            }

            if (config["mcpServers"] is not JsonObject servers)
            {
                servers = [];
                config["mcpServers"] = servers;
            }

            servers[ServerName] = new JsonObject
            {
                ["command"] = command,
                ["args"] = new JsonArray(),
                ["env"] = new JsonObject
                {
                    ["TENANT_ID"] = tenantId,
                    ["CLIENT_ID"] = clientId,
                    ["NATIONAL_CLOUD"] = nationalCloud,
                    ["GRAPH_SCOPES"] = string.Join(' ', graphScopes)
                }
            };

            File.WriteAllText(configPath, config.ToJsonString(JsonOptions));
        }

        private static bool WriteClaudeCodeConfig(string claudeCodeExe, string command, string tenantId, string clientId, string nationalCloud, string[] graphScopes, out string error)
        {
            _ = RemoveClaudeCodeConfig(claudeCodeExe, out _);

            var serverConfig = new JsonObject
            {
                ["type"] = "stdio",
                ["command"] = command,
                ["args"] = new JsonArray(),
                ["env"] = new JsonObject
                {
                    ["TENANT_ID"] = tenantId,
                    ["CLIENT_ID"] = clientId,
                    ["NATIONAL_CLOUD"] = nationalCloud,
                    ["GRAPH_SCOPES"] = string.Join(' ', graphScopes)
                }
            };

            var args = new List<string>
            {
                "mcp",
                "add-json",
                "--scope",
                "user",
                ServerName,
                serverConfig.ToJsonString()
            };

            return RunProcess(claudeCodeExe, args, out error);
        }

        private static bool RemoveClaudeCodeConfig(string claudeCodeExe, out string error)
        {
            return RunProcess(claudeCodeExe, ["mcp", "remove", "--scope", "user", ServerName], out error);
        }

        private static bool RunProcess(string fileName, IEnumerable<string> args, out string error)
        {
            using var process = new Process();
            process.StartInfo.FileName = fileName;
            foreach (var arg in args)
            {
                process.StartInfo.ArgumentList.Add(arg);
            }

            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardError = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.CreateNoWindow = true;

            process.Start();
            var stdout = process.StandardOutput.ReadToEnd();
            var stderr = process.StandardError.ReadToEnd();
            process.WaitForExit();

            error = string.IsNullOrWhiteSpace(stderr) ? stdout.Trim() : stderr.Trim();
            return process.ExitCode == 0;
        }

        private static void RemoveClaudeDesktopConfigEntry()
        {
            var configPath = GetClaudeDesktopConfigPath();
            if (!File.Exists(configPath))
            {
                return;
            }

            var existing = JsonNode.Parse(File.ReadAllText(configPath));
            if (existing is not JsonObject config || config["mcpServers"] is not JsonObject servers)
            {
                return;
            }

            servers.Remove(ServerName);
            File.WriteAllText(configPath, config.ToJsonString(JsonOptions));
        }

        private static string GetClaudeDesktopConfigPath()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Claude",
                "claude_desktop_config.json");
        }

        private static bool IsClaudeCodeInstalled()
        {
            return GetClaudeCodeExePath() is not null;
        }

        private static string? GetClaudeCodeExePath()
        {
            var path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                ".local",
                "bin",
                "claude.exe");

            return File.Exists(path) ? path : null;
        }

        private static string GetDefaultInstallRoot(CliOptions options)
        {
            if (options.HasFlag("machine"))
            {
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "McpServerGraphApi");
            }

            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Programs",
                "McpServerGraphApi");
        }

        private static string? ResolveInstallRoot(CliOptions options, TextWriter error)
        {
            var requestedRoot = options.Get("install-root", GetDefaultInstallRoot(options));
            var fullRoot = Path.GetFullPath(Environment.ExpandEnvironmentVariables(requestedRoot));
            var allowedRoots = GetAllowedInstallRoots();

            if (allowedRoots.Any(root => string.Equals(root, fullRoot, StringComparison.OrdinalIgnoreCase)))
            {
                return fullRoot;
            }

            error.WriteLine("Unsupported install root. This executable only installs to the approved per-user or machine-wide MCP folder.");
            error.WriteLine($"Per-user : {allowedRoots.UserRoot}");
            error.WriteLine($"Machine  : {allowedRoots.MachineRoot}");
            return null;
        }

        private static AllowedInstallRoots GetAllowedInstallRoots()
        {
            return new AllowedInstallRoots(
                Path.GetFullPath(Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Programs",
                    "McpServerGraphApi")),
                Path.GetFullPath(Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                    "McpServerGraphApi")));
        }

        private static string GetCurrentExecutablePath()
        {
            return Environment.ProcessPath
                ?? Process.GetCurrentProcess().MainModule?.FileName
                ?? Path.Combine(AppContext.BaseDirectory, ExeName);
        }

        private static string GetVersion()
        {
            return typeof(Program).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
                ?? typeof(Program).Assembly.GetName().Version?.ToString()
                ?? "unknown";
        }

        private static void WriteHelp(TextWriter writer)
        {
            writer.WriteLine("McpServerGraphApi");
            writer.WriteLine();
            writer.WriteLine("Default:");
            writer.WriteLine("  McpServerGraphApi.exe");
            writer.WriteLine("    Starts the MCP stdio server. MCP clients use this mode.");
            writer.WriteLine();
            writer.WriteLine("Install for Software Center or local use:");
            writer.WriteLine("  McpServerGraphApi.exe --install --tenant-id <tenant> --client-id <client> [--client auto|desktop|claude-code|both] [--national-cloud Global] [--graph-scopes \"User.Read Mail.Read Calendars.Read Files.Read.All\"]");
            writer.WriteLine();
            writer.WriteLine("Machine-wide install root:");
            writer.WriteLine("  McpServerGraphApi.exe --install --machine --tenant-id <tenant> --client-id <client>");
            writer.WriteLine();
            writer.WriteLine("Test sign-in:");
            writer.WriteLine("  McpServerGraphApi.exe --test-auth --tenant-id <tenant> --client-id <client> [--graph-scopes \"User.Read\"]");
            writer.WriteLine();
            writer.WriteLine("Uninstall:");
            writer.WriteLine("  McpServerGraphApi.exe --uninstall");
            writer.WriteLine();
            writer.WriteLine("Install roots are fixed to approved folders:");
            writer.WriteLine(@"  User    %LOCALAPPDATA%\Programs\McpServerGraphApi");
            writer.WriteLine(@"  Machine %ProgramFiles%\McpServerGraphApi");
        }

        private sealed record AllowedInstallRoots(string UserRoot, string MachineRoot)
        {
            public bool Any(Func<string, bool> predicate)
            {
                return predicate(UserRoot) || predicate(MachineRoot);
            }
        }
    }

    private sealed class CliOptions
    {
        private readonly Dictionary<string, string?> values;

        private CliOptions(Dictionary<string, string?> values)
        {
            this.values = values;
        }

        public static CliOptions Parse(string[] args)
        {
            var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

            for (var i = 0; i < args.Length; i++)
            {
                var arg = args[i];
                if (!arg.StartsWith("--", StringComparison.Ordinal))
                {
                    continue;
                }

                var keyValue = arg[2..].Split('=', 2);
                var key = keyValue[0];
                string? value = null;

                if (keyValue.Length == 2)
                {
                    value = keyValue[1];
                }
                else if (i + 1 < args.Length && !args[i + 1].StartsWith("--", StringComparison.Ordinal))
                {
                    value = args[++i];
                }

                values[key] = value;
            }

            return new CliOptions(values);
        }

        public bool HasFlag(string name)
        {
            return values.ContainsKey(name);
        }

        public string Get(string name, string defaultValue)
        {
            return values.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value)
                ? value
                : defaultValue;
        }

        public string? GetRequired(string name, TextWriter error)
        {
            if (values.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value))
            {
                return value;
            }

            error.WriteLine($"Missing required option: --{name}");
            return null;
        }
    }

    private sealed record GraphCloudSettings(string Name, string GraphEndpoint, Uri AuthorityHost)
    {
        public static GraphCloudSettings FromEnvironment(string? value)
        {
            var cloud = string.IsNullOrWhiteSpace(value) ? "Global" : value.Trim();
            var normalized = cloud.Replace("-", "_", StringComparison.Ordinal).ToUpperInvariant();

            return normalized switch
            {
                "GLOBAL" => new("Global", "https://graph.microsoft.com", AzureAuthorityHosts.AzurePublicCloud),
                "US_GOV" or "USGOV" => new("US_GOV", "https://graph.microsoft.us", AzureAuthorityHosts.AzureGovernment),
                "US_GOV_DOD" or "USGOVDOD" => new("US_GOV_DOD", "https://dod-graph.microsoft.us", AzureAuthorityHosts.AzureGovernment),
                "CHINA" => new("China", "https://microsoftgraph.chinacloudapi.cn", AzureAuthorityHosts.AzureChina),
                "GERMANY" => new("Germany", "https://graph.microsoft.de", new Uri("https://login.microsoftonline.de/")),
                _ => throw new InvalidOperationException(
                    "Unsupported NATIONAL_CLOUD value. Use one of: Global, US_GOV, US_GOV_DOD, China, Germany.")
            };
        }
    }

    [McpServerToolType]
    public class GraphApiTool
    {
        private const string FilterParam = "$filter";
        private const string SearchParam = "$search";
        private const string CountParam = "$count";

        [McpServerTool(Name = "graph-api", Title = "Access my Microsoft 365 data")]
        [Description("Call any Microsoft Graph API endpoint as the signed-in user. Use /me/ paths for personal data: /me/messages, /me/calendars, /me/drive/root/children, /me/mailFolders, etc.")]
        public static async Task<string> GetGraphApiData(HttpClient client,
            [Description("Microsoft Graph API path (e.g. '/me/messages', '/me/drive/root/children', '/me/calendar/events')")] string path,
            [Description("OData query parameters like $filter, $count, $search, $orderby, $select. Collection of keys and values")] Dictionary<string, object>? queryParameters = null,
            [Description("HTTP method: GET, POST, PUT, PATCH or DELETE")] string method = "GET",
            [Description("Request body (optional, for POST/PUT/PATCH)")] string body = "",
            [Description("Graph API version: 'v1.0' or 'beta'")] string graphVersion = "v1.0")
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path))
                {
                    return "Error: path is required.";
                }

                if (Uri.TryCreate(path, UriKind.Absolute, out _))
                {
                    return "Error: path must be a Microsoft Graph relative path, for example '/me/messages'.";
                }

                var version = graphVersion.Equals("beta", StringComparison.OrdinalIgnoreCase) ? "beta" : "v1.0";
                var graphPath = path.StartsWith('/') ? path : $"/{path}";
                var requestUrl = $"/{version}{graphPath}";
                if (queryParameters?.Count > 0)
                {
                    var queryString = string.Join("&", queryParameters
                        .Where(kvp => !string.IsNullOrWhiteSpace(kvp.Key) && kvp.Value is not null)
                        .Select(kvp => $"{Uri.EscapeDataString(kvp.Key)}={Uri.EscapeDataString(FormatQueryValue(kvp.Value))}"));
                    requestUrl += $"?{queryString}";
                }

                var httpMethod = ParseHttpMethod(method);
                if (httpMethod is null)
                {
                    return $"Error: unsupported HTTP method '{method}'. Use GET, POST, PUT, PATCH or DELETE.";
                }

                var requestMessage = new HttpRequestMessage(httpMethod, requestUrl)
                {
                    Headers = { { "Accept", graphPath.EndsWith(CountParam, StringComparison.OrdinalIgnoreCase) ? "text/plain" : "application/json" } }
                };

                if (ContainsQueryParameter(queryParameters, FilterParam) ||
                    ContainsQueryParameter(queryParameters, SearchParam) ||
                    ContainsQueryParameter(queryParameters, CountParam) ||
                    graphPath.EndsWith(CountParam, StringComparison.OrdinalIgnoreCase))
                {
                    requestMessage.Headers.Add("ConsistencyLevel", "eventual");
                }

                if ((httpMethod == HttpMethod.Post || httpMethod == HttpMethod.Put || httpMethod == HttpMethod.Patch) &&
                    !string.IsNullOrWhiteSpace(body))
                {
                    requestMessage.Content = new StringContent(body, Encoding.UTF8, "application/json");
                }

                using var response = await client.SendAsync(requestMessage);
                var content = await response.Content.ReadAsStringAsync();

                return response.IsSuccessStatusCode
                    ? content
                    : SafeGraphError(response, content);
            }
            catch (Exception ex)
            {
                return SafeExceptionMessage(ex);
            }
        }

        [McpServerTool(Name = "graph-me", Title = "Get my Microsoft 365 account")]
        [Description("Get the signed-in user's Microsoft 365 profile from Microsoft Graph /v1.0/me. Use this first to verify Graph authentication.")]
        public static async Task<string> GetMyAccount(HttpClient client)
        {
            try
            {
                using var requestMessage = new HttpRequestMessage(HttpMethod.Get, "/v1.0/me?$select=id,displayName,givenName,surname,userPrincipalName,mail,jobTitle,department,officeLocation,mobilePhone,businessPhones");
                requestMessage.Headers.Add("Accept", "application/json");

                using var response = await client.SendAsync(requestMessage);
                var content = await response.Content.ReadAsStringAsync();

                return response.IsSuccessStatusCode
                    ? content
                    : SafeGraphError(response, content);
            }
            catch (Exception ex)
            {
                return SafeExceptionMessage(ex);
            }
        }

        private static string SafeGraphError(HttpResponseMessage response, string content)
        {
            var details = TryReadGraphError(content);
            var requestId = TryGetHeader(response, "request-id") ?? details.RequestId;
            var clientRequestId = TryGetHeader(response, "client-request-id") ?? details.ClientRequestId;

            var builder = new StringBuilder();
            builder.Append("Graph request failed.");
            builder.Append($" Status: {(int)response.StatusCode} {response.ReasonPhrase ?? response.StatusCode.ToString()}.");

            if (!string.IsNullOrWhiteSpace(details.Code))
            {
                builder.Append($" Code: {details.Code}.");
            }

            if (!string.IsNullOrWhiteSpace(requestId))
            {
                builder.Append($" RequestId: {requestId}.");
            }

            if (!string.IsNullOrWhiteSpace(clientRequestId))
            {
                builder.Append($" ClientRequestId: {clientRequestId}.");
            }

            builder.Append(" Full Graph error content was withheld from the model. Run --test-auth locally or review Graph/Azure sign-in logs for details.");
            return builder.ToString();
        }

        private static GraphErrorDetails TryReadGraphError(string content)
        {
            if (string.IsNullOrWhiteSpace(content))
            {
                return new GraphErrorDetails(null, null, null);
            }

            try
            {
                using var document = JsonDocument.Parse(content);
                if (!document.RootElement.TryGetProperty("error", out var error))
                {
                    return new GraphErrorDetails(null, null, null);
                }

                var code = ReadJsonString(error, "code");
                var requestId = ReadJsonString(error, "request-id");
                var clientRequestId = ReadJsonString(error, "client-request-id");

                if (error.TryGetProperty("innerError", out var innerError))
                {
                    requestId ??= ReadJsonString(innerError, "request-id");
                    clientRequestId ??= ReadJsonString(innerError, "client-request-id");
                    clientRequestId ??= ReadJsonString(innerError, "clientRequestId");
                }

                return new GraphErrorDetails(SanitizeIdentifier(code), SanitizeIdentifier(requestId), SanitizeIdentifier(clientRequestId));
            }
            catch (JsonException)
            {
                return new GraphErrorDetails(null, null, null);
            }
        }

        private static string? ReadJsonString(JsonElement element, string propertyName)
        {
            return element.TryGetProperty(propertyName, out var property) && property.ValueKind == JsonValueKind.String
                ? property.GetString()
                : null;
        }

        private static string? TryGetHeader(HttpResponseMessage response, string headerName)
        {
            return response.Headers.TryGetValues(headerName, out var values)
                ? SanitizeIdentifier(values.FirstOrDefault())
                : null;
        }

        private static string? SanitizeIdentifier(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            var sanitized = Regex.Replace(value.Trim(), @"[^A-Za-z0-9_.:@-]", string.Empty);
            return sanitized.Length > 96 ? sanitized[..96] : sanitized;
        }

        private sealed record GraphErrorDetails(string? Code, string? RequestId, string? ClientRequestId);

        private static HttpMethod? ParseHttpMethod(string method)
        {
            return method?.Trim().ToUpperInvariant() switch
            {
                "GET" => HttpMethod.Get,
                "POST" => HttpMethod.Post,
                "PUT" => HttpMethod.Put,
                "PATCH" => HttpMethod.Patch,
                "DELETE" => HttpMethod.Delete,
                _ => null
            };
        }

        private static bool ContainsQueryParameter(Dictionary<string, object>? queryParameters, string key)
        {
            return queryParameters?.Keys.Any(k => string.Equals(k, key, StringComparison.OrdinalIgnoreCase)) == true;
        }

        private static string FormatQueryValue(object value)
        {
            return value switch
            {
                JsonElement element => element.ValueKind switch
                {
                    JsonValueKind.String => element.GetString() ?? string.Empty,
                    JsonValueKind.Number => element.GetRawText(),
                    JsonValueKind.True => "true",
                    JsonValueKind.False => "false",
                    JsonValueKind.Null => string.Empty,
                    _ => element.GetRawText()
                },
                bool boolValue => boolValue ? "true" : "false",
                _ => Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty
            };
        }
    }

    private static string SafeExceptionMessage(Exception ex)
    {
        var aadstsCode = Regex.Match(ex.ToString(), @"AADSTS\d{5,}", RegexOptions.IgnoreCase).Value;
        var exceptionName = ex.GetType().Name;

        if (!string.IsNullOrWhiteSpace(aadstsCode))
        {
            return $"Authentication failed. Code: {aadstsCode}. Detailed authentication output was withheld from the model. Run --test-auth locally and review Azure sign-in logs.";
        }

        if (ex is AuthenticationFailedException)
        {
            return "Authentication failed. Detailed authentication output was withheld from the model. Run --test-auth locally and review Azure sign-in logs.";
        }

        return $"MCP Graph server error: {exceptionName}. Detailed exception output was withheld from the model.";
    }
}
