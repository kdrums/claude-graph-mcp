using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;

namespace McpClient
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            // create MCP server config
            var clientTransport = new StdioClientTransport(new()
            {
                Name = "my-tenant-local",
                Command = Environment.GetEnvironmentVariable("MCP_SERVER_COMMAND") ?? "McpServerGraphApi.exe",
                Arguments = [],
                EnvironmentVariables = new Dictionary<string, string>()
                {
                    { "TENANT_ID", Environment.GetEnvironmentVariable("TENANT_ID") ?? "<tenant_id>" },
                    { "CLIENT_ID", Environment.GetEnvironmentVariable("CLIENT_ID") ?? "<public_client_application_id>" },
                    { "NATIONAL_CLOUD", Environment.GetEnvironmentVariable("NATIONAL_CLOUD") ?? "Global" }
                }
            });

            // create MCP client
            await using var mcpClient = await ModelContextProtocol.Client.McpClient.CreateAsync(clientTransport, new()
            {
                ClientInfo = new Implementation
                {
                    Name = "MCP Console App Client",
                    Version = "1.0.0"
                }
            });

            Console.WriteLine($"Connected to server: {mcpClient.ServerInfo.Name} ({mcpClient.ServerInfo.Version})");

            var tools = await mcpClient.ListToolsAsync();
            foreach (var tool in tools)
            {
                Console.WriteLine($"Tool: {tool.Name}, {tool.Description}");
            }

            using var anthropicService = new AnthropicService(tools);
            using var azureService = new AzureOpenAIService(tools);

            Console.WriteLine("Enter 1 for anthropic, 2 for azure (or 'exit' to quit):");

            while (Console.ReadLine() is string input && !"exit".Equals(input, StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Enter a message (or 'service' to change service or 'exit' to quit):");

                while (Console.ReadLine() is string message && !"service".Equals(message, StringComparison.OrdinalIgnoreCase))
                {
                    if (message.Equals("exit", StringComparison.OrdinalIgnoreCase))
                    {
                        input = "exit";
                        break;
                    }

                    if (!string.IsNullOrWhiteSpace(message))
                    {
                        try
                        {
                            if (input == "1")
                            {
                                var chatResponse = await anthropicService.GetResponseAsync(message);
                                Console.Write(chatResponse);
                            }
                            else if (input == "2")
                            {
                                var chatResponse = await azureService.GetResponseAsync(message);
                                Console.Write(chatResponse);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"An error occurred: {ex.Message}");
                        }
                        Console.WriteLine();
                    }

                    Console.WriteLine("Enter a message (or 'service' to change service):");
                }

                if (input.Equals("exit", StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }

                Console.WriteLine("Enter 1 for anthropic, 2 for azure (or 'exit' to quit):");
            }
        }
    }
}
