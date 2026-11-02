using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace BckPPT.Mcp;

// ═══════════════════════════════════════════════════════════════════════════
// MCP 클라이언트(Claude Code·Codex) 사용자 범위 등록. 셀박스 트레이가 묻고, 플러그인이 등록한다.
// 규약: bckhwp 레포 docs/mcp-client-registration.md. 원본은 bckhwp src/BckHwpMcp/McpClientRegistration.cs이고
// BckPPT는 사본을 쓴다 — 고칠 때 두 곳을 같이 맞춘다(쓰는 쪽은 속성으로만 바꾼다).
//
//   --mcp-status                 클라이언트별 설치·등록 상태(JSON)
//   --mcp-register <client>      등록. stdio면 프록시를 고정 경로에 복사하고 그 경로를 등록한다
//   --mcp-unregister <client>    등록 해제. stdio면 남은 등록이 없을 때 고정 경로의 프록시도 지운다
//   --mcp-refresh                업데이트 뒤: 고정 경로로 등록된 곳이 있으면 프록시를 새 판으로 바꾼다
//
// 두 방식: stdio(ProxyDirectory·ProxyFileName — BckHWP)와 http(Url — 앱 안의 서버, BckPPT).
// stdio에서 고정 경로를 쓰는 이유: 트레이는 업데이트 때 플러그인 폴더를 통째로 이름 바꿔 교체한다. 다른 레포의 세션이
// 그 폴더 안 프록시를 쓰고 있으면 교체가 막힌다. 고정 경로의 실행 중인 프록시는 이름만 바꿔 두고 새 파일을 놓는다.
// ═══════════════════════════════════════════════════════════════════════════
public sealed class McpClientRegistration
{
    public required string ClaudeName { get; init; }      // ~/.claude.json mcpServers 키
    public required string CodexName { get; init; }       // ~/.codex/config.toml [mcp_servers.<이름>]
    public string? ProxyDirectory { get; init; }          // stdio: 고정 경로 폴더
    public string? ProxyFileName { get; init; }
    public string? Url { get; init; }                     // http: 앱이 띄우는 MCP 주소
    /// <summary>고정 경로의 프록시가 읽는 옆 파일. 호스트처럼 플러그인 폴더에 남는 실행 파일의 위치를 적는다.</summary>
    public string? CompanionPath { get; init; }
    public string CompanionFileName { get; init; } = "host.txt";

    public string? ProxyPath => ProxyDirectory is null || ProxyFileName is null ? null : Path.Combine(ProxyDirectory, ProxyFileName);

    /// <summary>클라이언트 설정에 적는 값: http면 주소, stdio면 고정 경로의 프록시.</summary>
    public string Target => Url ?? ProxyPath ?? throw new InvalidOperationException("Url 또는 ProxyDirectory·ProxyFileName을 정하세요.");

    // 시험용: 사용자 폴더를 바꾸면 그 아래 설정만 읽고 쓰며 CLI는 쓰지 않는다.
    static string? TestHome => Environment.GetEnvironmentVariable("BCKSOFT_MCP_CONFIG_HOME");
    static string Home => TestHome ?? Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
    static string ClaudeJson => Path.Combine(Home, ".claude.json");
    static string CodexHome => TestHome is null && Environment.GetEnvironmentVariable("CODEX_HOME") is { Length: > 0 } h ? h : Path.Combine(Home, ".codex");
    static string CodexToml => Path.Combine(CodexHome, "config.toml");

    public static readonly string[] Clients = ["claude-code", "codex"];

    /// <summary>명령줄을 처리했으면 종료 코드, 이 파일의 명령이 아니면 null.</summary>
    public int? TryRun(string[] args)
    {
        if (args.Length == 0 || !args[0].StartsWith("--mcp-", StringComparison.Ordinal)) return null;
        try
        {
            switch (args[0])
            {
                case "--mcp-status" when args.Length == 1: break;
                case "--mcp-register" when args.Length == 2: Register(Client(args[1])); break;
                case "--mcp-unregister" when args.Length == 2: Unregister(Client(args[1])); break;
                case "--mcp-refresh" when args.Length == 1: Refresh(); break;
                default: throw new ArgumentException("사용법: --mcp-status | --mcp-register <client> | --mcp-unregister <client> | --mcp-refresh");
            }
            Print(Status());
            return 0;
        }
        catch (Exception ex)
        {
            Print(new JsonObject { ["ok"] = false, ["error"] = ex.Message });
            return 1;
        }
    }

    static string Client(string id) => Clients.Contains(id) ? id : throw new ArgumentException("알 수 없는 클라이언트: " + id);

    // 콘솔 코드 페이지가 아니라 언제나 UTF-8 (규약). 트레이가 한국어 오류 글을 그대로 보여 준다.
    static void Print(JsonNode node)
    {
        using var stdout = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false));
        stdout.WriteLine(node.ToJsonString(new JsonSerializerOptions { WriteIndented = true, Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping }));
    }

    public JsonObject Status()
    {
        var clients = new JsonArray();
        foreach (string id in Clients)
        {
            string? command = ReadTarget(id);
            clients.Add(new JsonObject
            {
                ["id"] = id,
                ["name"] = id == "claude-code" ? "Claude Code" : "Codex",
                ["available"] = Available(id),
                ["registered"] = command is not null,
                ["command"] = command,
                // false면 다른 경로로 등록되어 있다 — 트레이가 덮어쓸지 묻는다.
                ["current"] = command is not null && Same(command, Target),
            });
        }
        return new JsonObject { ["ok"] = true, ["transport"] = Url is null ? "stdio" : "http", ["target"] = Target, ["clients"] = clients };
    }

    public void Register(string client)
    {
        if (Url is null) InstallProxy();
        if (client == "claude-code") WriteClaude(Target);
        else WriteCodex(Target);
        if (!Same(ReadTarget(client) ?? "", Target)) throw new IOException(client + " 설정에 등록하지 못했습니다.");
    }

    public void Unregister(string client)
    {
        if (client == "claude-code") WriteClaude(null);
        else WriteCodex(null);
        if (ReadTarget(client) is not null) throw new IOException(client + " 설정에서 지우지 못했습니다.");
        if (Url is null && !Clients.Any(c => Same(ReadTarget(c) ?? "", Target))) RemoveProxy();
    }

    /// <summary>http는 주소가 그대로라 할 일이 없다.</summary>
    public void Refresh()
    {
        if (Url is null && Clients.Any(c => Same(ReadTarget(c) ?? "", Target))) InstallProxy();
    }

    // ── 고정 경로의 프록시 ───────────────────────────────────────────────

    void InstallProxy()
    {
        string self = Environment.ProcessPath!;
        string proxy = Target, directory = ProxyDirectory!;
        Directory.CreateDirectory(directory);
        foreach (string old in Directory.GetFiles(directory, ProxyFileName + ".old-*"))
            try { File.Delete(old); } catch { }   // 아직 도는 세션의 것은 남는다
        if (!SamePath(self, proxy) && !SameFile(self, proxy))
        {
            string staging = proxy + ".new";
            File.Copy(self, staging, overwrite: true);
            try { File.Move(staging, proxy, overwrite: true); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // 실행 중인 파일은 덮어쓸 수 없지만 이름은 바꿀 수 있다. 떠 있는 세션은 옛 파일로 계속 돈다.
                File.Move(proxy, proxy + ".old-" + DateTime.Now.ToString("yyyyMMddHHmmss"));
                File.Move(staging, proxy);
            }
        }
        string companion = Path.Combine(directory, CompanionFileName);
        if (CompanionPath is not null) File.WriteAllText(companion, CompanionPath, new UTF8Encoding(false));
    }

    void RemoveProxy()
    {
        if (SamePath(Environment.ProcessPath!, Target)) return;
        foreach (string file in Directory.Exists(ProxyDirectory) ? Directory.GetFiles(ProxyDirectory!) : [])
            try { File.Delete(file); } catch { }   // 도는 세션의 프록시는 다음 정리 때 지운다
    }

    /// <summary>고정 경로의 프록시가 옆 파일에서 읽는 위치. 없거나 가리키는 파일이 없으면 null.</summary>
    public static string? ReadCompanion(string directory, string fileName = "host.txt")
    {
        try
        {
            string path = File.ReadAllText(Path.Combine(directory, fileName)).Trim();
            return path.Length > 0 && File.Exists(path) ? path : null;
        }
        catch { return null; }
    }

    static bool SameFile(string a, string b)
    {
        if (!File.Exists(b) || new FileInfo(a).Length != new FileInfo(b).Length) return false;
        using var sha = System.Security.Cryptography.SHA256.Create();
        using var fa = File.OpenRead(a);
        using var fb = File.OpenRead(b);
        return sha.ComputeHash(fa).AsSpan().SequenceEqual(sha.ComputeHash(fb));
    }

    // 주소는 글자 그대로(끝의 / 무시), 경로는 정규화해서 비교한다.
    static bool Same(string a, string b) => IsUrl(a) || IsUrl(b)
        ? string.Equals(a.TrimEnd('/'), b.TrimEnd('/'), StringComparison.OrdinalIgnoreCase)
        : SamePath(a, b);

    static bool IsUrl(string value) =>
        value.StartsWith("http://", StringComparison.OrdinalIgnoreCase) || value.StartsWith("https://", StringComparison.OrdinalIgnoreCase);

    static bool SamePath(string a, string b)
    {
        try { return string.Equals(Path.GetFullPath(a), Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase); }
        catch { return false; }
    }

    // ── 클라이언트 설정 ─────────────────────────────────────────────────

    static bool Available(string client) => client == "claude-code"
        ? FindCli("claude") is not null || File.Exists(ClaudeJson) || Directory.Exists(Path.Combine(Home, ".claude"))
        : FindCli("codex") is not null || Directory.Exists(CodexHome);

    /// <summary>그 이름으로 등록된 명령(stdio) 또는 주소(http). 없으면 null.</summary>
    string? ReadTarget(string client) => client == "claude-code" ? ReadClaude() : ReadCodex();

    string? ReadClaude()
    {
        try
        {
            var entry = File.Exists(ClaudeJson) ? JsonNode.Parse(File.ReadAllText(ClaudeJson))?["mcpServers"]?[ClaudeName] : null;
            return (string?)entry?["command"] ?? (string?)entry?["url"];
        }
        catch (Exception ex) when (ex is JsonException or InvalidOperationException) { return null; }
    }

    // 떠 있는 Claude Code도 ~/.claude.json을 고쳐 쓴다. CLI가 있으면 CLI에 맡기고, 없거나 결과가 안 맞을 때만 직접 쓴다.
    void WriteClaude(string? command)
    {
        if (FindCli("claude") is { } cli)
        {
            RunCli(cli, "mcp", "remove", "--scope", "user", ClaudeName);
            if (command is null || (IsUrl(command)
                    ? RunCli(cli, "mcp", "add", "--scope", "user", "--transport", "http", ClaudeName, command)
                    : RunCli(cli, "mcp", "add", "--scope", "user", ClaudeName, "--", command)))
                if (SameOrBothNull(ReadClaude(), command)) return;
        }
        EditFile(ClaudeJson, raw =>
        {
            var config = (raw.Length > 0 ? JsonNode.Parse(raw) as JsonObject : null) ?? new JsonObject();
            var servers = config["mcpServers"] as JsonObject;
            if (command is null) servers?.Remove(ClaudeName);
            else
            {
                if (servers is null) config["mcpServers"] = servers = new JsonObject();
                servers[ClaudeName] = IsUrl(command)
                    ? new JsonObject { ["type"] = "http", ["url"] = command }
                    : new JsonObject { ["type"] = "stdio", ["command"] = command, ["args"] = new JsonArray() };
            }
            return config.ToJsonString(new JsonSerializerOptions { WriteIndented = true });
        });
    }

    // 그 서버의 구역과 하위 구역([mcp_servers.<이름>.env] 등)을 다음 다른 구역 머리 전까지.
    Regex CodexSection => new($@"^\[mcp_servers\.(?:{Regex.Escape(CodexName)}|""{Regex.Escape(CodexName)}"")\][^\S\n]*\r?\n(?:(?!\[(?!mcp_servers\.{Regex.Escape(CodexName)}\.))[^\n]*\n?)*", RegexOptions.Multiline);

    string? ReadCodex()
    {
        if (!File.Exists(CodexToml)) return null;
        var section = CodexSection.Match(File.ReadAllText(CodexToml));
        if (!section.Success) return null;
        var value = Regex.Match(section.Value, @"^\s*(?:command|url)\s*=\s*(?:'([^']*)'|""((?:[^""\\]|\\.)*)"")", RegexOptions.Multiline);
        if (!value.Success) return null;
        return value.Groups[1].Success ? value.Groups[1].Value : Regex.Unescape(value.Groups[2].Value);
    }

    void WriteCodex(string? command)
    {
        if (FindCli("codex") is { } cli)
        {
            RunCli(cli, "mcp", "remove", CodexName);
            if (command is null || (IsUrl(command)
                    ? RunCli(cli, "mcp", "add", CodexName, "--url", command)
                    : RunCli(cli, "mcp", "add", CodexName, "--", command)))
                if (SameOrBothNull(ReadCodex(), command)) return;
        }
        // 그 이름의 [mcp_servers.…] 구역만 바꾼다. 나머지 글은 그대로 둔다.
        EditFile(CodexToml, raw =>
        {
            string text = CodexSection.Replace(raw, "");
            if (command is null) return text;
            if (text.Length > 0 && !text.EndsWith('\n')) text += Environment.NewLine;
            string body = IsUrl(command) ? $"url = '{command}'" : $"command = '{command}'{Environment.NewLine}args = []";
            return text + $"[mcp_servers.{CodexName}]{Environment.NewLine}{body}{Environment.NewLine}";
        });
    }

    static bool SameOrBothNull(string? actual, string? expected) =>
        expected is null ? actual is null : actual is not null && Same(actual, expected);

    // 같은 폴더의 임시 파일에 쓰고 바꿔 끼운다 — 반쯤 쓴 설정 파일이 남지 않게.
    static void EditFile(string path, Func<string, string> edit)
    {
        string raw = File.Exists(path) ? File.ReadAllText(path) : "";
        string next = edit(raw);
        if (next == raw) return;
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        string temp = path + ".bcksoft-" + Guid.NewGuid().ToString("N")[..8];
        File.WriteAllText(temp, next, new UTF8Encoding(false));
        File.Move(temp, path, overwrite: true);
    }

    static string? FindCli(string name)
    {
        if (TestHome is not null) return null;
        string profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var dirs = (Environment.GetEnvironmentVariable("PATH") ?? "").Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries)
            .Append(Path.Combine(profile, ".local", "bin"))
            .Append(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "npm"));
        foreach (string dir in dirs)
            foreach (string ext in new[] { ".exe", ".cmd" })
            {
                try { string p = Path.Combine(dir.Trim('"'), name + ext); if (File.Exists(p)) return p; }
                catch (ArgumentException) { }
            }
        return null;
    }

    static bool RunCli(string cli, params string[] args)
    {
        var start = new ProcessStartInfo { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true, RedirectStandardInput = true };
        if (cli.EndsWith(".cmd", StringComparison.OrdinalIgnoreCase))
        {
            start.FileName = Environment.GetEnvironmentVariable("ComSpec") ?? "cmd.exe";
            start.ArgumentList.Add("/d"); start.ArgumentList.Add("/c"); start.ArgumentList.Add(cli);
        }
        else start.FileName = cli;
        foreach (string a in args) start.ArgumentList.Add(a);
        try
        {
            using var process = Process.Start(start)!;
            process.StandardInput.Close();
            _ = process.StandardOutput.ReadToEndAsync();   // 비워 두지 않으면 출력이 많을 때 멈춘다
            _ = process.StandardError.ReadToEndAsync();
            if (!process.WaitForExit(30_000)) { try { process.Kill(entireProcessTree: true); } catch { } return false; }
            return process.ExitCode == 0;
        }
        catch { return false; }
    }
}
