using System.Diagnostics;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using BckPPT.Mcp;

// ═══════════════════════════════════════════════════════════════════════════
// BckPptMcp — stdio → HTTP 프록시
//
// MCP 클라이언트(stdio) ←→ 이 프로세스 ←→ BckPPT 앱의 MCP 서버(http://localhost:38701/mcp)
//
// BckPPT가 떠 있지 않으면 띄우고 서버가 열릴 때까지 기다린다. 서버는 세션이 없으므로(요청마다 독립)
// 한 줄을 POST 하나로 넘기고 답을 한 줄로 돌려준다. 앱이 중간에 다시 시작해도 다시 맞출 상태가 없다.
// ═══════════════════════════════════════════════════════════════════════════

// --mcp-*: 셀박스 트레이가 부르는 Claude Code·Codex 사용자 범위 등록 (bckhwp docs/mcp-client-registration.md).
var registration = new McpClientRegistration
{
    ClaudeName = "bckppt",
    CodexName = "bckppt",
    ProxyDirectory = ProxyDirectory(),
    ProxyFileName = "BckPptMcp.exe",
    CompanionPath = FindHost(),
};
if (registration.TryRun(args) is int registrationExit)
{
    Environment.ExitCode = registrationExit;
    return;
}

// BckPPT McpServer.Endpoint와 같은 주소. 시험은 BCKPPT_MCP_URL로 바꾼다.
var endpoint = new Uri(Environment.GetEnvironmentVariable("BCKPPT_MCP_URL") is { Length: > 0 } url ? url : "http://localhost:38701/mcp");
using var http = new HttpClient { Timeout = Timeout.InfiniteTimeSpan };   // 렌더링 같은 긴 도구가 있다. 끊는 것은 클라이언트가 한다.
var stdout = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { AutoFlush = true };
var stdoutLock = new object();
var startLock = new SemaphoreSlim(1, 1);

var inFlight = new List<Task>();
using (var stdin = new StreamReader(Console.OpenStandardInput(), Encoding.UTF8))
{
    while (await stdin.ReadLineAsync() is { } line)
    {
        if (line.Trim().Length == 0) continue;
        inFlight.RemoveAll(t => t.IsCompleted);
        inFlight.Add(Task.Run(() => ForwardAsync(line)));   // 도구 호출은 동시에 올 수 있다
    }
}
// stdin이 닫혔다(클라이언트가 끝낸다). 거의 끝난 답은 보내 주고 나간다.
await Task.WhenAny(Task.WhenAll(inFlight), Task.Delay(5_000));

async Task ForwardAsync(string line)
{
    string? reply;
    try { reply = await PostAsync(line); }
    catch (Exception ex) { reply = ErrorReply(line, ex.Message); }
    if (string.IsNullOrEmpty(reply)) return;   // 알림이었다
    // 클라이언트는 한 줄에 메시지 하나를 읽는다. 서버 답은 이미 한 줄이지만 줄바꿈이 섞여도 깨지지 않게.
    if (reply.Contains('\n')) reply = JsonNode.Parse(reply)!.ToJsonString();
    lock (stdoutLock) stdout.WriteLine(reply);
}

async Task<string?> PostAsync(string line)
{
    for (int attempt = 0; ; attempt++)
    {
        try
        {
            using var content = new StringContent(line, Encoding.UTF8, "application/json");
            using var response = await http.PostAsync(endpoint, content);
            if (response.StatusCode == System.Net.HttpStatusCode.Accepted) return null;
            string body = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode) throw new IOException($"BckPPT MCP 서버가 HTTP {(int)response.StatusCode}로 답했습니다.");
            return body;
        }
        // 연결조차 못 했다 = 요청이 서버에 닿지 않았다. 앱을 띄우고 한 번 더 보낸다(두 번 실행될 일은 없다).
        catch (HttpRequestException ex) when (ex.HttpRequestError == HttpRequestError.ConnectionError)
        {
            if (attempt > 0) throw new IOException("BckPPT MCP 서버에 연결할 수 없습니다.", ex);
            await EnsureHostAsync();
        }
        catch (HttpRequestException ex) when (ex.HttpRequestError != HttpRequestError.ConnectionError)
        {
            throw new IOException("BckPPT와 연결이 끊겨 이 요청의 결과를 알 수 없습니다. 문서 상태를 확인한 뒤 다시 시도하세요.", ex);
        }
    }
}

// 앱을 띄우고 MCP 서버가 열릴 때까지 기다린다. 동시에 온 요청들은 한 번만 띄운다.
async Task EnsureHostAsync()
{
    await startLock.WaitAsync();
    try
    {
        if (await ListeningAsync()) return;
        string host = FindHost() ?? throw new IOException("BckPPT.exe를 찾을 수 없습니다. 셀박스 트레이에서 BckPPT 플러그인을 설치하세요.");
        bool running = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(host)).Length > 0;
        if (!running)
        {
            Log($"BckPPT가 꺼져 있어 띄웁니다: {host}");
            // 셸로 띄운다 — 직접 띄우면 앱이 이 프록시의 stdin·stdout 핸들을 물려받아 클라이언트가 끝을 못 본다.
            Process.Start(new ProcessStartInfo(host) { UseShellExecute = true, WorkingDirectory = Path.GetDirectoryName(host)! });
        }
        // 이미 떠 있으면 막 시작하는 중일 수 있다. 다시 띄우면 단일 인스턴스가 채팅창을 열 뿐이라 기다리기만 한다.
        var clock = Stopwatch.StartNew();
        var limit = TimeSpan.FromSeconds(running ? 10 : 60);
        while (clock.Elapsed < limit)
        {
            if (await ListeningAsync()) { Log("BckPPT MCP 서버에 연결했습니다"); return; }
            await Task.Delay(300);
        }
        throw new IOException(running
            ? "BckPPT는 실행 중이지만 MCP 서버가 열려 있지 않습니다. 트레이 아이콘에서 MCP 상태를 확인하세요(포트 사용 중일 수 있습니다)."
            : "BckPPT를 띄웠지만 MCP 서버가 열리지 않았습니다.");
    }
    finally { startLock.Release(); }
}

async Task<bool> ListeningAsync()
{
    try
    {
        using var tcp = new TcpClient();
        using var timeout = new CancellationTokenSource(1_000);
        await tcp.ConnectAsync(endpoint.Host, endpoint.Port, timeout.Token);
        return true;
    }
    catch { return false; }
}

// 요청(id가 있는 것)마다 JSON-RPC 오류로 답한다. 알림에는 답하지 않는다.
static string? ErrorReply(string line, string message)
{
    JsonNode? parsed;
    try { parsed = JsonNode.Parse(line); }
    catch (JsonException) { return Error(null, -32700, "JSON을 읽을 수 없습니다.").ToJsonString(); }
    if (parsed is JsonArray batch)
    {
        var replies = new JsonArray();
        foreach (var item in batch)
            if (item?["id"] is { } id && item["method"] is not null) replies.Add(Error(id.DeepClone(), -32000, message));
        return replies.Count == 0 ? null : replies.ToJsonString();
    }
    return parsed?["id"] is { } single && parsed["method"] is not null ? Error(single.DeepClone(), -32000, message).ToJsonString() : null;
}

static JsonObject Error(JsonNode? id, int code, string message) => new()
{
    ["jsonrpc"] = "2.0",
    ["id"] = id,
    ["error"] = new JsonObject { ["code"] = code, ["message"] = message },
};

// stderr는 MCP 클라이언트의 로그로 간다. 콘솔 코드 페이지가 아니라 UTF-8로.
static void Log(string message)
{
    lock (StdErr.Writer) StdErr.Writer.WriteLine($"[BckPPT proxy] {message}");
}

static string? FindHost()
{
    // 시험: 정한 파일만 본다(없으면 null) — 개발 PC의 진짜 BckPPT를 띄우지 않게.
    if (Environment.GetEnvironmentVariable("BCKPPT_HOST_PATH") is { Length: > 0 } fixedHost)
        return File.Exists(fixedHost) ? Path.GetFullPath(fixedHost) : null;

    string myDir = Path.GetDirectoryName(Environment.ProcessPath!) ?? ".";
    // 플러그인 폴더: <플러그인>\mcp\BckPptMcp.exe 옆의 위 폴더
    foreach (string candidate in new[] { Path.Combine(myDir, "BckPPT.exe"), Path.Combine(myDir, "..", "BckPPT.exe") })
        if (File.Exists(candidate)) return Path.GetFullPath(candidate);
    // 고정 경로에 복사된 프록시: 등록할 때 적어 둔 앱 위치
    if (McpClientRegistration.ReadCompanion(myDir) is { } registered) return registered;
    // 개발: 레포의 Debug 빌드(run.ps1)
    for (var dir = new DirectoryInfo(myDir); dir is not null; dir = dir.Parent)
    {
        string dev = Path.Combine(dir.FullName, "bin", "x64", "Debug", "net8.0-windows", "BckPPT.exe");
        if (File.Exists(dev)) return dev;
    }
    return null;
}

// Claude Code·Codex에 등록되는 프록시 자리. 플러그인 폴더 밖이라 트레이가 폴더를 바꿔 끼워도 막히지 않는다.
static string ProxyDirectory() =>
    Environment.GetEnvironmentVariable("BCKPPT_MCP_PROXY_DIR") is { Length: > 0 } dir ? dir : Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BckSoft", "BckPPT", "Proxy");

static class StdErr
{
    public static readonly StreamWriter Writer = new(Console.OpenStandardError(), new UTF8Encoding(false)) { AutoFlush = true };
}
