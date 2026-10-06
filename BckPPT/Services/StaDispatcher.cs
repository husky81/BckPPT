using System.Collections.Concurrent;

namespace BckPPT.Services;

/// <summary>
/// 앱 수명 동안 유지되는 전용 STA 스레드.
/// 매번 new Thread()를 생성하는 대신 이 싱글턴을 통해 COM 호출을 직렬화한다.
/// </summary>
public sealed class StaDispatcher : IDisposable
{
    public static readonly StaDispatcher Instance = new();
    public static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(45);

    private readonly Thread _thread;
    private readonly BlockingCollection<(Action action, TaskCompletionSource tcs)> _queue = new();
    private volatile bool _isHealthy = true;
    public bool IsHealthy => _isHealthy && !_queue.IsAddingCompleted;

    private StaDispatcher()
    {
        _thread = new Thread(Run)
        {
            Name         = "COM-STA",
            IsBackground = true,
        };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    private void Run()
    {
        foreach (var (action, tcs) in _queue.GetConsumingEnumerable())
        {
            try   { action(); tcs.SetResult(); }
            catch (Exception ex) { tcs.SetException(ex); }
        }
    }

    public Task InvokeAsync(Action action)
    {
        if (!_isHealthy)
            throw new PptComUnavailableException("이전 PowerPoint 작업이 아직 끝나지 않았습니다. 끝나면 자동으로 다시 쓸 수 있으니 잠시 뒤 다시 시도해 주세요.");
        if (_queue.IsAddingCompleted)
            throw new ObjectDisposedException(nameof(StaDispatcher));
        var tcs = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        _queue.Add((action, tcs));
        return tcs.Task;
    }

    public async Task<T> InvokeAsync<T>(Func<T> func)
    {
        T result = default!;
        await InvokeAsync(() => { result = func(); }).ConfigureAwait(false);
        return result;
    }

    /// <summary>
    /// 기다리지 않고 줄 끝에 넣는다. 시간 초과로 격리된 동안에도 받아, 멈춘 작업이 끝난 뒤 차례로 실행한다.
    /// 끊긴 세션이 열어 둔 문서를 닫는 정리용이다(격리되면 닫지 못해 다음 실행이 느려졌다, 2026-10-07).
    /// </summary>
    public void Post(Action action)
    {
        if (_queue.IsAddingCompleted) return;
        try { _queue.Add((action, new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously))); }
        catch (InvalidOperationException) { }
    }

    /// <summary>동기 래퍼 (호출 스레드를 블록한다).</summary>
    public T Invoke<T>(Func<T> func)   => WaitWithTimeout(InvokeAsync(func));

    /// <summary>동기 래퍼 (호출 스레드를 블록한다).</summary>
    public void Invoke(Action action)  => WaitWithTimeout(InvokeAsync(action));

    private T WaitWithTimeout<T>(Task<T> task)
    {
        try { return task.WaitAsync(DefaultTimeout).GetAwaiter().GetResult(); }
        catch (TimeoutException ex) { Quarantine(task); throw new PptComTimeoutException(DefaultTimeout, ex); }
    }

    private void WaitWithTimeout(Task task)
    {
        try { task.WaitAsync(DefaultTimeout).GetAwaiter().GetResult(); }
        catch (TimeoutException ex) { Quarantine(task); throw new PptComTimeoutException(DefaultTimeout, ex); }
    }

    /// <summary>
    /// 멈춘 작업이 끝날 때까지 새 호출을 막고, 끝나면(성공·실패 모두) 다시 연다. 전에는 앱을 다시 시작할 때까지 막혀 있었다.
    /// </summary>
    private void Quarantine(Task stuck)
    {
        _isHealthy = false;
        stuck.ContinueWith(_ => { _isHealthy = true; Recovered?.Invoke(); }, TaskScheduler.Default);
    }

    /// <summary>격리가 풀렸을 때(멈춘 작업이 끝났을 때).</summary>
    public event Action? Recovered;

    public void Dispose()
    {
        _queue.CompleteAdding();
        _thread.Join(TimeSpan.FromSeconds(5));
    }
}

public sealed class PptComTimeoutException(TimeSpan timeout, Exception? inner = null)
    : TimeoutException($"PowerPoint COM 작업이 {timeout.TotalSeconds:0}초 안에 완료되지 않았습니다. 그 작업이 끝날 때까지 새 작업을 받지 않습니다.", inner);

public sealed class PptComUnavailableException(string message) : InvalidOperationException(message);

/// <summary>열린 PowerPoint에서 바로 작업하는 중에 사용자가 도형·글을 선택(편집)했다. 작업을 멈춘다.</summary>
public sealed class PptUserInterventionException()
    : Exception("PowerPoint에서 도형이나 글을 선택(편집)해 작업을 멈췄습니다. 조작을 마친 뒤 이어서 요청해 주세요.");
