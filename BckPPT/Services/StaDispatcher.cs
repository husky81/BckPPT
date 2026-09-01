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
            throw new PptComUnavailableException("이전 PowerPoint COM 작업이 응답하지 않아 COM 서비스가 격리되었습니다. BckPPT를 다시 시작해 주세요.");
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

    /// <summary>동기 래퍼 (호출 스레드를 블록한다).</summary>
    public T Invoke<T>(Func<T> func)   => WaitWithTimeout(InvokeAsync(func));

    /// <summary>동기 래퍼 (호출 스레드를 블록한다).</summary>
    public void Invoke(Action action)  => WaitWithTimeout(InvokeAsync(action));

    private T WaitWithTimeout<T>(Task<T> task)
    {
        try { return task.WaitAsync(DefaultTimeout).GetAwaiter().GetResult(); }
        catch (TimeoutException ex)
        {
            _isHealthy = false;
            throw new PptComTimeoutException(DefaultTimeout, ex);
        }
    }

    private void WaitWithTimeout(Task task)
    {
        try { task.WaitAsync(DefaultTimeout).GetAwaiter().GetResult(); }
        catch (TimeoutException ex)
        {
            _isHealthy = false;
            throw new PptComTimeoutException(DefaultTimeout, ex);
        }
    }

    public void Dispose()
    {
        _queue.CompleteAdding();
        _thread.Join(TimeSpan.FromSeconds(5));
    }
}

public sealed class PptComTimeoutException(TimeSpan timeout, Exception? inner = null)
    : TimeoutException($"PowerPoint COM 작업이 {timeout.TotalSeconds:0}초 안에 완료되지 않았습니다. 작업을 중단하고 COM 서비스를 격리했습니다.", inner);

public sealed class PptComUnavailableException(string message) : InvalidOperationException(message);
