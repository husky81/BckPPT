using System.Collections.Concurrent;

namespace CellBox2.BckPPT.Services;

/// <summary>
/// 앱 수명 동안 유지되는 전용 STA 스레드.
/// 매번 new Thread()를 생성하는 대신 이 싱글턴을 통해 COM 호출을 직렬화한다.
/// </summary>
public sealed class StaDispatcher : IDisposable
{
    public static readonly StaDispatcher Instance = new();

    private readonly Thread _thread;
    private readonly BlockingCollection<(Action action, TaskCompletionSource tcs)> _queue = new();

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
    public T Invoke<T>(Func<T> func)   => InvokeAsync(func).GetAwaiter().GetResult();

    /// <summary>동기 래퍼 (호출 스레드를 블록한다).</summary>
    public void Invoke(Action action)  => InvokeAsync(action).GetAwaiter().GetResult();

    public void Dispose()
    {
        _queue.CompleteAdding();
        _thread.Join(TimeSpan.FromSeconds(5));
    }
}
