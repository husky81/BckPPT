namespace CellBox2.BckPPT.Services;

/// <summary>UI 스레드 디스패처 추상화 — ViewModel이 Application.Current에 직접 의존하지 않도록.</summary>
public interface IUiDispatcher
{
    void Post(Action action);
}
