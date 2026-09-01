namespace BckPPT.Infrastructure;

public sealed record EnvironmentHealth(bool IsPowerPointInstalled, string Message);

public static class EnvironmentHealthService
{
    public static EnvironmentHealth Check()
    {
        try
        {
            var installed = Type.GetTypeFromProgID("PowerPoint.Application", throwOnError: false) is not null;
            return installed
                ? new EnvironmentHealth(true, "PowerPoint 설치 확인됨")
                : new EnvironmentHealth(false, "Microsoft PowerPoint 데스크톱 앱이 설치되어 있지 않습니다. 채팅은 가능하지만 PPT 적용은 사용할 수 없습니다.");
        }
        catch (Exception ex) { return new EnvironmentHealth(false, $"PowerPoint 설치 상태 확인 실패: {ex.Message}"); }
    }
}
