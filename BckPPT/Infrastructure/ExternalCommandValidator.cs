using System.IO;

namespace BckPPT.Infrastructure;

public static class ExternalCommandValidator
{
    public static void Validate(ExternalCommandRequest request)
    {
        if (request.ProtocolVersion != "1.0") throw new ArgumentException("지원하지 않는 protocolVersion입니다.");
        if (!Guid.TryParse(request.RequestId, out _)) throw new ArgumentException("requestId는 UUID여야 합니다.");
        if (request.Command == "get_job_status") return;
        if (request.Command != "create_job") throw new ArgumentException("지원하지 않는 command입니다.");
        if (string.IsNullOrWhiteSpace(request.FilePath) || !Path.IsPathFullyQualified(request.FilePath))
            throw new ArgumentException("filePath는 절대 경로여야 합니다.");
        var fullPath = Path.GetFullPath(request.FilePath);
        if (!File.Exists(fullPath)) throw new FileNotFoundException("대상 PPT 파일이 없습니다.", fullPath);
        if (!new[] { ".pptx", ".pptm", ".ppt" }.Contains(Path.GetExtension(fullPath), StringComparer.OrdinalIgnoreCase))
            throw new ArgumentException("지원하지 않는 파일 형식입니다.");
        if (string.IsNullOrWhiteSpace(request.Instruction) || request.Instruction.Length > 4000)
            throw new ArgumentException("instruction은 1~4000자여야 합니다.");
        if (!string.Equals(request.SaveMode, "copy", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("현재 saveMode는 copy만 지원합니다.");
    }
}
