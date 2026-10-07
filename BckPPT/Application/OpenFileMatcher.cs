using System.IO;
using System.Text.RegularExpressions;
using BckPPT.Models;

namespace BckPPT.Application;

public enum FileMatchKind { None, One, Ambiguous }

public sealed record FileMatch(FileMatchKind Kind, OpenPresentationInfo? File, string Token, IReadOnlyList<OpenPresentationInfo> Candidates);

/// <summary>
/// 요청 문장에 적힌 파일 이름 조각으로 열린 PPT를 찾는다. 예: "착수보고 파일 40쪽에…", "추진일정.pptx에서".
/// "기획"처럼 흔한 낱말로 대상이 바뀌지 않도록, 이름 조각 바로 뒤에 "파일·문서·ppt"가 붙은 경우만 본다.
/// 여러 문서가 맞으면 바꾸지 않고 알린다.
/// </summary>
public static class OpenFileMatcher
{
    private static readonly Regex Mention = new(
        @"([^\s""'“”‘’()\[\]]{2,})\s*(?:\.pptx?|\.pptm|파일|문서|피피티|ppt|PPT)(?![A-Za-z])",
        RegexOptions.Compiled);

    public static FileMatch Match(string request, IReadOnlyList<OpenPresentationInfo> open)
    {
        var saved = open.Where(o => o.IsSaved).ToList();
        foreach (Match m in Mention.Matches(request))
        {
            string token = Regex.Replace(m.Groups[1].Value, @"(에서|에|의|을|를|은|는|이|가)$", "");
            if (token.Length < 2 || Regex.IsMatch(token, @"^\d+$") || token is "이" or "그" or "저" or "현재" or "열린" or "선택된") continue;
            var hits = saved.Where(o => Path.GetFileNameWithoutExtension(o.Path).Contains(token, StringComparison.OrdinalIgnoreCase)).ToList();
            if (hits.Count == 1) return new FileMatch(FileMatchKind.One, hits[0], token, hits);
            if (hits.Count > 1)
            {
                // 이름 전체가 같은 문서가 하나면 그것.
                var exact = hits.Where(o => string.Equals(Path.GetFileNameWithoutExtension(o.Path), token, StringComparison.OrdinalIgnoreCase)).ToList();
                return exact.Count == 1
                    ? new FileMatch(FileMatchKind.One, exact[0], token, hits)
                    : new FileMatch(FileMatchKind.Ambiguous, null, token, hits);
            }
        }
        return new FileMatch(FileMatchKind.None, null, "", []);
    }
}
