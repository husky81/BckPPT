using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using CellBox2.BckPPT.Models;
using System.Collections.ObjectModel;

namespace CellBox2.BckPPT.Services;

/// <summary>트리 구조를 JSON으로 저장/로드. %LOCALAPPDATA%\BckSoft\BckPPT\structures\</summary>
public static class TreeStructureStore
{
    private static readonly string StructureDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "BckSoft", "BckPPT", "structures");

    private static readonly JsonSerializerOptions _opts = new()
    {
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    public static void Save(string presentationName, ObservableCollection<PptNode> rootNodes)
    {
        try
        {
            Directory.CreateDirectory(StructureDir);
            var doc = new TreeDoc
            {
                PresentationName = presentationName,
                Nodes = rootNodes.Select(ToDto).ToList(),
            };
            var json = JsonSerializer.Serialize(doc, _opts);
            var path = GetPath(presentationName);
            System.Diagnostics.Debug.WriteLine($"[TreeStore] Saving {rootNodes.Count} nodes to {path}");
            File.WriteAllText(path, json);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[TreeStore] Save error: {ex}");
        }
    }

    /// <summary>저장된 구조를 로드하여 shapes에 매핑. 없으면 null 반환.</summary>
    public static ObservableCollection<PptNode>? Load(
        string presentationName, IReadOnlyList<PptSlideInfo> allShapes)
    {
        try
        {
            var path = GetPath(presentationName);
            if (!File.Exists(path)) return null;

            var json = File.ReadAllText(path);
            var doc = JsonSerializer.Deserialize<TreeDoc>(json, _opts);
            if (doc?.Nodes is null) return null;

            var shapeMap = new Dictionary<string, PptSlideInfo>();
            foreach (var s in allShapes)
            {
                var key = $"{s.SlideIndex}:{s.ShapeName}";
                if (!shapeMap.ContainsKey(key)) shapeMap[key] = s; // 중복 시 첫 번째 값 사용
            }

            var result = new ObservableCollection<PptNode>();
            var usedIds = new HashSet<string>();

            foreach (var dto in doc.Nodes)
                if (FromDto(dto, null, shapeMap, usedIds) is { } node)
                    result.Add(node);

            // 구조에 없는 새 shape는 루트에 추가 (중복 방지: usedIds 갱신 필수)
            foreach (var shape in allShapes)
            {
                var id = $"{shape.SlideIndex}:{shape.ShapeName}";
                if (usedIds.Add(id))  // Add가 false면 이미 처리된 id → 건너뜀
                    result.Add(new PptNode(shape));
            }

            return result;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[TreeStore] Load error: {ex}");
            return null;
        }
    }

    private static string GetPath(string presentationName)
    {
        var safeName = string.Join("_",
            presentationName.Split(Path.GetInvalidFileNameChars()));
        return Path.Combine(StructureDir, safeName + ".json");
    }

    // ── DTO ──────────────────────────────────────────────────────

    private record TreeDoc
    {
        public string PresentationName { get; init; } = "";
        public List<NodeDto> Nodes { get; init; } = [];
    }

    private record NodeDto
    {
        public string Id { get; init; } = "";
        public int Slide { get; init; }
        public string Shape { get; init; } = "";
        // Role은 사용자 라벨링된 경우만 저장 (null = 자동 분류 위임)
        public string? Role { get; init; }
        public List<NodeDto> Children { get; init; } = [];
    }

    private static NodeDto ToDto(PptNode node) => new()
    {
        Id = node.NodeId,
        Slide = node.Shape.SlideIndex,
        Shape = node.Shape.ShapeName,
        Role = node.Shape.IsUserLabeled ? node.Shape.Role.ToString() : null,
        Children = node.Children.Select(ToDto).ToList(),
    };

    private static PptNode? FromDto(NodeDto dto, PptNode? parent,
        Dictionary<string, PptSlideInfo> shapeMap, HashSet<string> usedIds)
    {
        System.Diagnostics.Debug.WriteLine($"[TreeStore] FromDto: id={dto.Id} children={dto.Children.Count} found={shapeMap.ContainsKey(dto.Id)}");
        if (!shapeMap.TryGetValue(dto.Id, out var shape)) return null;
        if (!usedIds.Add(dto.Id)) return null;  // 이미 처리된 id → 중복 노드 방지

        var node = new PptNode(shape, parent);

        // 구조 파일에 저장된 사용자 라벨 복원 (labels.jsonl 없이도 role 유지)
        if (!string.IsNullOrEmpty(dto.Role) && Enum.TryParse<ShapeRole>(dto.Role, out var role))
            node.SetRole(role);

        foreach (var childDto in dto.Children)
            if (FromDto(childDto, node, shapeMap, usedIds) is { } child)
                node.Children.Add(child);
        return node;
    }
}
