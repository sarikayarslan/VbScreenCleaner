namespace VbScreenCleaner.Core;

public enum VbFileKind { Form, UserControl, PropertyPage, UserDocument, Designer, Module, Class }

public sealed class VbSourceFile
{
    public required string FullPath { get; init; }
    /// <summary>Path relative to the project root; null when the file lives outside the root.</summary>
    public required string? RelativePath { get; init; }
    public required VbFileKind Kind { get; init; }
    public required List<string> Lines { get; init; }
    public string NewLine { get; init; } = "\r\n";
    public bool EndsWithNewLine { get; init; } = true;

    public string ObjectName { get; set; } = "";
    public ControlNode? Root { get; set; }
    public int CodeStart { get; set; }
    public List<LogicalLine> LogicalLines { get; } = new();
    public List<CodeUnit> Units { get; } = new();
    public List<ProcInfo> Procs { get; } = new();
    public Dictionary<string, ControlGroup> Groups { get; } = new(StringComparer.OrdinalIgnoreCase);

    public bool IsInsideRoot => RelativePath != null;
    public bool HasDesigner => Root != null;
    public bool CanRemoveControls =>
        IsInsideRoot && HasDesigner && Kind is VbFileKind.Form or VbFileKind.UserControl
            or VbFileKind.PropertyPage or VbFileKind.UserDocument;
    public string DisplayName => RelativePath ?? FullPath;

    public override string ToString() => DisplayName;
}

public sealed class ControlNode
{
    public required string TypeName { get; init; }
    public required string Name { get; init; }
    public int? Index { get; set; }
    public int BeginLine { get; init; }
    public int EndLine { get; set; }
    public ControlNode? Parent { get; init; }
    public List<ControlNode> Children { get; } = new();
    public Dictionary<string, string> Properties { get; } = new(StringComparer.OrdinalIgnoreCase);
    /// <summary>Line index of each top-level property.</summary>
    public Dictionary<string, int> PropertyLines { get; } = new(StringComparer.OrdinalIgnoreCase);
    /// <summary>Properties inside BeginProperty blocks (fonts, OCX sub-objects).</summary>
    public List<(string Name, string Value)> NestedProperties { get; } = new();
    public ControlGroup? Group { get; set; }

    // Analysis state
    public bool HiddenForever { get; set; }
    public bool EffHidden { get; set; }
    public bool WillRemove { get; set; }
    public bool IsRemovalRoot { get; set; }
    public string HiddenWhy { get; set; } = "";

    public bool IsRoot => Parent == null;
    public bool IsMenu => TypeName.Equals("VB.Menu", StringComparison.OrdinalIgnoreCase);
    public bool DesignVisible => !Properties.TryGetValue("Visible", out var v) || !PropertyValue.IsFalse(v);
    public string DisplayName => Index is null ? Name : $"{Name}({Index})";

    public IEnumerable<ControlNode> Descendants()
    {
        foreach (var c in Children)
        {
            yield return c;
            foreach (var d in c.Descendants()) yield return d;
        }
    }

    public IEnumerable<ControlNode> SelfAndDescendants() => Descendants().Prepend(this);

    public override string ToString() => $"{TypeName} {DisplayName}";
}

public static class PropertyValue
{
    public static string Strip(string raw)
    {
        // Values look like:  0   'False   or   "text"   or   "frm.frx":0000
        var t = raw.Trim();
        if (t.StartsWith('"')) return t;
        int q = t.IndexOf('\'');
        return (q >= 0 ? t[..q] : t).Trim();
    }

    public static bool IsFalse(string raw)
    {
        var v = Strip(raw);
        return v == "0" || v.Equals("False", StringComparison.OrdinalIgnoreCase);
    }

    public static bool IsTrue(string raw)
    {
        var v = Strip(raw);
        return v == "-1" || v == "1" || v.Equals("True", StringComparison.OrdinalIgnoreCase);
    }

    public static string? Unquote(string raw)
    {
        var t = raw.Trim();
        if (!t.StartsWith('"')) return null;
        int end = t.IndexOf('"', 1);
        return end < 0 ? null : t[1..end];
    }
}

/// <summary>All elements sharing one control name in a form (one element, or a control array).</summary>
public sealed class ControlGroup
{
    public required VbSourceFile File { get; init; }
    public required string Name { get; init; }
    public List<ControlNode> Elements { get; } = new();
    public List<ProcInfo> Handlers { get; } = new();

    public List<string> StaticReasons { get; } = new();
    public List<string> Reasons { get; } = new();
    public List<VisEffect> Shows { get; } = new();
    public List<VisEffect> LoadHides { get; } = new();
    public bool HasLiveNonVisibleWrite { get; set; }
    public bool Removed { get; set; }

    public string Key => MakeKey(File, Name);
    public static string MakeKey(VbSourceFile f, string name) => $"{f.DisplayName}|{name}".ToLowerInvariant();

    public void AddReason(string reason)
    {
        if (!Reasons.Contains(reason)) Reasons.Add(reason);
    }

    public override string ToString() => $"{File.ObjectName}.{Name}";
}

public sealed class LogicalLine
{
    public required int FirstLine { get; init; }
    public required int LastLine { get; init; }
    public required string Indent { get; init; }
    public required string Code { get; init; }
    public required string Comment { get; init; }
    public required List<Token> Tokens { get; init; }
    /// <summary>Leading line label exactly as written ("ErrHandler:", "10:" or "10"), if any.</summary>
    public string? Label { get; set; }
    public List<CodeUnit> Units { get; } = new();
}

public enum TargetKind { VisibleHide, VisibleShow, Write, Method }

public sealed record AssignMatch(
    IReadOnlyList<ControlGroup> Groups,
    TargetKind Kind,
    int? Index,
    bool IndexUnknown,
    HashSet<int> Excluded)
{
    public ControlGroup? Group => Groups.Count == 1 ? Groups[0] : null;
    public bool IsVisible => Kind is TargetKind.VisibleHide or TargetKind.VisibleShow;
}

public sealed record VisEffect(ControlGroup Group, bool Show, int? Index, bool IndexUnknown, bool Conditional, CodeUnit Unit);

public sealed record BlockRef(ControlGroup Group, string Reason);

/// <summary>One statement (or a whole single-line If) inside a logical line.</summary>
public sealed class CodeUnit
{
    public required VbSourceFile File { get; init; }
    public required LogicalLine Line { get; init; }
    public required int TokStart { get; init; }
    public required int TokEnd { get; init; }
    public bool IsSingleLineIf { get; init; }

    public ProcInfo? Proc { get; set; }
    public int Depth { get; set; }
    public bool AfterExit { get; set; }

    // Pre-computed analysis data
    public AssignMatch? Target { get; set; }
    public bool RemovableShape => Target != null;
    public List<BlockRef> BlockRefs { get; } = new();
    public List<VisEffect> VisEffects { get; } = new();
    public List<string> CalledNames { get; } = new();
    public List<string> Strings { get; } = new();
    /// <summary>Forms whose Controls collection this statement touches (Me.Controls, frmX.Controls).</summary>
    public List<VbSourceFile> ControlsForms { get; } = new();
    /// <summary>Touches the Controls collection of an object whose form cannot be determined.</summary>
    public bool ControlsUnknown { get; set; }

    public string Text
    {
        get
        {
            var t = Line.Tokens;
            int s = t[TokStart].Start, e = t[TokEnd - 1].End;
            return Line.Code[s..e].Trim();
        }
    }

    public string Location => $"{File.DisplayName}:{Line.FirstLine + 1}";

    public void ResetAnalysis()
    {
        Target = null;
        BlockRefs.Clear();
        VisEffects.Clear();
        CalledNames.Clear();
        Strings.Clear();
        ControlsForms.Clear();
        ControlsUnknown = false;
    }
}

public sealed class ProcInfo
{
    public required VbSourceFile File { get; init; }
    public required string Name { get; init; }
    public required string Kind { get; init; }
    public required int HeaderLine { get; init; }
    public bool IsPrivate { get; init; }
    public int EndLine { get; set; } = -1;
    public CodeUnit? Header { get; set; }
    public CodeUnit? End { get; set; }
    public List<CodeUnit> Units { get; } = new();
    public ControlGroup? HandlerOwner { get; set; }

    public bool HasBody => Units.Any(u => u != Header && u != End);
}
