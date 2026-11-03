using System.Text;

namespace VbScreenCleaner.Core;

public enum TokKind { Ident, String, Number, Dot, Bang, LParen, RParen, Colon, Equals, Comma, Other }

public readonly record struct Token(TokKind Kind, string Text, int Start, int End)
{
    public bool Is(string ident) =>
        Kind == TokKind.Ident && string.Equals(Text, ident, StringComparison.OrdinalIgnoreCase);

    public bool IsMemberAccess => Kind is TokKind.Dot or TokKind.Bang;
}

/// <summary>Minimal VB5/VB6 lexer: enough to find identifiers, strings and statement boundaries.</summary>
public static class VbLexer
{
    /// <summary>Splits a physical line into its code part and its comment part (' or Rem).</summary>
    public static (string Code, string Comment) SplitComment(string line)
    {
        bool inString = false, atStatementStart = true;
        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (inString)
            {
                if (c == '"') inString = false;
                continue;
            }
            if (c == '"') { inString = true; atStatementStart = false; continue; }
            if (c == '\'') return (line[..i], line[i..]);
            if (c == ':') { atStatementStart = true; continue; }
            if (char.IsWhiteSpace(c)) continue;
            if (atStatementStart && IsRem(line, i)) return (line[..i], line[i..]);
            atStatementStart = false;
        }
        return (line, "");
    }

    static bool IsRem(string s, int i) =>
        i + 3 <= s.Length
        && string.Compare(s, i, "Rem", 0, 3, StringComparison.OrdinalIgnoreCase) == 0
        && (i + 3 == s.Length || char.IsWhiteSpace(s[i + 3]));

    public static List<Token> Tokenize(string code)
    {
        var list = new List<Token>();
        int i = 0, n = code.Length;
        while (i < n)
        {
            char c = code[i];
            if (char.IsWhiteSpace(c)) { i++; continue; }
            int start = i;

            if (c == '"')
            {
                var sb = new StringBuilder();
                i++;
                while (i < n)
                {
                    if (code[i] == '"')
                    {
                        if (i + 1 < n && code[i + 1] == '"') { sb.Append('"'); i += 2; continue; }
                        i++;
                        break;
                    }
                    sb.Append(code[i++]);
                }
                list.Add(new Token(TokKind.String, sb.ToString(), start, i));
                continue;
            }

            if (char.IsLetter(c))
            {
                while (i < n && (char.IsLetterOrDigit(code[i]) || code[i] == '_')) i++;
                string text = code[start..i];
                // Type suffix characters; '!' is a suffix only when it is not the bang operator.
                if (i < n && "$%&#@".IndexOf(code[i]) >= 0) i++;
                else if (i < n && code[i] == '!' && !(i + 1 < n && (char.IsLetter(code[i + 1]) || code[i + 1] == '['))) i++;
                list.Add(new Token(TokKind.Ident, text, start, i));
                continue;
            }

            if (c == '[')
            {
                int close = code.IndexOf(']', i);
                if (close < 0) close = n - 1;
                list.Add(new Token(TokKind.Ident, code[(i + 1)..close], start, close + 1));
                i = close + 1;
                continue;
            }

            bool prevIsValue = list.Count > 0 && list[^1].Kind is TokKind.Ident or TokKind.RParen;
            if (char.IsDigit(c) || (c == '.' && i + 1 < n && char.IsDigit(code[i + 1]) && !prevIsValue))
            {
                i++;
                while (i < n && (char.IsLetterOrDigit(code[i]) || code[i] == '.')) i++;
                if (i < n && "#!@%&".IndexOf(code[i]) >= 0) i++;
                list.Add(new Token(TokKind.Number, code[start..i], start, i));
                continue;
            }

            if (c == '&' && i + 2 < n && "HhOo".IndexOf(code[i + 1]) >= 0 && char.IsLetterOrDigit(code[i + 2]))
            {
                i += 2;
                while (i < n && char.IsLetterOrDigit(code[i])) i++;
                if (i < n && code[i] == '&') i++;
                list.Add(new Token(TokKind.Number, code[start..i], start, i));
                continue;
            }

            switch (c)
            {
                case '.': list.Add(new Token(TokKind.Dot, ".", start, ++i)); break;
                case '!': list.Add(new Token(TokKind.Bang, "!", start, ++i)); break;
                case '(': list.Add(new Token(TokKind.LParen, "(", start, ++i)); break;
                case ')': list.Add(new Token(TokKind.RParen, ")", start, ++i)); break;
                case ',': list.Add(new Token(TokKind.Comma, ",", start, ++i)); break;
                case '=': list.Add(new Token(TokKind.Equals, "=", start, ++i)); break;
                case ':':
                    if (i + 1 < n && code[i + 1] == '=') { i += 2; list.Add(new Token(TokKind.Other, ":=", start, i)); }
                    else list.Add(new Token(TokKind.Colon, ":", start, ++i));
                    break;
                default: list.Add(new Token(TokKind.Other, c.ToString(), start, ++i)); break;
            }
        }
        return list;
    }

    /// <summary>Splits [from, to) at top-level ':' separators.</summary>
    public static List<(int Start, int End)> SplitColons(IReadOnlyList<Token> toks, int from, int to)
    {
        var result = new List<(int, int)>();
        int segStart = from;
        for (int i = from; i < to; i++)
        {
            if (toks[i].Kind != TokKind.Colon) continue;
            if (i > segStart) result.Add((segStart, i));
            segStart = i + 1;
        }
        if (to > segStart) result.Add((segStart, to));
        return result;
    }

    public static int FindIdent(IReadOnlyList<Token> toks, int from, int to, string ident)
    {
        for (int i = from; i < to; i++)
            if (toks[i].Is(ident)) return i;
        return -1;
    }

    public static int MatchParen(IReadOnlyList<Token> toks, int open, int to)
    {
        int depth = 0;
        for (int i = open; i < to; i++)
        {
            if (toks[i].Kind == TokKind.LParen) depth++;
            else if (toks[i].Kind == TokKind.RParen && --depth == 0) return i;
        }
        return -1;
    }
}
