using System;

namespace OutlookAIAddin
{
    /// <summary>解析「写邮件」动作返回的 JSON：{"subject":"...","body":"..."}。</summary>
    internal static class ComposeResultParser
    {
        internal sealed class Result
        {
            public string Subject = "";
            public string Body = "";
            public string Raw = "";
        }

        public static Result Parse(string content)
        {
            var result = new Result { Body = content ?? "", Raw = content ?? "" };
            if (string.IsNullOrWhiteSpace(content)) { return result; }

            string text = content.Trim();

            // 去掉可能的 ```json ... ``` 包裹
            if (text.StartsWith("```", StringComparison.Ordinal))
            {
                int firstNewLine = text.IndexOf('\n');
                if (firstNewLine >= 0)
                {
                    text = text.Substring(firstNewLine + 1);
                }
                int lastFence = text.LastIndexOf("```", StringComparison.Ordinal);
                if (lastFence >= 0) { text = text.Substring(0, lastFence); }
                text = text.Trim();
            }

            object parsed = TryParseObject(text) ?? TryParseFirstObject(text);
            var dict = parsed as System.Collections.Generic.Dictionary<string, object>;
            if (dict != null)
            {
                string subject = GetString(dict, "subject");
                string body = GetString(dict, "body");
                if (subject != null || body != null)
                {
                    result.Subject = subject ?? "";
                    result.Body = body ?? "";
                    return result;
                }
            }

            // 兜底：正则抽取字段
            string subjectValue = ExtractJsonString(text, "subject");
            string bodyValue = ExtractJsonString(text, "body");
            if (!string.IsNullOrEmpty(subjectValue) || !string.IsNullOrEmpty(bodyValue))
            {
                result.Subject = subjectValue ?? "";
                result.Body = bodyValue ?? result.Body;
            }
            return result;
        }

        private static object TryParseObject(string text)
        {
            try
            {
                return Json.Deserialize<System.Collections.Generic.Dictionary<string, object>>(text);
            }
            catch
            {
                return null;
            }
        }

        private static object TryParseFirstObject(string text)
        {
            int start = text.IndexOf('{');
            int end = text.LastIndexOf('}');
            if (start < 0 || end <= start) { return null; }
            return TryParseObject(text.Substring(start, end - start + 1));
        }

        private static string GetString(System.Collections.Generic.Dictionary<string, object> dict, string key)
        {
            object value;
            if (dict.TryGetValue(key, out value) && value != null)
            {
                return Convert.ToString(value);
            }
            return null;
        }

        private static string ExtractJsonString(string text, string key)
        {
            try
            {
                var match = System.Text.RegularExpressions.Regex.Match(
                    text, "\"" + key + "\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\"",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (!match.Success) { return null; }
                string raw = match.Groups[1].Value;
                return Unescape(raw);
            }
            catch
            {
                return null;
            }
        }

        private static string Unescape(string raw)
        {
            if (raw.IndexOf('\\') < 0) { return raw; }
            var sb = new System.Text.StringBuilder(raw.Length);
            for (int i = 0; i < raw.Length; i++)
            {
                char c = raw[i];
                if (c != '\\' || i + 1 >= raw.Length) { sb.Append(c); continue; }
                char next = raw[++i];
                switch (next)
                {
                    case 'n': sb.Append('\n'); break;
                    case 'r': sb.Append('\r'); break;
                    case 't': sb.Append('\t'); break;
                    case '"': sb.Append('"'); break;
                    case '\\': sb.Append('\\'); break;
                    case '/': sb.Append('/'); break;
                    case 'u':
                        if (i + 4 < raw.Length)
                        {
                            try
                            {
                                sb.Append((char)Convert.ToInt32(raw.Substring(i + 1, 4), 16));
                                i += 4;
                            }
                            catch { sb.Append("\\u"); }
                        }
                        else { sb.Append("\\u"); }
                        break;
                    default: sb.Append(next); break;
                }
            }
            return sb.ToString();
        }
    }
}
