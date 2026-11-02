using System;
using System.Text;

namespace OutlookAIAddin
{
    /// <summary>构造发给 DeepSeek 等聊天接口的提示词（中文指令，输出格式可控）。</summary>
    internal static class Prompts
    {
        public const string SystemPrompt =
            "你是一位资深的商务邮件助手，精通中英文商务写作。你的回复必须精准、专业、可以直接使用；" +
            "严格遵守用户要求的输出格式，不要添加额外解释、前言或 Markdown 代码块（除非明确要求）。" +
            "不要编造邮件中不存在的事实。";

        public static string Summary(string subject, string from, string body, string summaryLanguage)
        {
            string langLine = summaryLanguage == "source"
                ? "输出语言：与邮件正文相同的主要语言。"
                : "输出语言：中文。";
            var sb = new StringBuilder();
            sb.AppendLine("请阅读下面这封邮件，输出 3-6 条要点摘要，每条一行并以「- 」开头。");
            sb.AppendLine("要求：保留关键数字、日期、人名、金额和行动要求；不要编造正文没有的信息；如果正文没有实质内容，输出「- 邮件没有实质内容」。");
            sb.AppendLine(langLine);
            sb.AppendLine();
            sb.Append(MailBlock(subject, from, body));
            return sb.ToString();
        }

        public static string Translate(string text, string target)
        {
            string targetName;
            if (target == "en") { targetName = "英文"; }
            else if (target == "zh") { targetName = "中文"; }
            else { targetName = LooksChinese(text) ? "英文" : "中文"; }

            var sb = new StringBuilder();
            sb.AppendLine("请把下面内容翻译成" + targetName + "，保持原有段落结构，只输出译文，不要解释。");
            sb.AppendLine();
            sb.AppendLine("【需要翻译的内容】");
            sb.AppendLine(TrimTo(text, 15000));
            return sb.ToString();
        }

        /// <summary>mode: polish / shorten / formal</summary>
        public static string Rewrite(string text, string mode, string subject, string from)
        {
            string instruction;
            switch (mode)
            {
                case "shorten":
                    instruction = "请精简下面这封邮件中的文字：保留全部关键信息，去掉冗余表达，篇幅缩短约 30%。";
                    break;
                case "formal":
                    instruction = "请把下面这封邮件中的文字改写为更正式、专业的商务语气，保持原意和关键信息不变。";
                    break;
                default:
                    instruction = "请润色下面这封邮件中的文字：保持原意和事实不变，修正语法、拼写、标点和用词，使语气得体、表达专业流畅。";
                    break;
            }

            var sb = new StringBuilder();
            sb.AppendLine(instruction);
            sb.AppendLine("要求：不要翻译成其他语言，保持原文语言；只输出处理后的文字，不要解释、不要加引号或代码块。");
            if (!string.IsNullOrWhiteSpace(subject))
            {
                sb.AppendLine("（邮件上下文，仅供理解，不需要输出：主题「" + TrimTo(subject, 200) + "」"
                    + (string.IsNullOrWhiteSpace(from) ? "" : "，来自「" + TrimTo(from, 200) + "」") + "）");
            }
            sb.AppendLine();
            sb.AppendLine("【需要处理的文字】");
            sb.AppendLine(TrimTo(text, 15000));
            return sb.ToString();
        }

        /// <summary>润色 + 切换语言：把选中文字润色成中文或英文。</summary>
        public static string PolishToLanguage(string text, string targetLanguage, string subject, string from)
        {
            var sb = new StringBuilder();
            sb.AppendLine("请润色下面这封邮件中的文字，并把语言转换成" + targetLanguage + "：");
            sb.AppendLine("要求：保持原意和事实不变；修正语法、拼写、标点和不自然的表达；使语气专业、得体、流畅；");
            sb.AppendLine("只输出处理后的文字，不要解释、不要加引号或代码块。");
            if (!string.IsNullOrWhiteSpace(subject))
            {
                sb.AppendLine("（邮件上下文，仅供理解，不需要输出：主题「" + TrimTo(subject, 200) + "」"
                    + (string.IsNullOrWhiteSpace(from) ? "" : "，来自「" + TrimTo(from, 200) + "」") + "）");
            }
            sb.AppendLine();
            sb.AppendLine("【需要处理的文字】");
            sb.AppendLine(TrimTo(text, 15000));
            return sb.ToString();
        }

        public static string Compose(MailTemplate template, string instruction)
        {
            var sb = new StringBuilder();
            sb.AppendLine("请根据下面的模板和我的要求，写一封完整的商务邮件。");
            sb.AppendLine("要求：严格遵循模板的结构和语气风格；内容必须针对我的要求；称呼、署名等占位符如果我的要求中没有提供，就使用模板中的写法或留空。");
            sb.AppendLine();
            if (template != null && !string.IsNullOrWhiteSpace(template.name))
            {
                sb.AppendLine("【模板名称】" + template.name.Trim());
            }
            if (template != null && !string.IsNullOrWhiteSpace(template.description))
            {
                sb.AppendLine("【模板说明】" + template.description.Trim());
            }
            if (!string.IsNullOrWhiteSpace(instruction))
            {
                sb.AppendLine("【我的要求】" + instruction.Trim());
            }
            if (template != null && !string.IsNullOrWhiteSpace(template.defaultTo))
            {
                sb.AppendLine("【收件人】" + template.defaultTo.Trim());
            }
            if (template != null && !string.IsNullOrWhiteSpace(template.signature))
            {
                sb.AppendLine("【署名/签名】" + template.signature.Trim());
            }
            sb.AppendLine();
            sb.AppendLine("【模板结构 / 风格示例】");
            sb.AppendLine(template != null && !string.IsNullOrWhiteSpace(template.structure)
                ? TrimTo(template.structure, 8000)
                : "（无模板结构，请按通用商务邮件格式书写。）");
            sb.AppendLine();
            sb.AppendLine("【输出要求】");
            sb.AppendLine("1. 必须输出合法 JSON，不要 Markdown 代码块、不要任何其它文字。格式：{\"subject\":\"邮件主题\",\"body\":\"邮件正文\"}");
            sb.AppendLine("2. body 使用纯文本，段落之间用 \\n\\n 分隔，不要在 body 里重复主题。");
            sb.AppendLine("3. 语言：与我的要求相同的主要语言（我的要求是中文就用中文，否则用邮件正文中的语言）。");
            return sb.ToString();
        }

        public static string MailBlock(string subject, string from, string body)
        {
            var sb = new StringBuilder();
            sb.AppendLine("【邮件信息】");
            sb.AppendLine("主题：" + (string.IsNullOrWhiteSpace(subject) ? "（无主题）" : TrimTo(subject, 300)));
            if (!string.IsNullOrWhiteSpace(from)) { sb.AppendLine("发件人：" + TrimTo(from, 300)); }
            sb.AppendLine();
            sb.AppendLine("【正文】");
            sb.AppendLine(TrimTo(body, 15000));
            return sb.ToString();
        }

        public static bool LooksChinese(string text)
        {
            if (string.IsNullOrEmpty(text)) { return false; }
            int cjk = 0;
            int letters = 0;
            foreach (char c in text)
            {
                if (c >= '\u4e00' && c <= '\u9fff') { cjk++; }
                if (char.IsLetter(c)) { letters++; }
            }
            if (letters == 0) { return false; }
            return (double)cjk / letters > 0.25;
        }

        public static string TrimTo(string text, int max)
        {
            if (text == null) { return ""; }
            if (text.Length <= max) { return text; }
            return text.Substring(0, max) + "\n……（内容过长，已截断）";
        }
    }
}
