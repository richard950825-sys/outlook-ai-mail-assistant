using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OutlookAIAddin
{
    /// <summary>一封邮件的自定义模板：名称 + 说明 + 结构/示例 + 默认收件人 + 签名。</summary>
    public sealed class MailTemplate
    {
        public string id = "";
        public string name = "";
        public string description = "";
        public string structure = "";
        public string defaultTo = "";
        public string signature = "";
    }

    internal static class TemplateStore
    {
        private static readonly object Sync = new object();

        public static List<MailTemplate> Load()
        {
            lock (Sync)
            {
                List<MailTemplate> list = null;
                try
                {
                    if (File.Exists(AppPaths.TemplatesFile))
                    {
                        string text = File.ReadAllText(AppPaths.TemplatesFile, Encoding.UTF8);
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            list = Json.Deserialize<List<MailTemplate>>(text);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Error("读取模板失败，使用内置模板", ex);
                }

                if (list == null || list.Count == 0)
                {
                    list = Defaults();
                    try { Save(list); } catch { }
                }

                for (int i = 0; i < list.Count; i++)
                {
                    if (list[i] == null) { list[i] = new MailTemplate(); }
                    if (string.IsNullOrWhiteSpace(list[i].id)) { list[i].id = NewId(); }
                    if (string.IsNullOrWhiteSpace(list[i].name)) { list[i].name = "未命名模板"; }
                }
                return list;
            }
        }

        public static void Save(List<MailTemplate> templates)
        {
            lock (Sync)
            {
                if (templates == null) { templates = new List<MailTemplate>(); }
                Directory.CreateDirectory(AppPaths.RootDir);
                File.WriteAllText(AppPaths.TemplatesFile, Json.Serialize(templates), new UTF8Encoding(false));
            }
        }

        public static MailTemplate CreateEmpty()
        {
            return new MailTemplate
            {
                id = NewId(),
                name = "新模板",
                description = "",
                structure = "",
                defaultTo = "",
                signature = ""
            };
        }

        public static string NewId()
        {
            return Guid.NewGuid().ToString("N").Substring(0, 8);
        }

        public static List<MailTemplate> Defaults()
        {
            return new List<MailTemplate>
            {
                new MailTemplate
                {
                    id = "default-meeting",
                    name = "会议邀请 / Meeting Invitation",
                    description = "邀请对方开会，包含时间、地点、议程",
                    structure = "Subject: Meeting Invitation — <meeting topic>\r\n\r\nHi <name>,\r\n\r\n<one short sentence about why we need this meeting>\r\n\r\nTime: <date and time>\r\nLocation: <place / Teams link>\r\nAgenda:\r\n1. <item 1>\r\n2. <item 2>\r\n\r\nPlease let me know if this time works for you. Looking forward to the meeting.\r\n\r\nBest regards,\r\n<your name>",
                    defaultTo = "",
                    signature = ""
                },
                new MailTemplate
                {
                    id = "default-followup",
                    name = "跟进确认 / Follow-up",
                    description = "对之前的邮件/会议做跟进确认，语气礼貌简洁",
                    structure = "Subject: Re: <original subject>\r\n\r\nHi <name>,\r\n\r\nJust following up on <topic>. <what has been done / what is pending>\r\n\r\nCould you please confirm <the specific item you need confirmed>? If anything is unclear, feel free to let me know.\r\n\r\nThanks and best regards,\r\n<your name>",
                    defaultTo = "",
                    signature = ""
                },
                new MailTemplate
                {
                    id = "default-report",
                    name = "进度汇报 / Status Update",
                    description = "向上级或客户简要汇报进展、风险和下一步计划",
                    structure = "Subject: Status Update — <project name> (<date>)\r\n\r\nHi <name>,\r\n\r\nPlease find below the latest status of <project>:\r\n\r\n1. Completed: <...>\r\n2. In progress: <...>\r\n3. Risks / Blockers: <...>\r\n4. Next steps: <...>\r\n\r\nLet me know if you have any questions.\r\n\r\nBest regards,\r\n<your name>",
                    defaultTo = "",
                    signature = ""
                }
            };
        }
    }
}
