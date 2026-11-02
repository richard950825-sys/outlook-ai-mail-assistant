using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookAIAddin
{
    internal sealed class AiResult
    {
        public string Content = "";
        public string Model = "";
        public string FinishReason = "";
        public int PromptTokens;
        public int CompletionTokens;
        public int TotalTokens;
        public int ReasoningTokens;
    }

    internal sealed class AiException : Exception
    {
        public AiException(string message) : base(message) { }
    }

    /// <summary>
    /// 调用 OpenAI 兼容的 /chat/completions 接口（默认 DeepSeek）。
    /// 注意：deepseek-flash 是推理模型，reasoning 也占用 max_tokens；长文本时可能
    /// 把额度全部用于推理而返回空 content（finish_reason=length）。这里会自动
    /// 提高 max_tokens 重试一次，并把原因写进错误信息。
    /// </summary>
    internal static class AiClient
    {
        private const int RetryMaxTokens = 16000;
        private const int DefaultMaxTokens = 4000;
        private static readonly HttpClient Http = CreateClient();

        private sealed class RawResponse
        {
            public HttpStatusCode StatusCode;
            public string ReasonPhrase = "";
            public string Body = "";
        }

        private static HttpClient CreateClient()
        {
            try { ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12; }
            catch { /* ignore */ }

            var handler = new HttpClientHandler { AllowAutoRedirect = true };
            var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(240) };
            client.DefaultRequestHeaders.UserAgent.ParseAdd("OutlookAIAddin/1.0");
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            return client;
        }

        public static async Task<AiResult> ChatAsync(
            string baseUrl,
            string apiKey,
            string model,
            string systemPrompt,
            string userPrompt,
            double temperature,
            int maxTokens,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            if (string.IsNullOrWhiteSpace(baseUrl)) { throw new AiException("还没有配置 Base URL（设置 → AI 接口）。"); }
            if (string.IsNullOrWhiteSpace(apiKey)) { throw new AiException("还没有填写 API Key（设置 → AI 接口）。"); }
            if (string.IsNullOrWhiteSpace(model)) { throw new AiException("还没有填写模型名（设置 → AI 接口）。"); }

            int requestMaxTokens = maxTokens > 0 ? maxTokens : DefaultMaxTokens;

            for (int attempt = 1; ; attempt++)
            {
                Log.Info("调用 AI: model=" + model.Trim() + ", max_tokens=" + requestMaxTokens
                         + ", promptLen=" + (userPrompt == null ? 0 : userPrompt.Length)
                         + (attempt > 1 ? ", 第 " + attempt + " 次尝试" : ""));

                RawResponse raw = await SendAsync(
                    baseUrl, apiKey, model, systemPrompt, userPrompt,
                    temperature, requestMaxTokens, cancellationToken).ConfigureAwait(false);

                Dictionary<string, object> root = null;
                try { root = Json.Deserialize<Dictionary<string, object>>(raw.Body); }
                catch { root = null; }

                if ((int)raw.StatusCode >= 400)
                {
                    string message = ExtractErrorMessage(root);
                    if (string.IsNullOrWhiteSpace(message))
                    {
                        message = "HTTP " + (int)raw.StatusCode
                                  + (string.IsNullOrEmpty(raw.ReasonPhrase) ? "" : " " + raw.ReasonPhrase);
                    }
                    throw new AiException("AI 接口返回错误：" + message);
                }
                if (root == null)
                {
                    throw new AiException("AI 接口返回的内容无法解析。");
                }

                AiResult result = BuildResult(root, model);
                Log.Info("AI 返回: contentLen=" + (result.Content == null ? 0 : result.Content.Length)
                         + ", finish=" + result.FinishReason
                         + ", reasoningTokens=" + result.ReasoningTokens
                         + ", completionTokens=" + result.CompletionTokens);

                if (!string.IsNullOrWhiteSpace(result.Content))
                {
                    result.Content = result.Content.Trim();
                    return result;
                }

                // 推理模型把输出额度用完了：自动加大 max_tokens 再试一次
                if (attempt == 1 && requestMaxTokens < RetryMaxTokens && IsLengthLimited(result.FinishReason))
                {
                    requestMaxTokens = RetryMaxTokens;
                    Log.Warn("模型输出为空（finish_reason=" + result.FinishReason
                             + ", reasoningTokens=" + result.ReasoningTokens
                             + "），自动提高 max_tokens 到 " + requestMaxTokens + " 重试");
                    continue;
                }

                throw new AiException(BuildEmptyContentMessage(result));
            }
        }

        /// <summary>用一条短请求测试接口是否可用。</summary>
        public static async Task<string> TestAsync(
            string baseUrl, string apiKey, string model,
            CancellationToken cancellationToken = default(CancellationToken))
        {
            var result = await ChatAsync(
                baseUrl, apiKey, model,
                "你是一个测试助手，只需按要求回复，不要多说。",
                "请只回复两个字：可用",
                0.0, 1000, cancellationToken).ConfigureAwait(false);

            return "连接成功（模型：" + (string.IsNullOrWhiteSpace(result.Model) ? model : result.Model) + "）";
        }

        private static async Task<RawResponse> SendAsync(
            string baseUrl,
            string apiKey,
            string model,
            string systemPrompt,
            string userPrompt,
            double temperature,
            int maxTokens,
            CancellationToken cancellationToken)
        {
            string url = baseUrl.TrimEnd('/') + "/chat/completions";

            var payload = new Dictionary<string, object>
            {
                { "model", model.Trim() },
                {
                    "messages", new object[]
                    {
                        new Dictionary<string, object> { { "role", "system" }, { "content", systemPrompt } },
                        new Dictionary<string, object> { { "role", "user" }, { "content", userPrompt } }
                    }
                },
                { "temperature", temperature },
                { "max_tokens", maxTokens },
                { "stream", false }
            };

            string requestBody = Json.Serialize(payload);
            var raw = new RawResponse();

            using (var request = new HttpRequestMessage(HttpMethod.Post, url))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey.Trim());
                request.Content = new StringContent(requestBody, new UTF8Encoding(false), "application/json");

                HttpResponseMessage response;
                try
                {
                    response = await Http.SendAsync(request, cancellationToken).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    throw new AiException("请求超时或已取消，请稍后重试。");
                }
                catch (Exception ex)
                {
                    throw new AiException("无法连接 AI 接口（" + baseUrl + "）：" + ex.Message);
                }

                using (response)
                {
                    raw.Body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    raw.StatusCode = response.StatusCode;
                    raw.ReasonPhrase = response.ReasonPhrase;
                }
            }
            return raw;
        }

        private static AiResult BuildResult(Dictionary<string, object> root, string fallbackModel)
        {
            var result = new AiResult();

            object modelValue;
            if (root.TryGetValue("model", out modelValue) && modelValue != null)
            {
                result.Model = Convert.ToString(modelValue) ?? "";
            }
            if (string.IsNullOrWhiteSpace(result.Model)) { result.Model = fallbackModel ?? ""; }

            object usageValue;
            if (root.TryGetValue("usage", out usageValue))
            {
                var usage = AsDict(usageValue);
                if (usage != null)
                {
                    result.PromptTokens = GetInt(usage, "prompt_tokens");
                    result.CompletionTokens = GetInt(usage, "completion_tokens");
                    result.TotalTokens = GetInt(usage, "total_tokens");

                    object detailsValue;
                    if (usage.TryGetValue("completion_tokens_details", out detailsValue))
                    {
                        var details = AsDict(detailsValue);
                        if (details != null) { result.ReasoningTokens = GetInt(details, "reasoning_tokens"); }
                    }
                }
            }

            object choicesValue;
            if (!root.TryGetValue("choices", out choicesValue)) { return result; }
            var choices = AsList(choicesValue);
            if (choices == null || choices.Count == 0) { return result; }

            var first = AsDict(choices[0]);
            if (first == null) { return result; }

            object finishValue;
            if (first.TryGetValue("finish_reason", out finishValue) && finishValue != null)
            {
                result.FinishReason = Convert.ToString(finishValue) ?? "";
            }

            object messageValue;
            if (!first.TryGetValue("message", out messageValue)) { return result; }
            var message = AsDict(messageValue);
            if (message == null) { return result; }

            result.Content = ExtractContentValue(message);
            return result;
        }

        /// <summary>content 可能是字符串，也可能是 [{type,text}] 数组（兼容处理）。</summary>
        private static string ExtractContentValue(Dictionary<string, object> message)
        {
            object contentValue;
            if (!message.TryGetValue("content", out contentValue) || contentValue == null) { return ""; }

            var text = contentValue as string;
            if (text != null) { return text; }

            var parts = AsList(contentValue);
            if (parts != null)
            {
                var sb = new StringBuilder();
                foreach (object part in parts)
                {
                    var partDict = AsDict(part);
                    if (partDict != null)
                    {
                        object textValue;
                        if (partDict.TryGetValue("text", out textValue) && textValue != null)
                        {
                            if (sb.Length > 0) { sb.AppendLine(); }
                            sb.Append(Convert.ToString(textValue));
                        }
                    }
                    else if (part != null)
                    {
                        if (sb.Length > 0) { sb.AppendLine(); }
                        sb.Append(Convert.ToString(part));
                    }
                }
                return sb.ToString();
            }

            return Convert.ToString(contentValue) ?? "";
        }

        private static bool IsLengthLimited(string finishReason)
        {
            return string.IsNullOrEmpty(finishReason)
                   || string.Equals(finishReason, "length", StringComparison.OrdinalIgnoreCase);
        }

        private static string BuildEmptyContentMessage(AiResult result)
        {
            var sb = new StringBuilder();
            sb.Append("模型返回了空内容");
            if (!string.IsNullOrEmpty(result.FinishReason))
            {
                sb.Append("（finish_reason=").Append(result.FinishReason).Append("）");
            }
            if (result.ReasoningTokens > 0)
            {
                sb.Append("，其中 ").Append(result.ReasoningTokens).Append(" tokens 消耗在推理上");
            }
            sb.Append("。");
            sb.Append("deepseek-flash 是推理模型：长文本翻译/润色时推理会先占满输出额度，导致正式回复为空。");
            sb.Append("可以重试一次；如果反复出现，建议把模型改成 deepseek-chat（非推理模型），或缩短正文。");
            return sb.ToString();
        }

        private static System.Collections.IList AsList(object value)
        {
            return value as System.Collections.IList;
        }

        private static Dictionary<string, object> AsDict(object value)
        {
            if (value == null) { return null; }
            var dictionary = value as Dictionary<string, object>;
            if (dictionary != null) { return dictionary; }

            var generic = value as System.Collections.IDictionary;
            if (generic == null) { return null; }
            var converted = new Dictionary<string, object>();
            foreach (System.Collections.DictionaryEntry entry in generic)
            {
                converted[Convert.ToString(entry.Key)] = entry.Value;
            }
            return converted;
        }

        private static string ExtractErrorMessage(Dictionary<string, object> root)
        {
            if (root == null) { return null; }
            object errorValue;
            if (!root.TryGetValue("error", out errorValue)) { return null; }
            var error = AsDict(errorValue);
            if (error == null) { return Convert.ToString(errorValue); }
            object messageValue;
            if (error.TryGetValue("message", out messageValue) && messageValue != null)
            {
                return Convert.ToString(messageValue);
            }
            return Json.Serialize(error);
        }

        private static int GetInt(Dictionary<string, object> dict, string key)
        {
            object value;
            if (dict.TryGetValue(key, out value) && value != null)
            {
                try { return Convert.ToInt32(value); } catch { }
            }
            return 0;
        }
    }
}
