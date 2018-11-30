namespace EpicDocx
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading.Tasks;

    using Newtonsoft.Json;

    using Word = Microsoft.Office.Interop.Word;

    public class Program
    {
        private const string RelationsQueryPath = "wit/wiql?api-version=4.1-preview"; //"queries/Shared Queries/EFUs";
        private const string WorkItemsQueryPath = "wit/workitems?ids={0}&api-version=4.1-preview"; //"queries/Shared Queries/EFUs";
        private const string SecurityTokensUrl = "_details/security/tokens";

        private static readonly string Account = ConfigurationManager.AppSettings["Account"];
        private static readonly string Project = ConfigurationManager.AppSettings["Project"];
        private static readonly string AreaPath = ConfigurationManager.AppSettings["AreaPath"];
        private static readonly string WorkItemsQuery = ConfigurationManager.AppSettings["WorkItemsQuery"];
        private static readonly string WordTemplate = ConfigurationManager.AppSettings["WordTemplate"];
        private static readonly string PersonalAccessToken = ConfigurationManager.AppSettings["PersonalAccessToken"];
        private static readonly string HeadersColor = ConfigurationManager.AppSettings["HeadersColor"];
        private static readonly string[] ColorReplacements = ConfigurationManager.AppSettings["ColorReplacements"].Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
        private const string WorkItemsJson = "./WorkItems.json";
        private const string HtmlNewLine = "<br/>";
        private const string Dot = ". ";
        private const int Max = 200;
        private static List<EFU> efus = null;

        [STAThread]
        public static void Main(string[] args)
        {
            MainAsync(args).Wait();
            GenerateDoc(string.Join(HtmlNewLine, efus.Select(GetContent)));
            Console.ReadLine();
        }

        public static async Task MainAsync(string[] args)
        {
            try
            {
                if (args != null && args.Length > 0 && (args[0].EndsWith("-i", StringComparison.OrdinalIgnoreCase) || args[0].EndsWith("/i", StringComparison.OrdinalIgnoreCase)))
                {
                    if (File.Exists(WorkItemsJson))
                    {
                        await ProcessWorkItems(true);
                    }
                    else
                    {
                        Console.WriteLine($"{WorkItemsJson} not found! Fetching details from VSTS...");
                        await ProcessWorkItems();
                    }
                }
                else
                {
                    await ProcessWorkItems();
                }
            }
            catch (Exception ex)
            {
                WriteError(ex.Message);
            }
        }

        private static async Task ProcessWorkItems(bool local = false)
        {
            if (local)
            {
                efus = JsonConvert.DeserializeObject<List<EFU>>(File.ReadAllText(WorkItemsJson));
                Console.WriteLine($"Loaded {efus?.Count} Work-items from {WorkItemsJson}");
                await GetWorkItems();
            }
            else
            {
                if (string.IsNullOrWhiteSpace(PersonalAccessToken))
                {
                    var tokenUrl = $"{Account}{SecurityTokensUrl}";
                    Console.WriteLine($"PersonalAccessToken is blank in the config!\nHit ENTER to generate it (select 'All Scopes') at: {tokenUrl}");
                    Console.ReadLine();
                    Process.Start(tokenUrl);
                    Console.WriteLine($"Update 'PersonalAccessToken' value in EpicDocx.exe.config with the generated/copied token and restart the app.");
                }
                else
                {
                    await GetWorkItems();
                }
            }
        }

        private static void WriteError(string error)
        {
            Console.WriteLine("Error: " + error);
        }

        private static async Task GetWorkItems()
        {
            // GetWorkItemsByQuery(workItems);
            await GetWorkItemsByStoredQuery(); //.ContinueWith(ContinuationAction);
        }

        private static void ContinuationAction(Task task)
        {
            GenerateDoc(string.Join(HtmlNewLine, efus.Select(GetContent)));
        }

        private static async Task GetWorkItemsByStoredQuery()
        {
            if (efus == null)
            {
                efus = new List<EFU>();
                var wis = await GetData<WiqlRelationList>(RelationsQueryPath, Project, "{\"query\": \"" + string.Format(WorkItemsQuery, Project, string.IsNullOrWhiteSpace(AreaPath) ? Project : AreaPath) + "\"}");
                Console.WriteLine($"Work-item relations fetched: {wis.workItemRelations.Length}");
                var rootItems = wis.workItemRelations.Where(x => x.source == null).ToList();
                await IterateWorkItems(rootItems, null, wis);
                File.WriteAllText(WorkItemsJson, JsonConvert.SerializeObject(efus));
            }
        }

        private static async Task IterateWorkItems(List<WorkitemRelation> relations, EFU parent, WiqlRelationList wis)
        {
            if (relations.Count > 0)
            {
                var workitems = await GetWorkItems(relations.ToList());
                foreach (var wi in workitems)
                {
                    Console.WriteLine($" {wi.fields.SystemWorkItemType} ".PadRight(13) + wi.id.ToString().PadLeft(4) + Dot + wi.fields.SystemTitle + $" [{wi.fields.SystemTags}]");
                    var efu = new EFU
                    {
                        Id = wi.id,
                        Title = wi.fields.SystemTitle,
                        Description = wi.fields.SystemDescription,
                        Workitemtype = wi.fields.SystemWorkItemType,
                        AcceptanceCriteria = wi.fields.MicrosoftVSTSCommonAcceptanceCriteria,
                        Tags = wi.fields.SystemTags,
                        Parent = parent?.Id
                    };

                    efus.Add(efu);
                    parent?.Children.Add(efu.Id);
                    await IterateWorkItems(wis.workItemRelations.Where(x => x.source != null && x.source.id.Equals(wi.id)).ToList(), efu, wis);
                }
            }
        }

        private static async Task<WorkItem> GetWorkItem(WiqlWorkitem item)
        {
            var result = await GetData<WorkItem>(item.url, Project, string.Empty);
            if (result != null)
            {
                Console.WriteLine(result.id.ToString().PadLeft(4) + $" {result.fields.SystemWorkItemType} - ".PadLeft(14) + result.fields.SystemTitle);
                return result;
            }

            return null;
        }

        private static async Task<List<WorkItem>> GetWorkItems(List<WorkitemRelation> items)
        {
            var result = new List<WorkItem>();
            var splitItems = SplitList(items);
            if (splitItems?.Any() == true)
            {
                foreach (var relations in splitItems)
                {
                    var builder = new StringBuilder();
                    foreach (var item in relations.Select(x => x.target))
                    {
                        builder.Append(item.id.ToString()).Append(',');
                    }

                    var ids = builder.ToString().TrimEnd(',');
                    if (!string.IsNullOrWhiteSpace(ids))
                    {
                        var workItems = await GetData<WorkItems>(string.Format(WorkItemsQueryPath, ids), Project, string.Empty);
                        if (workItems != null)
                        {
                            result.AddRange(workItems.Items);
                        }
                    }
                }
            }

            return result;
        }

        // Credit: https://stackoverflow.com/a/11463800
        public static IEnumerable<List<T>> SplitList<T>(List<T> list, int limit = Max)
        {
            if (list?.Any() == true)
            {
                for (var i = 0; i < list.Count; i += limit)
                {
                    yield return list.GetRange(i, Math.Min(limit, list.Count - i));
                }
            }
        }

        private static async Task GetWorkItemsByQuery(List<WorkItem> workItems = null)
        {
            if (workItems == null)
            {
                workItems = new List<WorkItem>();
                // var wis = await GetData<WiqlList>(WorkItemsQueryPath, Project, string.Empty);
                var wis = await GetData<WiqlList>(RelationsQueryPath, Project, "{\"query\": \"" + string.Format(WorkItemsQuery, Project, string.IsNullOrWhiteSpace(AreaPath) ? Project : AreaPath) + "\"}");
                foreach (var wi in wis.workItems)
                {
                    var result = await GetData<WorkItem>(wi.url, Project, string.Empty);
                    if (result != null)
                    {
                        Console.WriteLine(result.id.ToString().PadLeft(4) + ": " + result.fields.SystemTitle);
                        workItems.Add(result);
                    }
                    else
                    {
                        Console.WriteLine($"Unable to fetch details for: {wi.id}");
                    }
                }

                File.WriteAllText(WorkItemsJson, JsonConvert.SerializeObject(workItems));
            }

            // TODO: GenerateDoc(string.Join(HtmlNewLine, workItems.Select(GetContent)));
        }

        private static void GenerateDoc(string content)
        {
            Console.WriteLine($"Generating Document from Work-items...");
            var wordApp = new Word.Application { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone, ScreenUpdating = false };
            object fileName = Path.Combine(Environment.CurrentDirectory, WordTemplate);
            object missing = Type.Missing;
            var wordDoc = wordApp.Documents.Open(
                ref fileName,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing);

            var value = "<html><body style=\"font-family:'segoe ui';font-size:14px\">" + content + "</body></html>";
            var bookmark = wordDoc.Bookmarks.get_Item(1);
            ReplaceBookmark(bookmark.Range, value);

            object saveTo = Path.Combine(Environment.CurrentDirectory, "FuncSpec (UserStories).docx");
            object format = Word.WdSaveFormat.wdFormatXMLDocument;

            object start = wordDoc.Content.Start;
            object end = wordDoc.Content.End;
            wordDoc.Range(ref start, ref end).Select();
            wordApp.Selection.Range.Font.Name = "Segoe UI";
            wordApp.Selection.Range.Font.Size = 10;

            wordDoc.SaveAs(ref saveTo, ref format);
            wordDoc.Close(ref missing, ref missing, ref missing);
            NAR(wordDoc);
            wordApp.Quit(ref missing, ref missing, ref missing);
            NAR(wordApp);
            Console.WriteLine($"Done creating {saveTo}");
        }

        private static string GetContent(EFU efu)
        {
            if (efu == null)
            {
                return string.Empty;
            }

            var desc = string.IsNullOrWhiteSpace(efu.Description) ? string.Empty : Trim(efu.Description);
            if (efu.Workitemtype.Equals("Epic", StringComparison.OrdinalIgnoreCase))
            {
                return $"<div style=\"color:#969696\"><b>" + efu.Id + ". " + (efu.Title?.ToUpperInvariant() ?? string.Empty) + "</b></div>" + desc;
            }

            if (efu.Workitemtype.Equals("Feature", StringComparison.OrdinalIgnoreCase))
            {
                return $"<div style=\"color:#969696\">" + efu.Id + ". " + (efu.Title?.ToUpperInvariant() ?? string.Empty) + "</div>" + desc;
            }

            var acceptance = string.IsNullOrWhiteSpace(efu.AcceptanceCriteria) ? string.Empty : Trim(efu.AcceptanceCriteria);
            return $"<div style=\"color:{HeadersColor}\">" + efu.Id + ". " + (efu.Title ?? string.Empty) + "</div>" + desc + "<b>Acceptance Criteria</b>: " + acceptance + HtmlNewLine;
        }

        private static string Trim(string content)
        {
            content = Regex.Replace(content, "</?(font|span)[^>]*>", string.Empty, RegexOptions.IgnoreCase).Trim('\n');
            // content = Regex.Replace(content, "(border-color)[^;]*", $"border-color:{HeadersColor}", RegexOptions.IgnoreCase);
            // content.Replace("rgb(0, 0, 0)", HeadersColor).Replace("black", HeadersColor).Replace("#f0f0f0", HeadersColor).Replace("windowtext", HeadersColor);
            foreach (var replace in ColorReplacements)
            {
                content = content.Replace(replace, HeadersColor);
            }

            return content;
        }

        public static void ReplaceBookmark(Word.Range rng, string html)
        {
            // var val = string.Format("Version:0.9\nStartHTML:80\nEndHTML:{0,8}\nStartFragment:80\nEndFragment:{0,8}\n", 80 + html.Length) + html + "<";
            // Clipboard.SetData(DataFormats.Html, val);
            //// Clipboard.SetText(val, TextDataFormat.Html);
            // rng.PasteSpecial(DataType: Word.WdPasteDataType.wdPasteHTML);

            rng.Font.Name = "Segoe UI";
            rng.Font.Size = 11;
            var file = Path.Combine(Path.GetDirectoryName(Assembly.GetCallingAssembly().CodeBase.Replace("file:///", string.Empty)), "temp.html");
            File.WriteAllText(file, html);
            rng.InsertFile(file);
        }

        protected static void NAR(object o)
        {
            try
            {
                if (o != null)
                {
                    Marshal.FinalReleaseComObject(o);
                }
            }
            finally
            {
                o = null;
            }
        }

        private static async Task<T> GetData<T>(string path, string project, string postData)
        {
            // https://www.visualstudio.com/en-us/docs/integrate/api/wit/samples
            using (var client = new HttpClient())
            {
                var credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes($":{PersonalAccessToken}"));
                client.BaseAddress = new Uri(Account);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", credentials);
                if (!path.StartsWith(Account, StringComparison.OrdinalIgnoreCase))
                {
                    path = $"{project}/_apis/{path}";
                }

                Trace.TraceInformation($"BaseAddress: {Account} | Path: {path} | Content: {postData}");
                HttpResponseMessage queryHttpResponseMessage;

                if (string.IsNullOrWhiteSpace(postData))
                {
                    queryHttpResponseMessage = await client.GetAsync(path);
                }
                else
                {
                    var content = new StringContent(postData, Encoding.UTF8, "application/json");
                    queryHttpResponseMessage = await client.PostAsync(path, content);
                }

                if (queryHttpResponseMessage.IsSuccessStatusCode)
                {
                    var result = await queryHttpResponseMessage.Content.ReadAsStringAsync();
                    return JsonConvert.DeserializeObject<T>(result);
                }
                else
                {
                    throw new Exception($"{queryHttpResponseMessage.ReasonPhrase}");
                }
            }
        }
    }

}
