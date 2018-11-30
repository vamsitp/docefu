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
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading.Tasks;
    using Newtonsoft.Json;
    using Word = Microsoft.Office.Interop.Word;
    using System.Windows.Forms;

    public class Program
    {
        private const string RelationsQueryPath = "wit/wiql?api-version=2.3"; //"queries/Shared Queries/EFUs";
        private const string WorkItemsQueryPath = "wit/workitems?ids={0}&api-version=2.3"; //"queries/Shared Queries/EFUs";
        private const string SecurityTokensUrl = "_details/security/tokens";

        private static readonly string TeamSite = ConfigurationManager.AppSettings["TeamSite"];
        private static readonly string TeamProject = ConfigurationManager.AppSettings["TeamProject"];
        private static readonly string WorkItemsQuery = ConfigurationManager.AppSettings["WorkItemsQuery"];
        private static readonly string WordTemplate = ConfigurationManager.AppSettings["WordTemplate"];
        private static readonly string PersonalAccessToken = ConfigurationManager.AppSettings["PersonalAccessToken"];
        private static readonly string HeadersColor = ConfigurationManager.AppSettings["HeadersColor"];
        private static readonly string[] ColorReplacements = ConfigurationManager.AppSettings["ColorReplacements"].Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
        private const string WorkItemsJson = "./WorkItems.json";
        private const string HtmlNewLine = "<br/>";
        private const string Dot = ". ";
        private static List<EFU> efus = null;

        [STAThread]
        public static void Main(string[] args)
        {
            MainAsync(args).Wait();
            GenerateDoc(string.Join(HtmlNewLine, efus.Select(GetContent)));
            Console.ReadLine();
            Process.Start(Environment.CurrentDirectory);
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
                Console.WriteLine($"Loaded {efus?.Count} work-items from {WorkItemsJson}");
                await GetWorkItems();
            }
            else
            {
                if (string.IsNullOrWhiteSpace(PersonalAccessToken))
                {
                    var tokenUrl = $"{TeamSite}{SecurityTokensUrl}";
                    Console.WriteLine($"PersonalAccessToken is blank in the config!\nHit ENTER to generate it (select 'All Scopes') at: {tokenUrl}");
                    Console.ReadLine();
                    Process.Start(tokenUrl);
                    Console.WriteLine($"Update 'PersonalAccessToken' value in VstsUwpPackageInstaller.config with the generated/copied token and restart the app.");
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

        private async static Task GetWorkItems()
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
                var wis = await GetData<WiqlRelationList>(RelationsQueryPath, string.Empty, "{\"query\": \"" + string.Format(WorkItemsQuery, TeamProject) + "\"}");
                Console.WriteLine($"Workitem relations fetched: {wis.workItemRelations.Length}");
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
                    Console.WriteLine($" {wi.fields.SystemWorkItemType} ".PadRight(13)  + wi.id.ToString().PadLeft(4) + Dot + wi.fields.SystemTitle + $" [{wi.fields.SystemTags}]");
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
            var result = await GetData<WorkItem>(item.url, string.Empty, string.Empty);
            if (result != null)
            {
                Console.WriteLine(result.id.ToString().PadLeft(4) + $" {result.fields.SystemWorkItemType} - ".PadLeft(14) + result.fields.SystemTitle);
                return result;
            }

            return null;
        }

        private static async Task<List<WorkItem>> GetWorkItems(List<WorkitemRelation> relations)
        {
            if (relations != null && relations.Any())
            {
                var builder = new StringBuilder();
                foreach (var item in relations.Select(x => x.target))
                {
                    builder.Append(item.id.ToString()).Append(',');
                }

                var ids = builder.ToString().TrimEnd(',');

                if (!string.IsNullOrWhiteSpace(ids))
                {
                    var workItems = await GetData<WorkItems>(string.Format(WorkItemsQueryPath, ids), string.Empty, string.Empty);
                    if (workItems != null)
                    {
                        return workItems.Items.ToList();
                    }
                }
            }

            return null;
        }

        private static async Task GetWorkItemsByQuery(List<WorkItem> workItems = null)
        {
            if (workItems == null)
            {
                workItems = new List<WorkItem>();
                // var wis = await GetData<WiqlList>(WorkItemsQueryPath, TeamProject, string.Empty);
                var wis = await GetData<WiqlList>(RelationsQueryPath, string.Empty, "{\"query\": \"" + string.Format(WorkItemsQuery, TeamProject) + "\"}");
                foreach (var wi in wis.workItems)
                {
                    var result = await GetData<WorkItem>(wi.url, string.Empty, string.Empty);
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
            Console.WriteLine($"Generating Document from Workitems...");
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
            if (efu == null) return string.Empty;

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
            var val = string.Format("Version:0.9\nStartHTML:80\nEndHTML:{0,8}\nStartFragment:80\nEndFragment:{0,8}\n", 80 + html.Length) + html + "<";
            Clipboard.SetData(DataFormats.Html, val);
            // Clipboard.SetText(val, TextDataFormat.Html);

            rng.PasteSpecial(DataType: Word.WdPasteDataType.wdPasteHTML);
            rng.Font.Name = "Segoe UI";
            rng.Font.Size = 11;
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
                client.BaseAddress = new Uri(TeamSite);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", credentials);
                if (!path.StartsWith(TeamSite, StringComparison.OrdinalIgnoreCase)) path = $"{project}/_apis/{path}";
                Trace.TraceInformation($"BaseAddress: {TeamSite} | Path: {path} | Content: {postData}");
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
