namespace EpicDoc
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Threading.Tasks;

    using ColoredConsole;

    using Newtonsoft.Json;

    internal class AzDO
    {
        internal static string Account => ConfigurationManager.AppSettings["Account"];
        internal static string Project => ConfigurationManager.AppSettings["Project"];
        internal static string Token => ConfigurationManager.AppSettings["PersonalAccessToken"];

        // private const string SecurityTokensUrl = "_details/security/tokens";
        private const string RelationsQueryPath = "wit/wiql?api-version=4.1-preview"; //"queries/Shared Queries/EFUs";
        private const string WorkItemsQueryPath = "wit/workitems?ids={0}&api-version=4.1-preview"; //"queries/Shared Queries/EFUs";

        private static readonly string WorkItemsQuery = ConfigurationManager.AppSettings["WorkItemsQuery"];

        private const string Dot = ". ";

        internal static async Task<List<EFU>> GetWorkItemsByStoredQuery(string workItemsJsonFile)
        {
            var efus = new List<EFU>();
            var wis = await GetData<WiqlRelationList>(RelationsQueryPath, Project, "{\"query\": \"" + string.Format(WorkItemsQuery, Project) + "\"}"); // AND [Source].[System.AreaPath] UNDER '{1}'
            ColorConsole.WriteLine($"Work-item relations fetched: {wis.workItemRelations.Length}");
            var rootItems = wis.workItemRelations.Where(x => x.source == null).ToList();
            await IterateWorkItems(efus, rootItems, null, wis);
            File.WriteAllText(workItemsJsonFile, JsonConvert.SerializeObject(efus));
            return efus;
        }

        private static async Task IterateWorkItems(List<EFU> efus, List<WorkitemRelation> relations, EFU parent, WiqlRelationList wis)
        {
            if (relations.Count > 0)
            {
                var workitems = await GetWorkItems(relations.ToList());
                foreach (var wi in workitems)
                {
                    ColorConsole.WriteLine($" {wi.fields.SystemWorkItemType} ".PadRight(13) + wi.id.ToString().PadLeft(4) + Dot + wi.fields.SystemTitle + $" [{wi.fields.SystemTags}]");
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
                    await IterateWorkItems(efus, wis.workItemRelations.Where(x => x.source != null && x.source.id.Equals(wi.id)).ToList(), efu, wis);
                }
            }
        }

        private static async Task<WorkItem> GetWorkItem(WiqlWorkitem item)
        {
            var result = await GetData<WorkItem>(item.url, Project, string.Empty);
            if (result != null)
            {
                ColorConsole.WriteLine(result.id.ToString().PadLeft(4) + $" {result.fields.SystemWorkItemType} - ".PadLeft(14) + result.fields.SystemTitle);
                return result;
            }

            return null;
        }

        private static async Task<List<WorkItem>> GetWorkItems(List<WorkitemRelation> items)
        {
            var result = new List<WorkItem>();
            var splitItems = items.SplitList();
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

        internal static async Task<List<WorkItem>> GetWorkItemsByQuery(string workItemsJsonFile)
        {
            var workItems = new List<WorkItem>();
            // var wis = await GetData<WiqlList>(WorkItemsQueryPath, Project, string.Empty);
            var wis = await GetData<WiqlList>(RelationsQueryPath, Project, "{\"query\": \"" + string.Format(WorkItemsQuery, Project) + "\"}");
            foreach (var wi in wis.workItems)
            {
                var result = await GetData<WorkItem>(wi.url, Project, string.Empty);
                if (result != null)
                {
                    ColorConsole.WriteLine(result.id.ToString().PadLeft(4) + ": " + result.fields.SystemTitle);
                    workItems.Add(result);
                }
                else
                {
                    ColorConsole.WriteLine($"Unable to fetch details for: {wi.id}".Red());
                }
            }

            File.WriteAllText(workItemsJsonFile, JsonConvert.SerializeObject(workItems));
            return workItems;

            // TODO: GenerateDoc(string.Join(HtmlNewLine, workItems.Select(GetContent)));
        }

        private static async Task<T> GetData<T>(string path, string project, string postData)
        {
            // https://www.visualstudio.com/en-us/docs/integrate/api/wit/samples
            using (var client = new HttpClient())
            {
                var credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes($":{Token}"));
                client.BaseAddress = new Uri(Account.IndexOf(".com", StringComparison.OrdinalIgnoreCase) > 0 ? Account : $"https://{Account}.visualstudio.com");
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
