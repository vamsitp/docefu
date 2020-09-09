namespace EpicDoc
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;

    using ColoredConsole;

    using Newtonsoft.Json;

    public class Program
    {
        private const string WorkItemsJsonFile = "./WorkItems.json";

        private static List<EFU> efus = null;

        [STAThread]
        public static void Main(string[] args)
        {
            Execute(args).Wait();
            if (args?.ContainsArg("w") == true)
            {
                Doc.Generate(efus);
            }
            else if (args?.ContainsArg("p") == true)
            {
                Deck.Generate(efus);
            }
            else
            {
                Doc.Generate(efus);
                Deck.Generate(efus);
            }

            Console.ReadLine();
        }

        public static async Task Execute(string[] args)
        {
            try
            {
                if (args?.ContainsArg("i") == true)
                {
                    if (File.Exists(WorkItemsJsonFile))
                    {
                        await ProcessWorkItems(true);
                    }
                    else
                    {
                        ColorConsole.WriteLine($"{WorkItemsJsonFile} not found! Fetching details from Azure DevOps...".Yellow());
                        await ProcessWorkItems();
                    }
                }
                else
                {
                    CheckSettings();
                    ColorConsole.WriteLine($"You can use an existing {WorkItemsJsonFile} file as an input using '/i'".Yellow());
                    await ProcessWorkItems();
                }
            }
            catch (Exception ex)
            {
                WriteError(ex.Message);
            }
        }

        private static void CheckSettings()
        {
            if (string.IsNullOrWhiteSpace(AzDO.Account) || string.IsNullOrWhiteSpace(AzDO.Project) || string.IsNullOrWhiteSpace(AzDO.Token))
            {
                ColorConsole.WriteLine("\n Please provide Azure DevOps details in the format (without braces): <Account> <Project> <PersonalAccessToken>".Black().OnCyan());
                var details = Console.ReadLine().Split(' ')?.ToList();
                var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var section = config.Sections.OfType<AppSettingsSection>().FirstOrDefault();
                var settings = section.Settings;

                for (var i = 0; i < 3; i++) // Only 4 values required
                {
                    var key = settings.AllKeys[i];
                    settings[key].Value = details[i];
                }

                config.Save(ConfigurationSaveMode.Minimal);
                ConfigurationManager.RefreshSection(section.SectionInformation.Name);
            }
        }

        private static async Task ProcessWorkItems(bool local = false)
        {
            if (local)
            {
                efus = JsonConvert.DeserializeObject<List<EFU>>(File.ReadAllText(WorkItemsJsonFile));
                ColorConsole.WriteLine($"Loaded {efus?.Count} Work-items from {WorkItemsJsonFile}");
                await GetWorkItems();
            }
            else
            {
                await GetWorkItems();
            }
        }

        private static void WriteError(string error)
        {
            ColorConsole.WriteLine($"Error: {error}".White().OnRed());
        }

        private static async Task GetWorkItems()
        {
            if (efus == null)
            {
                // efus = GetWorkItemsByQuery(workItems);
                efus = await AzDO.GetWorkItemsByStoredQuery(WorkItemsJsonFile); //.ContinueWith(ContinuationAction);
            }
        }

        private static void ContinuationAction(Task task)
        {
            Doc.Generate(efus);
        }
    }
}
