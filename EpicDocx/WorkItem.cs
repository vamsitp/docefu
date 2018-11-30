using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace EpicDocx
{
    public class WiqlList
    {
        public string queryType { get; set; }
        public string queryResultType { get; set; }
        public DateTime asOf { get; set; }
        public Column[] columns { get; set; }
        public Sortcolumn[] sortColumns { get; set; }
        public WiqlWorkitem[] workItems { get; set; }
    }

    public class WiqlRelationList
    {
        public string queryType { get; set; }
        public string queryResultType { get; set; }
        public DateTime asOf { get; set; }
        public Column[] columns { get; set; }
        public Sortcolumn[] sortColumns { get; set; }
        public WorkitemRelation[] workItemRelations { get; set; }
    }

    public class Column
    {
        public string referenceName { get; set; }
        public string name { get; set; }
        public string url { get; set; }
    }

    public class Sortcolumn
    {
        public Field field { get; set; }
        public bool descending { get; set; }
    }

    public class Field
    {
        public string referenceName { get; set; }
        public string name { get; set; }
        public string url { get; set; }
    }

    public class WiqlWorkitem
    {
        public int id { get; set; }
        public string url { get; set; }
    }


    public class WorkItem
    {
        public int id { get; set; }
        public int rev { get; set; }
        public Fields fields { get; set; }
        public Links _links { get; set; }
        public string url { get; set; }
    }

    public class Fields
    {
        [JsonProperty("System.AreaPath")]
        public string SystemAreaPath { get; set; }

        [JsonProperty("System.TeamProject")]
        public string SystemTeamProject { get; set; }

        [JsonProperty("System.IterationPath")]
        public string SystemIterationPath { get; set; }

        [JsonProperty("System.WorkItemType")]
        public string SystemWorkItemType { get; set; }

        [JsonProperty("System.State")]
        public string SystemState { get; set; }

        [JsonProperty("System.Reason")]
        public string SystemReason { get; set; }

        [JsonProperty("System.AssignedTo")]
        public string SystemAssignedTo { get; set; }

        [JsonProperty("System.CreatedDate")]
        public DateTime SystemCreatedDate { get; set; }

        [JsonProperty("System.CreatedBy")]
        public string SystemCreatedBy { get; set; }

        [JsonProperty("System.ChangedDate")]
        public DateTime SystemChangedDate { get; set; }

        [JsonProperty("System.ChangedBy")]
        public string SystemChangedBy { get; set; }

        [JsonProperty("System.Title")]
        public string SystemTitle { get; set; }

        [JsonProperty("System.BoardColumn")]
        public string SystemBoardColumn { get; set; }

        [JsonProperty("System.BoardColumnDone")]
        public bool SystemBoardColumnDone { get; set; }

        [JsonProperty("Microsoft.VSTS.Common.StateChangeDate")]
        public DateTime MicrosoftVSTSCommonStateChangeDate { get; set; }

        [JsonProperty("Microsoft.VSTS.Common.ActivatedDate")]
        public DateTime MicrosoftVSTSCommonActivatedDate { get; set; }

        [JsonProperty("Microsoft.VSTS.Common.ActivatedBy")]
        public string MicrosoftVSTSCommonActivatedBy { get; set; }

        [JsonProperty("Microsoft.VSTS.Common.Priority")]
        public int MicrosoftVSTSCommonPriority { get; set; }

        [JsonProperty("Microsoft.VSTS.Common.StackRank")]
        public float MicrosoftVSTSCommonStackRank { get; set; }

        [JsonProperty("Microsoft.VSTS.Common.ValueArea")]
        public string MicrosoftVSTSCommonValueArea { get; set; }

        [JsonProperty("WEF_27196AF0D18D4A859EDC1E112AAFAAD9_Kanban.Column")]
        public string WEF_27196AF0D18D4A859EDC1E112AAFAAD9_KanbanColumn { get; set; }

        [JsonProperty("WEF_27196AF0D18D4A859EDC1E112AAFAAD9_Kanban.Column.Done")]
        public bool WEF_27196AF0D18D4A859EDC1E112AAFAAD9_KanbanColumnDone { get; set; }

        [JsonProperty("System.Description")]
        public string SystemDescription { get; set; }

        [JsonProperty("Microsoft.VSTS.Common.AcceptanceCriteria")]
        public string MicrosoftVSTSCommonAcceptanceCriteria { get; set; }

        [JsonProperty("System.Tags")]
        public string SystemTags { get; set; }
    }

    public class Links
    {
        public Self self { get; set; }
        public Workitemupdates workItemUpdates { get; set; }
        public Workitemrevisions workItemRevisions { get; set; }
        public Workitemhistory workItemHistory { get; set; }
        public Html html { get; set; }
        public Workitemtype workItemType { get; set; }
        public LinkFields fields { get; set; }
    }

    public class Self
    {
        public string href { get; set; }
    }

    public class Workitemupdates
    {
        public string href { get; set; }
    }

    public class Workitemrevisions
    {
        public string href { get; set; }
    }

    public class Workitemhistory
    {
        public string href { get; set; }
    }

    public class Html
    {
        public string href { get; set; }
    }

    public class Workitemtype
    {
        public string href { get; set; }
    }

    public class LinkFields
    {
        public string href { get; set; }
    }
    
    public class WorkitemRelation
    {
        public WiqlWorkitem target { get; set; }
        public string rel { get; set; }
        public WiqlWorkitem source { get; set; }
    }

    public class EFU
    {
        public EFU()
        {
            this.Children = new List<int>();
        }

        public int Id { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public string AcceptanceCriteria { get; set; }
        public string Workitemtype { get; set; }
        public List<int> Children { get; set; }
        public int? Parent { get; set; }
        public string Tags { get; set; }
    }

    public class WorkItems
    {
        [JsonProperty("count")]
        public int Count { get; set; }
        [JsonProperty("value")]
        public WorkItem[] Items { get; set; }
    }
}
