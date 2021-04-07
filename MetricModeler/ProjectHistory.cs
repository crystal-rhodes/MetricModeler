using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetricModeler
{
    class ProjectHistory
    {
        public string ProjectId { get; internal set; }
        public string ProjectName { get; internal set; }
        public string ProjectDescription { get; internal set; }
        public string ProjectType { get; internal set; }
        public DateTime StartDate { get; internal set; }
        public DateTime EndDate { get; internal set; }
        public int EstDuration { get; internal set; }
        public int EstProjectCost { get; internal set; }
        public int ActualProjectCost { get; internal set; }
        public int LOC { get; internal set; }
        public int EstimatedFP { get; internal set; }
        public int ActualFP { get; internal set; }
        public int AveCostPerPersonHour { get; internal set; }
        public int SoftwareDevelopmentCapability { get; internal set; }
        public int DesignReviewHours { get; internal set; }
        public string DevelopmentLanguage { get; internal set; }
        public int LanguageProductivityFactor { get; internal set; }
        public int RequiredFunctionalitiesExpectation { get; internal set; }
        public int NumTables { get; internal set; }


        public ProjectHistory(string projectId, string projectName, string projectDescription,
            string projectType, DateTime startDate, DateTime endDate, int estDuration, int estProjectCost,
            int actualProjectCost, int loc, int estimatedFP, int actualFP, int aveCostPerPersonHour, 
            int softwareDevelopmentCapability, int designReviewHours, string developmentLanguage, 
            int languageProductivityFactor, int requiredFunctionalitiesExpectation, int numTables)
        {
            ProjectId = projectId;
            ProjectName = projectName;
            ProjectDescription = projectDescription;
            ProjectType = projectType;
            StartDate = startDate;
            EndDate = endDate;
            EstDuration = estDuration;
            EstProjectCost = estProjectCost;
            ActualProjectCost = actualProjectCost;
            LOC = loc;
            EstimatedFP = estimatedFP;
            ActualFP = actualFP;
            AveCostPerPersonHour = aveCostPerPersonHour;
            SoftwareDevelopmentCapability = softwareDevelopmentCapability;
            DesignReviewHours = designReviewHours;
            DevelopmentLanguage = developmentLanguage;
            LanguageProductivityFactor = languageProductivityFactor;
            RequiredFunctionalitiesExpectation = requiredFunctionalitiesExpectation;
            NumTables = numTables;
        }

        public override string ToString()
        {
            return String.Format("Project: {0}\tActual Function Point: {1}", ProjectId, ActualFP);
        }
    }
}
