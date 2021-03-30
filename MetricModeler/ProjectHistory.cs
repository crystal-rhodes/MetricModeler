﻿using System;
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
        public int EstEffort { get; internal set; }
        public int ActualEffort { get; internal set; }
        public int EstLOC { get; internal set; }
        public int ActualLOC { get; internal set; }
        public int EstimatedFP { get; internal set; }
        public int ActualFP { get; internal set; }
        public int ExpectedErrorRate { get; internal set; }
        public int AveCostPerPersonHour { get; internal set; }
        public int AverageStaffingLevel { get; internal set; }
        public int DesignReviewHours { get; internal set; }
        public int ErrorsFound { get; internal set; }
        public int DefectsReported { get; internal set; }
        public string DevelopmentLanguage { get; internal set; }
        public int LanguageProductivityFactor { get; internal set; }
        public int CPMTasksDefined { get; internal set; }
        public int ChangeOrdersIssued { get; internal set; }
        public int DocumentationPages { get; internal set; }

        public ProjectHistory(string projectId, string projectName, string projectDescription, 
            string projectType, DateTime startDate, DateTime endDate, int estDuration, int estProjectCost, 
            int actualProjectCost, int estEffort, int actualEffort, int estLOC, int actualLOC, int estimatedFP, 
            int actualFP, int expectedErrorRate, int aveCostPerPersonHour, int averageStaffingLevel, int designReviewHours, 
            int errorsFound, int defectsReported, string developmentLanguage, int languageProductivityFactor, int cPMTasksDefined, 
            int changeOrdersIssued, int documentationPages)
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
            EstEffort = estEffort;
            ActualEffort = actualEffort;
            EstLOC = estLOC;
            ActualLOC = actualLOC;
            EstimatedFP = estimatedFP;
            ActualFP = actualFP;
            ExpectedErrorRate = expectedErrorRate;
            AveCostPerPersonHour = aveCostPerPersonHour;
            AverageStaffingLevel = averageStaffingLevel;
            DesignReviewHours = designReviewHours;
            ErrorsFound = errorsFound;
            DefectsReported = defectsReported;
            DevelopmentLanguage = developmentLanguage;
            LanguageProductivityFactor = languageProductivityFactor;
            CPMTasksDefined = cPMTasksDefined;
            ChangeOrdersIssued = changeOrdersIssued;
            DocumentationPages = documentationPages;
        }

        public override string ToString()
        {
            return String.Format("Project: {0}\tActual Function Point: {1}", ProjectId, ActualFP);
        }
    }
}
