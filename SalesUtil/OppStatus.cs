using System;
using System.Collections.Generic;
using System.Linq;
//using System.Text;

using Google.Apis.Services;
using Google.Apis.Translate.v2;
//using Google.Apis.Translate.v2.Data;


namespace SalesUtil
{
    internal enum ProperyTypes { CaseName, Description }

    public class OppStatus
    {

        public int Revenue;

        public string CaseId;
        public string CaseOwner;
        public string CaseName;
        public string Customer;
        public DateTime EstCloseDate;
        public string Description;
        public string Currency;
        public string CaseType;
        public int Probability;
        public string ProcessType;
        public string ProcessStage;

        private static TranslateService service;

        public bool HasErrors()
        {
            foreach (var product in Products)
            {
                if (product.HasErrors())
                    return true;
            }
            return
                IsDescriptionFieldEmptyError ||
                IsProcessTypeSelectedError ||
                IsProbabilityValidError ||
                IsNOKWarning ||
                IsEstimateCloseDataValidError ||
                IsProcessStageIdentified ||
                IsProcessTypeCallOff ||
                IsOpportunityTypeFieldNotEmptyError ||
                IsBadDU ||
                IsBadProduct ||
                IsNotCaseNameEnglish ||
                IsNotDescriptionEnglish;
        }

        public List<OppProductStatus> Products;

        public List<BadCase> BadProducts;
        public List<BadCase> BadDUs;

        private OppStatus(List<OppProductStatus> products)
        {
            Products = products;
            Revenue = products.Sum(x => x.Revenue);
            var commonData = Products.FirstOrDefault();
            CaseId = commonData.CaseID;
            CaseOwner = commonData.CaseOwner;

            CaseName = commonData.CaseName;
            Customer = commonData.Customer;
            EstCloseDate = commonData.EstCloseDate;
            Description = commonData.Description;
            Currency = commonData.Currency;
            CaseType = commonData.CaseType;
            Probability = commonData.Probability;
            ProcessType = commonData.ProcessType;
            ProcessStage = commonData.ProcessStage;
        }

        public static List<OppStatus> Factory(List<OppProductStatus> input)
        {
            List<OppStatus> res = new List<OppStatus>();


            var result = from p in input
                         group p by p.CaseID into g
                         select g;
            foreach (var group in result)
            {
                res.Add(new OppStatus(group.ToList()));
            }


            ///Traffic optimization
            LanguageDetection(res.ToList());
            /// End of code smell.

            return res;
        }

        public bool IsDescriptionFieldEmptyError
        {
            get
            {
                return string.IsNullOrEmpty(Description);
            }
        }

        public bool IsEstimateCloseDataValidError
        {
            get
            {
                return DateTime.Compare(EstCloseDate, DateTime.Today) <= 0;
            }
        }

        public bool IsProcessTypeSelectedError
        {
            get
            {
                return ProcessType.Equals("*Must Select Type*");
            }
        }

        public bool IsProcessTypeCallOff
        {
            get
            {
                if (Revenue >= 5000000)
                    return ProcessType.Equals("Call Off");
                return false;
            }
        }

        public bool IsProbabilityValidError
        {
            get
            {
                bool isProbabilityValid = Probability == 1;
                string stage = ProcessStage;
                bool isStageValid = !string.IsNullOrEmpty(stage) || !stage.Equals("Identified(to be validated)");

                return isProbabilityValid && isStageValid;
            }
        }

        public bool IsOpportunityTypeFieldNotEmptyError
        {
            get
            {
                return !string.IsNullOrEmpty(CaseType);
            }
        }

        public bool IsNOKWarning
        {
            get
            {
                return !string.IsNullOrEmpty(Currency) && Currency.Equals("Norsk krone");
            }
        }

        public bool IsProcessStageIdentified
        {
            get
            {
                return ProcessStage.Equals("Identified");
            }
        }

        public bool IsBadProduct { get; set; }

        public bool IsBadDU { get; set; }

        //TODO: refactor, revenue factor is no longer assessed!
        public bool IsNotEnglish()
        {
            return true;
        }

        public bool IsNotCaseNameEnglish { get; private set; }        

        public bool IsNotDescriptionEnglish { get; private set; }

        private static void LanguageDetection(List<OppStatus> cases)
        {           
            var detectionList = new List<Tuple<string, ProperyTypes, string>>();
            var detectedLanguagesList = new List<Tuple<string, float?>>();

            foreach (var _case in cases)
            {
                var element1 = new Tuple<string, ProperyTypes, string>(_case.CaseId, ProperyTypes.CaseName, _case.CaseName);
                var element2 = new Tuple<string, ProperyTypes, string>(_case.CaseId, ProperyTypes.Description, _case.Description);
                detectionList.Add(element1);
                detectionList.Add(element2);
            }

            service = new TranslateService(new BaseClientService.Initializer()
            {
                ApiKey = Program.GoogleKey,
                ApplicationName = "SalesUtil v.1"
            });
            var respond = service.Detections.List(detectionList.Select(x => x.Item3).ToList()).Execute();

            foreach (var detection in respond.Detections)
            {
                detectedLanguagesList.Add(Tuple.Create(
                    detection[0].Language,
                    detection[0].Confidence));
            }

            var resultSequence = detectionList.Zip(detectedLanguagesList, (x, y) => new
            {
                caseId = x.Item1,
                PropertyType = x.Item2,
                propertyName = x.Item3,
                Language = y.Item1,
                Confidence = y.Item2
            });

            //tobe refactored
            //var query = (from p in resultSequence
            //             group p by p.caseId into g
            //             let isNotEnglish = g.Any(x => x.Language != "en")
            //             select new
            //             {
            //                 caseId = g.Key,
            //                 isEnglish = !isNotEnglish
            //             }).ToList();

            //foreach (var _opp in cases)
            //{
            //    var translation = resultSequence.Select(x => x.caseId == _opp.CaseId);
            //    var CaseNameLanguage = 
            //}

            //foreach (var _opp in cases)
            //{
            //    _opp.IsEnglish = query.Find(x => x.caseId == _opp.CaseId).isEnglish;
            //}

        }

    }
}
