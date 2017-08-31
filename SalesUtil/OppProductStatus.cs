using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace SalesUtil
{
    public class OppProductStatus : ExcelDataObject
    {
        public string CaseOwner;
        public string CaseID;
        public int Revenue;

        public string CaseName;
        public string Description;
        public string CaseId;


        public string Customer;
        public DateTime EstCloseDate;

        public string Currency;
        public string CaseType;
        public int Probability;
        public string ProcessType;
        public string ProcessStage;
        public string Product;
        public int Margin;
        public int NumberOfPeriods;
        public DateTime? RevenueStartDate;
        public string DeliveryUnit;
        public string RevenueType;

        public override void Init(Func<string, string> Getter)
        {
            GetValue = Getter;

            string columnFilter1 = "Revenue";
            string value = GetValue(columnFilter1);
            Revenue = !string.IsNullOrWhiteSpace(value) ? int.Parse(value) : 0;

            string columnFilter2 = "Opportunity Owner";
            CaseOwner = GetValue(columnFilter2);

            string columnFilter3 = "Opportunity Number";
            CaseID = GetValue(columnFilter3);

            string columnFilter4 = "Opportunity Name";
            CaseName = GetValue(columnFilter4);

            string columnFilter5 = "Description";
            Description = GetValue(columnFilter5);

            string columnFilter6 = "Customer";
            Customer = GetValue(columnFilter6);

            string columnFilter7 = "Est# Close Date";
            value = GetValue(columnFilter7);
            if (!string.IsNullOrWhiteSpace(value))
                EstCloseDate = DateTime.Parse(value);

            string columnFilter8 = "Currency";
            Currency = GetValue(columnFilter8);

            string columnFilter9 = "Opportunity Type";
            CaseType = GetValue(columnFilter9);

            string columnFilter10 = "Probability";
            value = GetValue(columnFilter10);
            Probability = !string.IsNullOrWhiteSpace(value) ? int.Parse(value) : 0;

            string columnFilter11 = "Process Type";
            ProcessType = GetValue(columnFilter11);

            string columnFilter12 = "Process  Stage";
            ProcessStage = GetValue(columnFilter12);

            string columnFilter13 = "Product/Service";
            Product = GetValue(columnFilter13);

            string columnFilter14 = "Margin";
            value = GetValue(columnFilter14);
            Margin = !string.IsNullOrWhiteSpace(value) ? int.Parse(value) : 0;

            string columnFilter15 = "Number of Periods";
            value = GetValue(columnFilter15);
            NumberOfPeriods = !string.IsNullOrWhiteSpace(value) ? int.Parse(value) : 0;

            string columnFilter16 = "Revenue Start Date";
            value = GetValue(columnFilter16);
            if (!string.IsNullOrWhiteSpace(value))
                RevenueStartDate = DateTime.Parse(value);

            string columnFilter17 = "Delivery Unit";
            DeliveryUnit = GetValue(columnFilter17);

            string columnFilter18 = "Revenue Type";
            RevenueType = GetValue(columnFilter18);

        }

        public bool IsRevanueStartDateValidError
        {
            get
            {
                if (RevenueStartDate.HasValue)
                    return DateTime.Compare(EstCloseDate, RevenueStartDate.Value) >= 0;
                return false;
            }
        }

        public bool IsRevanueStartDateOverdueError
        {
            get
            {
                if (RevenueStartDate.HasValue)
                    return DateTime.Compare(RevenueStartDate.Value, DateTime.Today) >= 0;
                return false;
            }
        }

        public bool IsRevenueStartEmptyError
        {
            get
            {
                return RevenueStartDate == null;
            }
        }

        public bool IsMarginValidError
        {
            get
            {
                return Revenue <= Margin;
            }
        }

        public bool IsProductFieldEmptyError
        {
            get
            {
                return string.IsNullOrEmpty(Product);
            }
        }

        public bool IsRevenueFieldEmptyError
        {
            get
            {
                return Revenue <= 0;
            }
        }

        public bool IsRevenueTypeFieldEmptyError
        {
            get
            {
                return string.IsNullOrEmpty(RevenueType);
            }
        }

        public bool IsMarginFieldEmptyError
        {
            get
            {
                return Margin <= 0;
            }
        }

        public bool IsNumberOfPeriodsFieldEmptyError
        {
            get
            {
                return NumberOfPeriods <= 0;
            }
        }

        public bool IsNumberOfPeriodsFieldNotValid
        {
            get
            {
                if (NumberOfPeriods == 1 && Revenue >= 1000000)
                    return true;
                return false;
            }
        }

        public bool IsNOKWarning
        {
            get
            {
                if (Program.Localization == Areas.NO)
                    return false;
                return Currency.Equals("Norsk krone");
            }
        }

        public bool IsDeliveryUnitFieldNotEmptyError
        {
            get
            {
                return string.IsNullOrEmpty(DeliveryUnit);
            }
        }

        public bool IsProductDiscontinuedError
        {
            get
            {
                return Product.Contains("Discontinued");
            }
        }

        public bool IsDeliverUnitTerminatedError
        {
            get
            {
                return DeliveryUnit.Contains("Terminated");
            }
        }

        public bool HasErrors()
        {
            return
                IsRevanueStartDateValidError ||
                IsMarginValidError ||
                IsProductFieldEmptyError ||
                IsRevenueFieldEmptyError ||
                IsMarginFieldEmptyError ||
                IsRevenueTypeFieldEmptyError ||
                IsNumberOfPeriodsFieldEmptyError ||
                IsRevanueStartDateOverdueError ||
                IsDeliveryUnitFieldNotEmptyError ||
                IsRevenueStartEmptyError ||
                IsProductDiscontinuedError ||
                IsDeliverUnitTerminatedError ||
                IsNOKWarning ||
                IsNumberOfPeriodsFieldNotValid;
        }
    }
}
