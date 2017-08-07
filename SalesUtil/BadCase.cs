using System;

namespace SalesUtil
{
    public class BadCase : ExcelDataObject
    {
        public string CaseID;

        public override void Init(Func<string, string> Getter)
        {
            GetValue = Getter;

            string columnFilter3 = "Opportunity Number";
            CaseID = GetValue(columnFilter3);
        }    
    }
}
