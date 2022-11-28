using ReactiveUI;


namespace Stage2WPF.MVVM.Model
{
    public class ExcelModel : ReactiveObject
    {   
        public string OrderDate { get; set; }
        
        public string Region { get; set; }
        
        public string Rep { get; set; }
        
        public string Item { get; set; }
        
        public int Units { get; set; }
        
        public double UnitCost { get; set; }
        
        public double Total { get; set; }
    }
}
