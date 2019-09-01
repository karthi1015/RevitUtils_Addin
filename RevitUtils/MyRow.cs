namespace RevitUtils
{
    public enum State
    {
        UsedFilter,

        UnUsedFilter
    }

    /// <summary>
    /// My customized model
    /// </summary>
    public class MyRow
    {
        public string Col1 { get; set; }

        public string Col2 { get; set; }

        public State State { get; set; }
    }
}