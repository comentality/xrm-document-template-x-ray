namespace DocumentTemplateXRay.Logic
{
    public class FieldInfo
    {
        public string FieldPath { get; set; }
        public string Tag { get; set; }
        public string Alias { get; set; }
        public string XPath { get; set; }
        public string StoreId { get; set; }
        public string Location { get; set; }
        public bool IsRepeatingSection { get; set; }
        public string RepeatingSectionName { get; set; }
        public string TableDisplayName { get; set; }
        public string ColumnDisplayName { get; set; }
    }
}
