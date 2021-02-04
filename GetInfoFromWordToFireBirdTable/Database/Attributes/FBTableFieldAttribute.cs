using System;
namespace GetInfoFromWordToFireBirdTable.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class FBTableFieldAttribute : Attribute
    {
        public string TableFieldName { get; set; }

        public string TypeName { get; set; }

        public bool IsPrymaryKey { get; set; }

        public bool IsNotNull { get; set; }
    }
}
