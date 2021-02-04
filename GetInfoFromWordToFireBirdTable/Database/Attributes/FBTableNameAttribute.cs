using System;
namespace GetInfoFromWordToFireBirdTable.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public class FBTableNameAttribute : Attribute
    {
        public string TableName { get; set; }

        public FBTableNameAttribute(string tableName)
        {
            TableName = tableName;
        }
    }
}
