using System;

namespace GetInfoFromWordToFireBirdTable.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class FBFieldAutoincrementAttribute : Attribute
    {
        /// <summary>
        /// Имя генератора автоинкремента
        /// </summary>
        public string GeneratorName { get; set; }
    }
}
