using System.Text;

namespace GetInfoFromWordToFireBirdTable.Database.Extensions
{
    public static class StringExtensions
    {
        public static bool FBIdentifierLengthIsTooLong(this string identifier)
        {
            return Encoding.UTF8.GetByteCount(identifier) > 31;
        }
    }
}
