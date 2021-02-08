namespace GetInfoFromWordToFireBirdTable.Database.Extensions
{
    public static class DecimalExtensions
    {
        public static string ToFBSqlString(this decimal number)
        {
            var str = number.ToString();
            return str.Replace(',', '.');
        }
    }
}
