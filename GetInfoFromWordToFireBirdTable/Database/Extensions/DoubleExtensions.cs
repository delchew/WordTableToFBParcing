namespace GetInfoFromWordToFireBirdTable.Database.Extensions
{
    public static class DoubleExtensions
    {
        public static string ToFBSqlString(this double number)
        {
            var str = number.ToString();
            return str.Replace(',', '.');
        }
    }
}
