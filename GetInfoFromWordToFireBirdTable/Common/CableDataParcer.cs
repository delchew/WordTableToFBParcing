using System;

namespace GetInfoFromWordToFireBirdTable.Common
{
    public abstract class CableDataParcer
    {
        /// <summary>
        /// Метод парсит данные из таблицы MSWord в базу данных Firebird
        /// </summary>
        /// <returns>Возвращает число добавленных записей</returns>
        public abstract int ParseDataToDatabase();

        public abstract event Action<int, int> ParseReport;
    }
}
