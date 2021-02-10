using System;

namespace CableDataParsing
{
    public interface ICableDataParcer
    {
        /// <summary>
        /// Метод парсит данные из таблицы MSWord в базу данных Firebird
        /// </summary>
        /// <returns>Возвращает число добавленных записей</returns>
        int ParseDataToDatabase();

        event Action<int, int> ParseReport;
    }
}
