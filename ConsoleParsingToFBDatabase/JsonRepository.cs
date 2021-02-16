using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace CablesCraftMobile
{
    public class JsonRepository
    {
        private static readonly string dataFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

        public bool SaveObject<T>(T obj, string fileName)
        {
            if(obj != null)
            {
                var jsonString = JsonConvert.SerializeObject(obj, Formatting.Indented);
                var filePath = Path.Combine(dataFolderPath, fileName);
                File.WriteAllText(filePath, jsonString, Encoding.UTF8);

                return true;
            }
            return false;
        }

        public T LoadObject<T>(string fileName, string specifiedJsonDataPath = "")
        {
            var loadedData = GetCurrentJsonDataToDeserialize(fileName, specifiedJsonDataPath);
            return JsonConvert.DeserializeObject<T>(loadedData);
        }

        public IList<T> GetObjects<T>(string fileName, string jsonDataPath = "")
        {
            var selectedJsonData = GetCurrentJsonDataToDeserialize(fileName, jsonDataPath);
            return JsonConvert.DeserializeObject<List<T>>(selectedJsonData);
        }

        public IList<T> GetObjects<T>(string fileName)
        {
            return GetObjects<T>(fileName, string.Empty);
        }

        public IDictionary<TKey, TValue> GetObjects<TKey, TValue>(string fileName, string specifiedJsonDataPath = "")
        {
            var selectedJsonData = GetCurrentJsonDataToDeserialize(fileName, specifiedJsonDataPath);
            return JsonConvert.DeserializeObject<Dictionary<TKey, TValue>>(selectedJsonData);
        }

        public IDictionary<TKey, TValue> GetObjects<TKey, TValue>(string fileName)
        {
            return GetObjects<TKey, TValue>(fileName, string.Empty);
        }

        private string GetCurrentJsonDataToDeserialize(string fileName, string specifiedJsonDataPath)
        {
            var filePath = Path.Combine(dataFolderPath, fileName);
            if (File.Exists(filePath))
            {
                var loadedData = File.ReadAllText(filePath, Encoding.UTF8);
                if (!string.IsNullOrEmpty(loadedData))
                {
                    if (!string.IsNullOrEmpty(specifiedJsonDataPath))
                    {
                        var jObject = JObject.Parse(loadedData);
                        var jsonSelectedData = jObject.SelectToken(specifiedJsonDataPath);
                        return jsonSelectedData.ToString();
                    }
                    return loadedData;
                }
                throw new FileLoadException($"Отсутствуют данные в файле {fileName}!");
            }
            throw new FileNotFoundException($"Файл {fileName} не найден! Проверьте правильность пути к файлу.");
        }
    }
}
