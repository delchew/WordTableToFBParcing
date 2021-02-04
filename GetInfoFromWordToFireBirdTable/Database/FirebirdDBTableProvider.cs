using GetInfoFromWordToFireBirdTable.Attributes;
using FirebirdSql.Data.FirebirdClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using GetInfoFromWordToFireBirdTable.Database.Extensions;
using System.Linq;
using Microsoft.Extensions.Configuration;

namespace GetInfoFromWordToFireBirdTable
{
    public class FirebirdDBTableProvider<T>
    {
        /// <summary>
        /// Table name in database
        /// </summary>
        private static readonly string _tableName;

        /// <summary>
        /// Key - FBTableNameAttribute, value - property name.
        /// </summary>
        private static readonly Dictionary<FBTableFieldAttribute, string> _tableFieldInfoDict;

        private static Type _tableEntityType;

        /// <summary>
        /// Key - Autoincremented table field name, value - generator name.
        /// </summary>
        private static readonly Dictionary<string, string> _tableGeneratorsNamesDict;

        private readonly string _dbConnectionString;
        private FbConnection _connection;
        private FbCommand _command;

        static FirebirdDBTableProvider()
        {
            _tableGeneratorsNamesDict = new Dictionary<string, string>();

            _tableEntityType = typeof(T);
            var attrs = _tableEntityType.GetCustomAttributes(typeof(FBTableNameAttribute), true);
            if (attrs.Length == 0)
                throw new Exception("Class is not marked as a \"Table\" by TableNameAttribute!");
            _tableName = ((FBTableNameAttribute)attrs[0]).TableName;

            _tableFieldInfoDict = new Dictionary<FBTableFieldAttribute, string>();
            var properties = _tableEntityType.GetProperties();

            string propertyName;
            FBTableFieldAttribute fieldAttribute;
            for (int i = 0; i < properties.Length; i++)
            {
                var propAttr = properties[i].GetCustomAttributes(typeof(FBTableFieldAttribute), true).First();
                if (propAttr != null)
                {
                    propertyName = properties[i].Name;
                    fieldAttribute = propAttr as FBTableFieldAttribute;
                    _tableFieldInfoDict.Add(fieldAttribute, propertyName);
                }
            }
        }

        public FirebirdDBTableProvider()
        {
            var builder = new ConfigurationBuilder();
            // установка пути к текущему каталогу
            builder.SetBasePath(Directory.GetCurrentDirectory());
            // получаем конфигурацию из файла appsettings.json
            builder.AddJsonFile("appsettings.json");
            // создаем конфигурацию
            var config = builder.Build();
            // возвращаем из метода строку подключения
            _dbConnectionString = config.GetConnectionString("DefaultConnection");
        }

        public void OpenConnection()
        {
            _connection = new FbConnection(_dbConnectionString);
            _connection.Open();
        }

        public void CloseConnection()
        {
            _connection.Close();
            _connection.Dispose();
        }

        public void CreateTableIfNotExists()
        {
            if (!TableExists(_tableName))
            {
                var createTableQueryBuilder = new StringBuilder($@"CREATE TABLE {_tableName} (");

                string fieldName, fieldDBType, notNull, primaryKey, comma = string.Empty;
                var autoincrementedTableFieldNames = new List<string>();
                foreach (var fieldAttr in _tableFieldInfoDict.Keys)
                {
                    fieldName = fieldAttr.TableFieldName;
                    fieldDBType = fieldAttr.TypeName;
                    notNull = fieldAttr.IsNotNull ? " NOT NULL" : string.Empty;
                    primaryKey = fieldAttr.IsPrymaryKey ? " PRIMARY KEY" : string.Empty;
                    createTableQueryBuilder.Append($"{comma}{fieldName} {fieldDBType}{notNull}{primaryKey}");
                    comma = ", ";
                    if (fieldAttr.Autoincrement)
                        autoincrementedTableFieldNames.Add(fieldName);
                }

                createTableQueryBuilder.Append(");");
                using (_command = new FbCommand(createTableQueryBuilder.ToString(), _connection))
                {
                    _command.ExecuteNonQuery();
                }

                foreach (var tableFieldName in autoincrementedTableFieldNames)
                    _tableGeneratorsNamesDict.Add(tableFieldName, CreateFieldAutoincrement(tableFieldName));
            }
        }

        public void AddItem(T item)
        {
            var sqlInsertRequestString = GetInsertRequestString(item);
            using (_command = new FbCommand(sqlInsertRequestString, _connection))
            {
                _command.ExecuteNonQuery();
            }
        }

        public object GetSingleObjBySQL(string sqlRequest)
        {
            object result;
            using (var command = new FbCommand(sqlRequest, _connection))
            {
                result = command.ExecuteScalar();
            }
            return result;
        }

        public void CreateBoolIntDBDomain(FileInfo databaseFile, string domainName)
        {
            var createBoolDomainSql = $@"CREATE DOMAIN {domainName} AS INTEGER DEFAULT 0 NOT NULL CHECK (VALUE IN(0,1));";
            var connection = new FbConnection(_dbConnectionString);
            connection.Open();
            using (var command = new FbCommand(createBoolDomainSql, connection))
            {
                command.ExecuteNonQuery();
            }
            connection.Close();
            connection.Dispose();
        }

        private string GetInsertRequestString(T item) //TODO Доделать!!
        {
            if (!TableExists(_tableName))
                throw new Exception("Such table name is not exists in current database!");

            foreach (var info in _tableFieldInfoDict)
            {
                var propV = _tableEntityType.GetProperty(info.Value).GetValue(item);
            }

            var sqlRequestStringBuilder = new StringBuilder($@"INSERT INTO {_tableName} (");
            var properties = _tableEntityType.GetProperties();
            var propValues = new List<string>();
            string fieldName, stringValue, comma = string.Empty;
            object propValue;
            for (int i = 0; i < properties.Length; i++)
            {
                var fieldPropAttrs = properties[i].GetCustomAttributes(typeof(FBTableFieldAttribute), true);
                if (fieldPropAttrs.Length == 0)
                    continue;
                var attr = (FBTableFieldAttribute)fieldPropAttrs[0];
                fieldName = attr.TableFieldName;
                propValue = properties[i].GetValue(item);
                stringValue = GetSqlTypeStringValue(propValue);
                propValues.Add(stringValue);
                sqlRequestStringBuilder.Append(comma + fieldName);
                comma = ", ";
            }
            sqlRequestStringBuilder.Append(") VALUES (");
            comma = string.Empty;
            for (int i = 0; i < propValues.Count; i++)
            {
                sqlRequestStringBuilder.Append(comma + propValues[i]);
                comma = ", ";
            }
            sqlRequestStringBuilder.Append(");");
            var sqlRequestString = sqlRequestStringBuilder.ToString();
            return sqlRequestString;
        }

        private bool TableExists(string tableName)
        {
            var sqlCheckTableExistString = $@"SELECT 1 FROM RDB$RELATIONS r WHERE r.RDB$RELATION_NAME = '{tableName}'";
            object result;
            using (_command = new FbCommand(sqlCheckTableExistString, _connection))
            {
                result = _command.ExecuteScalar();
            }
            if (result == null) return false;
            return true;
        }

        /// <summary>
        /// Set autoincrement to Table field
        /// </summary>
        /// <param name="tableFieldName">Table field name wich need an autoincrement</param>
        /// <returns>Database Generator name</returns>
        private string CreateFieldAutoincrement(string tableFieldName)
        {
            var genName = $@"{_tableName}_{tableFieldName}_GEN";
            using (var command = new FbCommand($@"CREATE GENERATOR {genName};", _connection))
            {
                command.ExecuteNonQuery();
                CreateFieldAutoincrement(tableFieldName, genName);
            }
            return genName;
        }

        private void CreateFieldAutoincrement(string tableFieldName, string existsGeneratorName)
        {
            var builder = new StringBuilder($@"CREATE TRIGGER TRG_{_tableName}_{tableFieldName} FOR {_tableName} ACTIVE BEFORE INSERT AS BEGIN ");
            builder.Append($@"IF (NEW.{tableFieldName} IS NULL) THEN NEW.{tableFieldName} = GEN_ID({existsGeneratorName}, 1); END;");
            using (var command = new FbCommand(builder.ToString(), _connection))
            {
                command.ExecuteNonQuery();
            }
        }

        private int GetNextValueAutoincrementTableField(string tableFieldName)
        {
            var generatorName = _tableGeneratorsNamesDict[tableFieldName];
            var sqlRequestString = $@"SELECT GEN_ID({generatorName}, 1) FROM RDB$DATABASE";
            using (var command = new FbCommand(sqlRequestString, _connection))
            {
                var result = command.ExecuteScalar();
                if (result != null)
                    return (int)result;
                throw new Exception("Не удалось получить следующий номер ID!");

            }
        }

        private string GetSqlTypeStringValue(object value)
        {
            var propType = value.GetType();
            string stringValue;
            if (propType.Name == typeof(bool).Name)
                stringValue = ((bool)value).ToFireBirdDBBoolInt().ToString();
            else
                if (propType.Name == typeof(string).Name)
                stringValue = $@"'{value}'";
            else
                if (propType.Name == typeof(double).Name)
                stringValue = ((double)value).ToFBSqlString();
            else
                stringValue = value.ToString();
            return stringValue;
        }
    }
}
