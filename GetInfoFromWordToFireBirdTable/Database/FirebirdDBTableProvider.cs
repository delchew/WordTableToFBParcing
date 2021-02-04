using GetInfoFromWordToFireBirdTable.Attributes;
using FirebirdSql.Data.FirebirdClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using GetInfoFromWordToFireBirdTable.Database.Extensions;
using System.Reflection;
using System.Linq;

namespace GetInfoFromWordToFireBirdTable
{
    public class FirebirdDBTableProvider<T>
    {
        private static Type _tableEntityType;
        private static readonly string _tableName;
        private static readonly Dictionary<FBTableFieldAttribute, string> _tableFieldInfoDict;

        private readonly string _connectionString;
        private FbConnection _connection;
        private FbCommand _command;

        static FirebirdDBTableProvider()
        {
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

        public FirebirdDBTableProvider(FileInfo databaseFile)
        {
            if (!databaseFile.Exists)
                throw new FileNotFoundException();

            _connectionString = GetConnectionString(databaseFile);
        }

        public void OpenConnection()
        {
            _connection = new FbConnection(_connectionString);
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
                var otherQueries = new List<(string tableFieldName, Action<string> queryAction)>();
                foreach (var fieldAttr in _tableFieldInfoDict.Keys)
                {
                    fieldName = fieldAttr.TableFieldName;
                    fieldDBType = fieldAttr.TypeName;
                    notNull = fieldAttr.IsNotNull ? " NOT NULL" : string.Empty;
                    primaryKey = fieldAttr.IsPrymaryKey ? " PRIMARY KEY" : string.Empty;
                    createTableQueryBuilder.Append($"{comma}{fieldName} {fieldDBType}{notNull}{primaryKey}");
                    comma = ", ";
                    if (fieldAttr.Autoincrement)
                    {
                        otherQueries.Add((fieldName, CreateFieldAutoincrement));
                    }
                }

                createTableQueryBuilder.Append(");");
                _command = new FbCommand(createTableQueryBuilder.ToString(), _connection);
                _command.ExecuteNonQuery();
                foreach(var pair in otherQueries)
                {
                    pair.queryAction.Invoke(pair.tableFieldName);
                }
            }
        }

        public void CreateFieldAutoincrement(string tableFieldName)
        {
            var genName = $@"{_tableName}_{tableFieldName}_GEN";
            var command = new FbCommand($@"CREATE GENERATOR {genName};", _connection);
            command.ExecuteNonQuery();
            CreateFieldAutoincrement(tableFieldName, genName);
        }

        public void CreateFieldAutoincrement(string tableFieldName, string existsGeneratorName)
        {
            var builder = new StringBuilder($@"CREATE TRIGGER TRG_{_tableName}_{tableFieldName} FOR {_tableName} ACTIVE BEFORE INSERT AS BEGIN ");
            builder.Append($@"IF (NEW.{tableFieldName} IS NULL) THEN NEW.{tableFieldName} = GEN_ID({existsGeneratorName}, 1); END;");
            var command = new FbCommand(builder.ToString(), _connection);
        }

        public void AddItem(T item)
        {
            var sqlInsertRequestString = GetInsertRequestString(item);
            _command = new FbCommand(sqlInsertRequestString, _connection);
            _command.ExecuteNonQuery();
        }

        public object GetSingleObjBySQL(string sqlRequest)
        {
            var command = new FbCommand(sqlRequest, _connection);
            var result = command.ExecuteScalar();
            return result;
        }

        public static void CreateBoolIntDBDomain(FileInfo databaseFile, string domainName)
        {
            var createBoolDomainSql = $@"CREATE DOMAIN {domainName} AS INTEGER DEFAULT 0 NOT NULL CHECK (VALUE IN(0,1));";
            var connection = new FbConnection(GetConnectionString(databaseFile));
            connection.Open();
            var command = new FbCommand(createBoolDomainSql, connection);
            command.ExecuteNonQuery();
            connection.Close();
            connection.Dispose();
        }

        private string GetInsertRequestString(T item)
        {
            if (!TableExists(_tableName))
                throw new Exception("Such table name is not exists in current database!");

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
            _command = new FbCommand(sqlCheckTableExistString, _connection);
            var result = _command.ExecuteScalar();
            if (result == null) return false;
            return true;
        }

        private static string GetConnectionString(FileInfo databaseFile)
        {
            var builder = new FbConnectionStringBuilder
            {
                Charset = "none",
                UserID = "SYSDBA",
                Password = "masterkey",
                Dialect = 3,
                DataSource = "localhost",
                Port = 3050,
                Database = databaseFile.FullName
            };
            return builder.ToString();
        }

        private int GetNewItemId()
        {
            var sqlRequestString = @"SELECT GEN_ID(BONDARENKO_ID_GEN, 1) FROM RDB$DATABASE";
            var command = new FbCommand(sqlRequestString, _connection);
            var reader = command.ExecuteReader();
            if (reader.HasRows)
            {
                reader.Read();
                var item = reader.GetValue(0);
                var newId = Convert.ToInt32(item);
                if (newId != 0)
                    return newId;
                throw new Exception("Не удалось получить следующий номер ID!");
            }
            throw new Exception("Не удалось получить следующий номер ID!");
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
