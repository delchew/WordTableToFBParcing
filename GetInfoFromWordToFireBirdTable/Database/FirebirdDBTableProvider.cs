using GetInfoFromWordToFireBirdTable.Attributes;
using FirebirdSql.Data.FirebirdClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using GetInfoFromWordToFireBirdTable.Database.Extensions;

namespace GetInfoFromWordToFireBirdTable
{
    public class FirebirdDBTableProvider<T>
    {
        private static Type _tableEntityType;

        private readonly string _connectionString;
        private FbConnection _connection;
        private FbCommand _command;
        private readonly static string _tableName;

        static FirebirdDBTableProvider()
        {
            _tableEntityType = typeof(T);
            var attrs = _tableEntityType.GetCustomAttributes(typeof(FBTableNameAttribute), true);
            if (attrs.Length == 0)
                throw new Exception("Class is not marked as a \"Table\" by TableNameAttribute!");
            _tableName = ((FBTableNameAttribute)attrs[0]).TableName;
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

        public void CreateTableIfNotExists() //TODO доделать
        {
            if (!TableExists(_tableName))
            {
                var sqlRequestStringBuilder = new StringBuilder($@"CREATE TABLE {_tableName} (");
                var properties = _tableEntityType.GetProperties();
                string fieldName, fieldDBType, notNull, primaryKey, comma = string.Empty;
                for (int i = 0; i < properties.Length; i++)
                {
                    var propAttrs = properties[i].GetCustomAttributes(typeof(FBTableFieldAttribute), true);
                    if (propAttrs.Length > 0)
                    {
                        var fieldAttr = (FBTableFieldAttribute)propAttrs[0];
                        fieldName = fieldAttr.TableFieldName;
                        fieldDBType = fieldAttr.TypeName;
                        notNull = fieldAttr.IsNotNull ? " NOT NULL" : string.Empty;
                        primaryKey = fieldAttr.IsPrymaryKey ? " PRIMARY KEY" : string.Empty;
                        sqlRequestStringBuilder.Append($"{comma}{fieldName} {fieldDBType}{notNull}{primaryKey}");
                        comma = ", ";
                    }
                }
                sqlRequestStringBuilder.Append(");");
                _command = new FbCommand(sqlRequestStringBuilder.ToString(), _connection);
                _command.ExecuteNonQuery();
            }
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
