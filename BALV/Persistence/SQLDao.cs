using System;
using System.Linq;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
using Common;
using Common.DataShapes;
using Common.Utility;
using BAL.Persistence.DataMappers;
using BAL.Services;

namespace BAL.Persistence
{
    public class SqlDao
    {
        private static NLog.Logger logger;

        #region "Database Helper Methods"
        private SqlConnection _sharedConnection;
        public SqlConnection SharedConnection
        {
            get
            {
                if (_sharedConnection == null)
                {
                    _sharedConnection = MasterConnection.dbConnection;
                }
                return _sharedConnection;
            }
            set
            {
                _sharedConnection = value;
            }
        }


        // Constructors
        public SqlDao()
        {
            NLog.LogManager.ThrowExceptions = true;
            logger = Loggers.MyLogManager.Instance.GetCurrentClassLogger();
        }

        public SqlDao(SqlConnection connection)
        {
            NLog.LogManager.ThrowExceptions = true;
            logger = Loggers.MyLogManager.Instance.GetCurrentClassLogger();
            this.SharedConnection = connection;
        }


        // GetDbSqlCommand
        public SqlCommand GetSqlCommand(string sqlQuery)
        {
            SqlCommand command = new SqlCommand();
            command.Connection = SharedConnection;
            command.CommandType = CommandType.Text;
            command.CommandText = sqlQuery;
            return command;
        }


        // GetDbSprocCommand
        public SqlCommand GetSprocCommand(string sprocName)
        {
            SqlCommand command = new SqlCommand(sprocName);
            command.Connection = SharedConnection;
            command.CommandType = CommandType.StoredProcedure;
            return command;
        }


        // CreateNullParameter
        public SqlParameter CreateNullParameter(string name, SqlDbType paramType)
        {
            SqlParameter parameter = new SqlParameter();
            parameter.SqlDbType = paramType;
            parameter.ParameterName = name;
            parameter.Value = DBNull.Value;
            parameter.Direction = ParameterDirection.Input;
            return parameter;
        }


        // CreateNullParameter - with size for nvarchars
        public SqlParameter CreateNullParameter(string name, SqlDbType paramType, int size)
        {
            SqlParameter parameter = new SqlParameter();
            parameter.SqlDbType = paramType;
            parameter.ParameterName = name;
            parameter.Size = size;
            parameter.Value = null;
            parameter.Direction = ParameterDirection.Input;
            return parameter;
        }


        // CreateOutputParameter
        public SqlParameter CreateOutputParameter(string name, SqlDbType paramType)
        {
            SqlParameter parameter = new SqlParameter();
            parameter.SqlDbType = paramType;
            parameter.ParameterName = name;
            parameter.Direction = ParameterDirection.Output;
            return parameter;
        }


        // CreateOuputParameter - with size for nvarchars
        public SqlParameter CreateOutputParameter(string name, SqlDbType paramType, int size)
        {
            SqlParameter parameter = new SqlParameter();
            parameter.SqlDbType = paramType;
            parameter.Size = size;
            parameter.ParameterName = name;
            parameter.Direction = ParameterDirection.Output;
            return parameter;
        }


        // CreateParameter - uniqueidentifier
        public SqlParameter CreateParameter(string name, Guid value)
        {
            if (value.Equals(NullValues.NullGuid))
            {
                // If value is null then create a null parameter
                return CreateNullParameter(name, SqlDbType.UniqueIdentifier);
            }
            else
            {
                SqlParameter parameter = new SqlParameter();
                parameter.SqlDbType = SqlDbType.UniqueIdentifier;
                parameter.ParameterName = name;
                parameter.Value = value;
                parameter.Direction = ParameterDirection.Input;
                return parameter;
            }
        }


        // CreateParameter - int
        public SqlParameter CreateParameter(string name, int value)
        {
            if (value == NullValues.NullInt)
            {
                // If value is null then create a null parameter
                return CreateNullParameter(name, SqlDbType.Int);
            }
            else
            {
                SqlParameter parameter = new SqlParameter();
                parameter.SqlDbType = SqlDbType.Int;
                parameter.ParameterName = name;
                parameter.Value = value;
                parameter.Direction = ParameterDirection.Input;
                return parameter;
            }
        }
        // CreateParameter - boolean
        public SqlParameter CreateParameter(string name, bool value, bool defaultValue)
        {
            SqlParameter parameter = new SqlParameter();
            parameter.SqlDbType = SqlDbType.Bit;
            parameter.ParameterName = name;
            parameter.Value = value;
            parameter.Direction = ParameterDirection.Input;
            return parameter;
        }



        // CreateParameter - datetime
        public SqlParameter CreateParameter(string name, DateTime value)
        {
            if (value == NullValues.NullDateTime)
            {
                // If value is null then create a null parameter
                return CreateNullParameter(name, SqlDbType.DateTime);
            }
            else
            {
                SqlParameter parameter = new SqlParameter();
                parameter.SqlDbType = SqlDbType.DateTime;
                parameter.ParameterName = name;
                parameter.Value = value;
                parameter.Direction = ParameterDirection.Input;
                return parameter;
            }
        }


        // CreateParameter - nvarchar
        public SqlParameter CreateParameter(string name, string value, int size)
        {
            if (String.IsNullOrEmpty(value))
            {
                // If value is null then create a null parameter
                return CreateNullParameter(name, SqlDbType.NVarChar);
            }
            else
            {
                SqlParameter parameter = new SqlParameter();
                parameter.SqlDbType = SqlDbType.NVarChar;
                parameter.Size = size;
                parameter.ParameterName = name;
                parameter.Value = value;
                parameter.Direction = ParameterDirection.Input;
                return parameter;
            }
        }

        #endregion



        #region "Data Projection Methods"


        // ExecuteNonQuery
        public void ExecuteNonQuery(SqlCommand command)
        {
            try
            {
                if (command.Connection.State != ConnectionState.Open)
                {
                    command.Connection.Open();
                }
                command.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                logger.ErrorException("Got exception.", e);
                throw new Exception("Error executing query", e);
            }
            finally
            {
                command.Connection.Close();
            }
        }


        // ExecuteScalar
        public Object ExecuteScalar(SqlCommand command)
        {
            try
            {
                if (command.Connection.State != ConnectionState.Open)
                {
                    command.Connection.Open();
                }
                return command.ExecuteScalar();
            }
            catch (Exception e)
            {
                logger.ErrorException("Got exception.", e);
                throw new Exception("Error executing query", e);
            }
            finally
            {
                command.Connection.Close();
            }
        }


        // GetSingleValue
        public T GetSingleValue<T>(SqlCommand command)
        {
            T returnValue = default(T);
            try
            {
                if (command.Connection.State != ConnectionState.Open)
                {
                    command.Connection.Open();
                }
                SqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();
                    if (!reader.IsDBNull(0)) { returnValue = (T)reader[0]; }
                    reader.Close();
                }
            }
            catch (Exception e)
            {
                logger.ErrorException("Got exception.", e);
                throw new Exception("Error populating data", e);
            }
            finally
            {
                command.Connection.Close();
            }
            return returnValue;
        }


        // GetSingleString
        public Int32 GetSingleInt32(SqlCommand command)
        {
            Int32 returnValue = default(int);
            try
            {
                if (command.Connection.State != ConnectionState.Open)
                {
                    command.Connection.Open();
                }
                SqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();
                    if (!reader.IsDBNull(0)) { returnValue = reader.GetInt32(0); }
                    reader.Close();
                }
            }
            catch (Exception e)
            {
                logger.ErrorException("Got exception.", e);
                throw new Exception("Error populating data", e);
            }
            finally
            {
                command.Connection.Close();
            }
            return returnValue;
        }


        // GetSingleString
        public string GetSingleString(SqlCommand command)
        {
            string returnValue = null;
            try
            {
                if (command.Connection.State != ConnectionState.Open)
                {
                    command.Connection.Open();
                }
                SqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();
                    if (!reader.IsDBNull(0)) { returnValue = reader.GetString(0); }
                    reader.Close();
                }
            }
            catch (Exception e)
            {
                logger.ErrorException("Got exception.", e);
                throw new Exception("Error populating data", e);
            }
            finally
            {
                command.Connection.Close();
            }
            return returnValue;
        }


        // GetStringList
        public List<string> GetStringList(SqlCommand command)
        {
            List<string> returnList = null;
            try
            {
                if (command.Connection.State != ConnectionState.Open)
                {
                    command.Connection.Open();
                }
                SqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    returnList = new List<string>();
                    while (reader.Read())
                    {
                        if (!reader.IsDBNull(0)) { returnList.Add(reader.GetString(0)); }
                    }
                    reader.Close();
                }
            }
            catch (Exception e)
            {
                logger.ErrorException("Got exception.", e);
                throw new Exception("Error populating data", e);
            }
            finally
            {
                command.Connection.Close();
            }
            return returnList;
        }


        // GetSingle
        public T GetSingle<T>(SqlCommand command) where T : class
        {
            T dto = null;
            try
            {
                if (command.Connection.State != ConnectionState.Open)
                {
                    command.Connection.Open();
                }
                SqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    reader.Read();
                    IDataMapper mapper = new DataMapperFactory().GetMapper(typeof(T));
                    dto = (T)mapper.GetData(reader);
                    reader.Close();
                }
            }
            catch (Exception e)
            {
                logger.ErrorException("Got exception.", e);
                throw new Exception("Error populating data", e);
            }
            finally
            {
                command.Connection.Close();
            }
            // return the DTO, it's either populated with data or null.
            return dto;
        }


        // GetList
        public List<T> GetList<T>(SqlCommand command) where T : class
        {
            List<T> dtoList = new List<T>();
            try
            {
                if (command.Connection.State != ConnectionState.Open)
                {
                    command.Connection.Open();
                }
                SqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    IDataMapper mapper = new DataMapperFactory().GetMapper(typeof(T));
                    while (reader.Read())
                    {
                        T dto = null;
                        dto = (T)mapper.GetData(reader);
                        dtoList.Add(dto);
                    }
                    reader.Close();
                }
            }
            catch (Exception e)
            {
                logger.ErrorException("Got exception.", e);
                // throw new Exception("Error populating data", e);
            }
            finally
            {
                command.Connection.Close();
            }
            // We return either the populated list if there was data,
            // or if there was no data we return an empty list.
            return dtoList;
        }




        // GetDataPage
        public DataPage<T> GetDataPage<T>(SqlCommand command, int pageIndex, int pageSize) where T : class
        {
            DataPage<T> page = new DataPage<T>();
            page.PageIndex = pageIndex;
            page.PageSize = pageSize;
            try
            {
                if (command.Connection.State != ConnectionState.Open)
                {
                    command.Connection.Open();
                }
                SqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    IDataMapper mapper = new DataMapperFactory().GetMapper(typeof(T));
                    while (reader.Read())
                    {
                        // get the data for this row
                        T dto = null;
                        dto = (T)mapper.GetData(reader);
                        page.Data.Add(dto);
                        // If we haven't set the RecordCount yet then set it
                        if (page.RecordCount == 0) { page.RecordCount = mapper.GetRecordCount(reader); }
                    }
                    reader.Close();
                }
            }
            catch (Exception e)
            {
                logger.ErrorException("Got exception.", e);
                throw new Exception("Error populating data", e);
            }
            finally
            {
                command.Connection.Close();
            }
            return page;
        }


        #endregion





    }
}


