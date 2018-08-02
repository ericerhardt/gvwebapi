using System;
using System.Data;
using System.Data.SqlClient;

namespace GVWebapi.Helpers
{
    public static class SqlExtensions
    {
        public static SqlParameter ToParameter(this DateTimeOffset source, string name)
        {
            return CreateParameter(name, SqlDbType.DateTimeOffset, source);
        }

        public static SqlParameter ToParameter(this DateTimeOffset? source, string name)
        {
            if (source.HasValue == false)
                return CreateNullParameter(name, SqlDbType.DateTimeOffset);
            return CreateParameter(name, SqlDbType.DateTimeOffset, source);
        }

        public static SqlParameter ToParameter(this int source, string name)
        {
            return CreateParameter(name, SqlDbType.Int, source);
        }

        public static SqlParameter ToParameter(this int? source, string name)
        {
            if (source.HasValue == false)
                return CreateNullParameter(name, SqlDbType.Int);
            return CreateParameter(name, SqlDbType.Int, source);
        }

        public static SqlParameter ToParameter(this decimal source, string name)
        {
            return CreateParameter(name, SqlDbType.Decimal, source);
        }

        public static SqlParameter ToParameter(this bool source, string name)
        {
            return CreateParameter(name, SqlDbType.Bit, source);
        }

        public static SqlParameter ToParameter(this long source, string name)
        {
            return CreateParameter(name, SqlDbType.BigInt, source);
        }

        public static SqlParameter ToParameter(this long? source, string name)
        {
            if (source.HasValue == false)
                return CreateNullParameter(name, SqlDbType.BigInt);
            return CreateParameter(name, SqlDbType.BigInt, source);
        }

        public static SqlParameter ToParameter(this string source, string name)
        {
            return CreateParameter(name, SqlDbType.NVarChar, source);
        }

        public static long ToLong(this object source)
        {
            return Convert.ToInt64(source);
        }

        public static int ToInt(this object source)
        {
            return Convert.ToInt32(source);
        }

        public static decimal ToDecimal(this object source)
        {
            return Convert.ToDecimal(source);
        }

        public static DateTimeOffset ToDateTimeOffset(this object source)
        {
            DateTimeOffset outValue;
            var answer = DateTimeOffset.TryParse(source.ToString(), out outValue);
            return answer ? outValue : DateTimeOffset.Now;
        }

        private static SqlParameter CreateNullParameter(string name, SqlDbType type)
        {
            return new SqlParameter(name, type)
            {
                Value = DBNull.Value
            };
        }

        private static SqlParameter CreateParameter(string name, SqlDbType type, object value)
        {
            return new SqlParameter(name, type)
            {
                Value = value
            };
        }
    }
}