using System;
using System.Collections.Generic;

namespace ExcelAiAssistant.Excel
{
    public static class ExcelToolHostHelpers
    {
        public static Dictionary<string, object> GetOperation(object[,] operations, int index)
        {
            return operations[index, 0] as Dictionary<string, object> ??
                   throw new ArgumentException("operations 必须是对象数组。");
        }

        public static int GetInt(Dictionary<string, object> args, string key)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                throw new ArgumentException($"缺少必填参数：{key}");
            }

            return value is Newtonsoft.Json.Linq.JValue jValue ? Convert.ToInt32(jValue.Value) : Convert.ToInt32(value);
        }

        public static int GetIntSafe(Dictionary<string, object> args, string key, int defaultValue)
        {
            if (!args.TryGetValue(key, out var value) || value == null)
            {
                return defaultValue;
            }

            return Convert.ToInt32(value is Newtonsoft.Json.Linq.JValue jValue ? jValue.Value : value);
        }

public static object GetOperationId(Core.Models.ToolResult result)
        {
            var data = result?.Data;
            if (data == null)
            {
                return null;
            }

            if (data is Newtonsoft.Json.Linq.JObject token)
            {
                return token["operation_id"]?.ToObject<object>();
            }

            // WriteRangeCore 等返回匿名对象——`as JObject` 必为 null，必须走反射读属性，
            // 否则 batch/clean/filter 的 operation_id 恒空（快照已写却报不出 id，用户无法回滚）。
            var property = data.GetType().GetProperty("operation_id");
            return property?.GetValue(data);
        }

        public static object[,] FlattenValues(dynamic source, int rows, int columns)
        {
            var result = new object[rows, columns];
            if (source == null || rows == 0 || columns == 0)
            {
                return result;
            }

            if (source is object[,] array)
            {
                for (var row = 0; row < Math.Min(rows, array.GetLength(0)); row++)
                {
                    for (var column = 0; column < Math.Min(columns, array.GetLength(1)); column++)
                    {
                        result[row, column] = array[row + array.GetLowerBound(0), column + array.GetLowerBound(1)];
                    }
                }

                return result;
            }

            if (rows == 1 && columns == 1)
            {
                result[0, 0] = source;
            }

            return result;
        }
    }
}
