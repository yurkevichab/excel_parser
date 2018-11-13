using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using ExcelMapper.Attribute;
using ExcelMapper.Model;
using NPOI.SS.UserModel;

namespace ExcelMapper
{
    public static class Helper
    {
        public static ParsingResult<T> ParseExcelData<T>(Stream excelFileStream) where T : new()
        {
            const int correctionIndexForExcel = 1;
            var result = new ParsingResult<T> { Validation = new ValidationResult() };

            var properties =
                typeof(T).GetProperties()
                    .Where(prop => System.Attribute.IsDefined(prop, typeof(ExcelColumnMapAttribute)))
                    .ToDictionary(o => o.GetCustomAttribute<ExcelColumnMapAttribute>().ColumnName, p => p);

            var dataList = new List<T>();

            var columnNameDict = new Dictionary<string, int>();

            var wb = WorkbookFactory.Create(excelFileStream);
            var sheet = wb.GetSheetAt(0);
            result.SheetName = sheet.SheetName;
            var firstrowCells = sheet.GetRow(0).Cells;

            foreach (var property in properties)
            {
                var cell =
                    firstrowCells.FirstOrDefault(
                        x => string.Equals(x.ToString(), property.Key, StringComparison.CurrentCultureIgnoreCase));
                if (cell == null)
                {
                    result.Validation.Errors.Add($"В файле отсутствует столбец \"{property.Key}\"");
                    return result;
                }
                columnNameDict.Add(property.Key, cell.ColumnIndex);
            }
            var errorTypeList = columnNameDict.ToDictionary(x => x.Key, y => new List<int>());
            var charErrorList = columnNameDict.ToDictionary(x => x.Key, y => new List<int>());
            var emptyCellErrorDic = columnNameDict.ToDictionary(x => x.Key, y => new List<int>());

            for (var rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                var correctRowIndexInExcel = rowIndex + correctionIndexForExcel;

                if (row == null || row.Cells.All(d => d.CellType == CellType.Blank))
                {
                    continue;
                }
                var data = new T();

                foreach (var propertyinfo in properties)
                {
                    var property = propertyinfo.Value;
                    var propertyType = property.PropertyType;
                    var excelAttrProperty = property.GetCustomAttribute<ExcelColumnMapAttribute>();
                    var isRequiredCell = excelAttrProperty.IsRequired;
                    var allowCharsRegex = excelAttrProperty.AllowCharsRegex;
                    var defaultValue = excelAttrProperty.DefaultStringValue;

                    var cell = row.GetCell(columnNameDict[propertyinfo.Key]);

                    var cellType = cell?.CellType ?? CellType.String;

                    var cellvalue = GetValueFromCell(cell, cellType, propertyType, defaultValue);

                    CheckIsEmptyCell(isRequiredCell, cellvalue, emptyCellErrorDic, propertyinfo.Key,
                        correctRowIndexInExcel);

                    CheckCharsInCell(allowCharsRegex, cellvalue, charErrorList, propertyinfo.Key, correctRowIndexInExcel);

                    if (propertyType.IsGenericType &&
                        propertyType.GetGenericTypeDefinition() == typeof(Nullable<>) &&
                        string.IsNullOrWhiteSpace(cellvalue))
                    {
                        property.SetValue(data, null);
                    }
                    else
                    {
                        try
                        {
                            propertyType = Nullable.GetUnderlyingType(propertyType) ?? propertyType;
                            var exceldata = Convert.ChangeType(cellvalue,
                                propertyType, CultureInfo.InvariantCulture);
                            property.SetValue(data, exceldata);
                        }
                        catch (Exception)
                        {
                            errorTypeList[propertyinfo.Key].Add(rowIndex + correctionIndexForExcel);
                        }
                    }
                }
                dataList.Add(data);
            }

            if (!dataList.Any())
            {
                result.Validation.Errors.Add("Документ пуст. Заполните документ.");
                return result;
            }
            result.Validation.Errors.AddRange(GetErrorMessages(errorTypeList, "содержится недопустимое значение"));
            result.Validation.Errors.AddRange(GetErrorMessages(charErrorList, "содержится недопустимый символ"));
            result.Validation.Errors.AddRange(GetErrorMessages(emptyCellErrorDic, "недопустимое значение"));
            result.ListData = dataList;
            return result;
        }

        private static string GetValueFromCell(ICell cell, CellType cellType, Type propertyType, string defaultValue)
        {

            var isNumericProperty = IsNumericType(propertyType);
            var result = defaultValue;

            if (cell == null || cell.CellType == CellType.Blank || cell.CellType == CellType.Error)
            {
                return result;
            }
            if (cellType == CellType.String)
            {
                var currentCellValue = cell.StringCellValue.ToString(CultureInfo.InvariantCulture);
                result = isNumericProperty ? currentCellValue.Replace(',', '.') : currentCellValue;
            }
            else
            {
                result = isNumericProperty ? cell.NumericCellValue.ToString(CultureInfo.InvariantCulture) : cell.ToString();
            }
            return result.Trim();
        }

        private static List<string> GetErrorMessages(Dictionary<string, List<int>> errorList,
            string message)
        {
            var errors = new List<string>();
            foreach (var error in errorList)
            {
                if (error.Value.Count > 0)
                {
                    errors.Add(
                        $"В столбце \"{error.Key}\" в строке(строках): {string.Join(", ", error.Value.Take(15))} {message}.<br>");
                }
            }
            return errors;
        }

        private static void CheckIsEmptyCell(bool isRequiredCell, string cellvalue,
            Dictionary<string, List<int>> emptyCellErrorDic,
            string columnName, int rowIndex)
        {
            if (isRequiredCell && string.IsNullOrEmpty(cellvalue))
            {
                emptyCellErrorDic[columnName].Add(rowIndex);
            }
        }

        private static void CheckCharsInCell(string allowCharsRegex, string cellvalue,
            Dictionary<string, List<int>> charErrorList,
            string columnName, int rowIndex)
        {
            var regex = new Regex(allowCharsRegex);
            if (allowCharsRegex != string.Empty && cellvalue != "" && !regex.IsMatch(cellvalue))
            {
                charErrorList[columnName].Add(rowIndex);
            }
        }

        private static bool IsNumericType(Type type)
        {
            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Single:
                    return true;
                case TypeCode.Object:
                    if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
                    {
                        return IsNumericType(Nullable.GetUnderlyingType(type));
                    }
                    return false;
                default:
                    return false;
            }
        }
    }
}