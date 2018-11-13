using System;

namespace ExcelMapper.Attribute
{ 
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
    public class ExcelColumnMapAttribute : System.Attribute
    {
        /// <summary>
        /// Название колонки в файле excel для мапинга 
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        /// Указывает, является ли поле обязательным к заполнению(не может быть пустой строкой) 
        /// </summary>
        /// <value>false</value>
        public bool IsRequired { get; set; } = false;

        /// <summary>
        /// Регекс выражение для проверки значения ячейки (для строго соотвестствия необходимо в уровнении задавать символы 
        /// начала ^ и конца $
        /// </summary>
        /// <remarks>^^[\\a-zA-Zа-яА-Я\d\/\|\.,\*\^~_&amp;\$#@!\?\(\)\{\}\[\]\+\-№=&lt;&gt;:]*$</remarks>
        public string AllowCharsRegex { get; set; } = "";

        /// <summary>
        /// Задает значение ячейки по умолчанию, в случае если оно пустое
        /// </summary>
        /// <value>string.Empty</value>
        public string DefaultStringValue { get; set; } = string.Empty;
    }
}
