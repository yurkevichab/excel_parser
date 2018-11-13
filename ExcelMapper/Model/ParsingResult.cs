using System.Collections.Generic;

namespace ExcelMapper.Model
{
    public class ParsingResult<T>
    {
        public List<T> ListData { get; set; }
        public string SheetName { get; set; }
        public ValidationResult Validation { get; set; }
    }
}