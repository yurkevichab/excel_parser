using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelMapper.Model
{
    public class ValidationResult
    {
        private readonly List<string> errors = new List<string>();

        public virtual bool IsValid => this.Errors.Count == 0;

        public List<string> Errors => this.errors;

        public string GetErrorMessage()
        {
            return string.Join("\n", Errors.Select<string, string>((Func<string, string>)(e => e)));
        }

        public ValidationResult()
        {
        }

        public ValidationResult(IEnumerable<string> failures)
        {
            this.errors.AddRange(failures.Where<string>((Func<string, bool>)(failure => failure != null)));
        }
    }
}
