using System;
using System.Globalization;

namespace Excel_Iterate
{
    // Immutable request: parameter validation happens before any host write.
    internal sealed class FormulaRequest
    {
        // Cell-slot marker embedded in a custom-function template. Control characters
        // never occur inside real formula bodies, so a single Replace is safe.
        internal const string Placeholder = "\x01";
        internal FormulaOperation Operation { get; private set; }
        internal string Argument { get; private set; }
        // Which empty argument of a custom template received the cell placeholder.
        // Only the editor reads it to rebuild the slot picker; the argument already
        // carries the placeholder at the chosen position.
        internal int Slot { get; private set; }
        private FormulaRequest(FormulaOperation operation, string argument, int slot = 0)
        { Operation = operation; Argument = argument; Slot = slot; }

        internal bool IsArithmetic { get { return Operation >= FormulaOperation.Add; } }
        internal bool ProcessesValues { get { return IsArithmetic || Operation == FormulaOperation.Custom; } }
        internal string Operator { get { return Operation == FormulaOperation.Add ? "+" : Operation == FormulaOperation.Subtract ? "-" : Operation == FormulaOperation.Multiply ? "*" : "/"; } }
        internal static string DefaultNumber(FormulaOperation operation)
        { return RibbonConfig.ArithmeticDefault(operation == FormulaOperation.Add ? "add" : operation == FormulaOperation.Subtract ? "subtract" : operation == FormulaOperation.Multiply ? "multiply" : "divide"); }
        internal static FormulaRequest Arithmetic(FormulaOperation operation, string value, bool expression)
        {
            if (operation < FormulaOperation.Add) throw new ArgumentException("不是数值计算操作");
            FormulaRequest parsed = expression ? IFErrorFormula(value) : IFErrorNumber(value);
            if (!expression && operation == FormulaOperation.Divide && double.Parse(parsed.Argument, CultureInfo.InvariantCulture) == 0)
                throw new FormatException("除数不能为零。");
            string argument = expression || parsed.Argument.StartsWith("-", StringComparison.Ordinal) ? "(" + parsed.Argument + ")" : parsed.Argument;
            return new FormulaRequest(operation, argument);
        }
        internal static bool IsNumeric(object value)
        {
            // VT_ERROR can return as a boxed HRESULT, not an ErrorWrapper.
            if (value is int && (((uint)(int)value & 0xFFFF0000u) == 0x800A0000u)) return false;
            if (value is uint && (((uint)value & 0xFFFF0000u) == 0x800A0000u)) return false;
            return value is double || value is float || value is decimal || value is int || value is long || value is short || value is byte || value is uint || value is ulong || value is ushort || value is sbyte;
        }
        internal static string NumericText(object value)
        {
            if (!IsNumeric(value)) throw new FormatException("单元格不是数值常量");
            double number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            if (double.IsNaN(number) || double.IsInfinity(number)) throw new FormatException("单元格数值不是有限数值");
            return number.ToString("R", CultureInfo.InvariantCulture);
        }
        internal static bool IsScalarValue(object value)
        { return IsNumeric(value) || value is string; }
        internal static string SerializeValue(object value)
        {
            if (value is string) return "\"" + ((string)value).Replace("\"", "\"\"") + "\"";
            if (IsNumeric(value)) return NumericText(value);
            throw new FormatException("单元格内容不是文本或数值，无法套用自定义函数。");
        }
        // Applies the stored custom template, replacing the cell slot with content.
        internal string ApplyCustom(string content)
        { return Argument.Replace(Placeholder, content); }
        internal static FormulaRequest Default(FormulaOperation operation)
        { if (operation >= FormulaOperation.Add) return Arithmetic(operation, DefaultNumber(operation), false);
          return new FormulaRequest(operation, operation == FormulaOperation.IFError ? "\"\"" : operation == FormulaOperation.Round ? "2" : null); }

        internal static FormulaRequest IFErrorText(string value)
        { return new FormulaRequest(FormulaOperation.IFError, "\"" + (value ?? "").Replace("\"", "\"\"") + "\""); }

        internal static FormulaRequest IFErrorNumber(string value)
        {
            double number;
            if (!double.TryParse((value ?? "").Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out number)
                || double.IsNaN(number) || double.IsInfinity(number))
                throw new FormatException("请输入有效数值，例如 0、-1 或 1.5；小数点使用英文句点。");
            return new FormulaRequest(FormulaOperation.IFError, number.ToString("R", CultureInfo.InvariantCulture));
        }

        internal static FormulaRequest IFErrorFormula(string value)
        {
            string expression = (value ?? "").Trim();
            if (expression.StartsWith("=", StringComparison.Ordinal)) expression = expression.Substring(1).Trim();
            if (expression.Length == 0) throw new FormatException("请输入公式或引用，例如 A1 或 SUM(B1:B3)，可省略开头的等号。");
            // Parse as one complete function argument. This rejects separators or
            // closing parentheses that would alter the wrapper's argument list.
            string unwrapped = FormulaWrapper.Transform("=__ARG(" + expression + ",0)", FormulaOperation.Unwrap);
            if (unwrapped != "=" + expression || expression.StartsWith("=", StringComparison.Ordinal))
                throw new FormatException("请输入一个完整的公式参数；函数内部可有逗号，参数外部不能有逗号或分号。");
            return new FormulaRequest(FormulaOperation.IFError, expression);
        }

        internal static FormulaRequest Round(string digits)
        {
            int count;
            if (!int.TryParse((digits ?? "").Trim(), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out count))
                throw new FormatException("位数必须是整数，例如 10、0 或 -2（32 位整数范围）。");
            return new FormulaRequest(FormulaOperation.Round, count.ToString(CultureInfo.InvariantCulture));
        }

        internal static FormulaRequest Custom(string template)
        { return Custom(template, 0); }

        internal static FormulaRequest Custom(string template, int slot)
        {
            string expression = (template ?? "").Trim();
            if (expression.StartsWith("=", StringComparison.Ordinal)) expression = expression.Substring(1).Trim();
            if (expression.Length == 0) throw new FormatException("请输入函数模板，例如 =TEXTSPLIT(,\"-\",1) 或 =ISBLANK()。");
            return new FormulaRequest(FormulaOperation.Custom, FormulaWrapper.PrepareCustomTemplate(expression, slot), slot);
        }
    }
}
