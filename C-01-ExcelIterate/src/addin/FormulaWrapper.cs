using System;
using System.Collections.Generic;

namespace Excel_Iterate
{
    internal enum FormulaOperation { IFError, Abs, Round, Unwrap, Custom, Add, Subtract, Multiply, Divide }

    // Formula (not FormulaLocal) syntax. Preserve the argument's original text.
    internal static class FormulaWrapper
    {
        internal static string Transform(string formula, FormulaOperation operation)
        { return Transform(formula, FormulaRequest.Default(operation)); }

        internal static string Transform(string formula, FormulaRequest request)
        {
            if (request == null) throw new ArgumentNullException("request");
            FormulaOperation operation = request.Operation;
            if (string.IsNullOrEmpty(formula) || formula[0] != '=')
                throw new FormatException("读取到的公式缺少等号。");
            string body = formula.Substring(1);
            int firstSeparator, outerClose, firstOpen;
            Scan(body, out firstOpen, out firstSeparator, out outerClose);
            if (request.IsArithmetic) return "=(" + body + ")" + request.Operator + request.Argument;
            if (operation == FormulaOperation.IFError) return "=IFERROR(" + body + "," + request.Argument + ")";
            if (operation == FormulaOperation.Abs) return "=ABS(" + body + ")";
            if (operation == FormulaOperation.Round) return "=ROUND(" + body + "," + request.Argument + ")";
            if (operation == FormulaOperation.Custom) return "=" + request.ApplyCustom(body);
            if (firstOpen < 1 || outerClose < firstOpen || body.Substring(outerClose + 1).Trim().Length != 0)
                return formula;
            string name = body.Substring(0, firstOpen).Trim();
            if (name.Length == 0 || !(char.IsLetter(name[0]) || name[0] == '_')) return formula;
            foreach (char c in name)
                if (!(char.IsLetterOrDigit(c) || c == '_' || c == '.')) return formula;
            int end = firstSeparator >= 0 ? firstSeparator : outerClose;
            string argument = body.Substring(firstOpen + 1, end - firstOpen - 1);
            return argument.Trim().Length == 0 ? formula : "=" + argument;
        }

        private static void Scan(string text, out int firstOpen, out int firstSeparator, out int outerClose)
        {
            firstOpen = firstSeparator = outerClose = -1;
            Stack<char> nesting = new Stack<char>();
            char quote = '\0';
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (quote != '\0')
                {
                    if (c == quote)
                    {
                        if (i + 1 < text.Length && text[i + 1] == quote) i++;
                        else quote = '\0';
                    }
                    continue;
                }
                // In structured references apostrophe escapes special characters,
                // rather than starting a sheet-name quote. All bracket text is opaque.
                if (nesting.Count > 0 && nesting.Peek() == '[')
                {
                    if (c == '\'' && i + 1 < text.Length && "[]#'@".IndexOf(text[i + 1]) >= 0) { i++; continue; }
                    if (c == '[') nesting.Push(c);
                    else if (c == ']') nesting.Pop();
                    continue;
                }
                if (c == '"' || c == '\'') { quote = c; continue; }
                if (c == '(' || c == '[' || c == '{')
                {
                    if (c == '(' && nesting.Count == 0 && firstOpen < 0) firstOpen = i;
                    nesting.Push(c); continue;
                }
                if (c == ')' || c == ']' || c == '}')
                {
                    char expected = c == ')' ? '(' : c == ']' ? '[' : '{';
                    if (nesting.Count == 0 || nesting.Pop() != expected) throw new FormatException("公式括号不完整，未修改选区。");
                    if (nesting.Count == 0 && c == ')' && outerClose < 0) outerClose = i;
                    continue;
                }
                if ((c == ',' || c == ';') && nesting.Count == 1 && nesting.Peek() == '(' && firstSeparator < 0 && outerClose < 0)
                    firstSeparator = i;
            }
            if (quote != '\0' || nesting.Count != 0) throw new FormatException("公式引号或括号不完整，未修改选区。");
        }

        // Validates a user-supplied custom-function template (no leading '=') and
        // replaces the chosen empty top-level argument with the cell placeholder.
        internal static string PrepareCustomTemplate(string expression)
        { return PrepareCustomTemplate(expression, 0); }

        internal static string PrepareCustomTemplate(string expression, int slot)
        {
            int firstOpen, firstSeparator, outerClose;
            Scan(expression, out firstOpen, out firstSeparator, out outerClose);
            if (firstOpen < 1)
                throw new FormatException("请输入单个函数模板，例如 =TEXTSPLIT(,\"-\",1) 或 =ISBLANK()。");
            string name = expression.Substring(0, firstOpen).Trim();
            if (name.Length == 0 || !(char.IsLetter(name[0]) || name[0] == '_'))
                throw new FormatException("函数名不合法，请使用英文函数名，例如 =TEXTSPLIT(,\"-\",1)。");
            foreach (char c in name)
                if (!(char.IsLetterOrDigit(c) || c == '_' || c == '.'))
                    throw new FormatException("函数名不合法，请使用英文函数名。");
            if (outerClose < 0 || expression.Substring(outerClose + 1).Trim().Length != 0)
                throw new FormatException("模板必须是一个完整的函数调用，例如 =TEXTSPLIT(,\"-\",1)。");
            List<int[]> slots = CollectEmptyArguments(expression, firstOpen + 1, outerClose);
            if (slots.Count == 0)
                throw new FormatException("请在参数列表中留一个空参数位表示每个单元格内容，例如 =TEXTSPLIT(,\"-\",1) 或 =ISBLANK()。");
            if (slot < 0 || slot >= slots.Count)
                throw new FormatException("选区填入的位置不存在，当前模板只有 " + slots.Count + " 个空参数位。");
            int[] picked = slots[slot];
            return expression.Substring(0, picked[0]) + FormulaRequest.Placeholder + expression.Substring(picked[1]);
        }

        // Shape check without throwing, so the editor can rescan on every keystroke
        // and keep its slot picker in step with the template being typed.
        internal static List<int[]> FindCustomSlots(string expression)
        {
            int firstOpen, firstSeparator, outerClose;
            try { Scan(expression, out firstOpen, out firstSeparator, out outerClose); }
            catch (FormatException) { return null; }
            if (firstOpen < 1) return null;
            string name = expression.Substring(0, firstOpen).Trim();
            if (name.Length == 0 || !(char.IsLetter(name[0]) || name[0] == '_')) return null;
            foreach (char c in name)
                if (!(char.IsLetterOrDigit(c) || c == '_' || c == '.')) return null;
            if (outerClose < 0 || expression.Substring(outerClose + 1).Trim().Length != 0) return null;
            return CollectEmptyArguments(expression, firstOpen + 1, outerClose);
        }

        internal static int CountEmptyArguments(string expression)
        { List<int[]> slots = FindCustomSlots(expression); return slots == null ? 0 : slots.Count; }

        // Keeps every top-level empty argument instead of stopping at the first one.
        // Quotes and nested brackets stay opaque; only depth-zero separators split
        // arguments, so a nested function's empty argument is deliberately excluded.
        private static List<int[]> CollectEmptyArguments(string text, int start, int end)
        {
            List<int[]> slots = new List<int[]>();
            int depth = 0; char quote = '\0'; int current = start;
            for (int i = start; i < end; i++)
            {
                char c = text[i];
                if (quote != '\0')
                {
                    if (c == quote)
                    {
                        if (i + 1 < end && text[i + 1] == quote) i++;
                        else quote = '\0';
                    }
                    continue;
                }
                if (c == '"' || c == '\'') { quote = c; continue; }
                if (c == '(' || c == '[' || c == '{') { depth++; continue; }
                if (c == ')' || c == ']' || c == '}') { depth--; continue; }
                if ((c == ',' || c == ';') && depth == 0)
                {
                    if (text.Substring(current, i - current).Trim().Length == 0) slots.Add(new int[] { current, i });
                    current = i + 1;
                }
            }
            if (text.Substring(current, end - current).Trim().Length == 0) slots.Add(new int[] { current, end });
            return slots;
        }
    }
}
