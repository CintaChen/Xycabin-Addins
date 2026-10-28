using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace Excel_Iterate
{
    /// <summary>源列号 → 目标列号（每个来源一个）。</summary>
    internal sealed class MergeColumnMapper : IMergeColumnMapper
    {
        private readonly Dictionary<int, int> map = new Dictionary<int, int>();
        internal void Add(int sourceColumn, int targetColumn) { map[sourceColumn] = targetColumn; }
        public int MapColumn(int sourceColumn)
        {
            int target;
            return map.TryGetValue(sourceColumn, out target) ? target : 0;
        }
    }
}
