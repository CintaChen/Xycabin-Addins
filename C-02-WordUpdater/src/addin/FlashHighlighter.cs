using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace WordUpdater
{
    // =====================================================================
    // 标黄“闪一下”的公共实现
    //   （「更新链接」一次标黄 N 条 + 「查看链接」单条跳转，共用同一套）
    //
    // 为什么单独抽这个文件：
    //   标黄原本只存在于 LinkViewEngine.JumpToItemEx 里（引擎打黄，UI 起
    //   Timer 还原）。「更新链接」现在一次要标黄 N 条，且要求"结果弹窗关闭
    //   之后才开始计时"，LinkViewForm 那套【单槽】装不下。更关键的是：
    //   两套 Timer 各还原各的，会把【对方刚打上的黄】读成"原值"，最后还原
    //   成 wdYellow —— 黄色永久残留在用户文档里，这是最严重的失败模式。
    //   所以统一成一个多槽 session：按引用去重，永远保留最早的那个原值。
    //
    // 分层（沿用既有的 JumpToItemEx 约定）：
    //   引擎侧负责"在哪个瞬间打黄"——只有刚更新完的那一次迭代里 COM Range
    //   才是新鲜的（Update() 之后旧引用会 null 绑定错误）；
    //   UI 侧只负责"什么时候还原"。
    // =====================================================================

    public static class FlashHighlighter
    {
        // WdColorIndex：wdNoHighlight = 0，wdYellow = 7
        internal const int WdYellow = 7;

        /// <summary>
        /// 把目标区域标黄，并把它【原来】的高亮值通过 out 交出去。
        /// 任何失败都静默返回 false —— 标黄只是提示，绝不能因此打断更新。
        /// </summary>
        public static bool Apply(object targetRange, out int prevHighlight)
        {
            prevHighlight = 0;
            if (targetRange == null) return false;
            try
            {
                dynamic r = targetRange;
                prevHighlight = (int)r.HighlightColorIndex;   // 先记原值
                r.HighlightColorIndex = WdYellow;
                return true;
            }
            catch
            {
                prevHighlight = 0;
                return false;
            }
        }

        /// <summary>
        /// 按【原值】还原（不是清零）。直接复用 LinkViewEngine 既有的
        /// ClearFlash —— 它已是 public static 且内部全静默，不再重写一份。
        /// </summary>
        public static void Clear(object range, int prevHighlight)
        {
            LinkViewEngine.ClearFlash(range, prevHighlight);
        }

        /// <summary>
        /// 把视图滚到目标区域，让用户立刻看见那一抹黄。
        /// ★ 不 Select、不 Activate —— 不抢选区、不抢焦点（弹窗纪律）。
        ///   三种调用形态逐级退让，全不认就静默放弃：宁可不滚，也绝不退化成
        ///   会动用户选区的 Select。
        /// </summary>
        public static bool ScrollIntoView(object application, object targetRange)
        {
            if (application == null || targetRange == null) return false;
            try
            {
                dynamic win = ((dynamic)application).ActiveWindow;
                // Word/WPS：Window.ScrollIntoView(Obj, Down)
                try { win.ScrollIntoView(targetRange, true); return true; }
                catch { }
                try { win.ScrollIntoView(targetRange); return true; }
                catch { }
                try
                {
                    dynamic pane = win.ActivePane;
                    pane.ScrollIntoView(targetRange, true);
                    return true;
                }
                catch { }
            }
            catch { }
            return false;
        }
    }

    /// <summary>一条待还原的标黄：被打黄的 Range + 它原来的高亮值。</summary>
    public sealed class FlashEntry
    {
        public object Range;
        public int PrevHighlight;
    }

    /// <summary>
    /// 多槽标黄容器 + 定时器。支持"晚一点再开始计时"——
    /// 「更新链接」要求结果弹窗关闭的那一刻才启动 2.5 秒倒计时。
    /// </summary>
    public sealed class FlashSession
    {
        private readonly List<FlashEntry> _entries = new List<FlashEntry>();
        private Timer _timer;
        private bool _started;

        public int Count { get { return _entries.Count; } }
        public bool HasPending { get { return _entries.Count > 0; } }

        /// <summary>
        /// 登记一条已打上的黄。★ 同一个 Range 被重复登记时【保留最早那条】
        /// 并直接忽略后者 —— 否则后写者会把 7（先来者打上的黄）当成原值，
        /// 最后还原成 7，黄色就永久留在文档里了。
        /// </summary>
        public void Add(object range, int prev)
        {
            if (range == null) return;
            for (int i = 0; i < _entries.Count; i++)
            {
                if (ReferenceEquals(_entries[i].Range, range)) return;
            }
            FlashEntry e = new FlashEntry();
            e.Range = range;
            e.PrevHighlight = prev;
            _entries.Add(e);
        }

        /// <summary>
        /// 启动倒计时（「更新链接」在结果弹窗关闭后才调）。
        /// 定时器起不来就【立刻还原】，绝不把黄留在文档里。
        /// </summary>
        public void Start(int ms)
        {
            if (_entries.Count == 0 || _started) return;
            _started = true;
            try
            {
                if (_timer == null)
                {
                    _timer = new Timer();
                    _timer.Tick += delegate { ClearNow(); };
                }
                _timer.Interval = ms;
                _timer.Start();
            }
            catch
            {
                ClearNow();
            }
        }

        /// <summary>重复跳转时重开计时：停掉旧计时，但不清掉已打上的黄。</summary>
        public void Restart(int ms)
        {
            try { if (_timer != null) _timer.Stop(); }
            catch { }
            _started = false;
            Start(ms);
        }

        /// <summary>停掉定时器 + 逐条按原值还原 + 清空槽位。全程静默。</summary>
        public void ClearNow()
        {
            try
            {
                if (_timer != null)
                {
                    _timer.Stop();
                    _timer.Dispose();
                    _timer = null;
                }
            }
            catch { }
            FlashEntry[] arr = _entries.ToArray();
            _entries.Clear();
            for (int i = 0; i < arr.Length; i++)
            {
                try { FlashHighlighter.Clear(arr[i].Range, arr[i].PrevHighlight); }
                catch { }
            }
        }
    }
}
