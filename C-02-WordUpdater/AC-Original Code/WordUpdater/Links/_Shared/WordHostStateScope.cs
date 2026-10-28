using System;
using System.Collections.Generic;
using System.Globalization;

namespace WordUpdater
{
    // Owns only state changes, never the borrowed Application reference.
    internal sealed class WordHostStateScope : IDisposable
    {
        private readonly object application, alerts, screen;
        private bool alertsTouched, screenTouched, restored;
        private WordHostStateScope(object app)
        {
            application = app;
            // No mutation is allowed until both originals have been captured.
            alerts = ((dynamic)app).DisplayAlerts;
            screen = ((dynamic)app).ScreenUpdating;
        }
        internal static WordHostStateScope Enter(object app)
        {
            var scope = new WordHostStateScope(app);
            try
            {
                scope.alertsTouched = true; ((dynamic)app).DisplayAlerts = 0;
                scope.screenTouched = true; ((dynamic)app).ScreenUpdating = false;
                return scope;
            }
            catch (Exception primary)
            {
                try { scope.Dispose(); }
                catch (Exception recovery) { throw new AggregateException("宿主状态设置及恢复失败。", primary, recovery); }
                throw;
            }
        }
        private static bool Same(object first, object second)
        { return Equals(first, second) || Convert.ToString(first, CultureInfo.InvariantCulture) == Convert.ToString(second, CultureInfo.InvariantCulture); }
        public void Dispose()
        {
            if (restored) return;
            restored = true;
            var errors = new List<Exception>();
            if (screenTouched) try
            {
                ((dynamic)application).ScreenUpdating = screen;
                if (!Same(((dynamic)application).ScreenUpdating, screen)) throw new InvalidOperationException("屏幕刷新恢复核对失败。");
            }
            catch (Exception error) { errors.Add(error); }
            if (alertsTouched) try
            {
                ((dynamic)application).DisplayAlerts = alerts;
                if (!Same(((dynamic)application).DisplayAlerts, alerts)) throw new InvalidOperationException("提示状态恢复核对失败。");
            }
            catch (Exception error) { errors.Add(error); }
            if (errors.Count > 0) throw new AggregateException("宿主状态恢复失败，请检查提示及屏幕刷新状态。", errors);
        }
    }
}
