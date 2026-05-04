using System.Collections.Generic;
using System.Windows.Forms;

namespace MailConverter
{
    /// <summary>
    /// TabControl 扩展方法 - 支持通过 SetVisible 实时显示/隐藏标签页
    /// </summary>
    public static class TabControlExtensions
    {
        // 存储被隐藏的 TabPage 及其(父级TabControl, 索引)
        private static readonly Dictionary<TabPage, (TabControl, int)> _hiddenTabs = new Dictionary<TabPage, (TabControl, int)>();

        /// <summary>
        /// 设置 TabPage 的可见性（实时切换）
        /// </summary>
        /// <param name="tabPage">要设置的 TabPage</param>
        /// <param name="visible">true=显示, false=隐藏</param>
        public static void SetVisible(this TabPage tabPage, bool visible)
        {
            if (tabPage == null) return;

            var tabControl = tabPage.Parent as TabControl;

            if (visible)
            {
                // 显示 TabPage
                if (tabControl != null && tabControl.TabPages.Contains(tabPage))
                {
                    // 已经显示了，无需处理
                    return;
                }

                // 尝试从隐藏记录中恢复
                if (_hiddenTabs.TryGetValue(tabPage, out var info))
                {
                    var (originalTabControl, index) = info;
                    if (originalTabControl != null)
                    {
                        index = System.Math.Min(index, originalTabControl.TabCount);
                        originalTabControl.TabPages.Insert(index, tabPage);
                        Serilog.Log.Information("SetVisible: Restored {TabName} to {TabControl} at index {Index}", tabPage.Text, originalTabControl.Name, index);
                    }
                    _hiddenTabs.Remove(tabPage);
                }
            }
            else
            {
                // 隐藏 TabPage
                if (tabControl != null && tabControl.TabPages.Contains(tabPage))
                {
                    int index = tabControl.TabPages.IndexOf(tabPage);
                    _hiddenTabs[tabPage] = (tabControl, index);
                    tabControl.TabPages.Remove(tabPage);
                    Serilog.Log.Information("SetVisible: Hidden {TabName} from {TabControl}, index={Index}", tabPage.Text, tabControl.Name, index);
                }
            }
        }

        /// <summary>
        /// 检查 TabPage 是否可见
        /// </summary>
        public static bool IsVisible(this TabPage tabPage)
        {
            if (tabPage == null) return false;
            var tabControl = tabPage.Parent as TabControl;
            return tabControl != null && tabControl.TabPages.Contains(tabPage);
        }

        /// <summary>
        /// 批量设置可见性（防止闪烁）
        /// </summary>
        public static void SetVisibleBatch(this TabControl tabControl, IEnumerable<TabPage> tabPages, bool visible)
        {
            if (tabControl == null) return;

            tabControl.SuspendLayout();

            foreach (var tabPage in tabPages)
            {
                tabPage.SetVisible(visible);
            }

            tabControl.ResumeLayout();
        }
    }
}