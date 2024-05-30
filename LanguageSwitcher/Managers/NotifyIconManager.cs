using System.Drawing;
using System.Windows.Forms;

namespace LanguageSwitcher
{
    public class NotifyIconManager : IDisposable
    {
        private readonly NotifyIcon notifyIcon;

        public NotifyIconManager()
        {
            notifyIcon = new NotifyIcon();
        }

        public void Initialize(EventHandler onExitMenuItemClicked)
        {
            notifyIcon.Icon = new Icon("Resources/tools.ico");
            notifyIcon.Visible = true;
            notifyIcon.ContextMenuStrip = new ContextMenuStrip();
            notifyIcon.ContextMenuStrip.Items.Add("Exit", null, onExitMenuItemClicked);
        }

        public void Dispose()
        {
            notifyIcon.Dispose();
        }
    }
}