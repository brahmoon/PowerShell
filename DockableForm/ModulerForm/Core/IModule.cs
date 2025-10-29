using System.Windows.Forms;

namespace DockableModularForm
{
    public interface IModule
    {
        string Name { get; }
        UserControl GetControl();
        void OnLoad();
        void OnUnload();
    }
}
