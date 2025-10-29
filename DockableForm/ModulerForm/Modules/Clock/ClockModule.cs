using System;
using System.Drawing;
using System.Windows.Forms;
using DockableModularForm;

public class ClockModule : IModule
{
    private readonly UserControl _control;
    private readonly Label _label;
    private readonly Timer _timer;

    public ClockModule()
    {
        _label = new Label
        {
            Dock = DockStyle.Fill,
            Font = new Font("Segoe UI", 24f, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleCenter
        };

        _control = new UserControl
        {
            Dock = DockStyle.Fill
        };
        _control.Controls.Add(_label);

        _timer = new Timer
        {
            Interval = 1000
        };
        _timer.Tick += (s, e) => UpdateTime();
    }

    public string Name => "Clock";

    public UserControl GetControl()
    {
        return _control;
    }

    public void OnLoad()
    {
        UpdateTime();
        _timer.Start();
    }

    public void OnUnload()
    {
        _timer.Stop();
        _timer.Dispose();
    }

    private void UpdateTime()
    {
        _label.Text = DateTime.Now.ToString("HH:mm:ss");
    }
}
