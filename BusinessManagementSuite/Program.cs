using System;
using System.Windows.Forms;
using BusinessManagementSuite.Data;

namespace BusinessManagementSuite;

internal static class Program
{
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();

        try
        {
            DbInitializer.Initialize();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Database initialization failed:\n{ex.Message}",
                "Startup Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return;
        }

        Application.Run(new MainForm());
    }
}
