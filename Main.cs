namespace PasteEquation
{
    public partial class Main
    {
        private void Main_Startup(object sender, System.EventArgs e)
        {
            KeyboardHook.SetHook();
        }


        private void Main_Shutdown(object sender, System.EventArgs e)
        {
            KeyboardHook.ReleaseHook();
        }

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Main_Startup);
            this.Shutdown += new System.EventHandler(Main_Shutdown);
        }
    }
}