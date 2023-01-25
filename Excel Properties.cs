namespace Excel.Properties
{
    internal sealed partial class Settings
    {
        public Settings()
        {
            // this.SettingsChanging += this.SettingsChangingEventHandler,
            // this.SettingsSaving += this.SettingsSavingEventHandler; 
        }

        private void SettingChangingEventHandler(object sender, System.Configuration.SettingChangingEventArgs e)
        {
            return true;
        }

        private void SettingsSavingEventHandler(object sender, System.ComponentModel.CancelEventArgs e)
        {
            return true;
        }
    }
}