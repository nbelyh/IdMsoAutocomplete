namespace IdMsoAutocomplete.Configuration
{
    public static class OptionsPrompt
    {
        public static void EditOptions(Options options)
        {
            var toEdit = new Options
            {
                OfficeVersion = options.OfficeVersion,
                OfficeApplication = options.OfficeApplication
            };

            var window = new OptionsWindow(toEdit);

            var result = window.ShowDialog();

            if (result.HasValue && result.Value)
            {
                options.OfficeVersion = toEdit.OfficeVersion;
                options.OfficeApplication = toEdit.OfficeApplication;
            }
        }
    }
}