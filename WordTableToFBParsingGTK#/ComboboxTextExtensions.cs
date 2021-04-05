using Gtk;
using System.Collections.Generic;

namespace WordTableToFBParsingGTK
{
    public static class ComboBoxTextExtensions
    {
        public static void AppendTextCollection(this ComboBoxText comboBox, IEnumerable<string> stringCollection)
        {
            foreach (var str in stringCollection)
                comboBox.AppendText(str);
        }
    }
}
