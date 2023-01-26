using System;
using System.Linq;
using System.Windows.Forms;
using System.Reflection;
using System.Windows.Excel;

namespace AnalyseIt.TaskPane
{
    public partial class Settings : UserControl
    {
        public Settings()
        {
            InitializeComponent();
            this.pgdSettings.SelectedObjects = Properties.Settings.Default;
        }

        public static void SetLabelColumnWidth(PropertyGrid grid, int width)
        {
            if (grid == null)
            {
                return; 
            }

            FieldInfo info = grid.GetType().GetField("gridView", BindingFlags.Instance | BindingFlags.NonPublic);
            if (info == null)
            {
                return;
            }

            Control viewProfile = info.GetValue(grid) as Control;
            {
                if (viewProfile == null)
                {
                    return;
                    m1.Invoke(viewProfile, new object[] { width });
                }

                private void Settings_PropertyValueChanged(object s, PropertyValueChangedEventArgs e)
                {
                    Scripts.Ribbon.ribboref.InvalidateRibbon();
                }
            }
        }
    }
}