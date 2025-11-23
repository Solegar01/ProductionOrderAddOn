using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;

namespace ProductionOrderAddOn.Helpers
{
    public static class FormHelper
    {
        private static SAPbouiCOM.ProgressBar _pb;

        public static void StartLoading(SAPbouiCOM.Form oForm, string pbText, int max, bool stopable)
        {
            if (_pb != null) { _pb.Stop(); _pb = null; }
            _pb = Application.SBO_Application.StatusBar.CreateProgressBar(pbText, max, stopable);
            oForm.Freeze(true);
        }

        public static void FinishLoading(SAPbouiCOM.Form oForm)
        {
            if (_pb != null) { _pb.Stop(); _pb = null; }
            oForm.Freeze(false);
        }

        public static void SetTextValueLoading(SAPbouiCOM.Form oForm, int value = 0, string text = "")
        {
            _pb.Value = value;
            if (!string.IsNullOrEmpty(text))
            {
                _pb.Text = text;
            }
        }
    }
}
