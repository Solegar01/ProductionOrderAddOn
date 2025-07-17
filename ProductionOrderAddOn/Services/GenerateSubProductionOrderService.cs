using SAPbouiCOM;
using System;

namespace ProductionOrderAddOn.Services
{
    public class GenerateSubProductionOrderService
    {
        private readonly Application _application;
        private const string ButtonUID = "btnGens";
        private const string TargetFormType = "4369"; // Example: Work Order, replace with correct type

        public GenerateSubProductionOrderService(Application application)
        {
            _application = application;
            _application.ItemEvent += OnItemEvent;
        }

        public void AddButtonToForm()
        {
            try
            {
                Form activeForm = _application.Forms.ActiveForm;

                // Avoid adding if already exists
                if (ItemExists(activeForm, ButtonUID))
                    return;

                Item refItem = activeForm.Items.Item("2");
                Item newItem = activeForm.Items.Add(ButtonUID, BoFormItemTypes.it_BUTTON);

                newItem.Top = refItem.Top;
                newItem.Left = refItem.Left;
                newItem.Width = 120;
                newItem.Height = 50;

                Button button = (Button)newItem.Specific;
                button.Caption = "Generate Sub-PO";
            }
            catch (Exception ex)
            {
                _application.MessageBox("Error adding button: " + ex.Message);
            }
        }

        private bool ItemExists(Form form, string itemUid)
        {
            try
            {
                var _ = form.Items.Item(itemUid);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void OnItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.ItemUID == ButtonUID &&
                pVal.EventType == BoEventTypes.et_ITEM_PRESSED &&
                !pVal.BeforeAction)
            {
                _application.MessageBox("Sub Production Order button clicked.");
                // Add your logic here
            }
        }
    }
}
