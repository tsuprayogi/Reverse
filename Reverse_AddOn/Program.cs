using System;
using SAPbouiCOM;
using SAPbobsCOM;
using Reverse_AddOn.Helper;
using Reverse_AddOn.Model;
using System.Collections.Generic;
using System.Linq;

namespace testSAPDesk
{
    internal static class Program
    {
        #region GLOBALS

        public static SAPbouiCOM.Application SBO_Application;
        public static SAPbobsCOM.Company DI_Company;

        // FORM TYPE
        private const string BP_FORM_TYPE = "143";

        // CUSTOM FORM UID
        private const string CUSTOM_FORM_UID = "frmCust";

        // ITEM IDS
        private const string BTN_GROUPING = "btnGrp";
        private const string ITEM_CARDCODE = "4";
        private const string REF_ITEM = "75";

        // MENU
        private const string MENU_ADD_ROW = "5966";

        // UDF
        private const string UDF_REVERSE = "U_T2_REVERSE";

        #endregion

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                InitializeApplication(args);
                InitializeCompany();

                RegisterEvents();

                SBO_Application.SetStatusBarMessage(
                    "Add-on connected",
                    BoMessageTime.bmt_Short,
                    false);

                System.Windows.Forms.Application.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString());
            }
            finally
            {
                Cleanup();
            }
        }

        #region INITIALIZATION

        private static void InitializeApplication(string[] args)
        {
            SboGuiApi guiApi = new SboGuiApi();

            string connectionString = args[0];

            guiApi.Connect(connectionString);

            SBO_Application = guiApi.GetApplication();
        }

        private static void InitializeCompany()
        {
            DI_Company = new SAPbobsCOM.Company();

            string cookie = DI_Company.GetContextCookie();

            string connectionContext =
                SBO_Application.Company.GetConnectionContext(cookie);

            DI_Company.SetSboLoginContext(connectionContext);

            int result = DI_Company.Connect();

            if (result != 0)
            {
                DI_Company.GetLastError(out int errCode, out string errMsg);

                throw new Exception(
                    $"DI API Connection Failed ({errCode}) : {errMsg}");
            }
        }

        private static void RegisterEvents()
        {
            SBO_Application.AppEvent += OnAppEvent;
            SBO_Application.ItemEvent += OnItemEvent;
            SBO_Application.MenuEvent += OnMenuEvent;
        }

        #endregion

        #region APP EVENTS

        private static void OnAppEvent(BoAppEventTypes eventType)
        {
            switch (eventType)
            {
                case BoAppEventTypes.aet_ShutDown:
                case BoAppEventTypes.aet_CompanyChanged:
                case BoAppEventTypes.aet_ServerTerminition:

                    Cleanup();

                    System.Windows.Forms.Application.Exit();

                    break;
            }
        }

        #endregion

        #region ITEM EVENTS

        private static void OnItemEvent(
            string formUID,
            ref ItemEvent pVal,
            out bool bubbleEvent)
        {
            bubbleEvent = true;

            try
            {
                if (pVal.FormTypeEx != BP_FORM_TYPE)
                    return;

                HandleBusinessPartnerFormEvents(formUID, pVal);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.ToString());
            }
        }

        private static void HandleBusinessPartnerFormEvents(
            string formUID,
            ItemEvent pVal)
        {
            // FORM LOAD
            if (pVal.EventType == BoEventTypes.et_FORM_LOAD &&
                !pVal.BeforeAction)
            {
                AddGroupingButton(formUID);
            }

            // BUTTON CLICK
            if (pVal.ItemUID == BTN_GROUPING &&
                pVal.EventType == BoEventTypes.et_ITEM_PRESSED &&
                !pVal.BeforeAction)
            {
                OpenCustomForm(formUID);
            }
        }

        #endregion

        #region FORM LOGIC

        private static void AddGroupingButton(string formUID)
        {
            Form form = SBO_Application.Forms.Item(formUID);

            if (IsItemExists(form, BTN_GROUPING))
                return;

            try
            {
                form.Freeze(true);

                Item newItem = form.Items.Add(
                    BTN_GROUPING,
                    BoFormItemTypes.it_BUTTON);

                Item refItem = form.Items.Item(REF_ITEM);

                newItem.Left = refItem.Left - 10;
                newItem.Top = refItem.Top - 40;
                newItem.Width = 100;
                newItem.Height = 19;

                Button button =
                    (Button)newItem.Specific;

                button.Caption = "Grouping";
            }
            finally
            {
                form.Freeze(false);
            }
        }
       

        private static List<GroupedItemModel> GetGroupedItemsFromGrpoMatrix(Form form)
        {
            SAPbouiCOM.Matrix matrix =
                (SAPbouiCOM.Matrix)form.Items.Item("38").Specific;

            // GROUP BY:
            // ItemCode + Warehouse
            Dictionary<string, GroupedItemModel> grouped =
                new Dictionary<string, GroupedItemModel>();

            for (int i = 1; i <= matrix.VisualRowCount; i++)
            {
                string itemCode =
                    ((SAPbouiCOM.EditText)
                    matrix.Columns.Item("1")
                    .Cells.Item(i).Specific)
                    .Value
                    .Trim();

                if (string.IsNullOrWhiteSpace(itemCode))
                    continue;

                string desc =
                    ((SAPbouiCOM.EditText)
                    matrix.Columns.Item("3")
                    .Cells.Item(i).Specific)
                    .Value
                    .Trim();

                string qtyStr =
                    ((SAPbouiCOM.EditText)
                    matrix.Columns.Item("11")
                    .Cells.Item(i).Specific)
                    .Value
                    .Trim();

                double qty = 0;

                double.TryParse(
                    qtyStr,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out qty);

                string uom =
                    ((SAPbouiCOM.EditText)
                    matrix.Columns.Item("212")
                    .Cells.Item(i).Specific)
                    .Value
                    .Trim();

                string whsCode =
                    ((SAPbouiCOM.EditText)
                    matrix.Columns.Item("24")
                    .Cells.Item(i).Specific)
                    .Value
                    .Trim();

                // =========================
                // GROUP KEY
                // =========================
                string groupKey = $"{itemCode}|{whsCode}";

                if (grouped.ContainsKey(groupKey))
                {
                    grouped[groupKey].TotalQty += qty;
                }
                else
                {
                    grouped[groupKey] = new GroupedItemModel
                    {
                        ItemCode = itemCode,
                        Description = desc,
                        TotalQty = qty,
                        UnitMsr = uom,
                        WhsCode = whsCode
                    };
                }
            }

            return grouped.Values.ToList();
        }

        private static void FillMatrix(
            Form form,
            List<GroupedItemModel> items)
        {
            SAPbouiCOM.DataTable dt =
                form.DataSources.DataTables.Item("DT_GRPO");

            SAPbouiCOM.Matrix matrix =
                (SAPbouiCOM.Matrix)
                form.Items.Item("mtxGrpo").Specific;

            try
            {
                form.Freeze(true);

                dt.Rows.Clear();

                for (int i = 0; i < items.Count; i++)
                {
                    dt.Rows.Add(1);

                    // =========================
                    // NUMBERING
                    // =========================
                    dt.SetValue("Num", i, (i + 1).ToString());

                    // =========================
                    // DATA
                    // =========================
                    dt.SetValue("ItemCode", i, items[i].ItemCode);

                    dt.SetValue(
                        "Dscription",
                        i,
                        items[i].Description);

                    dt.SetValue(
                        "Qty",
                        i,
                        items[i].TotalQty.ToString("0.####"));

                    dt.SetValue(
                        "unitMsr",
                        i,
                        items[i].UnitMsr);

                    dt.SetValue(
                        "WhsCode",
                        i,
                        items[i].WhsCode);
                }

                matrix.Clear();
                matrix.LoadFromDataSource();

                // =========================
                // ALIGNMENT
                // =========================
                SAPbouiCOM.Column qtyCol =
                    matrix.Columns.Item("colQty");

                qtyCol.RightJustified = true;

                // =========================
                // AUTO SIZE
                // =========================
                if (matrix.RowCount > 0)
                {
                    matrix.AutoResizeColumns();
                }
            }
            finally
            {
                form.Freeze(false);
            }
        }

        private static void OpenCustomForm(string formUID)
        {
            Form bpForm =
                SBO_Application.Forms.Item(formUID);

            Form customForm = null;

            try
            {
                bpForm.Freeze(true);

                string cardCode = GetCardCode(bpForm);

                if (string.IsNullOrWhiteSpace(cardCode))
                {
                    SBO_Application.MessageBox(
                        "Kode / Nama Vendor Kosong");

                    return;
                }

                LoadCustomForm();

                customForm =
                    SBO_Application.Forms.Item(CUSTOM_FORM_UID);

                customForm.Freeze(true);

                string docNum =
                    ((SAPbouiCOM.EditText)
                    bpForm.Items.Item("8").Specific)
                    .Value
                    .Trim();

                customForm.Title =
                    $"GRPO Item Grouping : {docNum}";

                // =========================
                // GET GROUPED DATA
                // =========================
                List<GroupedItemModel> groupedItems =
                    GetGroupedItemsFromGrpoMatrix(bpForm);

                // =========================
                // FILL MATRIX
                // =========================
                FillMatrix(customForm, groupedItems);

                customForm.Mode = BoFormMode.fm_OK_MODE;

                // tampilkan terakhir supaya tidak blink
                customForm.Visible = true;
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.ToString());
            }
            finally
            {
                if (customForm != null)
                    customForm.Freeze(false);

                bpForm.Freeze(false);
            }
        }

        private static string GetCardCode(Form form)
        {
            EditText txtCardCode =
                (EditText)form.Items.Item(ITEM_CARDCODE).Specific;

            return txtCardCode.Value.Trim();
        }

        private static void LoadCustomForm()
        {
            // already opened
            try
            {
                SBO_Application.Forms.Item(CUSTOM_FORM_UID);
                return;
            }
            catch
            {
            }

            string xml =
                FormLoader.LoadFromXML("testSapStudio.srf");

            SBO_Application.LoadBatchActions(xml);
        }

        #endregion

        #region MENU EVENTS

        private static void OnMenuEvent(
            ref MenuEvent pVal,
            out bool bubbleEvent)
        {
            bubbleEvent = true;

            try
            {
                if (pVal.BeforeAction)
                    return;

                if (pVal.MenuUID != MENU_ADD_ROW)
                    return;

                Form activeForm =
                    SBO_Application.Forms.ActiveForm;

                if (activeForm.TypeEx != "940")
                    return;

                bool success = SetReverseCheckbox(activeForm);

                if (!success)
                {
                    Form udfForm = FindUdfForm("-940");

                    if (udfForm != null)
                    {
                        success = SetReverseCheckbox(udfForm);
                    }
                }

                if (!success)
                {
                    SBO_Application.SetStatusBarMessage(
                        $"{UDF_REVERSE} not found",
                        BoMessageTime.bmt_Short,
                        true);
                }
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.ToString());
            }
        }

        #endregion

        #region UDF LOGIC

        private static bool SetReverseCheckbox(Form form)
        {
            for (int i = 0; i < 10; i++)
            {
                try
                {
                    Item item = form.Items.Item(UDF_REVERSE);

                    if (item == null)
                        continue;

                    item.Enabled = true;

                    CheckBox checkBox =
                        (CheckBox)item.Specific;

                    checkBox.Checked = true;

                    item.Enabled = false;

                    return true;
                }
                catch
                {
                    System.Threading.Thread.Sleep(200);
                }
            }

            return false;
        }

        private static Form FindUdfForm(string formType)
        {
            for (int i = 0; i < SBO_Application.Forms.Count; i++)
            {
                Form form = SBO_Application.Forms.Item(i);

                if (form.TypeEx == formType &&
                    form.Visible)
                {
                    return form;
                }
            }

            return null;
        }

        #endregion

        #region HELPERS

        private static bool IsItemExists(Form form, string itemId)
        {
            try
            {
                form.Items.Item(itemId);
                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region CLEANUP

        private static void Cleanup()
        {
            try
            {
                if (DI_Company != null &&
                    DI_Company.Connected)
                {
                    DI_Company.Disconnect();
                }
            }
            catch
            {
            }

            try
            {
                if (DI_Company != null)
                {
                    System.Runtime.InteropServices.Marshal
                        .ReleaseComObject(DI_Company);
                }
            }
            catch
            {
            }

            DI_Company = null;
            SBO_Application = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        #endregion
    }
}