using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;

namespace TB1300
{
    [FormAttribute("TB1300.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_4").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("Item_5").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_6").Specific));
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("Item_7").Specific));
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("Item_8").Specific));
            this.Button5 = ((SAPbouiCOM.Button)(this.GetItem("Item_9").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Item_10").Specific));
            this.EditText2.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.EditText2_KeyDownAfter);
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("Item_11").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_12").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_13").Specific));
            this.Button6 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button7 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_16").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private void OnCustomInitialize()
        {

        }
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Button Button3;
        private SAPbouiCOM.Button Button4;
        private SAPbouiCOM.Button Button5;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;

        private void EditText2_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            throw new System.NotImplementedException();

        }

        private SAPbouiCOM.Button Button6;
        private SAPbouiCOM.Button Button7;
        private SAPbouiCOM.Matrix Matrix0;
    }
}