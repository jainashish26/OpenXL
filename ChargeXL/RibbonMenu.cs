using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;


namespace ChargeXL
{
    public class RibbonMenu
    {
        [ExcelFunction(IsVolatile = true, Name = "TOMORROW")]
        public static DateTime Tomorrow()
        {
            return DateTime.Today.AddDays(1);
        }

        [ComVisible(true)]
        public class Ribbon : ExcelRibbon
        {
            private IRibbonUI ribbon = null;
            
            
            public void OnLogonPressed(IRibbonControl control)
            {
                if (ribbon != null)
                {
                    ribbon.InvalidateControl(control.Id);
                }

            }

            public void OnLoad(IRibbonUI ribbon)
            {
                this.ribbon = ribbon;
            }

            public static void HelloJi()
            {
                MessageBox.Show("Hello Ashish");
            }

            public void OnButtonPressed(IRibbonControl control)
            {
                MessageBox.Show("Hello from control " + control.Id);
            }

            public static void ShowHelloMessage()
            {
                MessageBox.Show("Hello from 'ShowHelloMessage'.");
            }

            public override string GetCustomUI(string uiName)
            {
                MessageBox.Show("Add-in 'Charge Excel' called!");
                String rbnGroup = @"<group 
					                id='grp'
					                imageMso='EnvelopesAndLabelsDialog'
					                label='My First Group'
					                visible='true'>
					                <button 
						                description='hello'
						                enabled='true'
						                id='btn'
						                imageMso='AnimationAudio'
						                label='Dynamic Button'
						                showImage='true'
						                size='large'
						                visible='true'
                                        onAction='RunTagMacro' 
                                        tag='ShowHelloMessage'/>
				                </group>";
                return
                @"<customUI  xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
	                <ribbon>
		                <tabs>
			                <tab 
				                id='chargeExcel'
				                insertAfterMso='TabAddin'
                        keytip='XL'
				                label='Charge Excel'
				                visible='true'> " + rbnGroup + @"
				                
			                </tab>
		                </tabs>
	                </ribbon>
                </customUI>";
            }
        }
    }
}
