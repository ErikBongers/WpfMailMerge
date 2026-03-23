using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfMailMerge
    {
    internal class Comm
        {

        public static void PutMachineNameOnClipboard()
            {
            var machine = System.Environment.MachineName;
            if (machine == null)
                return;
            var user = System.Environment.UserName;
            if (user == null)
                return;
            System.Windows.Clipboard.SetText($"DKO3PLUGIN:{machine}.{user}");
            }
        }
    }
