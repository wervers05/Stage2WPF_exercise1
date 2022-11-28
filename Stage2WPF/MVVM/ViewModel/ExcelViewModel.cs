using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Stage2WPF.MVVM.ViewModel
{
    class ExcelViewModel
    {
        public interface IOService
        {
            string OpenFileDialog(string path);

            Stream OpenFile(string path);
        }
    }
}
