using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;

namespace PrinterLib
{
    public abstract class PrintBase
    {
        protected string content;

        public string Content
        {
            get { return content; }
            set { content = value; }
        }

        protected int labelQty;

        public int LabelQty
        {
            set { labelQty = value; }
        }

        protected string message;

        public string Message
        {
            get { return message; }
        }

        public abstract bool Print();
        public abstract bool Print(Encoding encode);
    }
}
