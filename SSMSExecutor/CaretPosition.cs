using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SSMSExecutor
{
    public class CaretPosition
    {
        public int Line { get; set; }
        public int LineCharOffset { get; set; }
    }

    public class CaretCurrentStatement
    {
        public CaretPosition FirstToken { get; set; }
        public CaretPosition LastToken { get; set; }
    }
}
