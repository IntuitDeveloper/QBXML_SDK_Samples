// Copyright (c) 2021-2022 by Intuit Inc.
// All rights reserved
// Usage governed by the QuickBooks SDK Developer's License Agreement

using System;
using System.Collections.Generic;
using System.Text;

namespace MCInvoiceAddQBFC.Session_Framework
{
    public class QBException : Exception
    {
        private QBException() { }

        public QBException(string sMsg)
            : base(sMsg)
        {
        }

        public override string ToString()
        {
            return base.Message;
        }

    }
}
