// Copyright (c) 2007-2013 by Intuit Inc.
// All rights reserved
// Usage governed by the QuickBooks SDK Developer's License Agreement

using System;
using System.Collections.Generic;
using System.Text;

namespace MCInvoiceAddQBFC.Session_Framework
{
    public class QBResultException : Exception
    {
        private readonly int _index;

        private QBResultException() { }

        public QBResultException(int index, string errorMsg)
            : base(errorMsg)
        {
            _index = index;
        }

        /// <summary>
        /// Read-only property that provides the index of the offending IResponse object within
        /// the IResponseList object
        /// </summary>
        public int ErrorIndex
        {
            get
            {
                return _index;
            }
        }
    }
}
