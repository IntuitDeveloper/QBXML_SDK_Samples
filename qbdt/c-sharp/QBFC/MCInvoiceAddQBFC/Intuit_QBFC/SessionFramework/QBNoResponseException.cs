// Copyright (c) 2007-2013 by Intuit Inc.
// All rights reserved
// Usage governed by the QuickBooks SDK Developer's License Agreement

using System;
using System.Collections.Generic;
using System.Text;

namespace MCInvoiceAddQBFC.Session_Framework
{
    public class QBNoResponseException : Exception
    {
        private readonly int _index;

        private QBNoResponseException() { }

        /// <summary>
        /// Describes an error in which no IResponse object is found in the IResponseList
        /// </summary>
        /// <param name="index">Index within the IResponseList that was not found</param>
        /// <param name="errorMsg">Textual error message</param>
        public QBNoResponseException(int index, string errorMsg)
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

        public override string ToString()
        {
            return base.ToString();
        }
    }
}
