// Copyright (c) 2007-2013 by Intuit Inc.
// All rights reserved
// Usage governed by the QuickBooks SDK Developer's License Agreement

using System;

namespace MCInvoiceAddQBFC.Session_Framework
{
    public enum ENEdition
    {
        edUS = 0,
        edCA = 1,
        edUK = 2,
    }

    static class QBEdition
    {
        public static readonly string[] codes = { "US", "CA", "UK" };

        public static string getEdition(ENEdition ed)
        {
            return codes[(int)ed];
        }
    }
}
