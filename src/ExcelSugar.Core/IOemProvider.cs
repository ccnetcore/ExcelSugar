﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSugar.Core
{
    public interface IOemProvider
    {
        IOemExportable<T> CreateExportable<T>(IEnumerable<T> expObj);
        IOemQueryable<T> CreateQueryable<T>();
    }
}