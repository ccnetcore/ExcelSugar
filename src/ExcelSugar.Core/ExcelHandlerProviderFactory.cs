using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace ExcelSugar.Core
{
    public class ExcelHandlerProviderFactory
    {
        private OemConfig Config { get; }
        public ExcelHandlerProviderFactory(OemConfig oemConfig)
        {
            this.Config = oemConfig;
        }

        public IOemProvider CreateOemProvider()
        {
            IOemProvider result = default(IOemProvider);
            switch (Config.HandlerType)
            {
                case ExcelHandlerType.Npoi:
                    var targetType = Assembly.Load("ExcelSugar.Npoi").GetType("ExcelSugar.Npoi.NpoiOemProvider");
                    object instance = Activator.CreateInstance(targetType, new object[] { Config });
                    result = instance as IOemProvider;
                    break;
                default:
                    break;
            }
            return result;
        }
    }
}
