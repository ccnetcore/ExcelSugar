using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace ExcelSugar.Core
{
    public class ExcelHandlerProviderFactory
    {
        private ExcelSugarConfig Config { get; }
        public ExcelHandlerProviderFactory(ExcelSugarConfig oemConfig)
        {
            this.Config = oemConfig;
        }

        public IExcelSugarProvider CreateOemProvider()
        {
            IExcelSugarProvider result = default(IExcelSugarProvider);
            switch (Config.HandlerType)
            {
                case ExcelHandlerType.Npoi:
                    var targetType = Assembly.Load("ExcelSugar.Npoi").GetType("ExcelSugar.Npoi.NpoiOemProvider");
                    object instance = Activator.CreateInstance(targetType, new object[] { Config });
                    result = instance as IExcelSugarProvider;
                    break;
                default:
                    break;
            }
            return result;
        }
    }
}
