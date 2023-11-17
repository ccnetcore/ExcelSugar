using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Text;

namespace ExcelSugar.Core
{
    public class PipelineBuilder<T>
    {
        LinkedList<object> linkedList = new LinkedList<object>();

        LinkedList<string> From = new LinkedList<string>();

        LinkedList<Expression<Func<T, bool>>> Where = new LinkedList<Expression<Func<T, bool>>>();

    }
}
