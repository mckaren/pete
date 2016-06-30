using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace McKinsey.PowerPointGenerator
{
    public static class ParameterHelpers
    {
        public static void NotNull<T>(T parameter, string name)
        {
            if (parameter == null)
            {
                throw new ArgumentNullException(name);
            }
        }
    }
}
