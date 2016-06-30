using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using McKinsey.PowerPointGenerator.Core.Data;

namespace McKinsey.PowerPointGenerator.Commands
{
    public interface IUseIndexes
    {
        List<Index> UsedIndexes { get; set; }
    }
}
