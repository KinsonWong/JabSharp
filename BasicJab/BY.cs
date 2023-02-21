using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BasicJab
{
    [Guid("8A0B3453-BA6B-4628-969E-4BACE0FA6967"), ComVisible(true)]
    public enum BY   //BY 枚举
    {
        NAME = 1 << 1,
        DESCRIPTION = 1 << 2,
        ROLE = 1 << 3,
        STATES = 1 << 4,
        OBJECT_DEPTH = 1 << 5,
        CHILDREN_COUNT = 1 << 6,
        INDEX_IN_PARENT = 1 << 7,
        XPATH = 1 << 8,

    }
}
