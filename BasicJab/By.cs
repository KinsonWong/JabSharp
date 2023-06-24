using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BasicJab
{
    /// <summary>
    /// By枚举
    /// 同样需要暴露Com接口
    /// </summary>
    [Guid("8A0B3453-BA6B-4628-969E-4BACE0FA6967"), ComVisible(true)]
    public enum By
    {
        //这里用了位运算
        //左移
        
        NAME = 1 << 1, //2
        DESCRIPTION = 1 << 2, //4
        ROLE = 1 << 3, //8
        STATES = 1 << 4, //16
        OBJECT_DEPTH = 1 << 5, //32
        CHILDREN_COUNT = 1 << 6, //64
        INDEX_IN_PARENT = 1 << 7, //128
        XPATH = 1 << 8, //256
    }
}