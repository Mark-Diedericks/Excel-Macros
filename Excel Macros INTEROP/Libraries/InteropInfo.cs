/*
 * Mark Diedericks
 * 09/06/2018
 * Version 1.0.0
 * Assembly interoprability information data structures
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_INTEROP.Libraries
{
    public enum InteropMemberType
    {
        METHOD = 0,
        PROPERTY = 1
    }

    public struct InteropTypeInfo
    {
        public string nameregion;
        public string name;

        public Type type;
        public InteropMemberInfo[] members;

        public InteropTypeInfo(string n, string s, Type t, InteropMemberInfo[] m)
        {
            name = n;
            nameregion = s;
            type = t;
            members = m;
        }
    }

    public struct InteropParamInfo
    {
        public string name;
        public string actualname;

        public Type type;

        public InteropParamInfo(string n, string a, Type t)
        {
            name = n;
            actualname = a;
            type = t;
        }
    }

    public struct InteropMemberInfo
    {
        public InteropMemberType type;
        public string name;
        public bool accessmod;

        public InteropTypeInfo returnType;
        public InteropParamInfo[] paramTypes;

        public InteropMemberInfo(InteropMemberType t, string n, bool a, InteropTypeInfo r, InteropParamInfo[] p)
        {
            type = t;
            name = n;
            accessmod = a;
            returnType = r;
            paramTypes = p;
        }
    }
}
