// Cad/ComRelease.cs
using System.Runtime.InteropServices;

namespace PlanProfilYeni.Cad
{
    public static class ComRelease
    {
        public static void Release(object com)
        {
            if (com != null && Marshal.IsComObject(com))
                Marshal.ReleaseComObject(com);
        }
    }
}
