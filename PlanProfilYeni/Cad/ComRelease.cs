using System.Runtime.InteropServices;

namespace PlanProfilYeni.Cad
{
    public static class ComRelease
    {
        public static void Release(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
                Marshal.ReleaseComObject(obj);
        }
    }
}
