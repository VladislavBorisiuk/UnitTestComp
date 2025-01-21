using System.Collections.Generic;

namespace UnitTestComp.Services.Interface
{
    public interface IRegistryService
    {
        List<string> GetRecentDllPaths();
        void SaveDllPath(string path);
    }
}