using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UnitTestComp.Services.Interface;

namespace UnitTestComp.Services
{
    public class RegistryService : IRegistryService
    {
        private const string RegistryKeyPath = @"SOFTWARE\UnitTestComp";
        private const string RecentDllPathsValueName = "RecentDllPaths";

        public void SaveDllPath(string newPath)
        {
            using (RegistryKey userKey = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
            {
                string recentPaths = userKey.GetValue(RecentDllPathsValueName) as string;
                var paths = string.IsNullOrEmpty(recentPaths)
                    ? new List<string>()
                    : recentPaths.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries).ToList();

                // Добавляем новый путь, если его еще нет в списке
                if (!paths.Contains(newPath))
                {
                    paths.Insert(0, newPath);
                    if (paths.Count > 5)
                    {
                        paths = paths.Take(5).ToList();
                    }
                }

                userKey.SetValue(RecentDllPathsValueName, string.Join("|", paths));
            }
        }

        public List<string> GetRecentDllPaths()
        {
            using (RegistryKey userKey = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
            {
                if (userKey != null)
                {
                    string recentPaths = userKey.GetValue(RecentDllPathsValueName) as string;
                    return recentPaths?.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries).ToList() ?? new List<string>();
                }
            }

            return new List<string>();
        }
    }
}
