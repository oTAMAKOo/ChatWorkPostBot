
using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Extensions;
using IniFileParser;
using IniFileParser.Model;

namespace ChatWorkPostBot
{
    public abstract class IniFile<TInstance> : Singleton<TInstance> where TInstance : IniFile<TInstance>
    {
        //----- params -----

        //----- field -----

        protected IniData iniData = null;

        //----- property -----

        public abstract string FileName { get; }

        //----- method -----

        public string GetConfigFilePath()
        {
            var fileName = FileName + ".ini";

            var configFileDirectory = ConfigUtility.GetConfigFolderDirectory();

            var configFilePath = PathUtility.Combine(configFileDirectory, fileName);

            return configFilePath;
        }

        public async Task Load()
        {
            var configFilePath = GetConfigFilePath();

            if (!File.Exists(configFilePath))
            {
                CreateDefaultConfigFile(configFilePath);

                throw new Exception($"Require edit generated {configFilePath}.");
            }

            var iniDataString = await File.ReadAllTextAsync(configFilePath);

            var parser = new IniStringParser();

            iniData = parser.Parse(iniDataString);

            OnLoad();
        }

        private void CreateDefaultConfigFile(string configFilePath)
        {
            var parser = new IniFileParser.IniFileParser();

            var data = new IniData();

            SetDefaultData(ref data);

            parser.WriteFile(configFilePath, data, Encoding.UTF8);
        }

        protected T GetData<T>(string section, string key, T defaultValue = default)
        {
            var value = iniData[section][key];

            if (!iniData[section].ContainsKey(key))
            {
                return defaultValue;
            }

            return (T)Convert.ChangeType(value, typeof(T));
        }

        protected abstract void SetDefaultData(ref IniData iniData);

        protected virtual void OnLoad(){  }
    }
}
