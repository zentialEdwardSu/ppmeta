using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Office = Microsoft.Office.Core;
using Newtonsoft.Json;

namespace ppmeta
{
    internal class Config
    {
        public int PositionX { get; set; }
        public int PositionY { get; set; }
        public int FontSize { get; set; }
        public string FontFamily { get; set; }
        public bool AlwaysConfirm { get; set; }
        public bool AlwaysMiddle { get; set; }
        public Office.MsoTextOrientation TextOrientation { get; set; }
        public int TextBoxWidth { get; set; }
        public int TextBoxHeight { get; set; }

        public Config()
        {
            PositionX = 0;
            PositionY = 0;
            FontSize = 12;
            AlwaysMiddle = true; // 默认居中
            AlwaysConfirm = true;
            TextOrientation = Office.MsoTextOrientation.msoTextOrientationHorizontal;
            TextBoxWidth = 500;
            TextBoxHeight = 50;
            FontFamily = "微软雅黑";
        }

        private static string ConfigFilePath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),"ppmeta_config.json");

        public void Save()
        {
            var json = JsonConvert.SerializeObject(this, Newtonsoft.Json.Formatting.Indented);
            File.WriteAllText(ConfigFilePath, json);
        }

        public static Config Load()
        {
            var config = new Config(); 
            if (File.Exists(ConfigFilePath))
            {
                var json = File.ReadAllText(ConfigFilePath);
                JsonConvert.PopulateObject(json, config);
            }
            return config;
        }
    }
}
