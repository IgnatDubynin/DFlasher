using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;

namespace DFlasher
{
    public enum ExperimentMode
    {
        Staircase,
        Sequential
    }
    public class Configuration
    {
        public int StartingFreq { get; set; } = 60;
        public int StepJumpFreq { get; set; } = 10;
        public int TimeStepFreqChngMs { get; set; } = 500;
        public int StepClarifyThrshldHz { get; set; } = 20;
        public int StdDevIn3Itrtn { get; set; } = 3;

        public int Brightness { get; set; } = 2;

        // Режим эксперимента
        public ExperimentMode Mode { get; set; } = ExperimentMode.Staircase;

        // Sequential режим (минимальные настройки)
        public int SequentialStartingFreq { get; set; } = 60;
        public int SequentialTrials { get; set; } = 20;
        public int SequentialBrightness { get; set; } = 2;
        public int SequentialDirection { get; set; } = -1; // -1 = вниз, +1 = вверх (пока без UI)


        public static Configuration Load(string fileName)
        {
            if (!File.Exists(fileName))
                return new Configuration();

            string json = File.ReadAllText(fileName);
            return JsonSerializer.Deserialize<Configuration>(json)
                   ?? new Configuration();
        }
        public void Save(string fileName)
        {
            string dir = Path.GetDirectoryName(fileName);
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir);

            string json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(fileName, json);
        }
    }
}
