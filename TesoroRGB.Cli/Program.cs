using System;
using System.Threading.Tasks;
using TesoroRgb.Core;

namespace TesoroRGB.Cli
{
    public static class Program
    {
        public static async Task Main()
        {
            using var keyboard = new Keyboard();

            keyboard.Initialize();

            await keyboard.SetProfileAsync(TesoroProfile.Pc);

            await keyboard.SetLightingModeAsync(LightingMode.SpectrumColors, SpectrumMode.SpectrumShine, TesoroProfile.Pc);

            var random = new Random();

            for (var i = 0; i < 254; i++)
            {
                var r = random.Next(0, 255);
                var g = random.Next(0, 255);
                var b = random.Next(0, 255);

                await keyboard.SetKeyColorAsync((TesoroLedId)Convert.ToByte(i), r, g, b, TesoroProfile.Pc);
            }

            await keyboard.SaveSpectrumColorsAsync(TesoroProfile.Pc);
        }
    }
}
