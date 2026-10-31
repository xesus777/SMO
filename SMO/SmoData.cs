using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMO
{
    public class SmoData
    {
        public bool IsMultiChannel { get; set; } = false;
        public int ModelTypeIndex { get; set; } = 0;

        public double LambdaRaw { get; set; } = 1.0;
        public double MuRaw { get; set; } = 2.0;

        public int LambdaUnitIndex { get; set; } = 0;
        public int MuUnitIndex { get; set; } = 0;

        public double Lambda => ConvertToMinutes(LambdaRaw, LambdaUnitIndex, isIntensity: true);
        public double Mu => ConvertToMinutes(MuRaw, MuUnitIndex, isIntensity: true);

        public int Channels { get; set; } = 1;
        public int QueueLength { get; set; } = 1;

        public bool AutoCalculateChannels { get; set; } = false;
        public double TargetProbability { get; set; } = 0.97;

        private static double ConvertToMinutes(double value, int unitIndex, bool isIntensity)
        {
            double factor = unitIndex == 1 ? 60.0 : (unitIndex == 2 ? 1440.0 : 1.0);
            return isIntensity ? (value / factor) : (value * factor);
        }
    }
}

