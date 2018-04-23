using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLExportExampleSalesReport {
    class SalesData 
    {
        public SalesData(string state, string product, double q1, double q2, double q3, double q4) 
        {
            State = state;
            Product = product;
            Q1 = q1;
            Q2 = q2;
            Q3 = q3;
            Q4 = q4;
        }

        public string State { get; private set; }
        public string Product { get; private set; }
        public double Q1 { get; private set; }
        public double Q2 { get; private set; }
        public double Q3 { get; private set; }
        public double Q4 { get; private set; }
    }

    static class SalesDataRepository {
        static Random random = new Random();
        static string[] products = new string[] { "HD Video Player", "SuperLED 42", "SuperLED 50", "DesktopLED 19", "DesktopLED 21", "Projector Plus HD"};

        public static List<SalesData> CreateSalesData() {
            List<SalesData> result = new List<SalesData>();
            GenerateData(result, "Arizona");
            GenerateData(result, "California");
            GenerateData(result, "Colorado");
            GenerateData(result, "Florida");
            GenerateData(result, "Idaho");
            return result;
        }

        static void GenerateData(List<SalesData> data, string state) {
            foreach (string product in products) {
                SalesData item = new SalesData(state, product,
                    Math.Round(random.NextDouble() * 5000 + 3000),
                    Math.Round(random.NextDouble() * 4000 + 5000),
                    Math.Round(random.NextDouble() * 6000 + 5500),
                    Math.Round(random.NextDouble() * 5000 + 4000));
                data.Add(item);
            }
        }
    }
}
