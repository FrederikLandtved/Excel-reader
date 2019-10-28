using ExcelDataReader;
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;

namespace ExcelReader1
{
	public class Reader
	{
		public void ReadExcel()
		{
			string path = @"C:\Users\Frederik\ServiceUploadExampleV1.xlsx";
			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

			var stream = File.Open(path, FileMode.Open, FileAccess.Read);
			var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
			int i = 0;
			var treatments = new List<Treatment>();
			while (reader.Read())
			{
				if (i > 5)
				{
					var treatment = new Treatment
					{
						SKSCode = reader.GetString(0),
						Category = reader.GetString(1),
						Speciality = reader.GetString(2),
						ServiceType = reader.GetString(3),
						Service = reader.GetString(4),
						Price = "Ingen pris"
					};
					if (!reader.IsDBNull(5))
					{
						treatment.Price = reader.GetDouble(5).ToString();
					}
					treatments.Add(treatment);
				}
				i++;
			}

			foreach (var item in treatments)
			{
				Console.WriteLine(item.SKSCode + "  " + item.Category + "  " + item.Speciality + "  " + item.ServiceType + "  " + item.Service + "  " + item.Price);
			}
		}
	}
}
